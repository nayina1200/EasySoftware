from __future__ import annotations

import importlib.util
import json
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock
from zipfile import ZipFile

import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
SPEC = importlib.util.spec_from_file_location(
    "combined_document_workflow_title_body", ROOT / "scripts" / "combined_document_workflow.py"
)
MODULE = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = MODULE
assert SPEC.loader
SPEC.loader.exec_module(MODULE)


def make_docx(path: Path, paragraphs: list[str]) -> None:
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ET.register_namespace("w", ns)
    document = ET.Element(f"{{{ns}}}document")
    body = ET.SubElement(document, f"{{{ns}}}body")
    for text in paragraphs:
        paragraph = ET.SubElement(body, f"{{{ns}}}p")
        run = ET.SubElement(paragraph, f"{{{ns}}}r")
        ET.SubElement(run, f"{{{ns}}}t").text = text
    path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>")
        archive.writestr("word/document.xml", ET.tostring(document, xml_declaration=True, encoding="UTF-8"))


class FakeRuntime:
    def __init__(self, config_path, *, app_root):
        self.model_id = "fake-model-v1"
        self.status = SimpleNamespace(state="stopped", diagnostic_code=None, reason=None, server_version=None)
        self.started = False
        self.closed = False
        self.config_path = Path(config_path)

    def start(self):
        self.started = True
        self.status = SimpleNamespace(state="ready", diagnostic_code=None, reason=None, server_version="fake")
        return self.status

    def classify(self, title, paragraphs, metadata):
        if not self.started:
            raise AssertionError("runtime must start before the first uncached classification")
        self.seen = (title, paragraphs, metadata)
        return {
            "status": "MISMATCH",
            "reason": "正文明确描述另一类主要用途",
            "core_claims": [{
                "id": "claim-1", "claim": "软件用途", "supported": "no",
                "evidence_ids": ["evidence-1"],
            }],
            "evidence": [{
                "id": "evidence-1", "paragraph_id": paragraphs[0]["paragraph_id"],
                "quote": "库存出入库", "claim_id": "claim-1", "relation": "contradicts",
            }],
            "diagnostics": {"generation_requests": 1},
        }

    def close(self):
        self.closed = True


class CombinedTitleBodyTests(unittest.TestCase):
    def test_default_config_loads_template_from_scripts_directory(self):
        with tempfile.TemporaryDirectory() as directory:
            app_root = Path(directory)
            scripts = app_root / "scripts"
            scripts.mkdir()
            template = scripts / "title_body_config.example.json"
            template.write_bytes((ROOT / "scripts" / template.name).read_bytes())
            path, config = MODULE._title_body_config(app_root)
        self.assertEqual(path, template)
        self.assertEqual(config["mode"], "enforced")
        self.assertTrue(config["skip_mismatch_requested"])
        self.assertEqual(config["prompt_version"], "title-body-jev-domain-fit-score-v3")

    def test_scan_starts_local_runtime_and_uses_unique_txt_title_and_exact_evidence(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "materials"
            app_root = Path(directory) / "app"
            work_root = Path(directory) / "work"
            app_root.mkdir()
            project = root / "Software"
            (project / "Software.txt").parent.mkdir(parents=True)
            (project / "Software.txt").write_text("Software\n", encoding="utf-8")
            make_docx(project / "UnrelatedFileStem说明.docx", ["本系统用于库存出入库管理"])
            config = app_root / "title_body_config.json"
            config.write_text(json.dumps({
                "mode": "shadow", "prompt_version": "test-prompt", "policy_version": "test-policy",
                "provider": "local_llama", "runtime": {"context_size": 4096},
            }), encoding="utf-8")
            fake = FakeRuntime(config, app_root=app_root)
            import local_model_runtime
            with mock.patch.object(local_model_runtime, "LocalModelRuntime", return_value=fake):
                result = MODULE._scan_title_body_qualification(
                    root, work_path=work_root, app_root=app_root, config_path=config,
                )

            self.assertTrue(fake.started)
            self.assertTrue(fake.closed)
            self.assertEqual(fake.seen[0], "Software")
            self.assertEqual(result["projects"][0]["semantic_status"], "MISMATCH")
            self.assertEqual(result["projects"][0]["action"], "CONTINUE_WITH_REVIEW")
            self.assertEqual(result["model_request_count"], 1)
            self.assertEqual(result["failure_count"], 0)
            self.assertEqual(result["failures"], [])
            self.assertEqual(
                set(result["projects"][0]["associated_files"]),
                {"Software/Software.txt", "Software/UnrelatedFileStem说明.docx"},
            )

    def test_exact_associated_files_do_not_exclude_neighboring_material(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            selected = root / "one" / "Same.txt"
            neighbor = root / "two" / "Same.txt"
            selected.parent.mkdir()
            neighbor.parent.mkdir()
            excluded = [{"associated_files": ["one/Same.txt"]}]
            self.assertTrue(MODULE._is_excluded_project(selected, root, excluded))
            self.assertFalse(MODULE._is_excluded_project(neighbor, root, excluded))

    def test_mismatch_names_are_ordered_deduplicated_and_exclude_not_evaluated(self):
        names = MODULE._mismatch_names([
            {"name": "项目甲", "status": "MISMATCH"},
            {"name": "项目乙", "status": "NOT_EVALUATED"},
            {"name": "项目甲", "semantic_status": "MISMATCH"},
            {"name": "项目丙", "semantic_status": "MATCH", "status": "NOT_EVALUATED"},
            {"name": "项目丁", "semantic_status": "MISMATCH"},
        ])
        self.assertEqual(names, ["项目甲", "项目丁"])
        self.assertEqual(
            MODULE._format_mismatch_clipboard(names),
            "其余材料正在处理，以下材料不匹配：\n项目甲\n项目丁",
        )

    def test_empty_mismatch_list_never_touches_clipboard(self):
        writer = mock.Mock()
        result = MODULE._copy_mismatch_summary([], apply=True, clipboard_writer=writer)
        writer.assert_not_called()
        self.assertEqual(result["status"], "no_mismatch")
        self.assertFalse(result["attempted"])
        self.assertEqual(result["text"], "")

    def test_preview_reports_message_without_touching_clipboard(self):
        writer = mock.Mock()
        result = MODULE._copy_mismatch_summary(["项目甲"], apply=False, clipboard_writer=writer)
        writer.assert_not_called()
        self.assertEqual(result["status"], "preview_not_copied")
        self.assertFalse(result["attempted"])
        self.assertEqual(result["text"], "其余材料正在处理，以下材料不匹配：\n项目甲")

    def test_clipboard_error_is_reported_without_raising(self):
        def fail_writer(_text):
            raise OSError("clipboard busy")

        result = MODULE._copy_mismatch_summary(["项目甲"], apply=True, clipboard_writer=fail_writer)
        self.assertEqual(result["status"], "failed")
        self.assertTrue(result["failure"])
        self.assertIn("clipboard busy", result["error"])

    def test_clipboard_success_is_recorded(self):
        writer = mock.Mock()
        result = MODULE._copy_mismatch_summary(["项目甲"], apply=True, clipboard_writer=writer)
        writer.assert_called_once_with("其余材料正在处理，以下材料不匹配：\n项目甲")
        self.assertEqual(result["status"], "copied")
        self.assertTrue(result["attempted"])

    def test_windows_clipboard_uses_owner_window_and_transfers_unicode_memory(self):
        import ctypes

        user32 = mock.Mock()
        kernel32 = mock.Mock()
        owner = 0x1234
        allocation = 0x5678
        memory = ctypes.create_string_buffer(256)
        user32.CreateWindowExW.return_value = owner
        user32.OpenClipboard.return_value = 1
        user32.EmptyClipboard.return_value = 1
        user32.SetClipboardData.return_value = allocation
        user32.CloseClipboard.return_value = 1
        user32.DestroyWindow.return_value = 1
        kernel32.GlobalAlloc.return_value = allocation
        kernel32.GlobalLock.return_value = ctypes.addressof(memory)
        kernel32.GlobalUnlock.return_value = 0
        kernel32.GlobalFree.return_value = None

        with mock.patch.object(MODULE.os, "name", "nt"), mock.patch.object(
            MODULE.ctypes, "WinDLL", side_effect=[user32, kernel32], create=True,
        ):
            MODULE._write_windows_clipboard("其余材料正在处理，以下材料不匹配：\n项目甲")

        user32.CreateWindowExW.assert_called_once()
        self.assertEqual(user32.CreateWindowExW.call_args.args[1], "STATIC")
        user32.OpenClipboard.assert_called_once_with(owner)
        user32.SetClipboardData.assert_called_once_with(13, allocation)
        user32.DestroyWindow.assert_called_once_with(owner)
        kernel32.GlobalFree.assert_not_called()
        expected = "其余材料正在处理，以下材料不匹配：\n项目甲".encode("utf-16-le") + b"\x00\x00"
        self.assertEqual(memory.raw[:len(expected)], expected)

    def test_windows_clipboard_failure_releases_memory_and_owner_window(self):
        import ctypes

        user32 = mock.Mock()
        kernel32 = mock.Mock()
        owner = 0x1234
        allocation = 0x5678
        memory = ctypes.create_string_buffer(256)
        user32.CreateWindowExW.return_value = owner
        user32.OpenClipboard.return_value = 1
        user32.EmptyClipboard.return_value = 1
        user32.SetClipboardData.return_value = None
        user32.CloseClipboard.return_value = 1
        user32.DestroyWindow.return_value = 1
        kernel32.GlobalAlloc.return_value = allocation
        kernel32.GlobalLock.return_value = ctypes.addressof(memory)
        kernel32.GlobalUnlock.return_value = 0
        kernel32.GlobalFree.return_value = None

        with mock.patch.object(MODULE.os, "name", "nt"), mock.patch.object(
            MODULE.ctypes, "WinDLL", side_effect=[user32, kernel32], create=True,
        ):
            with self.assertRaises(OSError):
                MODULE._write_windows_clipboard("项目甲")

        user32.OpenClipboard.assert_called_once_with(owner)
        user32.CloseClipboard.assert_called_once()
        user32.DestroyWindow.assert_called_once_with(owner)
        kernel32.GlobalFree.assert_called_once_with(allocation)

    def test_windows_clipboard_does_not_transfer_locked_memory(self):
        import ctypes

        user32 = mock.Mock()
        kernel32 = mock.Mock()
        allocation = 0x5678
        memory = ctypes.create_string_buffer(256)
        kernel32.GlobalAlloc.return_value = allocation
        kernel32.GlobalLock.return_value = ctypes.addressof(memory)
        kernel32.GlobalUnlock.return_value = 1  # A nonzero result means it remains locked.
        kernel32.GlobalFree.return_value = allocation  # Locked blocks cannot be freed.

        with mock.patch.object(MODULE.os, "name", "nt"), mock.patch.object(
            MODULE.ctypes, "WinDLL", side_effect=[user32, kernel32], create=True,
        ):
            with self.assertRaisesRegex(OSError, "仍处于锁定状态"):
                MODULE._write_windows_clipboard("项目甲")

        user32.SetClipboardData.assert_not_called()
        kernel32.GlobalFree.assert_called_once_with(allocation)

    def test_apply_copies_mismatch_list_before_code_stage_and_skips_exact_files(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            source = base / "materials"
            work = base / "work"
            app_root = base / "app"
            source.mkdir()
            app_root.mkdir()
            (source / "A.txt").write_text("A\n", encoding="utf-8")
            for name in ("A说明.docx", "A代码.docx", "Neighbor.docx", "B说明.docx"):
                (source / name).write_bytes(b"fixture")
            (app_root / "title_body_config.json").write_text(json.dumps({
                "enabled": True, "mode": "enforced", "enforce": True,
                "provider": "local_llama", "skip_mismatch_requested": True,
                "not_evaluated_action": "hold_changes",
            }), encoding="utf-8")

            mismatch_item = {
                "name": "项目甲", "status": "MISMATCH", "semantic_status": "MISMATCH",
                "relative_path": "A说明.docx", "associated_files": ["A.txt", "A说明.docx", "A代码.docx"],
            }
            not_evaluated_item = {
                "name": "项目乙", "status": "NOT_EVALUATED", "semantic_status": "NOT_EVALUATED",
                "relative_path": "B说明.docx", "associated_files": ["B说明.docx"],
            }
            qualification = {
                "status": "COMPLETED", "mode": "enforced", "provider": "local_llama",
                "scanned_count": 2, "failure_count": 0, "failures": [], "cache_hit_count": 0,
                "model_request_count": 2, "projects": [mismatch_item, not_evaluated_item],
                "disqualified_projects": [mismatch_item], "_inputs_by_relative_path": {},
            }
            event_log = []
            progress_events = []

            class Progress:
                def update(self, phase, current=0, name="", **details):
                    progress_events.append({"phase": phase, "current": current, "name": name, **details})

            def clipboard_writer(text):
                event_log.append(("clipboard", text))
                raise OSError("clipboard busy")

            def code_stage(root, _app_root, _work_path, **kwargs):
                event_log.append(("code", ""))
                excluded = kwargs["excluded"]
                self.assertTrue(MODULE._is_excluded_project(root / "A.txt", root, excluded))
                self.assertTrue(MODULE._is_excluded_project(root / "A说明.docx", root, excluded))
                self.assertTrue(MODULE._is_excluded_project(root / "A代码.docx", root, excluded))
                self.assertFalse(MODULE._is_excluded_project(root / "Neighbor.docx", root, excluded))
                return {"status": "READY_FOR_REVIEW", "failures": 0, "documents": []}, 0

            import validation_isolation
            with (
                mock.patch.object(validation_isolation, "require_validation_path", side_effect=lambda path, **_kwargs: Path(path)),
                mock.patch.object(MODULE.process_batch, "recover_pending_cover_transactions", return_value=[]),
                mock.patch.object(MODULE, "_scan_title_body_qualification", return_value=qualification),
                mock.patch.object(MODULE, "_code_inventory", return_value=[]),
                mock.patch.object(MODULE, "_run_code_stage", side_effect=code_stage),
            ):
                result = MODULE.run(
                    source, work, app_root, apply=True, progress=Progress(),
                    document_stages=("match_title_body", "match_code"),
                    clipboard_writer=clipboard_writer,
                )

            self.assertEqual([event[0] for event in event_log], ["clipboard", "code"])
            self.assertEqual(
                event_log[0][1], "其余材料正在处理，以下材料不匹配：\n项目甲",
            )
            progress_item = next(item for item in progress_events if item["phase"] == "标题正文资格预检完成")
            self.assertIn(event_log[0][1], progress_item["name"])
            report = json.loads(Path(result["report_json"]).read_text(encoding="utf-8"))
            self.assertEqual(report["qualification"]["mismatch_names"], ["项目甲"])
            self.assertEqual(report["qualification"]["mismatch_clipboard"]["status"], "failed")
            self.assertTrue(report["qualification"]["mismatch_clipboard"]["failure"])
            self.assertIn("clipboard busy", report["qualification"]["mismatch_clipboard"]["error"])
            self.assertEqual(report["qualification"]["mismatch_clipboard"]["text"], event_log[0][1])
            html_report = Path(result["report"]).read_text(encoding="utf-8")
            self.assertIn("项目甲", html_report)
            self.assertIn("剪贴板复制失败，处理继续", html_report)

    def test_apply_copies_only_match_projects_and_names_qualified_and_total_counts(self):
        import re
        import validation_isolation

        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            source = base / "客户甲10件"
            app_root = base / "app"
            work = base / "work"
            app_root.mkdir()
            states = (("甲", "MATCH"), ("乙", "MISMATCH"), ("丙", "NOT_EVALUATED"))
            projects = []
            originals = {}
            for name, status in states:
                project_root = source / name
                project_root.mkdir(parents=True)
                project_files = {
                    f"{name}.txt": f"{name} title\n".encode(),
                    f"{name}说明.docx": f"{name} manual".encode(),
                    f"{name}说明.pdf": f"{name} manual pdf".encode(),
                    f"{name}代码.docx": f"{name} code".encode(),
                    f"{name}代码.pdf": f"{name} code pdf".encode(),
                }
                for filename, content in project_files.items():
                    path = project_root / filename
                    path.write_bytes(content)
                    originals[path] = content
                relative_files = [f"{name}/{filename}" for filename in project_files]
                projects.append({
                    "name": name,
                    "status": status,
                    "semantic_status": status,
                    "relative_path": f"{name}/{name}说明.docx",
                    "associated_files": relative_files,
                    "reason": "synthetic result",
                })
            config = app_root / "title_body_config.json"
            config.write_text(json.dumps({
                "mode": "enforced", "enforce": True, "skip_mismatch_requested": True,
                "not_evaluated_action": "hold_changes", "provider": "local_llama",
            }), encoding="utf-8")
            qualification = {
                "status": "COMPLETED", "mode": "enforced", "provider": "local_llama",
                "model_id": "fixture", "scanned_count": 3, "failure_count": 0,
                "failures": [], "cache_hit_count": 0, "model_request_count": 3,
                "projects": projects, "disqualified_projects": [projects[1]],
                "_inputs_by_relative_path": {},
            }
            with (
                mock.patch.object(validation_isolation, "require_validation_path", side_effect=lambda path, **_kwargs: Path(path)),
                mock.patch.object(MODULE.process_batch, "recover_pending_cover_transactions", return_value=[]),
                mock.patch.object(MODULE, "_scan_title_body_qualification", return_value=qualification),
            ):
                result = MODULE.run(
                    source, work, app_root, apply=True,
                    document_stages=("match_title_body",),
                    clipboard_writer=lambda _text: None,
                )

            output = Path(result["output"])
            self.assertRegex(
                output.name,
                re.compile(r"^客户甲1件（共3件）_已处理_\d{8}_\d{6}$"),
            )
            self.assertTrue((output / "甲" / "甲说明.docx").is_file())
            self.assertTrue((output / "甲" / "甲代码.docx").is_file())
            self.assertFalse((output / "乙").exists())
            self.assertFalse((output / "丙").exists())
            self.assertEqual({path: path.read_bytes() for path in originals}, originals)
            report = json.loads(Path(result["report_json"]).read_text(encoding="utf-8"))
            self.assertEqual(report["output_counts"], {
                "qualified_projects": 1, "total_projects": 3, "selection": "MATCH_ONLY",
            })
            self.assertEqual(report["qualification"]["qualified_project_count"], 1)
            self.assertEqual(report["qualification"]["input_project_count"], 3)
            self.assertEqual(report["qualification"]["projects"][1]["action"], "SKIP_TITLE_BODY_MISMATCH")
            self.assertEqual(report["qualification"]["projects"][2]["action"], "HOLD_NOT_EVALUATED")
            html_report = Path(result["report"]).read_text(encoding="utf-8")
            self.assertIn("客户甲1件（共3件）", html_report)
            self.assertIn("MATCH 纳入 1/3 件", html_report)

    def test_output_qualified_count_excludes_match_records_without_copyable_files(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory)
            (source / "甲说明.docx").write_bytes(b"manual")
            projects = [
                {
                    "semantic_status": "MATCH", "relative_path": "甲说明.docx",
                    "associated_files": ["甲说明.docx"],
                },
                {
                    "semantic_status": "MATCH", "relative_path": "missing/乙说明.docx",
                    "associated_files": ["missing/乙说明.docx"],
                },
                {
                    "semantic_status": "NOT_EVALUATED", "relative_path": "丙说明.docx",
                    "associated_files": ["丙说明.docx"],
                },
            ]

            plan = MODULE._qualified_material_plan(source, projects)

            self.assertEqual(plan[3:], (1, 3))


if __name__ == "__main__":
    unittest.main()

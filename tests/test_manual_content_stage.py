import importlib.util
import sys
import json
import base64
import io
import tempfile
import unittest
import zipfile
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
from unittest import mock
from pathlib import Path
from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
# 被测模块按文件路径加载，需要把 scripts 加入 sys.path，
# 否则被测模块内部的同级 import（如 archive_types）无法解析。
sys.path.insert(0, str(ROOT / "scripts"))
import process_batch as BATCH_MODULE
SPEC = importlib.util.spec_from_file_location("manual_content_stage", ROOT / "scripts" / "manual_content_stage.py")
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader
SPEC.loader.exec_module(MODULE)


def make_docx(path: Path, texts=None, image_count=0):
    path.parent.mkdir(parents=True, exist_ok=True)
    paragraphs = "".join(
        f'<w:p><w:r><w:t>{text}</w:t></w:r></w:p>' for text in (texts or ["示例系统说明"])
    )
    xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{paragraphs}</w:body></w:document>'
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", xml)
        for index in range(image_count):
            archive.writestr(f"word/media/image{index + 1}.png", f"image-{index}".encode())


class SamplingTests(unittest.TestCase):
    def test_quarter_sampling_is_equidistant_and_at_least_four(self):
        self.assertEqual(MODULE.equidistant_sample_indexes(3), [0, 1, 2])
        self.assertEqual(MODULE.equidistant_sample_indexes(8), [0, 2, 5, 7])
        self.assertEqual(len(MODULE.equidistant_sample_indexes(20)), 5)


class RuntimeIsolationTests(unittest.TestCase):
    def test_date_tools_follow_the_isolated_runtime_override(self):
        with tempfile.TemporaryDirectory() as directory:
            tool_root = Path(directory) / "date-tool"
            internal = tool_root / "_internal"
            internal.mkdir(parents=True)
            (internal / "python.exe").write_bytes(b"test runtime placeholder")
            (tool_root / "app.py").write_text("", encoding="utf-8")

            with mock.patch.dict("os.environ", {"EASYSOFTWARE_DATE_TOOL_ROOT": str(tool_root)}):
                ocr = MODULE.DefaultOcrAdapter(ROOT / "scripts")
                rewriter = MODULE.DefaultDateRewriter(ROOT / "scripts")

            self.assertEqual(ocr.tool_root, tool_root.resolve())
            self.assertEqual(rewriter.tool_root, tool_root.resolve())
            self.assertEqual(ocr.backend, "rapidocr_subprocess")
            self.assertTrue(rewriter.available)

    def test_ocr_batch_performance_metrics_are_opt_in(self):
        adapter = MODULE.DefaultOcrAdapter(ROOT / "scripts")
        adapter._batch_images = mock.Mock(return_value=[{"text": "ok"}])
        with mock.patch.dict("os.environ", {"EASYSOFTWARE_PERF_TRACE": "1"}):
            rows = adapter.batch([(b"image", ".png")])
        self.assertEqual(rows, [{"text": "ok"}])
        self.assertEqual(adapter._performance["batch_calls"], 1)
        self.assertEqual(adapter._performance["image_count"], 1)
        self.assertGreaterEqual(adapter._performance["batch_seconds"], 0)


class TxtPunctuationRepairTests(unittest.TestCase):
    def test_default_mode_removes_parenthesized_ordinal_and_counts_it(self):
        repaired, counts = BATCH_MODULE.repair_main_function_text("（一）设备管理")

        self.assertEqual(repaired, "设备管理")
        self.assertEqual(counts["ordinal_markers_removed"], 1)

    def test_parenthesized_ordinal_and_other_leading_punctuation_are_cleaned(self):
        repaired, counts = BATCH_MODULE.repair_main_function_text(
            "（一）设备管理（支持录入）,,并查询！？",
            remove_ordinal_markers=False,
        )

        self.assertEqual(repaired, "一，设备管理，支持录入，并查询！")
        self.assertTrue(repaired.startswith("一"))
        self.assertNotRegex(repaired, r"^[，、；。！？!?]")
        self.assertEqual(counts["leading_punctuation_removed"], 1)
        self.assertEqual(counts["duplicate_punctuation_removed"], 3)
        self.assertEqual(counts["parentheses_to_commas"], 4)

    def test_leading_ascii_and_chinese_punctuation_are_counted(self):
        repaired, counts = BATCH_MODULE.repair_main_function_text(
            "  ,；功能管理",
            remove_ordinal_markers=False,
        )

        self.assertEqual(repaired, "  功能管理")
        self.assertEqual(counts["leading_punctuation_removed"], 1)
        self.assertEqual(counts["duplicate_punctuation_removed"], 1)


class YearTests(unittest.TestCase):
    def test_three_way_baseline_and_replacement(self):
        self.assertEqual(MODULE.classify_year_context("截至2023年为最新版本", "2023")["decision"], "UPDATE_TO_2025")
        self.assertEqual(MODULE.classify_year_context("GB/T 2023 标准", "2023")["decision"], "KEEP")
        self.assertEqual(MODULE.classify_year_context("记录时间为2023年", "2023")["decision"], "REVIEW")
        replaced, decisions = MODULE.replace_update_years("截至2023年，历史数据为2020年度数据。")
        self.assertEqual(replaced, "截至2025年，历史数据为2020年度数据。")
        self.assertEqual([item["decision"] for item in decisions], ["UPDATE_TO_2025", "KEEP"])

    def test_low_confidence_update_is_demoted_to_review(self):
        replaced, decisions = MODULE.replace_update_years(
            "截至2023年",
            classifier=lambda *_: {"decision": "UPDATE_TO_2025", "confidence": 0.84},
        )
        self.assertEqual(replaced, "截至2023年")
        self.assertEqual(decisions[0]["decision"], "REVIEW")

    def test_future_year_is_kept_even_if_classifier_requests_update(self):
        replaced, decisions = MODULE.replace_update_years(
            "截至2033年，2025年为目标年。",
            classifier=lambda *_: {"decision": "UPDATE_TO_2025", "confidence": 1.0},
        )
        self.assertEqual(replaced, "截至2033年，2025年为目标年。")
        self.assertEqual([item["decision"] for item in decisions], ["KEEP", "KEEP"])


class BodyRelevanceTests(unittest.TestCase):
    def test_repeated_title_does_not_hide_conflicting_body(self):
        title = "太阳能LED路灯储能管理与充放电控制软件"
        paragraphs = [("p:1", title * 8), ("p:2", "散热、热泵、制热和供暖由水泵控制。" * 8)]
        result = MODULE.assess_body_relevance(title, paragraphs)
        self.assertEqual(result["status"], "review_topic_mismatch")
        self.assertEqual(result["title_terms"]["路灯"], 0)
        self.assertGreater(result["other_topics"]["散热供暖"]["热泵"], 0)

    def test_short_or_unmapped_text_stays_unknown(self):
        self.assertEqual(MODULE.assess_body_relevance("甲系统", [])["status"], "unknown")
        self.assertEqual(MODULE.assess_body_relevance("智能路灯", [("p:1", "开灯")])["status"], "unknown")

    def test_later_domain_evidence_prevents_generic_operation_review(self):
        title = "毫米波组件参数校验管理软件"
        paragraphs = [("p:1", title)] + [(f"p:{i}", "新增、删除、导出和查询毫米波组件的参数校验记录。") for i in range(2, 18)]
        self.assertEqual(MODULE.assess_body_relevance(title, paragraphs)["status"], "unknown")

    def test_competing_topic_requires_sustained_evidence(self):
        title = "射频芯片管理系统"
        paragraphs = [("p:1", title)] + [(f"p:{i}", "射频芯片参数监测与校准。") for i in range(2, 18)]
        paragraphs.append(("p:18", "协议栈也可用于终端软件。"))
        self.assertEqual(MODULE.assess_body_relevance(title, paragraphs)["status"], "unknown")

    def test_real_output_topic_review_without_false_flags(self):
        root = ROOT.parent / "order" / "北京知识产权10件_已处理_20260922_213656"
        if not root.is_dir():
            self.skipTest("本机真机输出样本不可用")
        results = {}
        for path in root.rglob("*说明.docx"):
            _images, paragraphs = MODULE._docx_parts(path)
            title = path.stem.removesuffix("说明")
            results[title] = MODULE.assess_body_relevance(title, paragraphs)["status"]
        self.assertEqual(len(results), 10)
        for title in (
            "太阳能LED路灯储能管理与充放电控制软件",
            "弘元绿能金刚线切片机控制系统软件",
            "智能LED路灯自适应调光控制系统",
            "无源射频芯片数据调控管理系统",
            "毫米波组件参数校验管理软件",
        ):
            self.assertEqual(results[title], "review_topic_mismatch", title)
        for title in (
            "医院人力资源综合管理系统",
            "建筑项目数据综合分析平台",
            "弘元绿能金刚线蓝宝石切片机控制系统软件",
            "智慧照明管理平台",
            "电子信息数据采集处理系统",
        ):
            self.assertNotEqual(results[title], "review_topic_mismatch", title)


class WorkflowTests(unittest.TestCase):
    def run_stage(self, root, *, ocr, relevance, date_rewriter):
        return MODULE.run_manual_content_stage(
            root, root / "out", root / "report.json", ocr=ocr,
            relevance_classifier=relevance, date_rewriter=date_rewriter,
        )

    @staticmethod
    def qualification_item(title, status, *, paragraph_id="paragraph:1", text="正文中的业务功能结果"):
        return {
            "project_id": f"project:{title}",
            "title": title,
            "txt_path": f"{title}.txt",
            "semantic_input_hash": f"hash:{title}",
            "paragraphs": [{
                "paragraph_id": paragraph_id, "text": text,
                "section_id": "section:front", "style_id": "BodyText",
                "outline_level": None, "is_toc": False, "in_table": False,
                "in_textbox": False, "is_title_repeat": False, "structure_uncertain": False,
            }],
            "metadata": {"title_source": "txt_filename_stem"},
            "result": {"status": status, "reason": f"语义结果 {status}"},
        }

    def test_semantic_input_reuses_txt_title_paragraphs_and_never_calls_legacy_rule(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            make_docx(root / "不同文件名说明.docx", image_count=1)
            relevance_calls = []
            qualification = self.qualification_item("TXT 文件名标题", "MATCH", text="TXT 主标题对应的真实业务正文")
            with (
                mock.patch.object(MODULE, "assess_body_relevance", side_effect=AssertionError("legacy topic rule called")),
                mock.patch.object(MODULE, "read_docx_paragraphs", side_effect=AssertionError("body was re-read")),
            ):
                report = MODULE.run_manual_content_stage(
                    root, root / "out", root / "report.json", apply=False,
                    title_body_inputs_by_relative_path={"不同文件名说明.docx": qualification},
                    title_body_mode="shadow", skip_mismatch=True,
                    ocr=lambda *_: {"text": "文字"},
                    relevance_classifier=lambda text, title, context: relevance_calls.append((title, context)) or {"status": "relevant", "confidence": 0.9},
                    check_image_vision=False,
                )
            project = report["projects"][0]
            self.assertEqual(project["title"], "TXT 文件名标题")
            self.assertEqual(project["txt_path"], "TXT 文件名标题.txt")
            self.assertEqual(project["project_id"], qualification["project_id"])
            self.assertEqual(project["semantic_status"], "MATCH")
            self.assertFalse(project["qualification_pending_review"])
            self.assertEqual(relevance_calls[0][0], "TXT 文件名标题")
            self.assertEqual(relevance_calls[0][1], "TXT 主标题对应的真实业务正文")

    def test_shadow_mismatch_and_not_evaluated_continue_with_review_despite_skip_request(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for name in ("甲说明.docx", "乙说明.docx"):
                make_docx(root / name, image_count=1)
            inputs = {
                "甲说明.docx": self.qualification_item("甲 TXT 标题", "MISMATCH"),
                "乙说明.docx": self.qualification_item("乙 TXT 标题", "NOT_EVALUATED"),
            }
            ocr_calls = []
            with mock.patch.object(MODULE, "assess_body_relevance", side_effect=AssertionError("legacy topic rule called")):
                report = MODULE.run_manual_content_stage(
                    root, root / "out", root / "report.json", apply=False,
                    title_body_inputs_by_relative_path=inputs, title_body_mode="shadow",
                    skip_mismatch=True, not_evaluated_action="hold_changes",
                    ocr=lambda *_: ocr_calls.append(1) or {"text": "文字"},
                    relevance_classifier=lambda *_: {"status": "relevant", "confidence": 0.9},
                    check_image_vision=False,
                )
            self.assertEqual(len(ocr_calls), 2)
            self.assertEqual([item["qualification_action"] for item in report["projects"]], [
                "CONTINUE_WITH_REVIEW", "CONTINUE_WITH_REVIEW",
            ])
            self.assertTrue(all(item["qualification_pending_review"] for item in report["projects"]))
            self.assertTrue(all(item["status"] == "PREFLIGHT_PASSED_WITH_REVIEW" for item in report["projects"]))
            self.assertFalse(report["skip_effective"])
            self.assertEqual(report["skip_requested"], True)
            self.assertEqual(report["qualification_review_count"], 2)

    def test_enforced_mismatch_skip_and_not_evaluated_hold_skip_all_changes(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            output = root / "out"
            names = {"甲": "MISMATCH", "乙": "NOT_EVALUATED", "丙": "MATCH"}
            originals = {}
            inputs = {}
            for name, status in names.items():
                path = root / f"{name}说明.docx"
                make_docx(path, ["截至2023年为最新版本"], image_count=2)
                originals[name] = path.read_bytes()
                inputs[path.name] = self.qualification_item(f"{name} TXT 标题", status)
            date_calls = []
            report = MODULE.run_manual_content_stage(
                root, output, root / "report.json", apply=True,
                title_body_inputs_by_relative_path=inputs, title_body_mode="enforced",
                skip_mismatch=True, not_evaluated_action="hold_changes",
                check_image_relevance=False, check_image_vision=False,
                date_rewriter=lambda path, year: date_calls.append(Path(path).name) or {"status": "UNCHANGED", "changed": False},
            )
            projects = {Path(item["source"]).stem.removesuffix("说明"): item for item in report["projects"]}
            self.assertEqual(projects["甲"]["status"], "SKIPPED_TOPIC_MISMATCH")
            self.assertEqual(projects["甲"]["qualification_action"], "SKIP_MISMATCH")
            self.assertEqual(projects["乙"]["status"], "HELD_QUALIFICATION_NOT_EVALUATED")
            self.assertEqual(projects["乙"]["qualification_action"], "HOLD_NOT_EVALUATED")
            self.assertEqual(report["held_not_evaluated_count"], 1)
            self.assertEqual(len(date_calls), 1)
            self.assertTrue(date_calls[0].startswith(".丙说明.date-year-stage-"))
            self.assertEqual((root / "甲说明.docx").read_bytes(), originals["甲"])
            self.assertEqual((root / "乙说明.docx").read_bytes(), originals["乙"])
            self.assertEqual(report["skipped_project_count"], 1)
            self.assertEqual(report["qualification_review_count"], 1)

    def test_missing_per_project_qualification_is_not_evaluated_and_does_not_fall_back(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            make_docx(root / "甲说明.docx", image_count=1)
            with mock.patch.object(MODULE, "assess_body_relevance", side_effect=AssertionError("legacy topic rule called")), mock.patch.object(
                MODULE, "_docx_images", side_effect=AssertionError("held project media must stay unloaded"),
            ):
                report = MODULE.run_manual_content_stage(
                    root, root / "out", root / "report.json", apply=False,
                    title_body_inputs_by_relative_path={}, title_body_mode="enforced",
                    not_evaluated_action="hold_changes", check_image_relevance=True,
                )
            project = report["projects"][0]
            self.assertEqual(project["semantic_status"], "NOT_EVALUATED")
            self.assertEqual(project["diagnostic_code"], "TITLE_BODY_RESULT_MISSING")
            self.assertEqual(project["status"], "HELD_QUALIFICATION_NOT_EVALUATED")
            self.assertEqual(project["image_count"], 1)

    def test_enforced_can_include_mismatch_and_continue_not_evaluated_with_review(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            inputs = {}
            for name, status in (("甲", "MISMATCH"), ("乙", "NOT_EVALUATED")):
                path = root / f"{name}说明.docx"
                make_docx(path, ["没有可改写年份"])
                inputs[path.name] = self.qualification_item(f"{name} TXT 标题", status)
            calls = []
            report = MODULE.run_manual_content_stage(
                root, root / "out", root / "report.json", apply=True,
                title_body_inputs_by_relative_path=inputs, title_body_mode="enforced",
                skip_mismatch=False, not_evaluated_action="continue_with_review",
                check_image_relevance=False, check_image_vision=False,
                date_rewriter=lambda *_: calls.append(1) or {"status": "UNCHANGED", "changed": False},
            )
            self.assertEqual(calls, [1, 1])
            self.assertTrue(all(item["qualification_pending_review"] for item in report["projects"]))
            self.assertTrue(all(item["qualification_action"] == "CONTINUE_WITH_REVIEW" for item in report["projects"]))
            self.assertFalse(report["skip_effective"])

    def test_unknown_ocr_does_not_block_or_fuse(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            make_docx(root / "甲说明.docx", image_count=8)
            calls = []
            report = self.run_stage(
                root,
                ocr=lambda *_: {"text": ""},
                relevance=lambda *_: {"status": "irrelevant", "confidence": 1},
                date_rewriter=lambda path, year: calls.append((path, year)) or {"status": "UNCHANGED", "changed": False},
            )
            self.assertFalse(report["batch_fused"])
            self.assertFalse(report["projects"][0]["blocked"])
            self.assertEqual(report["projects"][0]["body_relevance"]["status"], "unknown")
            self.assertEqual(len(calls), 1)

    def test_project_is_blocked_only_when_more_than_half_decisive_samples_are_irrelevant(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            make_docx(root / "甲说明.docx", image_count=4)
            verdicts = iter(["irrelevant", "irrelevant", "irrelevant", "relevant"])
            calls = []
            report = self.run_stage(
                root, ocr=lambda *_: {"text": "文字"},
                relevance=lambda *_: {"status": next(verdicts), "confidence": 0.9},
                date_rewriter=lambda *_: calls.append(1) or {"changed": False},
            )
            self.assertTrue(report["projects"][0]["blocked"])
            self.assertEqual(calls, [])

    def test_batch_fuse_above_twenty_percent_prevents_all_date_writes(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for name in "甲乙丙丁":
                make_docx(root / f"{name}说明.docx", image_count=4)
            calls = []
            report = self.run_stage(
                root, ocr=lambda *_: {"text": "文字"},
                relevance=lambda text, title, context: {"status": "irrelevant" if title == "甲" else "relevant", "confidence": 0.9},
                date_rewriter=lambda *_: calls.append(1) or {"changed": False},
            )
            self.assertTrue(report["batch_fused"])
            self.assertEqual(calls, [])
            self.assertEqual(json.loads((root / "report.json").read_text(encoding="utf-8"))["blocked_project_count"], 1)

    def test_every_preflight_finishes_before_first_write(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for name in "甲乙":
                make_docx(root / f"{name}说明.docx", image_count=4)
            events = []

            def ocr(data, suffix):
                events.append("ocr")
                return {"text": "相关"}

            def rewrite(path, year):
                events.append("write")
                self.assertEqual(events.count("ocr"), 8)
                return {"status": "UNCHANGED", "changed": False}

            self.run_stage(root, ocr=ocr, relevance=lambda *_: {"status": "relevant", "confidence": 0.9}, date_rewriter=rewrite)
            self.assertEqual(events[:8], ["ocr"] * 8)

    def test_batch_failure_keeps_original_and_later_document_completes(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for name in ("甲", "乙"):
                make_docx(root / f"{name}说明.docx", ["截至2023年为最新版本"], image_count=4)
            adapter = MODULE.DefaultDateRewriter(ROOT / "scripts")
            adapter.available = True
            original_popen = MODULE.subprocess.Popen
            process_instances = []

            def fake_child(args, **kwargs):
                child_code = (
                    "import json,pathlib,sys,time; "
                    "print(json.dumps({'ready':True,'performance':{'engine_load_seconds':0.12}}),flush=True); "
                    "request=json.loads(sys.stdin.readline()); "
                    "path=pathlib.Path(request['path']); "
                    "path.with_suffix(path.suffix+'.date-year.tmp').write_text('interrupted temp'); "
                    "time.sleep(30)"
                )
                process = original_popen(
                    [sys.executable, "-c", child_code], stdin=kwargs["stdin"], stdout=kwargs["stdout"],
                    stderr=kwargs["stderr"], text=kwargs["text"], encoding=kwargs["encoding"],
                    errors=kwargs["errors"], bufsize=kwargs["bufsize"],
                )
                process_instances.append(process)
                return process

            fallback_results = [
                {"status": "FAILED", "changed": False, "reason": "模拟单件重试仍失败"},
                {"status": "CHANGED", "changed": True, "effective_target_year": 2025},
            ]
            fallback_calls = []

            def run_fallback(_args, **_kwargs):
                fallback_calls.append(1)
                return MODULE.subprocess.CompletedProcess(
                    [], 0, stdout=json.dumps(fallback_results[len(fallback_calls) - 1]), stderr=""
                )

            original_first = (root / "甲说明.docx").read_bytes()
            with (
                mock.patch.dict("os.environ", {"EASYSOFTWARE_PERF_TRACE": "1"}),
                mock.patch.object(MODULE, "DATE_REWRITER_TIMEOUT_SECONDS", 0.2),
                mock.patch.object(MODULE.subprocess, "Popen", side_effect=fake_child),
                mock.patch.object(MODULE.subprocess, "run", side_effect=run_fallback),
            ):
                report = self.run_stage(
                    root, ocr=lambda *_: {"text": "相关"},
                    relevance=lambda *_: {"status": "relevant", "confidence": 0.9},
                    date_rewriter=adapter,
                )

            self.assertEqual(len(process_instances), 1)
            self.assertIsNotNone(process_instances[0].poll())
            self.assertEqual(len(fallback_calls), 2)
            self.assertEqual(report["projects"][0]["date_rewrite"]["status"], "FAILED")
            self.assertEqual(report["projects"][1]["date_rewrite"]["status"], "CHANGED")
            failed_output = Path(report["projects"][0]["output"])
            successful_output = Path(report["projects"][1]["output"])
            self.assertEqual(failed_output.read_bytes(), original_first)
            with zipfile.ZipFile(successful_output) as archive:
                success_xml = archive.read("word/document.xml").decode("utf-8")
            self.assertIn("2025", success_xml)
            self.assertFalse(list((root / "out").rglob("*.date-year-stage-*.docx")))
            self.assertFalse(list((root / "out").rglob("*.date-year-stage-*.docx.date-year.tmp")))
            date_perf = report["performance"]["date_rewriter"]
            self.assertEqual(date_perf["documents_attempted"], 2)
            self.assertEqual(date_perf["subprocess_calls"], 3)
            self.assertEqual(date_perf["batch_count"], 1)
            self.assertEqual(date_perf["engine_load_seconds"], 0.12)
            self.assertGreaterEqual(date_perf["batch_wall_seconds"], 0)

    def test_year_update_is_written_to_output_docx(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            make_docx(root / "甲说明.docx", ["截至2023年为最新版本"])
            report = self.run_stage(root, ocr=lambda *_: "", relevance=lambda *_: {}, date_rewriter=lambda *_: {"changed": False})
            output = Path(report["projects"][0]["output"])
            with zipfile.ZipFile(output) as archive:
                xml = archive.read("word/document.xml").decode("utf-8")
            self.assertIn("2025", xml)
            self.assertNotIn("2023", xml)

    def test_preview_never_creates_or_modifies_output(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            make_docx(root / "甲说明.docx", ["截至2023年为最新版本"])
            output = root / "out"
            report = MODULE.run_manual_content_stage(
                root, output, root / "preview.json", apply=False,
                ocr=lambda *_: "", date_rewriter=lambda *_: self.fail("preview must not rewrite"),
            )
            self.assertFalse(report["applied"])
            self.assertFalse(output.exists())
            with zipfile.ZipFile(root / "甲说明.docx") as archive:
                self.assertIn("2023", archive.read("word/document.xml").decode("utf-8"))

    def test_document_without_updatable_year_is_copied_byte_for_byte(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "甲说明.docx"
            make_docx(source, ["这里没有需要更新的年份"])
            before = source.read_bytes()
            report = self.run_stage(
                root, ocr=lambda *_: "", relevance=lambda *_: {},
                date_rewriter=lambda *_: {"status": "UNCHANGED", "changed": False},
            )
            self.assertEqual(Path(report["projects"][0]["output"]).read_bytes(), before)

    def test_visual_review_is_bounded_and_never_changes_ocr_gate(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            make_docx(root / "甲说明.docx", image_count=8)
            # The fixture payloads are not actual PNGs, so replace media with a
            # valid signature to exercise the vision path without image decoding.
            source = root / "甲说明.docx"
            with zipfile.ZipFile(source) as archive:
                contents = [(item, archive.read(item.filename)) for item in archive.infolist()]
            with zipfile.ZipFile(source, "w") as archive:
                for item, data in contents:
                    archive.writestr(item, b"\x89PNG\r\n\x1a\n" if item.filename.startswith("word/media/") else data)
            calls = []

            def vision(*_):
                calls.append(1)
                return {"status": "irrelevant", "confidence": 1.0, "reason": "图片明确显示其他项目主题"}

            report = MODULE.run_manual_content_stage(
                root, root / "out", root / "report.json", apply=False,
                ocr=lambda *_: {"text": ""}, vision_classifier=vision,
            )
            project = report["projects"][0]
            self.assertEqual(len(calls), 2)
            self.assertEqual(project["visual_call_count"], 2)
            self.assertEqual(project["visual_selected_indexes"], [0, 7])
            self.assertEqual(project["visual_candidate_count"], 4)
            self.assertEqual(project["visual_decision_count"], 2)
            self.assertEqual(project["irrelevant_count"], 0)
            self.assertFalse(project["blocked"])
            self.assertFalse(report["batch_fused"])
            self.assertEqual(report["visual_call_count"], 2)


class VisionAdapterTests(unittest.TestCase):
    def test_local_chat_completions_contract(self):
        requests = []

        class Handler(BaseHTTPRequestHandler):
            def do_POST(self):
                body = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
                requests.append((self.path, self.headers.get("Authorization"), body))
                response = {"choices": [{"message": {"content": '```json\n{"status":"irrelevant","confidence":0.94,"evidence":"图片明确显示仓储货架及库存盘点界面"}\n```'}}], "usage": {"total_tokens": 23}}
                encoded = json.dumps(response).encode()
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(encoded)))
                self.end_headers()
                self.wfile.write(encoded)

            def log_message(self, *_):
                pass

        server = HTTPServer(("127.0.0.1", 0), Handler)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        image = io.BytesIO()
        Image.new("RGB", (8, 8), "white").save(image, format="PNG")
        try:
            adapter = MODULE.LocalVisionAdapter(base_url=f"http://127.0.0.1:{server.server_port}/v1", api_key="proxy", timeout=2)
            result = adapter(image.getvalue(), ".png", "医院人力资源系统", "", "")
        finally:
            server.shutdown()
            server.server_close()
            thread.join(timeout=2)
        self.assertEqual(result["status"], "irrelevant")
        self.assertEqual(result["usage"]["total_tokens"], 23)
        path, auth, body = requests[0]
        self.assertEqual(path, "/v1/chat/completions")
        self.assertEqual(auth, "Bearer proxy")
        self.assertEqual(body["model"], "sensenova-6.8-flash-lite")
        image_url = body["messages"][0]["content"][1]["image_url"]["url"]
        self.assertTrue(image_url.startswith("data:image/jpeg;base64,"))
        self.assertTrue(base64.b64decode(image_url.partition(",")[2]).startswith(b"\xff\xd8"))

    def test_missing_evidence_is_unknown(self):
        with mock.patch.object(MODULE.urllib.request, "urlopen") as post:
            post.return_value.__enter__.return_value.read.return_value = json.dumps({
                "choices": [{"message": {"content": '{"status":"irrelevant","confidence":1,"evidence":""}'}}]
            }).encode()
            result = MODULE.LocalVisionAdapter()(b"png", ".png", "标题", "", "")
        self.assertEqual(result["status"], "unknown")


if __name__ == "__main__":
    unittest.main()

from __future__ import annotations

import contextlib
import importlib.util
import json
import subprocess
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
# 被测模块按文件路径加载，需要把 scripts 加入 sys.path，
# 否则被测模块内部的同级 import（如 archive_types）无法解析。
sys.path.insert(0, str(ROOT / "scripts"))
SPEC = importlib.util.spec_from_file_location("unified_workflow_progress_tests", ROOT / "scripts" / "unified_workflow.py")
WORKFLOW = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(WORKFLOW)


class ProgressReportingTests(unittest.TestCase):
    def assert_failed_progress(self, reporter, expected_total=None):
        state = json.loads(reporter.path.read_text(encoding="utf-8"))
        self.assertEqual(state["status"], "FAILED")
        self.assertFalse(state["current_project_active"])
        self.assertGreaterEqual(state["failures"], 1)
        self.assertGreaterEqual(state["total"], 1)
        self.assertEqual(state["current"], state["total"])
        self.assertTrue(state.get("error"))
        if expected_total is not None:
            self.assertEqual(state["total"], expected_total)
        return state

    def test_pdf_failure_sets_failed_terminal_state_and_stable_total_before_lock(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            modified = base / "modified"
            work = base / "work"
            modified.mkdir()
            (modified / "a.docx").write_bytes(b"docx")
            (modified / "b.docx").write_bytes(b"docx")
            reporter = WORKFLOW.ProgressReporter(work)
            observed_totals = []

            @contextlib.contextmanager
            def fake_word_lock(progress, module, timeout_seconds=180):
                observed_totals.append(progress.total)
                yield

            def fake_run(command, **kwargs):
                report_path = Path(command[command.index("-ReportPath") + 1])
                report_path.parent.mkdir(parents=True, exist_ok=True)
                report_path.write_text(json.dumps({
                    "status": "EXPORT_FAILED", "exported": 1, "skipped": 0,
                    "failed": 1, "files": [],
                }), encoding="utf-8")
                return subprocess.CompletedProcess(command, 0, stdout=b"")

            with mock.patch.object(WORKFLOW.shutil, "which", return_value="powershell.exe"), \
                    mock.patch.object(WORKFLOW, "word_automation_lock", fake_word_lock), \
                    mock.patch.object(WORKFLOW.subprocess, "run", side_effect=fake_run):
                result = WORKFLOW.export_modified_pdfs(modified, work, reporter)

            state = json.loads(reporter.path.read_text(encoding="utf-8"))
            self.assertEqual(observed_totals, [2])
            self.assertEqual(result["failed"], 1)
            self.assertEqual(state["total"], 2)
            self.assertEqual(state["failures"], 1)
            self.assertEqual(state["status"], "FAILED")

    def test_missing_modified_directory_writes_failed_terminal_progress(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            reporter = WORKFLOW.ProgressReporter(base / "work")
            result = WORKFLOW.export_modified_pdfs(base / "missing", base / "work", reporter)
            self.assertEqual(result["status"], "FAILED")
            self.assertIn("未找到修改版目录", result["error"])
            self.assert_failed_progress(reporter, 1)

    def test_missing_powershell_writes_failed_terminal_progress(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            modified = base / "modified"
            modified.mkdir()
            (modified / "a.docx").write_bytes(b"docx")
            (modified / "b.docx").write_bytes(b"docx")
            reporter = WORKFLOW.ProgressReporter(base / "work")
            with mock.patch.object(WORKFLOW.shutil, "which", return_value=None):
                result = WORKFLOW.export_modified_pdfs(modified, base / "work", reporter)
            self.assertEqual(result["status"], "FAILED")
            self.assertIn("未找到 PowerShell", result["error"])
            self.assert_failed_progress(reporter, 2)

    def test_duplicate_export_lock_preserves_total_and_writes_failed_terminal_progress(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            modified = base / "modified"
            work = base / "work"
            modified.mkdir()
            work.mkdir()
            (modified / "a.docx").write_bytes(b"docx")
            (modified / "b.docx").write_bytes(b"docx")
            lock_path = work / ".modified-pdf-export.running"
            lock_path.write_text(json.dumps({"pid": 12345}), encoding="utf-8")
            reporter = WORKFLOW.ProgressReporter(work)
            with mock.patch.object(WORKFLOW.shutil, "which", return_value="powershell.exe"), \
                    mock.patch.object(WORKFLOW, "process_is_running", return_value=True):
                result = WORKFLOW.export_modified_pdfs(modified, work, reporter)
            self.assertEqual(result["status"], "FAILED")
            self.assertIn("正在导出中", result["error"])
            self.assertTrue(lock_path.exists())
            self.assert_failed_progress(reporter, 2)

    def test_subprocess_timeout_writes_failed_terminal_progress(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            modified = base / "modified"
            work = base / "work"
            modified.mkdir()
            (modified / "a.docx").write_bytes(b"docx")
            reporter = WORKFLOW.ProgressReporter(work)
            with mock.patch.object(WORKFLOW.shutil, "which", return_value="powershell.exe"), \
                    mock.patch.object(WORKFLOW, "word_automation_lock", return_value=contextlib.nullcontext()), \
                    mock.patch.object(WORKFLOW.subprocess, "run", side_effect=subprocess.TimeoutExpired("powershell", 3600)):
                result = WORKFLOW.export_modified_pdfs(modified, work, reporter)
            self.assertEqual(result["status"], "FAILED")
            self.assertIn("超时", result["error"])
            self.assert_failed_progress(reporter, 1)

    def test_missing_subprocess_report_writes_failed_terminal_progress(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            modified = base / "modified"
            work = base / "work"
            modified.mkdir()
            (modified / "a.docx").write_bytes(b"docx")
            reporter = WORKFLOW.ProgressReporter(work)
            completed = subprocess.CompletedProcess(["powershell"], 1, stdout=b"")
            with mock.patch.object(WORKFLOW.shutil, "which", return_value="powershell.exe"), \
                    mock.patch.object(WORKFLOW, "word_automation_lock", return_value=contextlib.nullcontext()), \
                    mock.patch.object(WORKFLOW.subprocess, "run", return_value=completed):
                result = WORKFLOW.export_modified_pdfs(modified, work, reporter)
            self.assertEqual(result["status"], "FAILED")
            self.assertIn("导出未返回报告", result["error"])
            self.assert_failed_progress(reporter, 1)

    def test_code_cleaner_outer_exception_writes_failed_terminal_progress(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            archive = base / "batch.rar"
            work = base / "work"
            archive.write_bytes(b"rar")
            failing_module = types.SimpleNamespace(run=mock.Mock(side_effect=RuntimeError("scan exploded")))
            arguments = [
                "unified_workflow.py", str(archive), "--clean-code-comments",
                "--work-dir", str(work),
            ]
            with mock.patch.dict(sys.modules, {"code_comment_cleaner": failing_module}), \
                    mock.patch.object(sys, "argv", arguments), \
                    contextlib.redirect_stdout(__import__("io").StringIO()):
                exit_code = WORKFLOW.main()

            state = json.loads((work / "progress.json").read_text(encoding="utf-8"))
            self.assertEqual(exit_code, 2)
            self.assertEqual(state["status"], "FAILED")
            self.assertGreaterEqual(state["total"], 1)
            self.assertEqual(state["current"], state["total"])
            self.assertGreaterEqual(state["failures"], 1)
            self.assertFalse(state["current_project_active"])

    def test_powershell_export_contract_never_reports_success_with_failures(self):
        script = (ROOT / "scripts" / "export_modified_pdfs.ps1").read_text(encoding="utf-8-sig")
        self.assertIn('failures = @($results | Where-Object status -eq "FAILED").Count', script)
        self.assertIn('if ($fatalError -or $failedCount -gt 0) { "EXPORT_FAILED" }', script)
        self.assertIn('if ($finalStatus -eq "EXPORT_OK") { "DONE" } else { "FAILED" }', script)


if __name__ == "__main__":
    unittest.main()

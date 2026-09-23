from __future__ import annotations

import contextlib
import importlib.util
import io
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
SPEC = importlib.util.spec_from_file_location("unified_document_entry_tests", ROOT / "scripts" / "unified_workflow.py")
WORKFLOW = importlib.util.module_from_spec(SPEC)
assert SPEC.loader
SPEC.loader.exec_module(WORKFLOW)


class UnifiedDocumentEntryTests(unittest.TestCase):
    def invoke(self, flag: str, result: dict, extra: list[str] | None = None):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "materials"
            source.mkdir()
            work = Path(directory) / "work"
            fake_run = mock.Mock(return_value=result)
            fake_module = types.SimpleNamespace(run=fake_run)
            arguments = ["unified_workflow.py", str(source), "--work-dir", str(work), flag, *(extra or [])]
            with mock.patch.dict(sys.modules, {"combined_document_workflow": fake_module}), \
                    mock.patch.object(sys, "argv", arguments), contextlib.redirect_stdout(io.StringIO()):
                exit_code = WORKFLOW.main()
            return exit_code, fake_run

    def test_preview_entry_uses_combined_workflow_without_apply(self):
        code, run = self.invoke("--process-documents", {"status": "NEEDS_DOCUMENT_PROCESSING_REVIEW"})
        self.assertEqual(code, 2)
        self.assertFalse(run.call_args.kwargs["apply"])

    def test_apply_entry_uses_same_combined_workflow(self):
        code, run = self.invoke("--apply-document-processing", {"status": "DOCUMENT_PROCESSING_OK"})
        self.assertEqual(code, 0)
        self.assertTrue(run.call_args.kwargs["apply"])

    def test_document_stages_are_forwarded_to_preview_and_apply(self):
        for flag, status in (("--process-documents", "NEEDS_DOCUMENT_PROCESSING_REVIEW"),
                             ("--apply-document-processing", "DOCUMENT_PROCESSING_OK")):
            with self.subTest(flag=flag):
                _, run = self.invoke(flag, {"status": status}, ["--document-stages", "manual,txt_placeholder"])
                self.assertEqual(run.call_args.kwargs["document_stages"], "manual,txt_placeholder")


if __name__ == "__main__":
    unittest.main()

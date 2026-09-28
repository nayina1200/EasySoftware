from __future__ import annotations

import os
import contextlib
import io
import sys
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import validation_isolation
import refurbish_workflow


class ValidationIsolationTests(unittest.TestCase):
    def test_unconfigured_guard_preserves_existing_runtime_behavior(self):
        with tempfile.TemporaryDirectory() as directory:
            outside = Path(directory) / "formal-orders"
            no_marker = Path(directory) / "ordinary-project" / ".optimization-dev"
            with mock.patch.object(validation_isolation, "DEVELOPMENT_MARKER", no_marker), mock.patch.dict(os.environ, {}, clear=True):
                self.assertEqual(validation_isolation.require_validation_path(outside, kind="material"), outside.resolve())

    def test_materials_can_be_anywhere_but_workspaces_stay_under_validation_root(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "validation"
            runs = root / "runs"
            sample = runs / "case-001" / "input"
            sample.mkdir(parents=True)
            work = root / "temp" / "case-001"
            work.mkdir(parents=True)
            environment = {
                "EASYSOFTWARE_VALIDATION_ROOT": str(root),
                "EASYSOFTWARE_VALIDATION_RUNS_ROOT": str(runs),
            }
            no_marker = Path(directory) / "ordinary-project" / ".optimization-dev"
            with mock.patch.object(validation_isolation, "DEVELOPMENT_MARKER", no_marker):
                self.assertEqual(validation_isolation.require_validation_path(sample, kind="material", environ=environment), sample.resolve())
                self.assertEqual(validation_isolation.require_validation_path(root / "inputs" / "clean-corpus", kind="material", environ=environment), (root / "inputs" / "clean-corpus").resolve())
                self.assertEqual(validation_isolation.require_validation_path(Path(directory) / "formal-orders", kind="material", environ=environment), (Path(directory) / "formal-orders").resolve())
                self.assertEqual(validation_isolation.require_validation_path(work, kind="workspace", environ=environment), work.resolve())
                with self.assertRaises(validation_isolation.ValidationIsolationError):
                    validation_isolation.require_validation_path(Path(directory) / "formal-output", kind="workspace", environ=environment)
                with self.assertRaises(validation_isolation.ValidationIsolationError):
                    validation_isolation.require_validation_path(runs / "case-001" / "input", kind="workspace", environ=environment)
                self.assertEqual(validation_isolation.require_validation_path(runs / "case-001", kind="material", environ=environment), (runs / "case-001").resolve())

    def test_development_marker_fixes_workspace_root_but_accepts_external_materials(self):
        with tempfile.TemporaryDirectory() as directory:
            outside = Path(directory) / "formal-orders"
            marker = Path(directory) / "development-copy" / ".optimization-dev"
            marker.parent.mkdir(parents=True)
            marker.touch()
            environment = {}
            with mock.patch.object(validation_isolation, "DEVELOPMENT_MARKER", marker):
                self.assertTrue(validation_isolation.activate_optimization_isolation(environment))
            fixed_root = marker.parent.parent / "EasySoftware_优化验证材料"
            self.assertEqual(Path(environment["TEMP"]), fixed_root / "temp")
            self.assertEqual(Path(environment["TMP"]), fixed_root / "temp")
            self.assertEqual(Path(environment["PYTHONPYCACHEPREFIX"]), fixed_root / "cache" / "python")
            self.assertEqual(Path(environment["EASYSOFTWARE_DATE_TOOL_ROOT"]), fixed_root / "dependencies" / "date_ocr")
            with mock.patch.object(validation_isolation, "DEVELOPMENT_MARKER", marker):
                self.assertEqual(validation_isolation.require_validation_path(outside, kind="material", environ=environment), outside.resolve())
            with mock.patch.object(validation_isolation, "DEVELOPMENT_MARKER", marker), \
                    mock.patch.dict(os.environ, {"EASYSOFTWARE_VALIDATION_ROOT": directory}, clear=True):
                with self.assertRaises(validation_isolation.ValidationIsolationError):
                    validation_isolation.activate_optimization_isolation()

    def test_development_refurbish_workflow_refuses_submit_mode_before_api_calls(self):
        with tempfile.TemporaryDirectory() as directory:
            marker = Path(directory) / ".optimization-dev"
            marker.touch()
            api = mock.Mock()
            with mock.patch.object(refurbish_workflow, "DEVELOPMENT_MARKER", marker), \
                    self.assertRaises(refurbish_workflow.WorkflowError):
                refurbish_workflow.run_workflow(api, "test-operator", dry_run=False)
            api.list_batches.assert_not_called()

    def test_development_refurbish_cli_rejects_external_report_before_writing(self):
        with tempfile.TemporaryDirectory() as directory:
            report = ROOT.parent / f"outside-{Path(directory).name}.html"
            marker = Path(directory) / "development-copy" / ".optimization-dev"
            marker.parent.mkdir(parents=True)
            marker.touch()
            with mock.patch.object(validation_isolation, "DEVELOPMENT_MARKER", marker), \
                    mock.patch.object(refurbish_workflow, "DEVELOPMENT_MARKER", marker), \
                    mock.patch.object(sys, "argv", ["refurbish_workflow.py", "--dry-run", "--report", str(report)]), \
                    mock.patch.dict(os.environ, {}, clear=True), \
                    contextlib.redirect_stdout(io.StringIO()) as output:
                code = refurbish_workflow.main()
            self.assertEqual(code, 2)
            self.assertIn("ISOLATION_BLOCKED", output.getvalue())
            self.assertFalse(report.exists())


if __name__ == "__main__":
    unittest.main()

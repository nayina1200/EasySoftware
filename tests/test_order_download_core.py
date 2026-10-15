from __future__ import annotations

import contextlib
import io
import json
import sys
import subprocess
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import order_download_core
from order_download_core import finalize_download, safe_component, validate_zip


class OrderDownloadCoreTests(unittest.TestCase):
    def test_cli_success_receipt_is_ascii_safe_for_powershell(self):
        output = io.StringIO()
        with (
            mock.patch.object(sys, "argv", ["order_download_core.py", "finalize", "--source", "download.zip", "--destination", "deliver", "--customer", "客户"]),
            mock.patch.object(Path, "resolve", autospec=True, side_effect=lambda path, strict=False: path),
            mock.patch.object(order_download_core, "validate_zip", return_value=5),
            mock.patch.object(order_download_core, "_copy_validated_zip", return_value=Path("D:/订单/客户_5件/客户_5件.zip")),
            contextlib.redirect_stdout(output),
        ):
            self.assertEqual(order_download_core.main(), 0)
        receipt = output.getvalue()
        receipt.encode("ascii")
        self.assertEqual(json.loads(receipt)["count"], 5)

    def make_zip(self, path: Path, content: bytes = b"material") -> Path:
        with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("order/document.txt", content)
        return path

    def test_safe_component_replaces_windows_path_characters(self):
        self.assertEqual(safe_component('客户:甲/乙?*'), "客户_甲_乙__")
        self.assertEqual(safe_component("CON"), "_CON")
        with self.assertRaises(ValueError):
            safe_component("... ")

    def test_finalize_creates_same_name_folder_and_preserves_zip(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = self.make_zip(root / "download.zip")
            output = finalize_download(source, root / "deliver", "北方知识产权")

            self.assertEqual(output.relative_to(root / "deliver").as_posix(), "北方知识产权_1件/北方知识产权_1件.zip")
            self.assertEqual(output.read_bytes(), source.read_bytes())
            validate_zip(output)
            self.assertEqual(list(output.parent.glob("*.part")), [])

    def test_finalize_refuses_to_overwrite_existing_download(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = self.make_zip(root / "download.zip")
            existing = root / "deliver" / "客户_1件" / "客户_1件.zip"
            existing.parent.mkdir(parents=True)
            existing.write_bytes(b"keep me")

            with self.assertRaisesRegex(FileExistsError, "未覆盖"):
                finalize_download(source, root / "deliver", "客户")
            self.assertEqual(existing.read_bytes(), b"keep me")
            self.assertEqual(list(existing.parent.glob("*.part")), [])

    def test_zip_file_count_excludes_directory_entries_and_counts_nested_files(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "download.zip"
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("folder/", b"")
                archive.writestr("folder/nested/material.txt", b"nested")
                archive.writestr("readme.txt", b"root")

            output = finalize_download(source, root / "deliver", "客户")
            self.assertEqual(output.name, "客户_2件.zip")

    def test_validate_zip_accepts_single_member_archive(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "single.zip"
            self.make_zip(source)
            self.assertEqual(validate_zip(source), 1)

    def test_finalize_rejects_empty_archive_and_invalid_zip(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            empty = root / "empty.zip"
            with zipfile.ZipFile(empty, "w"):
                pass
            with self.assertRaisesRegex(ValueError, "没有可下载文件"):
                finalize_download(empty, root / "deliver", "客户")

            broken = root / "broken.zip"
            broken.write_bytes(b"not a zip")
            with self.assertRaisesRegex(ValueError, "有效 ZIP"):
                validate_zip(broken)

    def test_finalize_keeps_original_uppercase_extension(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = self.make_zip(root / "download.ZIP")
            output = finalize_download(source, root / "deliver", "客户")
            self.assertEqual(output.name, "客户_1件.ZIP")

    def test_cli_returns_json_with_actual_archive_file_count(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = self.make_zip(root / "download.zip")
            completed = subprocess.run(
                [
                    sys.executable,
                    str(ROOT / "scripts" / "order_download_core.py"),
                    "finalize",
                    "--source", str(source),
                    "--destination", str(root / "deliver"),
                    "--customer", "客户",
                ],
                capture_output=True,
                text=True,
                check=False,
            )

            self.assertEqual(completed.returncode, 0, completed.stderr)
            completed.stdout.encode("ascii")
            result = json.loads(completed.stdout)
            self.assertEqual(result["status"], "OK")
            self.assertEqual(result["count"], 1)
            self.assertEqual(result["filename"], "客户_1件.zip")

    def test_cli_marks_invalid_zip_as_validation_failure(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "broken.zip"
            source.write_bytes(b"not a zip")
            completed = subprocess.run(
                [
                    sys.executable,
                    str(ROOT / "scripts" / "order_download_core.py"),
                    "finalize",
                    "--source", str(source),
                    "--destination", str(root / "deliver"),
                    "--customer", "客户",
                ],
                capture_output=True,
                text=True,
                check=False,
            )

            self.assertEqual(completed.returncode, 1)
            result = json.loads(completed.stdout)
            self.assertEqual(result["status"], "FAILED")
            self.assertEqual(result["stage"], "validation")

    def test_cli_marks_archive_publish_failure_separately(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = self.make_zip(root / "download.zip")
            existing = root / "deliver" / "客户_1件" / "客户_1件.zip"
            existing.parent.mkdir(parents=True)
            existing.write_bytes(b"keep me")
            completed = subprocess.run(
                [
                    sys.executable,
                    str(ROOT / "scripts" / "order_download_core.py"),
                    "finalize",
                    "--source", str(source),
                    "--destination", str(root / "deliver"),
                    "--customer", "客户",
                ],
                capture_output=True,
                text=True,
                check=False,
            )

            self.assertEqual(completed.returncode, 1)
            result = json.loads(completed.stdout)
            self.assertEqual(result["status"], "FAILED")
            self.assertEqual(result["stage"], "save")
            self.assertEqual(existing.read_bytes(), b"keep me")


if __name__ == "__main__":
    unittest.main()

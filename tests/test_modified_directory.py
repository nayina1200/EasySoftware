from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import unified_workflow  # noqa: E402


def make_tree(relative_dirs: list[str], root: Path) -> None:
    for relative in relative_dirs:
        folder = root / relative
        folder.mkdir(parents=True, exist_ok=True)
        (folder / "代码.docx").write_bytes(b"material")


class ResolveModifiedDirectoryTests(unittest.TestCase):
    """修改版目录的定位不能依赖已解压文件夹命名为“解压版”。"""

    def resolve(self, folders: list[str], source_name: str):
        tmp = Path(tempfile.mkdtemp(prefix="modified-dir-"))
        self.addCleanup(lambda: __import__("shutil").rmtree(tmp, ignore_errors=True))
        make_tree(folders, tmp)
        source = tmp / source_name
        return unified_workflow.resolve_modified_directory(source.parent, source), tmp

    def test_folder_without_extracted_suffix_finds_sibling_modified(self):
        got, _ = self.resolve(["材料A", "材料A_修改版"], "材料A")
        self.assertEqual(got.name, "材料A_修改版")
        self.assertTrue(got.is_dir())

    def test_folder_named_extracted_still_resolves(self):
        got, _ = self.resolve(["解压版", "修改版"], "解压版")
        self.assertEqual(got.name, "修改版")
        self.assertTrue(got.is_dir())

    def test_prefixed_extracted_suffix_strips_only_suffix(self):
        got, _ = self.resolve(["X_解压版", "X_修改版"], "X_解压版")
        self.assertEqual(got.name, "X_修改版")
        self.assertTrue(got.is_dir())

    def test_modified_at_batch_root(self):
        got, _ = self.resolve(["我的材料", "修改版"], "我的材料")
        self.assertEqual(got.name, "修改版")
        self.assertTrue(got.is_dir())

    def test_selected_folder_is_itself_modified(self):
        got, _ = self.resolve(["修改版"], "修改版")
        self.assertEqual(got.name, "修改版")
        self.assertTrue(got.is_dir())

    def test_modified_nested_inside_source_folder(self):
        got, _ = self.resolve(["我的材料/修改版"], "我的材料")
        self.assertEqual(got.name, "修改版")
        self.assertTrue(got.is_dir())

    def test_empty_modified_dir_is_skipped(self):
        # 空目录不算可用修改版，继续往后找。
        got, _ = self.resolve(["材料A/修改版", "修改版"], "材料A")
        self.assertEqual(got.name, "修改版")
        self.assertTrue(got.is_dir())

    def test_falls_back_to_name_convention_guess(self):
        got, _ = self.resolve(["材料A"], "材料A")
        self.assertEqual(got.name, "材料A_修改版")


class LauncherModifiedDirectoryGuardTests(unittest.TestCase):
    """启动器里的修改版目录定位必须与 Python 侧共用同一套查找顺序。"""

    def setUp(self):
        self.ps1 = (ROOT / "启动EasySoftware.ps1").read_text(encoding="utf-8-sig")

    def test_launcher_defines_shared_resolver(self):
        self.assertIn("function Resolve-ModifiedDirectory", self.ps1)

    def test_launcher_stops_handing_back_a_hardcoded_modified_path(self):
        # 旧实现直接在 fromExtracted 分支里返回固定路径，这就是
        # “文件夹不叫解压版就导不出 PDF”的门槛。
        self.assertNotIn('return (Join-Path $source.Parent.FullName "修改版")', self.ps1)

    def test_launcher_uses_the_shared_resolver(self):
        self.assertIn("Resolve-ModifiedDirectory $batchRoot $source.FullName", self.ps1)

    def test_launcher_strips_extracted_suffix_with_substring(self):
        # 后缀必须按字面长度剥离；GetFileNameWithoutExtension 不会去掉
        # 没有点的后缀，会把 X_解压版 猜成 X_解压版_修改版。
        self.assertIn('$name.Substring(0, $name.Length - 4) + "_修改版"', self.ps1)
        self.assertNotIn("GetFileNameWithoutExtension($name)", self.ps1)


if __name__ == "__main__":
    unittest.main()

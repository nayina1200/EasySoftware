from __future__ import annotations

import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from title_body_cache import TitleBodyCache
from title_body_qualification import extract_docx_paragraphs, index_by_relative_path, qualify_materials, scan_title_body_inputs, validate_evidence_result


def write_docx(path: Path, paragraphs: list[str], *, toc: bool = False, heading_indexes: set[int] | None = None, hidden_indexes: set[int] | None = None, hidden_style_indexes: set[int] | None = None, default_hidden: bool = False, doc_defaults_hidden: bool = False) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    heading_indexes = heading_indexes or set()
    hidden_indexes = hidden_indexes or set()
    hidden_style_indexes = hidden_style_indexes or set()
    pxml = []
    for index, text in enumerate(paragraphs):
        if toc and index == 0:
            pxml.append('<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>%s</w:t></w:r></w:p>' % text)
        elif index == 1:
            pxml.append('<w:tbl><w:tr><w:tc><w:p><w:r><w:t>%s</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' % text)
        elif index in heading_indexes:
            pxml.append('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>%s</w:t></w:r></w:p>' % text)
        elif index in hidden_indexes:
            pxml.append('<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>隐藏内容不得进入判定</w:t></w:r><w:r><w:t>%s</w:t></w:r></w:p>' % text)
        elif index in hidden_style_indexes:
            pxml.append('<w:p><w:r><w:rPr><w:rStyle w:val="HiddenChar"/></w:rPr><w:t>%s</w:t></w:r></w:p>' % text)
        else:
            pxml.append('<w:p><w:r><w:t>%s</w:t></w:r></w:p>' % text)
    document = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body>%s</w:body></w:document>' % "".join(pxml)
    )
    default_style = (
        '<w:style w:type="paragraph" w:default="1" w:styleId="DefaultHidden"><w:name w:val="Default Hidden"/><w:rPr><w:vanish/></w:rPr></w:style>'
        if default_hidden else ""
    )
    document_defaults = (
        '<w:docDefaults><w:rPrDefault><w:rPr><w:vanish/></w:rPr></w:rPrDefault></w:docDefaults>'
        if doc_defaults_hidden else ""
    )
    styles = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">%s%s'
        '<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="目录 1"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>'
        '<w:style w:type="character" w:styleId="HiddenChar"><w:name w:val="Hidden Character"/><w:rPr><w:vanish/></w:rPr></w:style>'
        '</w:styles>'
    ) % (document_defaults, default_style)
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", document)
        archive.writestr("word/styles.xml", styles)


class StrictPairingTests(unittest.TestCase):
    def test_root_level_single_pair_with_unrelated_stems_is_ambiguous(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "正确标题.txt").write_text("", encoding="utf-8")
            write_docx(root / "无关说明.docx", ["正文业务内容"])

            [item] = scan_title_body_inputs(root)

        self.assertEqual(item["status"], "NOT_EVALUATED")
        self.assertEqual(item["metadata"]["pairing_reason"], "PAIRING_AMBIGUOUS")
        self.assertIsNone(item["txt_relative_path"])

    def test_hidden_run_text_is_not_extracted(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "hidden.docx"
            write_docx(path, ["可见业务内容"], hidden_indexes={0})

            [paragraph] = extract_docx_paragraphs(path)

        self.assertEqual(paragraph["text"], "可见业务内容")

    def test_text_hidden_through_character_style_inheritance_is_not_extracted(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "hidden-style.docx"
            write_docx(path, ["继承隐藏的内容", "可见正文内容"], hidden_style_indexes={0})

            rows = extract_docx_paragraphs(path)

        self.assertEqual([row["text"] for row in rows], ["可见正文内容"])

    def test_text_hidden_through_default_paragraph_style_is_not_extracted(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "hidden-default-style.docx"
            write_docx(path, ["默认段落样式隐藏正文"], default_hidden=True)

            rows = extract_docx_paragraphs(path)

        self.assertEqual(rows, [])

    def test_text_hidden_through_doc_defaults_is_not_extracted(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "hidden-doc-defaults.docx"
            write_docx(path, ["文档默认格式隐藏正文"], doc_defaults_hidden=True)

            rows = extract_docx_paragraphs(path)

        self.assertEqual(rows, [])

    def test_section_titles_mark_overview_and_main_function_roles(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "sectioned.docx"
            write_docx(path, ["项目概述", "概述正文", "主要功能模块", "具体业务功能结果"], heading_indexes={0, 2})

            rows = extract_docx_paragraphs(path)

        overview_body = rows[1]
        function_body = rows[3]
        self.assertEqual(overview_body["section_title"], "项目概述")
        self.assertEqual(overview_body["section_role"], "overview")
        self.assertTrue(overview_body["is_overview"])
        self.assertEqual(function_body["section_title"], "主要功能模块")
        self.assertEqual(function_body["section_role"], "main_function")
        self.assertTrue(function_body["is_functional"])
        self.assertTrue(function_body["is_main_function"])

    def test_single_pair_uses_txt_filename_as_title_and_preserves_docx_structure(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            project = root / "TXT 真正标题"
            project.mkdir()
            (project / "TXT 真正标题.txt").write_text("字段内容不得作为标题", encoding="utf-8")
            write_docx(project / "其他名称说明.docx", ["封面", "表格内的业务字段", "正文中的具体功能结果"])
            (project / "其他名称代码.pdf").write_bytes(b"code pdf")
            [item] = scan_title_body_inputs(root)
            self.assertEqual(item["title"], "TXT 真正标题")
            self.assertEqual(item["metadata"]["title_source"], "txt_filename_stem")
            table_paragraph = item["paragraphs"][1]
            self.assertEqual(table_paragraph["paragraph_id"], "paragraph:2")
            self.assertTrue(table_paragraph["in_table"])
            self.assertIn("表格内的业务字段", table_paragraph["text"])
            self.assertEqual(item["relative_path"], "TXT 真正标题/其他名称说明.docx")
            self.assertEqual(set(item["associated_files"]), {
                "TXT 真正标题/其他名称说明.docx", "TXT 真正标题/TXT 真正标题.txt", "TXT 真正标题/其他名称代码.pdf",
            })
            self.assertIs(index_by_relative_path([item])[item["relative_path"]], item)

    def test_missing_and_ambiguous_txt_never_call_classifier(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            missing = root / "缺失"
            write_docx(missing / "项目说明.docx", ["有内容"])
            ambiguous = root / "歧义"
            write_docx(ambiguous / "项目说明.docx", ["有内容"])
            (ambiguous / "项目.txt").write_text("", encoding="utf-8")
            (ambiguous / "项目　.txt").write_text("", encoding="utf-8")
            calls = []
            items = qualify_materials(root, classify=lambda *args: calls.append(args) or {"status": "MATCH"})
            by_reason = {item["metadata"]["pairing_reason"]: item for item in items}
            self.assertEqual(by_reason["TXT_MISSING"]["status"], "NOT_EVALUATED")
            self.assertEqual(by_reason["PAIRING_AMBIGUOUS"]["status"], "NOT_EVALUATED")
            self.assertEqual(by_reason["TXT_MISSING"]["associated_files"], ["缺失/项目说明.docx"])
            self.assertEqual(calls, [])

    def test_multi_materials_require_unique_normalized_filename_mapping(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "系统一.txt").write_text("", encoding="utf-8")
            (root / "系统二.txt").write_text("", encoding="utf-8")
            write_docx(root / "系统一说明.docx", ["正文 A"])
            write_docx(root / "系统二说明.docx", ["正文 B"])
            items = scan_title_body_inputs(root)
            self.assertEqual([item["metadata"]["pairing_status"] for item in items], ["PAIRED", "PAIRED"])
            self.assertNotEqual(items[0]["project_id"], items[1]["project_id"])

    def test_associated_file_sets_do_not_include_neighbor_projects(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for name in ("甲", "乙"):
                (root / f"{name}.txt").write_text("", encoding="utf-8")
                write_docx(root / f"{name}说明.docx", [f"{name}正文"])
                write_docx(root / f"{name}代码.docx", [f"{name}源代码"])
                (root / f"{name}说明.pdf").write_bytes(b"manual pdf")
                (root / f"{name}代码.pdf").write_bytes(b"code pdf")
            items = {item["title"]: item for item in scan_title_body_inputs(root)}
            self.assertEqual(set(items["甲"]["associated_files"]), {
                "甲.txt", "甲说明.docx", "甲代码.docx", "甲说明.pdf", "甲代码.pdf",
            })
            self.assertNotIn("乙说明.docx", items["甲"]["associated_files"])

    def test_nested_toc_and_title_repeat_metadata_are_preserved(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            project = root / "主题系统"
            project.mkdir()
            (project / "主题系统.txt").write_text("", encoding="utf-8")
            write_docx(project / "说明书说明.docx", ["主题系统", "主题系统", "表格正文", "实质业务段落"], toc=True)
            item = scan_title_body_inputs(root)[0]
            self.assertTrue(item["paragraphs"][0]["is_toc"])
            self.assertTrue(item["paragraphs"][0]["is_title_repeat"])
            self.assertTrue(item["paragraphs"][1]["in_table"])


class EvidenceValidationTests(unittest.TestCase):
    def setUp(self):
        self.paragraphs = [
            {"paragraph_id": "paragraph:1", "text": "目录：业务处理流程", "is_toc": True, "is_title_repeat": False},
            {"paragraph_id": "paragraph:2", "text": "项目标题", "is_toc": False, "is_title_repeat": True},
            {"paragraph_id": "paragraph:3", "text": "系统采集用户报表并生成风险分析结果。", "is_toc": False, "is_title_repeat": False},
        ]

    def test_wrong_paragraph_id_or_non_substring_quote_is_not_evaluated(self):
        for evidence in (
            [{"paragraph_id": "paragraph:9", "quote": "生成风险分析"}],
            [{"paragraph_id": "paragraph:3", "quote": "生成天气预报"}],
        ):
            result = validate_evidence_result(self._result("MATCH", evidence, "yes"), self.paragraphs)
            self.assertEqual(result["status"], "NOT_EVALUATED")

    def test_toc_and_repeated_title_cannot_be_the_only_evidence(self):
        result = validate_evidence_result(self._result(
            "MISMATCH", [
                {"paragraph_id": "paragraph:1", "quote": "业务处理流程"},
                {"paragraph_id": "paragraph:2", "quote": "项目标题"},
            ], "no"
        ), self.paragraphs)
        self.assertEqual(result["status"], "NOT_EVALUATED")

    def test_valid_substantive_quote_is_accepted(self):
        result = validate_evidence_result(self._result(
            "MISMATCH", [{"paragraph_id": "paragraph:3", "quote": "生成风险分析结果"}], "no"
        ), self.paragraphs)
        self.assertEqual(result["status"], "MISMATCH")
        self.assertTrue(result["evidence_valid"])

    def test_nested_claim_evidence_is_normalized_before_exact_quote_validation(self):
        result = validate_evidence_result({
            "status": "MISMATCH",
            "core_claims": [{
                "id": "claim-1",
                "claim": "项目必须支持业务审核",
                "supported": "no",
                "evidence": [{
                    "paragraph_id": "paragraph:3",
                    "quote": "生成风险分析结果",
                    "relation": "contradicts",
                }],
            }],
        }, self.paragraphs)

        self.assertEqual(result["status"], "MISMATCH")
        self.assertTrue(result["evidence_valid"])
        self.assertEqual(result["core_claims"][0]["evidence_ids"], ["evidence-1-1"])
        self.assertEqual(result["evidence"][0]["claim_id"], "claim-1")

    @staticmethod
    def _result(status, citations, supported):
        evidence = []
        evidence_ids = []
        for index, citation in enumerate(citations, 1):
            evidence_id = f"evidence-{index}"
            evidence_ids.append(evidence_id)
            evidence.append({
                **citation,
                "id": evidence_id,
                "claim_id": "claim-1",
                "relation": "contradicts" if supported == "no" else "supports",
            })
        return {
            "status": status,
            "core_claims": [{
                "id": "claim-1", "claim": "标题的必要主张", "supported": supported,
                "evidence_ids": evidence_ids,
            }],
            "evidence": evidence,
        }


class NonEvaluatedCacheTests(unittest.TestCase):
    def _make_input(self, root: Path) -> None:
        (root / "项目.txt").write_text("", encoding="utf-8")
        write_docx(root / "项目说明.docx", ["正文包含具体业务功能结果"])

    def test_content_unknown_becomes_not_evaluated_and_is_not_cached(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self._make_input(root)
            calls = []

            def classify(title, paragraphs, _metadata):
                calls.append(title)
                return {
                    "status": "UNKNOWN", "reason": "正文同时支持多个解释",
                    "evidence": [{"paragraph_id": paragraphs[0]["paragraph_id"], "quote": "具体业务功能结果"}],
                }

            cache_path = root / "title_body_cache.json"
            first = qualify_materials(root, classify=classify, cache=cache_path)[0]
            second = qualify_materials(root, classify=lambda *args: calls.append(args[0]) or {"status": "NOT_EVALUATED", "reason": "重试仍无法评估"}, cache=cache_path)[0]
            self.assertEqual(first["result"]["status"], "NOT_EVALUATED")
            self.assertEqual(second["result"]["status"], "NOT_EVALUATED")
            self.assertFalse(second["cache_hit"])
            self.assertEqual(calls, ["项目", "项目"])

    def test_failure_diagnostic_and_invalid_evidence_are_not_cached(self):
        responses = [
            {"status": "UNKNOWN", "failure": True},
            {"status": "UNKNOWN", "diagnostic_code": "TOKENIZER_FAILED"},
            {"status": "UNKNOWN", "evidence_valid": False},
            {"status": "UNKNOWN", "evidence": [{"paragraph_id": "paragraph:99", "quote": "不存在的引文"}]},
        ]
        for index, response in enumerate(responses):
            with self.subTest(index=index), tempfile.TemporaryDirectory() as directory:
                root = Path(directory)
                self._make_input(root)
                cache_path = root / "title_body_cache.json"
                item = qualify_materials(root, classify=lambda *_: response, cache=cache_path)[0]
                self.assertEqual(item["result"]["status"], "NOT_EVALUATED")
                self.assertIsNone(TitleBodyCache(cache_path).get(item["semantic_input_hash"], item["paragraphs"]))

    def test_cache_write_failure_preserves_the_classification(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self._make_input(root)
            result = EvidenceValidationTests._result(
                "MATCH", [{"paragraph_id": "paragraph:1", "quote": "正文包含具体业务功能结果"}], "yes"
            )
            with mock.patch.object(TitleBodyCache, "put", side_effect=OSError("disk full")):
                [item] = qualify_materials(root, classify=lambda *_: result, cache=root / "cache.json")

        self.assertEqual(item["result"]["status"], "MATCH")
        self.assertTrue(item["cache_warning"].startswith("CACHE_WRITE_FAILED"))

    def test_progress_callback_reports_active_and_terminal_state_per_item(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self._make_input(root)
            events = []
            qualify_materials(
                root,
                classify=lambda *_: {"status": "UNKNOWN", "reason": "需复核"},
                progress_callback=lambda phase, index, total, item: events.append((phase, index, total, item["title"])),
            )

        self.assertEqual(events, [("active", 1, 1, "项目"), ("complete", 1, 1, "项目")])


if __name__ == "__main__":
    unittest.main()

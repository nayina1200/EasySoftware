import importlib.util
import sys
import json
import tempfile
import unittest
import zipfile
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
from unittest import mock
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
# 被测模块按文件路径加载，需要把 scripts 加入 sys.path，
# 否则被测模块内部的同级 import（如 archive_types）无法解析。
sys.path.insert(0, str(ROOT / "scripts"))
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
        try:
            adapter = MODULE.LocalVisionAdapter(base_url=f"http://127.0.0.1:{server.server_port}/v1", api_key="proxy", timeout=2)
            result = adapter(b"\x89PNG\r\n\x1a\n", ".png", "医院人力资源系统", "", "")
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
        self.assertTrue(body["messages"][0]["content"][1]["image_url"]["url"].startswith("data:image/png;base64,"))

    def test_missing_evidence_is_unknown(self):
        with mock.patch.object(MODULE.urllib.request, "urlopen") as post:
            post.return_value.__enter__.return_value.read.return_value = json.dumps({
                "choices": [{"message": {"content": '{"status":"irrelevant","confidence":1,"evidence":""}'}}]
            }).encode()
            result = MODULE.LocalVisionAdapter()(b"png", ".png", "标题", "", "")
        self.assertEqual(result["status"], "unknown")


if __name__ == "__main__":
    unittest.main()

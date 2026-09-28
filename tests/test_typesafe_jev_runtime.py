from __future__ import annotations

import importlib.util
import json
import os
import sys
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location(
    "typesafe_jev_runtime_under_test", ROOT / "scripts" / "typesafe_jev_runtime.py"
)
RUNTIME_MODULE = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = RUNTIME_MODULE
assert SPEC.loader
SPEC.loader.exec_module(RUNTIME_MODULE)


class FakeResponse:
    def __init__(self, value: dict[str, object]):
        self.value = value

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    def read(self) -> bytes:
        return json.dumps(self.value).encode("utf-8")


def score_response(*, score: float = 3.0, confidence: float = 0.1) -> dict[str, object]:
    return {
        "model": "jev-1.13.0",
        "answers": {
            "title_body_fit": {
                "type": "score",
                "score": score,
                "confidence": confidence,
                "legend": {str(level): f"level {level}" for level in range(5)},
                "probabilities": {"0": 0, "1": 0, "2": 0, "3": 1, "4": 0},
            }
        },
        "usage": {"input_tokens": 120, "output_tokens": 12},
    }


class TypeSafeJevRuntimeTests(unittest.TestCase):
    def setUp(self):
        self.environment = mock.patch.dict(os.environ, {"TYPESAFE_API_KEY": "test-key"})
        self.environment.start()
        self.addCleanup(self.environment.stop)

    def make_runtime(self, **provider_options):
        return RUNTIME_MODULE.TypeSafeJevRuntime({
            "score_threshold": 70,
            "typesafe_jev": {"max_retry_after_seconds": 0, **provider_options},
        })

    def test_sends_full_named_body_and_threshold_alone_decides(self):
        runtime = self.make_runtime()
        self.assertEqual(runtime.start().state, "ready")
        paragraphs = [
            {
                "paragraph_id": "paragraph:1", "text": "完整正文段落甲",
                "section_id": "paragraph:1", "section_title": "主要功能",
                "section_role": "main_function", "is_heading": True,
                "in_table": False, "structure_uncertain": False, "is_title_repeat": True,
            },
            {
                "paragraph_id": "paragraph:2", "text": "表格中的完整正文乙",
                "section_id": "paragraph:1", "section_title": "主要功能",
                "section_role": "main_function", "in_table": True,
                "is_heading": False,
            },
            {
                "paragraph_id": "paragraph:3", "text": "目录中的完整正文丙",
                "is_toc": True, "section_id": "paragraph:1", "section_title": "目录",
            },
        ]
        captured = {}

        def fake_urlopen(request, timeout):
            captured["url"] = request.full_url
            captured["headers"] = request.headers
            captured["timeout"] = timeout
            captured["payload"] = json.loads(request.data)
            return FakeResponse(score_response())

        with mock.patch.object(RUNTIME_MODULE.urllib.request, "urlopen", side_effect=fake_urlopen):
            result = runtime.classify("文件名标题", paragraphs, {})

        self.assertEqual(captured["url"], RUNTIME_MODULE.API_URL)
        self.assertEqual(captured["payload"]["model"], "jev-1.13.0")
        self.assertEqual(captured["payload"]["state"]["title"], "文件名标题")
        self.assertEqual(
            captured["payload"]["state"]["body_paragraphs"],
            ["完整正文段落甲", "表格中的完整正文乙", "目录中的完整正文丙"],
        )
        self.assertEqual(
            captured["payload"]["state"]["non_evidentiary_paragraph_numbers"], [1, 3]
        )
        self.assertTrue(all(isinstance(value, str) for value in captured["payload"]["state"]["body_paragraphs"]))
        self.assertNotIn("section_title", json.dumps(captured["payload"]))
        instructions = captured["payload"]["questions"]["title_body_fit"]["instructions"]
        self.assertIn("公司名称", instructions)
        self.assertIn("核心领域", instructions)
        self.assertEqual(captured["payload"]["state"]["body_paragraphs"][1], "表格中的完整正文乙")
        self.assertEqual(result["status"], "MATCH")
        self.assertEqual(result["score_0_100"], 75)
        self.assertEqual(result["confidence"], 0.1)
        self.assertEqual(result["usage"], {"input_tokens": 120, "output_tokens": 12})
        self.assertEqual(result["diagnostics"]["complete_body_paragraph_count"], 3)
        self.assertNotIn("test-key", json.dumps(result))

    def test_full_17500_character_body_fits_compact_default_request_without_loss(self):
        runtime = self.make_runtime()
        runtime.start()
        body = "完整正文内容业务功能。" * 1591  # 17,501 source characters.
        captured = {}

        def fake_urlopen(request, timeout):
            captured["payload"] = json.loads(request.data)
            captured["request_characters"] = len(request.data.decode("utf-8"))
            return FakeResponse(score_response())

        with mock.patch.object(RUNTIME_MODULE.urllib.request, "urlopen", side_effect=fake_urlopen):
            result = runtime.classify("完整文件名标题", [{"paragraph_id": "p1", "text": body}], {})

        self.assertEqual(result["status"], "MATCH")
        self.assertLess(captured["request_characters"], runtime.max_request_characters)
        self.assertEqual(captured["payload"]["state"]["body_paragraphs"], [body])
        self.assertEqual(len(captured["payload"]["state"]["body_paragraphs"][0]), len(body))

    def test_rubric_and_selection_policy_versions_invalidate_old_cache_contract(self):
        self.assertEqual(RUNTIME_MODULE.RUBRIC_VERSION, "title-body-jev-domain-fit-rubric-v3")
        self.assertEqual(RUNTIME_MODULE.SELECTION_POLICY_VERSION, "title-body-complete-body-input-v2")
        config = json.loads((ROOT / "scripts" / "title_body_config.example.json").read_text(encoding="utf-8"))
        self.assertEqual(config["prompt_version"], "title-body-jev-domain-fit-score-v3")

    def test_oversized_complete_input_is_not_sent_or_truncated(self):
        runtime = self.make_runtime(max_request_characters=1000)
        runtime.start()
        with mock.patch.object(RUNTIME_MODULE.urllib.request, "urlopen") as urlopen:
            result = runtime.classify("标题", [{"paragraph_id": "p1", "text": "正文" * 1000}], {})
        urlopen.assert_not_called()
        self.assertEqual(result["status"], "NOT_EVALUATED")
        self.assertEqual(result["diagnostic_code"], "INPUT_TOO_LARGE")
        self.assertIn("未裁剪", result["reason"])

    def test_only_429_and_529_are_retried_with_a_bounded_attempt_count(self):
        runtime = self.make_runtime(max_retries=1, max_retry_after_seconds=0)
        runtime.start()
        errors = [
            __import__("urllib.error", fromlist=["HTTPError"]).HTTPError(
                RUNTIME_MODULE.API_URL, 429, "rate limited", {"Retry-After": "999"}, None
            ),
            __import__("urllib.error", fromlist=["HTTPError"]).HTTPError(
                RUNTIME_MODULE.API_URL, 429, "rate limited", {}, None
            ),
        ]
        with mock.patch.object(RUNTIME_MODULE.urllib.request, "urlopen", side_effect=errors) as urlopen:
            result = runtime.classify("标题", [{"paragraph_id": "p1", "text": "正文"}], {})
        self.assertEqual(urlopen.call_count, 2)
        self.assertEqual(result["status"], "NOT_EVALUATED")
        self.assertEqual(result["diagnostic_code"], "MODEL_UNAVAILABLE")
        self.assertEqual(result["diagnostics"]["generation_requests"], 2)

    def test_http_413_is_an_input_limit_not_a_service_outage(self):
        runtime = self.make_runtime()
        runtime.start()
        error = __import__("urllib.error", fromlist=["HTTPError"]).HTTPError(
            RUNTIME_MODULE.API_URL, 413, "payload too large", {}, None
        )
        with mock.patch.object(RUNTIME_MODULE.urllib.request, "urlopen", side_effect=error):
            result = runtime.classify("标题", [{"paragraph_id": "p1", "text": "正文"}], {})
        self.assertEqual(result["status"], "NOT_EVALUATED")
        self.assertEqual(result["diagnostic_code"], "INPUT_TOO_LARGE")
        self.assertEqual(runtime.status.state, "ready")

    def test_missing_key_is_operational_failure_and_not_cacheable(self):
        with mock.patch.dict(os.environ, {}, clear=True):
            runtime = RUNTIME_MODULE.TypeSafeJevRuntime({})
            self.assertIsNone(runtime.cache_identity())
            self.assertEqual(runtime.start().state, "unavailable")
            result = runtime.classify("标题", [{"paragraph_id": "p1", "text": "正文"}], {})
        self.assertEqual(result["status"], "NOT_EVALUATED")
        self.assertEqual(result["failure"], "MODEL_UNAVAILABLE")

    def test_cache_identity_never_contains_api_key(self):
        runtime = self.make_runtime()
        identity = runtime.cache_identity()
        self.assertIsNotNone(identity)
        self.assertNotIn("test-key", json.dumps(identity))
        self.assertEqual(identity["model_id"], "jev-1.13.0")


if __name__ == "__main__":
    unittest.main()

from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
SPEC = importlib.util.spec_from_file_location(
    "local_model_runtime", ROOT / "scripts" / "local_model_runtime.py"
)
MODULE = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class FakeProcess:
    def __init__(self, exit_code: int | None = None):
        self.exit_code = exit_code
        self.terminated = False
        self.killed = False
        self.wait_calls: list[float | None] = []

    def poll(self):
        return self.exit_code

    def terminate(self):
        self.terminated = True
        self.exit_code = 0

    def wait(self, timeout=None):
        self.wait_calls.append(timeout)
        return self.exit_code

    def kill(self):
        self.killed = True
        self.exit_code = 1


def ready_runtime(**overrides):
    config = {
        "mode": "shadow",
        "request_timeout_seconds": 1,
        "task_deadline_seconds": 2,
        "context_size": 4096,
        "output_reserve_tokens": 512,
        "safety_margin_tokens": 128,
    }
    config.update(overrides)
    runtime = MODULE.LocalModelRuntime(config)
    runtime._process = FakeProcess()
    runtime.base_url = "http://127.0.0.1:49123"
    runtime.server_version = "llama.cpp build 9442"
    runtime._set_status("ready")
    return runtime


def token_count_result(tokens=32):
    return MODULE.RuntimeResult(
        True, {"tokens": tokens}, input_tokens=tokens, elapsed_ms=2, tokenization_ms=2
    )


def semantic_output(status: str, quote: str, relation: str):
    supported = "no" if relation == "contradicts" else "yes"
    return {
        "status": status,
        "reason": "基于标题必要主张与正文证据",
        "core_claims": [{
            "id": "claim-1", "claim": "标题的一项必要主张", "supported": supported,
            "evidence_ids": ["evidence-1"],
        }],
        "evidence": [{
            "id": "evidence-1", "paragraph_id": "p1", "quote": quote,
            "claim_id": "claim-1", "relation": relation,
        }],
    }


class LocalModelRuntimeConfigTests(unittest.TestCase):
    def test_defaults_are_shadow_and_reranker_disabled(self):
        runtime = MODULE.LocalModelRuntime()
        self.assertEqual(runtime.mode, "shadow")
        self.assertFalse(runtime.config["enforce"])
        self.assertFalse(runtime.config["reranker"]["enabled"])
        self.assertEqual(runtime.config["request_timeout_seconds"], 20)

    def test_strategy_config_keeps_runtime_settings_nested(self):
        config_path = ROOT / "scripts" / "title_body_config.example.json"
        runtime = MODULE.LocalModelRuntime(config_path)
        self.assertEqual(runtime.config["mode"], "shadow")
        self.assertEqual(runtime.config["unknown_action"], "hold_changes")
        self.assertFalse(runtime.config["skip_mismatch_requested"])
        self.assertEqual(runtime.config["runtime"]["expected_server_version"], "9442")
        self.assertEqual(runtime.config["request_timeout_seconds"], 20)

    def test_invalid_utf8_config_fails_closed_without_raising(self):
        with tempfile.TemporaryDirectory() as directory:
            config = Path(directory) / "title_body_config.json"
            config.write_bytes(b"{\xff}")
            runtime = MODULE.LocalModelRuntime(config)
            status = runtime.start()

        self.assertEqual(status.state, "unavailable")
        self.assertEqual(status.diagnostic_code, "CONFIG_UNAVAILABLE")

    def test_service_lease_serializes_runtime_instances(self):
        first = MODULE.LocalModelRuntime({"startup_lock_timeout_seconds": 1})
        second = MODULE.LocalModelRuntime({"startup_lock_timeout_seconds": 0.05})
        first._acquire_service_lease()
        outcome = []

        def attempt_second():
            try:
                second._acquire_service_lease()
            except MODULE._RuntimeFailure as exc:
                outcome.append(exc.code)
            else:
                outcome.append("acquired")
                second._release_service_lease()

        thread = threading.Thread(target=attempt_second)
        thread.start()
        thread.join(timeout=1)
        first._release_service_lease()

        self.assertFalse(thread.is_alive())
        self.assertEqual(outcome, ["SERVICE_LOCK_TIMEOUT"])

    def test_close_without_owned_process_does_not_touch_external_server(self):
        runtime = MODULE.LocalModelRuntime()
        runtime._process = None
        runtime.close()
        self.assertEqual(runtime.status.state, "stopped")

    def test_start_reuses_resolver_checks_version_and_owns_only_spawned_process(self):
        runtime = MODULE.LocalModelRuntime({
            "expected_server_version": "build 9442",
            "startup_timeout_seconds": 2,
            "startup_lock_timeout_seconds": 1,
            "context_size": 4096,
        })
        child = FakeProcess()
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            server = root / "llama-server.exe"
            model = root / "QwenPaw-Flash-2B-Q4_K_M.gguf"
            server.touch()
            model.touch()
            with mock.patch.object(MODULE, "_resolve_runtime_paths", return_value=(server, model)) as resolve, \
                    mock.patch.object(MODULE.subprocess, "run", return_value=subprocess.CompletedProcess([], 0, "llama.cpp build 9442", "")), \
                    mock.patch.object(MODULE.subprocess, "Popen", return_value=child) as popen, \
                    mock.patch.object(runtime, "_probe", side_effect=[{"status": "ok"}, {"model_alias": runtime._service_alias}]):
                with mock.patch.object(MODULE, "_startup_lock") as startup_lock:
                    startup_lock.return_value.__enter__.return_value = None
                    status = runtime.start()

        self.assertEqual(status.state, "ready")
        resolve.assert_called_once()
        command = popen.call_args.args[0]
        self.assertIn("--parallel", command)
        self.assertEqual(command[command.index("--parallel") + 1], "1")
        self.assertEqual(command[command.index("--host") + 1], "127.0.0.1")
        runtime.close()
        self.assertTrue(child.terminated)

    def test_server_build_is_configurable_and_mismatch_never_spawns(self):
        runtime = MODULE.LocalModelRuntime({
            "expected_server_version": "release-x",
            "startup_timeout_seconds": 1,
            "startup_lock_timeout_seconds": 1,
        })
        with tempfile.TemporaryDirectory() as folder:
            server = Path(folder) / "llama-server.exe"
            model = Path(folder) / "QwenPaw-Flash-2B-Q4_K_M.gguf"
            server.touch()
            model.touch()
            with mock.patch.object(MODULE, "_resolve_runtime_paths", return_value=(server, model)), \
                    mock.patch.object(MODULE.subprocess, "run", return_value=subprocess.CompletedProcess([], 0, "build 9442", "")), \
                    mock.patch.object(MODULE.subprocess, "Popen") as popen, \
                    mock.patch.object(MODULE, "_startup_lock") as startup_lock:
                startup_lock.return_value.__enter__.return_value = None
                status = runtime.start()
        self.assertEqual(status.diagnostic_code, "SERVER_VERSION_MISMATCH")
        popen.assert_not_called()


class LocalModelRuntimeRequestTests(unittest.TestCase):
    def test_nested_claim_evidence_is_normalized_before_validation(self):
        runtime = ready_runtime()
        result = MODULE.RuntimeResult(True, {"choices": [{"message": {"content": json.dumps({
            "status": "MISMATCH", "reason": "业务用途明确不同",
            "core_claims": [{"claim": "项目审核", "supported": "no", "evidence": [
                {"paragraph_id": "p1", "quote": "库存出入库管理", "relation": "contradicts"},
            ]}],
        }, ensure_ascii=False)}}]}, generation_requests=1)

        parsed, valid = runtime._parse_classification_result(
            result, [{"id": "p1", "text": "本系统用于库存出入库管理"}], 0
        )

        self.assertTrue(valid)
        self.assertEqual(parsed["evidence"][0]["claim_id"], parsed["core_claims"][0]["id"])
        self.assertEqual(parsed["core_claims"][0]["evidence_ids"], [parsed["evidence"][0]["id"]])

    def test_health_probe_timeout_retries_without_terminating_owned_server(self):
        runtime = ready_runtime()
        child = FakeProcess()
        runtime._process = child

        with mock.patch.object(MODULE.urllib.request, "urlopen", side_effect=TimeoutError("still loading")):
            with self.assertRaises(MODULE._RuntimeFailure) as raised:
                runtime._probe("http://127.0.0.1:49123/health", time.monotonic() + 1)

        self.assertEqual(raised.exception.code, "MODEL_INTERFACE_ERROR")
        self.assertFalse(child.terminated)
        self.assertFalse(child.killed)

    def test_token_count_uses_rendered_model_template_then_tokenize(self):
        runtime = ready_runtime()
        calls = []

        def request(path, payload, deadline):
            calls.append((path, payload))
            if path == "/apply-template":
                return {"prompt": "<|user|>标题正文<|assistant|>"}
            return {"tokens": [1, 2, 3, 4]}

        with mock.patch.object(runtime, "_request", side_effect=request):
            result = runtime.count_tokens([{"role": "user", "content": "正文"}])
        self.assertTrue(result.ok)
        self.assertEqual(result.input_tokens, 4)
        self.assertEqual(calls[0][0], "/apply-template")
        self.assertEqual(calls[1][0], "/tokenize")
        self.assertEqual(calls[1][1]["content"], "<|user|>标题正文<|assistant|>")

    def test_chat_returns_usage_and_tokenization_timing(self):
        runtime = ready_runtime()

        def request(path, payload, deadline):
            if path == "/apply-template":
                return {"prompt": "rendered"}
            if path == "/tokenize":
                return {"tokens": [1, 2, 3]}
            self.assertEqual(path, "/v1/chat/completions")
            self.assertEqual(payload["model"], runtime._service_alias)
            self.assertFalse(payload["stream"])
            return {"choices": [{"message": {"content": "{}"}, "finish_reason": "stop"}],
                    "usage": {"completion_tokens": 7}}

        with mock.patch.object(runtime, "_request", side_effect=request):
            result = runtime.chat([{"role": "user", "content": "hello"}], response_format={"type": "json_object"})
        self.assertTrue(result.ok)
        self.assertEqual(result.input_tokens, 3)
        self.assertEqual(result.output_tokens, 7)
        self.assertIsNotNone(result.tokenization_ms)
        self.assertIsNotNone(result.request_ms)

    def test_input_budget_overflow_stops_before_generation(self):
        runtime = ready_runtime(context_size=10, output_reserve_tokens=2, safety_margin_tokens=1)
        paths = []

        def request(path, payload, deadline):
            paths.append(path)
            if path == "/apply-template":
                return {"prompt": "long"}
            if path == "/tokenize":
                return {"tokens": list(range(8))}
            self.fail("overflow must not reach chat completion")

        with mock.patch.object(runtime, "_request", side_effect=request):
            result = runtime.chat([{"role": "user", "content": "long"}])
        self.assertFalse(result.ok)
        self.assertEqual(result.diagnostic_code, "INPUT_TOO_LONG")
        self.assertEqual(paths, ["/apply-template", "/tokenize"])

    def test_classify_validates_exact_evidence_and_includes_diagnostics(self):
        runtime = ready_runtime()
        data = {"choices": [{"message": {"content": json.dumps({
            "status": "MISMATCH", "reason": "业务功能不同",
            "core_claims": [{"id": "claim-1", "claim": "项目审核", "supported": "no", "evidence_ids": ["evidence-1"]}],
            "evidence": [{"id": "evidence-1", "paragraph_id": "p1", "quote": "库存出入库", "claim_id": "claim-1", "relation": "contradicts"}],
        }, ensure_ascii=False)}, "finish_reason": "stop"}]}
        result = MODULE.RuntimeResult(True, data, input_tokens=44, output_tokens=18,
                                      elapsed_ms=321, tokenization_ms=5, request_ms=300,
                                      generation_requests=1)
        with mock.patch.object(runtime, "count_tokens", return_value=token_count_result()), \
                mock.patch.object(runtime, "chat", side_effect=[result, result]) as chat:
            classification = runtime.classify("项目审核", [{"id": "p1", "text": "本系统用于库存出入库管理"}])
        self.assertEqual(classification["status"], "MISMATCH")
        self.assertEqual(classification["diagnostics"]["input_tokens"], 88)
        self.assertGreaterEqual(classification["diagnostics"]["elapsed_ms"], 0)
        self.assertEqual(classification["diagnostics"]["request_ms"], 600)
        self.assertEqual(classification["diagnostics"]["generation_requests"], 2)
        self.assertEqual(chat.call_count, 2)

    def test_classify_invalid_quote_degrades_to_unknown(self):
        runtime = ready_runtime()
        data = {"choices": [{"message": {"content": json.dumps({
            "status": "MATCH", "reason": "支持",
            "core_claims": [{"id": "claim-1", "claim": "x", "supported": "yes", "evidence_ids": ["evidence-1"]}],
            "evidence": [{"id": "evidence-1", "paragraph_id": "p1", "quote": "没有的引文", "claim_id": "claim-1", "relation": "supports"}],
        })}, "finish_reason": "stop"}]}
        invalid = MODULE.RuntimeResult(True, data, generation_requests=1)
        with mock.patch.object(runtime, "count_tokens", return_value=token_count_result()), \
                mock.patch.object(runtime, "chat", side_effect=[invalid, invalid]) as chat:
            classification = runtime.classify("title", [{"id": "p1", "text": "正文段落"}])
        self.assertEqual(classification["status"], "UNKNOWN")
        self.assertEqual(classification["diagnostic_code"], "INVALID_MODEL_EVIDENCE")
        self.assertEqual(classification["diagnostics"]["generation_requests"], 2)
        self.assertEqual(chat.call_count, 2)

    def test_classify_unserializable_metadata_degrades_to_unknown(self):
        runtime = ready_runtime()
        with mock.patch.object(runtime, "chat") as chat:
            classification = runtime.classify("title", ["正文"], {"unsupported": object()})
        self.assertEqual(classification["status"], "UNKNOWN")
        self.assertEqual(classification["diagnostic_code"], "INPUT_INVALID")
        chat.assert_not_called()

    def test_selection_policy_version_is_stable_and_documents_budget_rules(self):
        self.assertEqual(
            MODULE.SELECTION_POLICY_VERSION,
            "title-body-selection-v3-rendered-prompt-fit-deadline-bounded",
        )

    def test_section_roles_use_titles_and_conditional_overview_coverage(self):
        runtime = ready_runtime()
        normalized, _content, _spans, structure_known = runtime._normalize_paragraphs([
            {"id": "p1", "text": "产品简介内容", "section_id": "s1", "section_title": "产品简介", "section_role": "overview"},
            {"id": "p2", "text": "功能模块内容", "section_id": "s2", "section_title": "主要功能模块", "section_role": "main_function"},
        ])

        self.assertTrue(structure_known)
        self.assertTrue(normalized[0]["is_overview"])
        self.assertTrue(normalized[1]["is_function"])
        self.assertTrue(normalized[1]["is_main_function"])

    def test_title_repeat_is_removed_before_body_coverage_and_chunk_selection(self):
        runtime = ready_runtime()
        normalized, content, spans, structure_known = runtime._normalize_paragraphs([
            {"id": "repeat", "text": "项目标题重复行", "section_id": "s1", "is_title_repeat": True},
            {"id": "body", "text": "正文包含实质功能结果", "section_id": "s1"},
        ])

        self.assertTrue(structure_known)
        self.assertEqual([item["id"] for item in normalized], ["body"])
        self.assertEqual(content, "正文包含实质功能结果")
        self.assertEqual([item["id"] for item in spans], ["body"])

    def test_classification_prompt_keeps_section_context_but_omits_paths_and_toc(self):
        runtime = ready_runtime()
        messages = runtime._classification_messages(
            "配对 TXT 标题",
            [
                {"paragraph_id": "heading-1", "text": "产品简介", "section_id": "s1", "section_title": "产品简介", "section_role": "overview", "is_heading": True},
                {"paragraph_id": "toc-1", "text": "目录中的主要功能", "section_id": "s2", "is_toc": True},
                {"paragraph_id": "body-1", "text": "正文功能结果", "section_id": "s1", "section_title": "产品简介", "section_role": "overview"},
            ],
            {"docx_path": r"D:\private\manual.docx", "txt_path": r"D:\private\title.txt"},
        )
        payload = json.loads(messages[1]["content"])

        self.assertEqual(payload["sections"], [{"id": "s1", "title": "产品简介", "role": "overview"}])
        self.assertEqual(payload["paragraphs"], [{"id": "body-1", "text": "正文功能结果", "section_id": "s1"}])
        self.assertNotIn("D:\\private", messages[1]["content"])
        self.assertNotIn("toc-1", messages[1]["content"])

    def test_candidate_chunks_use_384_tokens_and_48_sentence_overlap(self):
        runtime = ready_runtime()
        sentence = "x" * 47 + "!"
        content = sentence * 40
        normalized, _, source_spans, _ = runtime._normalize_paragraphs([
            {"id": "p1", "text": content, "section_id": "s1"}
        ])
        self.assertEqual(normalized[0]["id"], "p1")
        token_offsets = list(range(len(content) + 1))

        chunks, fallback_count = runtime._candidate_chunks(
            "unrelated title", content, source_spans, token_offsets
        )

        self.assertGreater(len(chunks), 2)
        self.assertEqual(fallback_count, 0)
        self.assertLessEqual(max(chunk["tokens"] for chunk in chunks), 384)
        self.assertEqual(chunks[0]["tokens"], 384)
        self.assertEqual(chunks[1]["start"], chunks[0]["end"] - 48)

    def test_long_paragraph_prompt_excerpts_keep_source_id_and_exact_text(self):
        runtime = ready_runtime()
        text = ("长段落内容。" * 150)
        messages = runtime._classification_messages(
            "标题", [{"id": "source-1", "content": text}], None
        )
        rows = json.loads(messages[1]["content"])["paragraphs"]

        self.assertGreater(len(rows), 1)
        self.assertTrue(all(row["id"] == "source-1" for row in rows))
        self.assertTrue(all(len(row["text"]) <= 700 for row in rows))
        self.assertEqual("".join(row["text"] for row in rows), text)

    def test_selection_uses_40_30_30_group_budgets_and_deduplicates_overlap(self):
        runtime = ready_runtime()
        candidates = []
        index = 0
        for kind, count in (("overview", 120), ("cross", 90), ("evidence", 90)):
            for offset in range(count):
                candidates.append({
                    "start": index * 500,
                    "end": index * 500 + 384,
                    "is_overview": kind == "overview",
                    "is_function": False,
                    "sections": {f"section-{index}"} if kind == "cross" else set(),
                    "score": 1.0 if kind == "evidence" else 0.0,
                })
                index += 1

        chosen, used = runtime._choose_candidate_intervals(candidates, 300 * 384)

        self.assertEqual(used, {
            "overview_function": 120 * 384,
            "cross_section": 90 * 384,
            "title_evidence": 90 * 384,
        })
        self.assertEqual(runtime._interval_token_count(chosen), 300 * 384)
        self.assertTrue(all(left[1] < right[0] for left, right in zip(chosen, chosen[1:])))

        overlapping = [
            {"start": 0, "end": 384, "is_overview": True, "is_function": False,
             "sections": set(), "score": 0.0},
            {"start": 336, "end": 720, "is_overview": True, "is_function": False,
             "sections": set(), "score": 0.0},
        ]
        deduped, _ = runtime._choose_candidate_intervals(overlapping, 720)
        self.assertEqual(deduped, [(0, 720)])
        self.assertEqual(runtime._interval_token_count(deduped), 720)

    def test_p5_coverage_failure_never_calls_generation(self):
        runtime = ready_runtime(context_size=200, output_reserve_tokens=10, safety_margin_tokens=10)
        selection = {
            "paragraphs": [{"id": "p1", "text": "正文片段"}],
            "diagnostics": {
                "coverage_sufficient": False,
                "body_token_coverage_ratio": 0.2,
                "section_coverage_ratio": 0.1,
            },
        }
        with mock.patch.object(runtime, "count_tokens", side_effect=[
                token_count_result(1000), token_count_result(10), token_count_result(100)]), \
                mock.patch.object(runtime, "_select_long_text", return_value=selection), \
                mock.patch.object(runtime, "chat") as chat:
            classification = runtime.classify("title", [{"id": "p1", "text": "长正文"}])

        self.assertEqual(classification["status"], "UNKNOWN")
        self.assertEqual(classification["diagnostic_code"], "P5_COVERAGE_INSUFFICIENT")
        self.assertEqual(classification["diagnostics"]["generation_requests"], 0)
        chat.assert_not_called()

    def test_long_prompt_selection_retries_with_a_tighter_rendered_budget(self):
        runtime = ready_runtime()
        selection = {
            "paragraphs": [{"id": "p1", "text": "正文片段"}],
            "review_paragraphs": [],
            "diagnostics": {
                "coverage_sufficient": True,
                "body_token_coverage_ratio": 0.7,
                "section_coverage_ratio": 0.5,
                "selection_budget_tokens": 1200,
            },
        }
        match = {"choices": [{"message": {"content": json.dumps(
            semantic_output("MATCH", "正文片段", "supports"), ensure_ascii=False
        )}, "finish_reason": "stop"}]}
        with mock.patch.object(runtime, "count_tokens", side_effect=[
                token_count_result(5000), token_count_result(100), token_count_result(4000), token_count_result(3200)]), \
                mock.patch.object(runtime, "_select_long_text", side_effect=[dict(selection), dict(selection)]) as select, \
                mock.patch.object(runtime, "chat", return_value=MODULE.RuntimeResult(True, match, generation_requests=1)):
            result = runtime.classify("标题", [{"id": "p1", "text": "长正文包含正文片段"}])

        self.assertEqual(result["status"], "UNKNOWN")
        self.assertEqual(result["diagnostic_code"], "CROPPED_MATCH_UNVERIFIED")
        self.assertEqual(select.call_count, 2)
        self.assertIsNone(select.call_args_list[0].kwargs["max_selection_tokens"])
        self.assertLess(select.call_args_list[1].kwargs["max_selection_tokens"], 1200)

    def test_cropped_mismatch_is_review_only_and_at_most_two_generations(self):
        runtime = ready_runtime()
        mismatch = {
            "choices": [{"message": {"content": json.dumps({
                "status": "MISMATCH", "reason": "业务不同",
                "core_claims": [{"id": "claim-1", "claim": "审核", "supported": "no", "evidence_ids": ["evidence-1"]}],
                "evidence": [{"id": "evidence-1", "paragraph_id": "p1", "quote": "库存出入库", "claim_id": "claim-1", "relation": "contradicts"}],
            }, ensure_ascii=False)}, "finish_reason": "stop"}]
        }
        response = MODULE.RuntimeResult(
            True, mismatch, input_tokens=100, output_tokens=20, generation_requests=1
        )
        selection = {
            "paragraphs": [{"id": "p1", "text": "库存出入库"}],
            "review_paragraphs": [{"id": "p1", "text": "库存出入库"}],
            "diagnostics": {
                "coverage_sufficient": True,
                "body_token_coverage_ratio": 0.9,
                "section_coverage_ratio": 1.0,
            },
        }
        with mock.patch.object(runtime, "count_tokens", side_effect=[
                token_count_result(5000), token_count_result(20), token_count_result(100),
                token_count_result(100)]), \
                mock.patch.object(runtime, "_select_long_text", return_value=selection), \
                mock.patch.object(runtime, "chat", side_effect=[response, response]) as chat:
            classification = runtime.classify(
                "审核项目", [{"id": "p1", "text": "库存出入库管理"}]
            )

        self.assertEqual(classification["status"], "UNKNOWN")
        self.assertEqual(classification["diagnostic_code"], "CROPPED_MISMATCH_REVIEW")
        self.assertEqual(classification["candidate_status"], "MISMATCH")
        self.assertEqual(classification["diagnostics"]["generation_requests"], 2)
        self.assertEqual(chat.call_count, 2)

    def test_request_deadline_terminates_only_owned_server(self):
        runtime = ready_runtime()
        child = FakeProcess()
        runtime._process = child
        release = threading.Event()

        class Response:
            def __enter__(self):
                return self

            def __exit__(self, *args):
                return None

            def read(self):
                release.wait(1)
                return b"{}"

        with mock.patch.object(MODULE.urllib.request, "urlopen", return_value=Response()):
            with self.assertRaisesRegex(MODULE._RuntimeFailure, "超时"):
                runtime._http_json("http://127.0.0.1:49123/slow", {}, time.monotonic() + 0.03)
        self.assertTrue(child.terminated)
        self.assertEqual(runtime.status.state, "timed_out")
        release.set()


if __name__ == "__main__":
    unittest.main()

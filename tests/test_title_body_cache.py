from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from title_body_cache import TitleBodyCache, semantic_input_hash


class TitleBodyCacheTests(unittest.TestCase):
    def setUp(self):
        self.paragraphs = [{
            "paragraph_id": "paragraph:2", "text": "系统支持传感器数据采集和分析。",
            "section_id": "paragraph:1", "style_id": "BodyText", "outline_level": None,
            "is_toc": False, "in_table": True, "in_textbox": False,
            "is_title_repeat": False, "structure_uncertain": False,
        }]

    def digest(self, paragraphs=None, **overrides):
        args = {"strategy_version": "s1", "model_id": "local:1", "prompt_version": "p1"}
        args.update(overrides)
        return semantic_input_hash("传感器数据分析系统", paragraphs or self.paragraphs, **args)

    def test_hash_covers_title_body_structure_and_dependencies_but_not_paragraph_ids(self):
        original = self.digest()
        moved = [dict(self.paragraphs[0], paragraph_id="paragraph:19")]
        self.assertEqual(original, self.digest(moved))
        self.assertNotEqual(original, self.digest(model_id="local:2"))
        self.assertNotEqual(original, self.digest(prompt_version="p2"))
        self.assertNotEqual(original, self.digest(policy_version="2"))
        changed = [dict(self.paragraphs[0], text="完全不同的正文")]
        self.assertNotEqual(original, self.digest(changed))
        changed_section = [dict(self.paragraphs[0], section_title="新的主要功能", section_role="main_function")]
        self.assertNotEqual(original, self.digest(changed_section))

    def test_atomic_json_cache_round_trip(self):
        with tempfile.TemporaryDirectory() as directory:
            cache = TitleBodyCache(Path(directory) / "cache.json")
            key = self.digest()
            result = {
                "status": "MATCH",
                "core_claims": [{"id": "claim-1", "claim": "数据采集", "supported": "yes", "evidence_ids": ["evidence-1"]}],
                "evidence": [{"id": "evidence-1", "paragraph_id": "paragraph:2", "quote": "传感器数据采集", "claim_id": "claim-1", "relation": "supports"}],
            }
            cache.put(key, result)
            self.assertEqual(cache.get(key, self.paragraphs)["status"], "MATCH")
            stored = json.loads(cache.path.read_text(encoding="utf-8"))
            self.assertEqual(stored["schema_version"], 1)
            self.assertEqual(list(Path(directory).glob("*.tmp")), [])

    def test_corrupt_cache_is_quarantined_and_treated_as_miss(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cache.json"
            path.write_text("{broken", encoding="utf-8")
            cache = TitleBodyCache(path)
            self.assertIsNone(cache.get("x", self.paragraphs))
            self.assertFalse(path.exists())
            self.assertEqual(len(list(Path(directory).glob("cache.json.corrupt.*"))), 1)

    def test_non_unique_cached_quote_is_not_reused(self):
        with tempfile.TemporaryDirectory() as directory:
            cache = TitleBodyCache(Path(directory) / "cache.json")
            key = self.digest()
            cache.put(key, {
                "status": "MATCH",
                "core_claims": [{"id": "claim-1", "claim": "数据", "supported": "yes", "evidence_ids": ["evidence-1"]}],
                "evidence": [{"id": "evidence-1", "paragraph_id": "paragraph:2", "quote": "数据", "claim_id": "claim-1", "relation": "supports"}],
            })
            changed = [
                dict(self.paragraphs[0], paragraph_id="paragraph:3", text="数据采集"),
                dict(self.paragraphs[0], paragraph_id="paragraph:4", text="数据分析"),
            ]
            self.assertIsNone(cache.get(key, changed))

    def test_unique_quote_is_rebound_when_the_original_paragraph_id_moves(self):
        with tempfile.TemporaryDirectory() as directory:
            cache = TitleBodyCache(Path(directory) / "cache.json")
            key = self.digest()
            cache.put(key, {
                "status": "MATCH",
                "core_claims": [{"id": "claim-1", "claim": "数据采集", "supported": "yes", "evidence_ids": ["evidence-1"]}],
                "evidence": [{"id": "evidence-1", "paragraph_id": "paragraph:2", "quote": "数据采集", "claim_id": "claim-1", "relation": "supports"}],
            })
            moved = [dict(self.paragraphs[0], paragraph_id="paragraph:19")]
            result = cache.get(key, moved)
            self.assertEqual(result["evidence"][0]["paragraph_id"], "paragraph:19")


if __name__ == "__main__":
    unittest.main()

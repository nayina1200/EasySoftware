from __future__ import annotations

import sys
import json
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from evaluate_title_body import evaluate_report


def _report(scores: dict[str, tuple[str, float]], threshold: float = 70) -> dict:
    return {"qualification": {"projects": [
        {
            "name": name,
            "semantic_status": status,
            "result": {"score_0_100": score, "decision_threshold": threshold},
        }
        for name, (status, score) in scores.items()
    ]}}


class TitleBodyEvaluationTests(unittest.TestCase):
    def test_anonymized_13_case_fixture_uses_one_score_cutoff(self) -> None:
        fixture_path = Path(__file__).resolve().parent / "fixtures" / "title_body" / "binary_acceptance_13_labels.json"
        labels = json.loads(fixture_path.read_text(encoding="utf-8"))
        scores = {
            name: (status, 85 if status == "MATCH" else 30)
            for name, status in labels["labels"].items()
        }

        result = evaluate_report(_report(scores), labels)

        self.assertEqual(result["expected_count"], 13)
        self.assertEqual(len(result["mismatches_expected"]), 2)
        self.assertTrue(result["score_threshold_evaluation"]["configured_threshold_passed"])

    def test_one_configured_threshold_passes_binary_fixture(self) -> None:
        labels = {"fixture_id": "sample", "labels": {"One": "MATCH", "Two": "MISMATCH"}}
        report = _report({"One": ("MATCH", 90), "Two": ("MISMATCH", 20)})

        result = evaluate_report(report, labels)

        self.assertTrue(result["acceptance_passed"])
        self.assertEqual(result["confusion_matrix"]["MISMATCH"]["MISMATCH"], 1)
        threshold_report = result["score_threshold_evaluation"]
        self.assertTrue(threshold_report["labels_separable_by_one_threshold"])
        self.assertEqual(threshold_report["suggested_threshold_fitted_to_this_fixture"], 55)
        self.assertTrue(threshold_report["configured_threshold_passed"])

    def test_no_common_threshold_means_fixture_is_not_accepted(self) -> None:
        labels = {"fixture_id": "sample", "labels": {"One": "MATCH", "Two": "MISMATCH"}}
        report = _report({"One": ("MISMATCH", 60), "Two": ("MATCH", 80)})

        result = evaluate_report(report, labels)

        self.assertFalse(result["acceptance_passed"])
        self.assertFalse(result["score_threshold_evaluation"]["labels_separable_by_one_threshold"])

    def test_missing_per_item_threshold_prevents_acceptance(self) -> None:
        labels = {"fixture_id": "sample", "labels": {"One": "MATCH", "Two": "MISMATCH"}}
        report = _report({"One": ("MATCH", 90), "Two": ("MISMATCH", 20)})
        report["qualification"]["projects"][1]["result"].pop("decision_threshold")

        result = evaluate_report(report, labels)

        threshold_report = result["score_threshold_evaluation"]
        self.assertTrue(threshold_report["score_coverage_complete"])
        self.assertFalse(threshold_report["threshold_coverage_complete"])
        self.assertFalse(result["acceptance_passed"])

    def test_not_evaluated_is_an_operational_outcome_not_a_content_label(self) -> None:
        labels = {"fixture_id": "sample", "labels": {"One": "MATCH", "Two": "MISMATCH"}}
        report = _report({"One": ("MATCH", 90), "Two": ("NOT_EVALUATED", 0)})

        result = evaluate_report(report, labels)

        self.assertFalse(result["acceptance_passed"])
        self.assertEqual(result["not_evaluated_count"], 1)
        self.assertEqual(result["confusion_matrix"]["MISMATCH"]["NOT_EVALUATED"], 1)

    def test_conflicting_report_status_fields_cannot_pass(self) -> None:
        labels = {"fixture_id": "sample", "labels": {"One": "MATCH"}}
        report = {"qualification": {"projects": [{
            "name": "One", "status": "NOT_EVALUATED", "semantic_status": "MATCH",
            "result": {"score_0_100": 90, "decision_threshold": 70},
        }]}}

        result = evaluate_report(report, labels)

        self.assertFalse(result["acceptance_passed"])
        self.assertEqual(result["status_field_conflict_count"], 1)
        self.assertEqual(result["rows"][0]["predicted"], "NOT_EVALUATED")

    def test_normalized_name_collisions_are_reported(self) -> None:
        labels = {"fixture_id": "sample", "labels": {"软件 A": "MATCH"}}
        report = _report({"软件 A": ("MATCH", 90), "软件A": ("MATCH", 91)})

        result = evaluate_report(report, labels)

        self.assertFalse(result["acceptance_passed"])
        self.assertEqual(result["duplicate_predictions"], ["软件A"])

    def test_labels_must_be_binary(self) -> None:
        labels = {"fixture_id": "sample", "labels": {"One": "UNKNOWN"}}
        with self.assertRaises(ValueError):
            evaluate_report(_report({"One": ("MATCH", 90)}), labels)


if __name__ == "__main__":
    unittest.main()

"""Evaluate Jev score decisions against a reviewed binary-label fixture."""
from __future__ import annotations

import argparse
import json
import math
import unicodedata
from pathlib import Path
from typing import Any


STATUSES = {"MATCH", "MISMATCH"}
PREDICTION_STATUSES = (*sorted(STATUSES), "NOT_EVALUATED")


def _name_key(value: str) -> str:
    return "".join(unicodedata.normalize("NFKC", value).split()).casefold()


def _number(value: Any) -> float | None:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    value = float(value)
    return value if math.isfinite(value) and 0 <= value <= 100 else None


def evaluate_report(report: dict[str, Any], labels_fixture: dict[str, Any]) -> dict[str, Any]:
    labels = labels_fixture.get("labels")
    if not isinstance(labels, dict) or not labels:
        raise ValueError("标签夹具必须包含非空 labels 对象")
    expected_by_key: dict[str, tuple[str, str]] = {}
    for name, status in labels.items():
        if not isinstance(name, str) or not isinstance(status, str) or status.upper() not in STATUSES:
            raise ValueError("验收标签只允许 MATCH 或 MISMATCH")
        key = _name_key(name)
        if key in expected_by_key:
            raise ValueError(f"标签名规范化后重复：{name}")
        expected_by_key[key] = (name, status.upper())

    qualification = report.get("qualification", {})
    projects = qualification.get("projects", []) if isinstance(qualification, dict) else []
    predicted_by_key: dict[str, dict[str, Any]] = {}
    duplicate_predictions: list[str] = []
    for item in projects if isinstance(projects, list) else []:
        if not isinstance(item, dict) or not isinstance(item.get("name"), str):
            continue
        name = item["name"]
        key = _name_key(name)
        semantic_raw = item.get("semantic_status")
        status_raw = item.get("status")
        semantic_status = str(semantic_raw or status_raw or "NOT_EVALUATED").upper()
        item_status = str(status_raw or semantic_raw or "NOT_EVALUATED").upper()
        semantic_status = semantic_status if semantic_status in STATUSES else "NOT_EVALUATED"
        item_status = item_status if item_status in STATUSES else "NOT_EVALUATED"
        status_conflict = semantic_raw is not None and status_raw is not None and semantic_status != item_status
        status = "NOT_EVALUATED" if status_conflict else semantic_status
        if key in predicted_by_key:
            duplicate_predictions.append(name)
        else:
            result = item.get("result") if isinstance(item.get("result"), dict) else {}
            predicted_by_key[key] = {
                "name": name,
                "status": status,
                "status_conflict": status_conflict,
                "score": _number(result.get("score_0_100")),
                "threshold": _number(result.get("decision_threshold")),
            }

    confusion = {expected: {predicted: 0 for predicted in PREDICTION_STATUSES} for expected in sorted(STATUSES)}
    rows: list[dict[str, Any]] = []
    scored_rows: list[dict[str, Any]] = []
    for key, (expected_name, expected_status) in expected_by_key.items():
        prediction = predicted_by_key.get(key)
        predicted_status = prediction["status"] if prediction else "NOT_EVALUATED"
        confusion[expected_status][predicted_status] += 1
        row = {
            "name": expected_name,
            "expected": expected_status,
            "predicted": predicted_status,
            "correct": expected_status == predicted_status,
            "present": prediction is not None,
            "status_conflict": bool(prediction and prediction["status_conflict"]),
            "score_0_100": prediction["score"] if prediction else None,
        }
        rows.append(row)
        if prediction is not None and prediction["score"] is not None:
            scored_rows.append({**row, "decision_threshold": prediction["threshold"]})

    missing = [name for key, (name, _) in expected_by_key.items() if key not in predicted_by_key]
    extra = [name for key, prediction in predicted_by_key.items() if key not in expected_by_key for name in [prediction["name"]]]
    mismatch_keys_expected = {key for key, (_name, status) in expected_by_key.items() if status == "MISMATCH"}
    mismatch_keys_predicted = {key for key, prediction in predicted_by_key.items() if prediction["status"] == "MISMATCH"}
    mismatches_expected = sorted(expected_by_key[key][0] for key in mismatch_keys_expected)
    mismatches_predicted = sorted(predicted_by_key[key]["name"] for key in mismatch_keys_predicted)
    correct = sum(row["expected"] == row["predicted"] for row in rows)
    all_present = not missing and not extra and not duplicate_predictions

    score_complete = len(scored_rows) == len(expected_by_key)
    threshold_complete = score_complete and all(row["decision_threshold"] is not None for row in scored_rows)
    used_thresholds = {row["decision_threshold"] for row in scored_rows if row["decision_threshold"] is not None}
    one_threshold = threshold_complete and len(used_thresholds) == 1
    used_threshold = next(iter(used_thresholds)) if one_threshold else None
    threshold_rows = []
    if used_threshold is not None:
        for row in scored_rows:
            score_prediction = "MATCH" if row["score_0_100"] >= used_threshold else "MISMATCH"
            threshold_rows.append({
                "name": row["name"], "expected": row["expected"],
                "predicted_from_score": score_prediction,
                "correct": row["expected"] == score_prediction,
            })
    threshold_correct = sum(row["correct"] for row in threshold_rows)
    matches = [row["score_0_100"] for row in scored_rows if row["expected"] == "MATCH"]
    mismatches = [row["score_0_100"] for row in scored_rows if row["expected"] == "MISMATCH"]
    max_mismatch = max(mismatches) if mismatches else None
    min_match = min(matches) if matches else None
    separable = max_mismatch is not None and min_match is not None and max_mismatch < min_match
    suggested_threshold = (max_mismatch + min_match) / 2 if separable else None
    status_conflict_count = sum(bool(prediction["status_conflict"]) for prediction in predicted_by_key.values())
    status_acceptance = (
        all_present and status_conflict_count == 0 and correct == len(expected_by_key)
        and mismatch_keys_expected == mismatch_keys_predicted
    )
    score_acceptance = (
        score_complete and one_threshold and len(threshold_rows) == len(expected_by_key)
        and threshold_correct == len(expected_by_key)
    )
    return {
        "fixture_id": labels_fixture.get("fixture_id"),
        "expected_count": len(expected_by_key),
        "predicted_count": len(predicted_by_key),
        "correct_count": correct,
        "accuracy": correct / len(expected_by_key),
        "not_evaluated_count": sum(row["predicted"] == "NOT_EVALUATED" for row in rows),
        "status_field_conflict_count": status_conflict_count,
        "confusion_matrix": confusion,
        "exact_mismatch_set": mismatch_keys_expected == mismatch_keys_predicted,
        "mismatches_expected": mismatches_expected,
        "mismatches_predicted": mismatches_predicted,
        "missing_projects": missing,
        "extra_projects": extra,
        "duplicate_predictions": duplicate_predictions,
        "status_acceptance_passed": status_acceptance,
        "score_threshold_evaluation": {
            "score_coverage_complete": score_complete,
            "threshold_coverage_complete": threshold_complete,
            "one_common_threshold": one_threshold,
            "configured_threshold": used_threshold,
            "correct_from_score_at_configured_threshold": threshold_correct,
            "configured_threshold_passed": score_acceptance,
            "max_expected_mismatch_score": max_mismatch,
            "min_expected_match_score": min_match,
            "labels_separable_by_one_threshold": separable,
            "suggested_threshold_fitted_to_this_fixture": suggested_threshold,
            "decision_rule": "MATCH iff score_0_100 >= threshold; otherwise MISMATCH",
            "rows": threshold_rows,
        },
        "acceptance_passed": status_acceptance and score_acceptance,
        "rows": rows,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("report", type=Path, help="统一工作流 JSON 预览报告")
    parser.add_argument("--labels", type=Path, required=True, help="人工复核标签 JSON")
    parser.add_argument("--output", type=Path, help="可选验收结果 JSON 输出路径")
    args = parser.parse_args()
    report = json.loads(args.report.read_text(encoding="utf-8"))
    labels = json.loads(args.labels.read_text(encoding="utf-8"))
    result = evaluate_report(report, labels)
    serialized = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(serialized + "\n", encoding="utf-8")
    print(serialized)
    return 0 if result["acceptance_passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())

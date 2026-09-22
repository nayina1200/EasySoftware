from __future__ import annotations

import argparse
import json
import os
from dataclasses import dataclass
from datetime import datetime
from html import escape
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen


VERSION = "1.0.0"
DEFAULT_BASE_URL = "http://36.212.60.8:10000"


class WorkflowError(RuntimeError):
    pass


@dataclass
class ApiResult:
    data: Any
    status_code: int = 200


class RefurbishApi:
    def __init__(self, base_url: str, token: str, timeout: float = 30.0):
        self.base_url = base_url.rstrip("/") + "/prod-api"
        self.token = token.strip()
        if self.token.lower().startswith("bearer "):
            self.token = self.token[7:].strip()
        self.timeout = timeout

    def request(self, path: str, method: str = "GET", params: dict[str, Any] | None = None, data: Any = None) -> ApiResult:
        query = urlencode({key: value for key, value in (params or {}).items() if value not in (None, "", [])})
        url = self.base_url + path + (("?" + query) if query else "")
        body = None
        headers = {"Accept": "application/json", "Authorization": "Bearer " + self.token}
        if data is not None:
            body = json.dumps(data, ensure_ascii=False).encode("utf-8")
            headers["Content-Type"] = "application/json;charset=utf-8"
        request = Request(url, data=body, method=method, headers=headers)
        try:
            with urlopen(request, timeout=self.timeout) as response:
                payload = response.read()
                status_code = response.status
        except HTTPError as exc:
            payload = exc.read()
            status_code = exc.code
        except (URLError, TimeoutError, OSError) as exc:
            raise WorkflowError(f"接口连接失败：{exc}") from exc
        try:
            result = json.loads(payload.decode("utf-8-sig"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise WorkflowError(f"接口返回不是可读取的 JSON：HTTP {status_code}") from exc
        code = result.get("code", 200) if isinstance(result, dict) else 200
        if code != 200:
            message = result.get("msg", "接口返回失败") if isinstance(result, dict) else "接口返回失败"
            if code == 401:
                raise WorkflowError("Edge 登录令牌未接入或已过期，请重新提供登录令牌")
            raise WorkflowError(f"接口失败（{code}）：{message}")
        return ApiResult(result.get("data", result), status_code)

    def groups(self, operator: str, begin: str = "", end: str = "", status: str = "") -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        page = 1
        while True:
            data = self.request("/doc/refurbish/batch/groups", params={
                "pageNum": page,
                "pageSize": 100,
                "createBy": operator,
                "beginTime": begin or None,
                "endTime": end + " 23:59:59" if end else None,
                "status": status or None,
            }).data or {}
            page_rows = data.get("rows", []) if isinstance(data, dict) else []
            rows.extend(page_rows)
            total = int(data.get("total", len(rows))) if isinstance(data, dict) else len(rows)
            if not page_rows or len(rows) >= total:
                return rows
            page += 1

    def orders(self, batch_no: str) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        page = 1
        while True:
            data = self.request("/doc/refurbish/list", params={
                "pageNum": page,
                "pageSize": 500,
                "batchNo": batch_no,
            }).data or {}
            page_rows = data.get("rows", []) if isinstance(data, dict) else []
            rows.extend(page_rows)
            total = int(data.get("total", len(rows))) if isinstance(data, dict) else len(rows)
            if not page_rows or len(rows) >= total:
                return rows
            page += 1

    def resume_code_generation(self, order_ids: list[Any]) -> Any:
        return self.request("/doc/refurbish/batch/resumeCodeGen", method="POST", data=order_ids).data


def _date_value(value: str, end: bool = False) -> str:
    if not value:
        return ""
    try:
        return datetime.strptime(value, "%Y-%m-%d").strftime("%Y-%m-%d")
    except ValueError as exc:
        label = "结束日期" if end else "开始日期"
        raise WorkflowError(f"{label}格式应为 YYYY-MM-DD：{value}") from exc


def _group_key(row: dict[str, Any]) -> str:
    return str(row.get("batchNo") or row.get("batchNO") or row.get("batch_no") or "")


def _order_label(row: dict[str, Any]) -> str:
    return str(row.get("orderName") or row.get("orderNo") or row.get("orderId") or "")


def _status(row: dict[str, Any]) -> str:
    return str(row.get("status") or "").strip().lower()


def _step(row: dict[str, Any]) -> str:
    return str(row.get("currentStep") or "").strip().lower()


def _append_batch(report: dict[str, Any], row: dict[str, Any], workflow: str, orders: list[dict[str, Any]], action: str, result: Any = None) -> None:
    item = {
        "batch_no": _group_key(row),
        "batch_name": str(row.get("orderName") or row.get("batchNo") or ""),
        "workflow": workflow,
        "order_count": len(orders),
        "orders": [{"order_id": item.get("orderId"), "name": _order_label(item)} for item in orders],
        "action": action,
    }
    if result is not None:
        item["result"] = result
    report["batches"].append(item)


def run_workflow(api: RefurbishApi, operator: str, begin: str = "", end: str = "", dry_run: bool = False) -> dict[str, Any]:
    begin = _date_value(begin)
    end = _date_value(end, end=True)
    if begin and end and begin > end:
        raise WorkflowError("开始日期不能晚于结束日期")
    report: dict[str, Any] = {
        "version": VERSION,
        "status": "OK",
        "operator": operator,
        "begin": begin,
        "end": end,
        "dry_run": dry_run,
        "batches": [],
        "processed_batches": 0,
        "paused_orders": 0,
        "resume_candidates": 0,
        "submitted_orders": 0,
        "resumed_orders": 0,
        "skipped_orders": 0,
    }
    seen: set[str] = set()
    for group_status, workflow in (("in_progress", "进行中批次：暂停订单待进入流程"), ("partial", "部分完成批次：恢复代码生成")):
        for group in api.groups(operator, begin, end, group_status):
            batch_no = _group_key(group)
            if not batch_no or batch_no in seen:
                continue
            seen.add(batch_no)
            orders = api.orders(batch_no)
            if group_status == "in_progress":
                paused = [item for item in orders if _status(item) == "paused"]
                report["paused_orders"] += len(paused)
                if paused:
                    _append_batch(report, group, workflow, paused, "WAIT_MANUAL_FLOW")
                continue
            failed_candidates = [
                item for item in orders
                if _status(item) == "failed"
                and _step(item) == "code_gen"
                and item.get("codeOrderNo")
                and item.get("orderId")
            ]
            report["resume_candidates"] += len(failed_candidates)
            if not failed_candidates:
                continue
            if dry_run:
                _append_batch(report, group, workflow, failed_candidates, "PLAN_ONLY")
                continue
            result = api.resume_code_generation([item.get("orderId") for item in failed_candidates])
            _append_batch(report, group, workflow, failed_candidates, "RESUME_CODE_GEN", result)
            report["processed_batches"] += 1
            report["submitted_orders"] += len(failed_candidates)
            if isinstance(result, dict):
                report["resumed_orders"] += int(result.get("resumed", 0) or 0)
                report["skipped_orders"] += int(result.get("skipped", 0) or 0)
    return report


def emit(result: dict[str, Any], code: int = 0) -> int:
    print(json.dumps(result, ensure_ascii=False, separators=(",", ":")))
    return code


def _report_action(action: str) -> tuple[str, str, str]:
    if action == "RESUME_CODE_GEN":
        return "已提交代码恢复", "success", "已向系统提交代码生成恢复请求，等待系统继续执行。"
    if action == "PLAN_ONLY":
        return "预览发现可恢复订单", "info", "本次仅预览，没有向系统提交任何恢复请求。"
    return "需要人工处理", "warning", "发现已暂停订单，需要人工进入“流程”继续处理。"


def _html_report(result: dict[str, Any]) -> str:
    status = str(result.get("status") or "FAILED")
    is_ok = status == "OK"
    title = "翻新失败业务处理报告"
    mode = "预览结果" if result.get("dry_run") else "正式处理结果"
    status_text = "处理完成" if is_ok else "处理失败"
    range_text = "全部提交时间" if not (result.get("begin") or result.get("end")) else f"{result.get('begin') or '开始不限'} 至 {result.get('end') or '结束不限'}"
    metrics = [
        ("相关批次", len(result.get("batches") or []), "neutral"),
        ("暂停待人工", result.get("paused_orders", 0), "warning"),
        ("代码恢复候选", result.get("resume_candidates", 0), "info"),
        ("已提交恢复", result.get("submitted_orders", 0), "success"),
    ]
    metric_html = "".join(
        f'<div class="metric {style}"><span>{escape(str(label))}</span><strong>{escape(str(value))}</strong></div>'
        for label, value, style in metrics
    )
    batches: list[str] = []
    for item in result.get("batches") or []:
        label, style, explanation = _report_action(str(item.get("action") or ""))
        order_items = "".join(
            f"<li>{escape(str(order.get('name') or order.get('order_id') or '未命名订单'))}</li>"
            for order in item.get("orders") or []
        ) or "<li>没有符合本次处理条件的订单。</li>"
        response = item.get("result")
        response_html = ""
        if response is not None:
            response_html = f"<details><summary>系统返回信息</summary><pre>{escape(json.dumps(response, ensure_ascii=False, indent=2))}</pre></details>"
        batches.append(
            "<section class=\"batch\">"
            f"<div class=\"batch-head\"><h2>{escape(str(item.get('batch_name') or item.get('batch_no') or '未命名批次'))}</h2>"
            f"<span class=\"badge {style}\">{label}</span></div>"
            f"<p class=\"batch-no\">批次号：{escape(str(item.get('batch_no') or '未提供'))}</p>"
            f"<p>{explanation}</p><h3>涉及订单（{escape(str(item.get('order_count') or 0))} 个）</h3><ul>{order_items}</ul>{response_html}</section>"
        )
    batch_html = "".join(batches) or "<section class=\"empty\">本次筛选范围内没有需要处理的批次。</section>"
    error_html = "" if is_ok else f"<section class=\"error\"><h2>失败原因</h2><p>{escape(str(result.get('error') or '未提供失败原因'))}</p></section>"
    return f"""<!doctype html>
<html lang=\"zh-CN\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width, initial-scale=1\"><title>{title}</title>
<style>
body{{margin:0;background:#f4f7fb;color:#1d2939;font:15px/1.6 \"Microsoft YaHei UI\",Arial,sans-serif}}main{{max-width:1040px;margin:0 auto;padding:28px}}header{{background:#fff;padding:24px 28px;border-left:6px solid #2563eb;box-shadow:0 1px 3px #d8e0eb}}h1{{margin:0;font-size:26px}}h2{{margin:0;font-size:18px}}h3{{font-size:15px;margin:18px 0 8px}}.sub{{color:#667085;margin:8px 0 0}}.status{{display:inline-block;margin-top:14px;padding:5px 12px;border-radius:4px;background:{'#dcfae6' if is_ok else '#fee4e2'};color:{'#087443' if is_ok else '#b42318'};font-weight:700}}.metrics{{display:grid;grid-template-columns:repeat(4,minmax(130px,1fr));gap:12px;margin:18px 0}}.metric{{background:#fff;padding:16px;border-top:4px solid #98a2b3}}.metric span{{display:block;color:#667085;font-size:13px}}.metric strong{{display:block;font-size:28px;margin-top:4px}}.metric.warning{{border-color:#f79009}}.metric.info{{border-color:#2e90fa}}.metric.success{{border-color:#12b76a}}.batch,.error,.empty{{background:#fff;margin:14px 0;padding:20px 24px;border:1px solid #e4e7ec}}.batch-head{{display:flex;gap:12px;align-items:center;justify-content:space-between}}.badge{{padding:3px 10px;border-radius:4px;font-size:13px;font-weight:700;white-space:nowrap}}.badge.success{{background:#dcfae6;color:#087443}}.badge.info{{background:#d1e9ff;color:#175cd3}}.badge.warning{{background:#fef0c7;color:#b54708}}.batch-no{{color:#667085;font-size:13px}}ul{{margin:0;padding-left:22px}}details{{margin-top:14px}}pre{{background:#f8fafc;padding:12px;overflow:auto;white-space:pre-wrap}}.error{{border-left:5px solid #d92d20}}@media(max-width:700px){{main{{padding:14px}}.metrics{{grid-template-columns:repeat(2,1fr)}}.batch-head{{align-items:flex-start;flex-direction:column}}}}
</style></head><body><main><header><h1>{title}</h1><p class=\"sub\">{mode} · 操作人：{escape(str(result.get('operator') or 'admin'))} · 提交时间：{escape(range_text)}</p><span class=\"status\">{status_text}</span></header><div class=\"metrics\">{metric_html}</div>{error_html}{batch_html}</main></body></html>"""


def write_report(path_value: str, result: dict[str, Any]) -> None:
    requested_path = Path(path_value).expanduser().resolve()
    html_path = requested_path if requested_path.suffix.lower() == ".html" else requested_path.with_suffix(".html")
    json_path = html_path.with_suffix(".json")
    html_path.parent.mkdir(parents=True, exist_ok=True)
    result["report_path"] = str(html_path)
    result["report_data_path"] = str(json_path)
    json_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    html_path.write_text(_html_report(result), encoding="utf-8")


def emit_with_report(path_value: str, result: dict[str, Any], code: int) -> int:
    if path_value.strip():
        try:
            write_report(path_value, result)
        except OSError as exc:
            result["report_error"] = f"报告保存失败：{exc}"
    return emit(result, code)


def main() -> int:
    parser = argparse.ArgumentParser(description="翻新批次失败业务顺序恢复助手")
    parser.add_argument("--base-url", default=DEFAULT_BASE_URL)
    parser.add_argument("--operator", default="admin")
    parser.add_argument("--begin", default="")
    parser.add_argument("--end", default="")
    parser.add_argument("--token", default="")
    parser.add_argument("--token-env", default="EASYSOFTWARE_REFURBISH_TOKEN")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--report", default="")
    args = parser.parse_args()
    token = args.token.strip() or os.environ.get(args.token_env, "").strip()
    if not token:
        return emit_with_report(args.report, {"version": VERSION, "status": "AUTH_REQUIRED", "error": "未接入 Edge 或 Chrome 登录令牌"}, 2)
    try:
        result = run_workflow(RefurbishApi(args.base_url, token), args.operator.strip() or "admin", args.begin.strip(), args.end.strip(), args.dry_run)
        return emit_with_report(args.report, result, 0)
    except WorkflowError as exc:
        return emit_with_report(args.report, {"version": VERSION, "status": "FAILED", "error": str(exc)}, 2)
    except Exception as exc:
        return emit_with_report(args.report, {"version": VERSION, "status": "FAILED", "error": f"程序异常：{exc}"}, 2)


if __name__ == "__main__":
    raise SystemExit(main())

# 当前架构导航

状态：Confirmed

| 职责 | 当前权威入口 |
|---|---|
| Windows 双击启动 | `启动EasySoftware.cmd` → `启动EasySoftware.ps1` |
| GUI、确认弹窗、进度展示 | `启动EasySoftware.ps1` |
| CLI 路由和公共进度 | `scripts/unified_workflow.py` |
| 说明 + 代码 + TXT 统一编排 | `scripts/combined_document_workflow.py` |
| 说明 OCR、相关性、正文年份、截图年份 | `scripts/manual_content_stage.py` |
| RapidOCR 批处理桥接 | `scripts/rapid_ocr_bridge.py` |
| 旧日期引擎桥接 | `scripts/date_year_rewriter_bridge.py` |
| 封面修复、Word 保存与说明 PDF | `scripts/word_manual_pipeline_stable.ps1` |
| 说明预审与动作计划 | `scripts/process_batch.py` |
| 代码中文注释清理 | `scripts/code_comment_cleaner.py` |
| 说明阶段测试 | `tests/test_manual_content_stage.py`、`tests/test_manual_cover_rules.py` |
| 统一入口测试 | `tests/test_combined_document_workflow.py`、`tests/test_unified_document_entry.py` |

行为边界以 `contracts/document-stage.schema.json` 和 `contracts/code-phase.md` 为准。

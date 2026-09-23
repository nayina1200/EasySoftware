# EasySoftware 任务交接（2026-09-23）

项目目录：`D:\360MoveData\Users\ban\Desktop\work_about\EasySoftware_绿色便携版`。请使用中文交流。**权威交接入口**是 [.project-memory/START-HERE.md](.project-memory/START-HERE.md)；具体事实与验证命令以该入口链接的契约、测试和证据为准。本文件只给下一对话一条最短的接续路线。

## 可直接发给下一对话

> 请接续 `D:\360MoveData\Users\ban\Desktop\work_about\EasySoftware_绿色便携版`。先读 `HANDOFF.md`、`.project-memory/START-HERE.md`、`.project-memory/evidence/corrected-real-batch-20260923.md` 和 `.project-memory/gaps.md`。2026-09-23 修复批次已通过验收；不要把旧批次 `北京知识产权10件_已处理_20260922_213656` 当作成果。请以我在新对话提出的具体目标继续，保留既有输入、输出和未提交改动，做完受影响验证并更新交接证据。TXT 三期规则尚未确定，最终 RAR 尚未打包。

## 当前可用结果（Confirmed）

- 原始输入：`D:\360MoveData\Users\ban\Desktop\work_about\order\北京知识产权10件.rar`。
- **本轮输出**：`D:\360MoveData\Users\ban\Desktop\work_about\order\北京知识产权10件_已处理_20260923_114920`。旧的 `..._20260922_213656` 含年份误改，不可交付。
- [HTML 报告](.omc/full-batch-corrected-20260923/combined_document_result.html)及[机器可读结果](.omc/full-batch-corrected-20260923/combined_document_result.json)：`DOCUMENT_PROCESSING_OK`、失败 0。报告展示 24 张年份图片的原稿/处理后全图及局部对比；卡片保留“待核验”提示，便于逐张检查。
- 10/10 份说明书封面 Word QA 为 `OK`。首次全跑发现 7 份仍有独立旧封面页，修复分页/分节边界删除规则后，对现有 DOCX 重跑这 7 份并重导 PDF；[重跑结果](.omc/full-batch-corrected-20260923/cover_retry_seven_result.json)均为 `OK`。[首次失败记录](.omc/full-batch-corrected-20260923/combined_cover_result_initial.json)仅用于追溯。
- 图片日期**仅把早于 2025 的有效日期年份改为 2025**，月、日、时间保留；2025 和未来年份不动。医院说明 11 张（14 个 `2022→2025` 字段）、蓝宝石切片机说明 13 张（13 个 `2024→2025` 字段）。原本误改的编号、图表刻度、毫米波 2026 日期及医院 2033 日期保持原样。[详细复验证据](.project-memory/evidence/corrected-real-batch-20260923.md)。
- 用户标出的 5 份标题与正文疑似冲突的说明全部在报告中提示**人工复核**；这仍是可解释的主题冲突提示，并非完整语义判定。图片相关性由 OCR 预检和本地多模态模型抽查辅助判断；视觉模型结果只供复核，不参与自动拦截。本地路由为 `http://127.0.0.1:3458/v1/chat/completions`，模型名必须是 `sensenova-6.8-flash-lite`，Key 可为 `proxy`。
- GUI 的“处理说明和代码”入口可在处理前勾选说明书、代码、TXT 占位；CLI 使用 `--document-stages manual,code,txt_placeholder`。默认全选，预览与执行使用同一选择；未选阶段为 `SKIPPED` 且不扫描/处理，进度只计所选阶段。至少选说明书或代码；TXT 单选无可执行动作。

## 验证与边界

- 全套回归：在项目目录运行 `.\python\python.exe -m unittest discover -s .\tests -q`，132 项通过。最终输出 20 个 DOCX 均通过 ZIP 校验，20 个 PDF 均可解析（共 857 页）；代码阶段 10/10 成功，清理 1100 处中文注释；10 个 TXT 与原稿哈希相同，无临时文件。独立验证见[修复批次证据](.project-memory/evidence/corrected-real-batch-20260923.md)。
- OCR 无有效文字为 `unknown`，不会自动拦截；只有既定高置信不相关比例达到阈值才拦截或熔断。不要因本轮模型接入而改变该契约。[用户确认边界](.project-memory/evidence/session-decisions.md)。
- TXT 三期只登记路径，不读、不改；其处理规则和样例仍是 `Unknown`。[未决事项](.project-memory/gaps.md)还包括大文档 Word 性能、日期扫描细粒度进度、正文相关性覆盖范围。
- 最终 RAR **未重新打包**。仓库仍有大量未提交改动（包含本轮实现、测试和项目记忆），下一对话不要重置或覆盖它们。
- 项目记忆结构检查 `memoryctl.py check .` 通过；`fresh .` 提示 457 个快照条目已变化（多数为本轮 `.omc` 验证产物），所以 `manifest.json` 不是本轮最终文件哈希清单。需要提交或做快照时，先决定哪些临时证据应排除，再更新快照。

## 代码入口

[GUI](启动EasySoftware.ps1) → [CLI/进度](scripts/unified_workflow.py) → [阶段编排与报告](scripts/combined_document_workflow.py)。说明 OCR、正文相关性、视觉抽查与年份策略在 [manual_content_stage.py](scripts/manual_content_stage.py)，截图日期桥接在 [date_year_rewriter_bridge.py](scripts/date_year_rewriter_bridge.py)，封面清理与说明 PDF 在 [word_manual_pipeline_stable.ps1](scripts/word_manual_pipeline_stable.ps1)。接口边界见 [.project-memory/CURRENT.md](.project-memory/CURRENT.md) 和 [.project-memory/contracts/document-stage.schema.json](.project-memory/contracts/document-stage.schema.json)。

# 下一期任务交接入口

状态：可接续（说明 + 代码统一工作流已实现并通过 10 件全量材料验证）

2026-09-23 真机复核发现并修复了封面残片，同时新增标题与正文的人工复核提示；结果见 [evidence/real-output-cover-and-relevance-20260923.md](evidence/real-output-cover-and-relevance-20260923.md)。旧全量批测中的“封面已清理干净”结论以这次真机复核为准。

同日补充了图片变化前后对比，并用本地视觉模型抽查 OCR 疑难图；逐图核验发现旧批次 10 张变化图中有 9 张日期误改，旧输出不可直接交付。详见 [evidence/year-gallery-and-local-vision-20260923.md](evidence/year-gallery-and-local-vision-20260923.md)。视觉结果只供复核，未改变 OCR 拦截契约。

修复后已从原稿重跑真机 10 件：新输出 `order/北京知识产权10件_已处理_20260923_114920`，10/10 封面 QA=OK、图片有效旧日期修改 24 张、未来日期保持原样、统一报告失败 0。GUI/CLI 新增执行前阶段勾选，未选阶段跳过。验收细节见 [evidence/corrected-real-batch-20260923.md](evidence/corrected-real-batch-20260923.md)。

## 先读

1. [evidence/session-decisions.md](evidence/session-decisions.md)：用户确认的处理边界。
2. [contracts/document-stage.schema.json](contracts/document-stage.schema.json)：说明阶段契约。
3. [contracts/code-phase.md](contracts/code-phase.md)：冻结的代码阶段契约。
4. [CURRENT.md](CURRENT.md)：入口与模块地图。
5. [gaps.md](gaps.md)：仍未完成的 TXT 三期和性能项。

## 当前结论

- GUI 的“处理说明和代码”与 CLI `--process-documents` / `--apply-document-processing` 已形成统一入口；执行前可勾选说明书、代码、TXT 占位，未选阶段跳过。
- 所选阶段的顺序为：说明内容预检与年份处理 → 串行 Word 封面/PDF → 代码注释清理/PDF → TXT 占位。Word COM 不并发。
- OCR 等距抽取约四分之一图片；无有效文字只记 `unknown`，不拦截。只有高置信不相关超过项目阈值才拦截；批量异常比例过高时熔断日期写入。
- 正文年份采用 `UPDATE_TO_2025 / KEEP / REVIEW`；只有高置信且早于 2025 的候选会写入。截图中只有早于 2025 的有效日期才改年份为 2025，月日不变；2025 和未来日期保留。
- 封面保留原页面尺寸，不补写缺失副标题；能清除独立重复封面、残片页和封面后的空白页。
- TXT 三期尚未实现：当前只登记路径，不读取、不修改。
- 最终 RAR 重新打包仍未执行，属于所有阶段完成后的最后动作。

## 历史全量批测（旧年份结果已被新复验取代）

测试材料：`北京知识产权10件.rar`，SHA-256 `00818AFDC39E2863D757BAC94F9718644D3387AA790EE843EA9E2B5972034E38`。

- 10 份说明、414 张图、抽检 112 张：相关 20、不相关 1、未知 91；0 项目拦截、未熔断。
- 3 份说明共改 10 张截图年份；正文 7 个年份候选均未达到自动修改条件。
- 10 份代码文档清理 1100 条中文注释，代码阶段 0 失败。
- 10 个 TXT 与输入哈希一致。
- 最终 20 DOCX、20 PDF（857 页）全部可读，无临时文件；83 项测试通过。

详见 [evidence/document-stage-full-batch.md](evidence/document-stage-full-batch.md)。

## 常用入口

- 双击：[启动EasySoftware.cmd](../启动EasySoftware.cmd)
- GUI：[启动EasySoftware.ps1](../启动EasySoftware.ps1)
- CLI：[scripts/unified_workflow.py](../scripts/unified_workflow.py)
- 统一编排：[scripts/combined_document_workflow.py](../scripts/combined_document_workflow.py)
- 回归：`.\python\python.exe -m unittest discover -s .\tests -q`

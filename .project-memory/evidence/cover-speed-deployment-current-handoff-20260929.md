# 封面提速任务当前交接（2026-09-29）

## 接手目标

继续完成“封面提速部署与验收”任务。主 Codex 线程 `01a0e84d-cfde-7b91-8c91-cf0677d2b49f` 当前为 `systemError`：最近两次续跑均在远端上下文压缩时因 `stream disconnected ... error decoding response body` 失败，未留下新的完整助手回合。线程不是在 Word 自动化中等待。2026-09-29 交接检查时没有活动 `WINWORD` 进程；开始新 Word 操作前仍须重新检查。

先读本文件、原始项目交接 [cover-speed-deployment-20260928.md](cover-speed-deployment-20260928.md) 和验证交接 `EasySoftware_优化验证材料/reports/cover-upgrade-20260928-benchmark-handoff.md`。**本文件是 2026-09-29 的最新状态覆盖说明；旧交接中的“第二轮正在运行”“外部 Word PID 9232 仍在运行”等描述已过期。**

## 已完成及有效结论

- 基线 Word/PDF 诊断三轮在 `EasySoftware_优化验证材料/runs/cover-upgrade-20260928/baseline-resume03-{1,2,3}`：耗时 491.136、523.901、474.285 秒。
- 候选 Word/PDF 诊断三轮在 `.../optimized-resume04-{1,2,3}`：耗时 338.100、304.458、305.222 秒。各轮输入哈希保护为真。
- 六轮均完成生产全书 PDF QA，六轮均未整体通过。基线前三轮和候选三轮都确认“供电可视化分工系统”PDF 第 2 页正文空白；基线第 3 轮还确认第 1 页空白。基线第三轮地中海材料也有 Word 正文保护/封面 QA 失败。关键报告在各轮 `production-fullbook-qa.json`、`word_report.json`，空页图像位于轮次目录下 `供电可视化分工系统/production-qa/`。因此这些耗时**不能作为等质量性能对照，也不能据此宣称提速收益**。
- 公共入口预览已在隔离输入副本运行：`.../public-entry-work-resume05/combined_document_preview.json` 状态 `NEEDS_DOCUMENT_PROCESSING_REVIEW`、`output=null`、`failures=0`；输入在 `.../public-entry-input-resume05`。这是预览，不是正式处理或事务闭环。
- 最新全量测试日志 `EasySoftware_优化验证材料/reports/cover-upgrade-candidate-tests-resume04.log`：295 项通过、1 项跳过。Python 编译、PowerShell 解析、契约 JSON 和 `git diff --check` 通过。近期又有封面表格首元素分页回归修复，新增表格用例和契约字段；代码文件时间晚于此前 Word 合成测试，须按下方步骤重跑。
- 合成测试旧证据 `optimized-synthetic02-1` 和 `baseline-synthetic02-1` 早于最新候选修复及 harness 媒体引用检查修改，不能当作当前版本验收。旧结果暴露过候选回归：文档首元素为表格时封面标题被插入首个单元格、表格发生改动且版面推至第二页。随后代码已修复并补测试，但**修复后的真实 Word 合成复测尚未完成**。

## 线程停止时的代码与工作区状态

- 候选项目：`D:\360MoveData\Users\ban\Desktop\work_about\EasySoftware_候选_审查版_20260928`。验证材料：`D:\360MoveData\Users\ban\Desktop\work_about\EasySoftware_优化验证材料`。
- 相关改动涉及 `scripts/word_manual_pipeline_stable.ps1`、`tests/test_cover_performance.py`、`tests/test_manual_cover_rules.py`、`.project-memory/contracts/document-stage.schema.json` 等。验证 harness 在 `EasySoftware_优化验证材料/runs/cover-upgrade-20260928/benchmark.py`。
- `git status --short` 显示仓库有多项未提交改动，包含订单下载器等其他正在进行的工作。它们不是本交接的清理对象。不要 reset、checkout、stash、覆盖或回滚任何既有改动；先检查完整 diff，区分封面任务与其他工作。
- 老 PID 记录不代表进程所有权。检查时没有 WINWORD；之后也只能操作本轮明确启动且确认归属的 Word 实例，绝不结束用户 Word。

## 接下来按此顺序执行

1. 只读检查 `git status --short` 和封面相关 diff；确认最近代码确实包含首表格封面修复以及测试/契约同步。保留所有已有未提交内容。
2. 对最终候选源码重跑完整测试、Python 编译、PowerShell 解析、JSON 契约校验和 `git diff --check`，保存到**新的**日志名。
3. 检查 Word 全局状态和桌面会话可用性。若没有活动实例，以新的 label 在隔离合成夹具上分别跑候选和基线 Word（不要覆盖 `*synthetic02*`），重点核对：表格标题仍留在第一页、表格结构/文字不变；非连续标题标签不错误合并正文；末尾空白页保留并报告失败；实际引用媒体哈希不变。失败需看渲染图和报告，不能仅看退出码。
4. 修复或查明供电材料封面后第 2 页空白。使用生产全书 QA 与 Word 页面映射复核，空白页仍失败就不计入性能通过样本。之后才有意义用相同输入、版本、字体和阶段重跑基线/候选三轮，报告中位数及质量结果。
5. 在隔离输入副本上完成公共统一入口处理、事务输出、最终审核页、恢复副本启动/预览；检查原件哈希未变。预览 `resume05` 不算完成。
6. 将修复后证据与真实剩余事项追加到原交接记录，保留旧记录作为历史；最终汇总时明确区分诊断样本和质量合格性能样本。

## 接手时的验收底线（续接前状态）

截至本轮接手时，任务仍未完成：真实文档质量失败未解决、修复后的 Word 合成回归未复测、公共入口只做预览、尚无等质量速度收益结论。之后的续接结果见本文末尾；等质量总体速度收益仍未验证。

## 2026-09-29 本轮续接结果

### 候选修复

- Word 规则版本、进度/结果报告、`process_batch.VERSION` 与 `document-stage` 契约已同步为 `3.4.28`。
- 空白页修复仅在页面审计、下一页分节起点、段落控制字符和对象保护检查全部匹配时，删除该起点的单个显式分页符字符；保留段落标记与分节。其它未分类空白页继续报告未解决。
- 首表格封面作为整体保护；普通首表格无法识别为封面时，封面插入和边界回退均安全跳过并要求人工处理。边界回退同时排除表格内标题与标签，避免在单元格写入分页符；去版本号步骤也不会改写已识别的表格封面。
- 窄范围复审确认首表格高危路径及“标题在表格外、标签在表格内”的边界均已关闭。相关 mock 焦点用例、PowerShell 语法解析、契约 JSON 解析及 `git diff --check` 通过；按用户要求，本轮没有在最终版本上重跑全量测试套件。

### 隔离公共入口验收

- 输入副本：`EasySoftware_优化验证材料/runs/cover-upgrade-20260929-acceptance-resume06/public-entry-input-resume06`。通过 `unified_workflow.py --apply-document-processing --document-stages manual_cover --include-topic-mismatch` 实际运行公共入口；结束状态 `DOCUMENT_PROCESSING_OK`，进程失败数为 0，两个项目的 Word/PDF 阶段状态均为 `OK`。
- 输出：`public-entry-input-resume062件（共2件）_已处理_20260929_142712`；最终审核页为 `public-entry-work-resume06/combined_document_result.html`。
- 供电材料原第 2 页空白的根因确认是下一页分节起点的重复显式分页符。此轮报告记录仅删除起点分页符字符、保留分节；Word 未解决空白页为空，导出 PDF 共 40 页，第 2 页过滤页眉页脚后有 217 个正文字符，未确认空白页。
- 地中海材料输出 PDF 共 13 页，第 2 页有正文文本，Word 未解决空白页为空；两个项目的封面及 PDF 阶段报告均无失败。
- 处理前记录的 6 个输入文件 SHA-256 在公共入口处理后全部相同。处理输出在恢复预览前后的 6 个文件哈希也全部相同。

### 恢复预览与性能结论

- 恢复/预览从原隔离输入根启动，状态 `NEEDS_DOCUMENT_PROCESSING_REVIEW`，`applied=false`、`output=null`、失败数为 0、恢复事务数为 0；预览 HTML 已生成。输入与处理输出哈希均未改变。
- 本轮没有重跑三轮等质量基准，因此不据此宣称总体提速比例。历史三轮报告不是等质量对照；如需发布性能收益结论，应另行用同输入与相同质量条件补做基准。
- 原始业务材料未作为处理输入，其他未提交工作仍保留。


## 2026-09-29 范围收敛后的真机验证结果

### 运行边界与当前状态

- 按同一份 `full_rule_corpus` 和修正后的 benchmark 计划完成 baseline 三轮 `resume05`；候选第 1 轮开始后，用户明确要求不重复全书验证。候选第 1 轮作为唯一候选完整输出验收；第 2 轮在收到该范围要求前已自动启动，为避免中断隔离副本中的 Word 文档而让其正常完成，并在轮次边界阻止第 3 轮。没有启动新的全书流程。
- 本轮启动的 Word 均已退出；当前无 `WINWORD` 进程。用于边界停止的 `STOP_AFTER_ROUND` 标记已在 benchmark 退出后移除。
- 候选源码仍为 3.4.28；本次只改了验证材料中的 `benchmark.py`，没有改产品处理逻辑。项目既有其他未提交工作保持不变。

### 质量与版面结果

- baseline 三轮、候选两轮的 Word 返回码均为 0，输入源文件未变。各轮两份 DOCX 的正文文字、表格文字、drawing 数量、引用图片哈希和关系检查均通过保护断言。
- 五份输出均为供电 40 页、地中海 13 页；Word/PDF 页数一致，修正后的 `fullbook_qa.py` 五份均通过，没有确认或未解决的正文空白页，封面标题/标签和书签检查通过。
- 供电第 2 页已从真实 Word 页面映射确认有正文：baseline 的独立显式分页符被删除；候选的下一页分节起点重复分页符被删除，同时保留段落标记和分节。不可把早期旧轮次的第 2 页空白结论套用到这组新输入与修正计划。
- `reports/resume05-output-diff.json` 对 baseline 第 1 轮与候选第 1 轮做了全 PDF 页文本及空白迹象比较：两项目页数一致，所有页提取文本均相同，无可疑空白页；首页与封面后正文起始页的 PNG 栅格 SHA-256 完全相同。人工查看候选两项目的第 1、2 页，封面标题及“使用说明书”标签正确，第二页正文正常。该页文本比较与首页/第 2 页像素核对是版面证据，不代表 DOCX XML 完全相同。

### 用时拆分与解释

- baseline benchmark 总耗时为 501.441、504.670、499.484 秒，中位数 501.441 秒；候选完整 Word/PDF 流程两轮为 364.342、369.125 秒，均值 366.733 秒。候选样本只有两轮，故这里只报告观察值，不当作三轮统计结论。总计时覆盖 Word PowerShell 处理、逐页审计和 PDF 导出；后置结构保护和 PDF QA 不在该计时内。
- `word_report.json` 的单文档秒数显示：供电 baseline 394.132–397.925 秒，候选 273.643–277.247 秒；地中海 baseline 74.259–76.078 秒，候选 60.912–62.089 秒。它们是每份文件完整 Word 阶段耗时，不是单独的封面修改耗时。
- `word_trace_stable.log` 显示最大耗时在供电全页空白审计：baseline 每轮 221.210–224.976 秒，候选 114.635–115.900 秒；地中海对应 16.392–17.787 秒与约 5.5 秒。PDF 导出分别约 4 秒和 2.4–2.5 秒。后置生产 PDF QA 的独立批次用时约 7.3 秒（baseline 三份）和 6.1 秒（候选两份），不应混入封面编辑时间。

### 前一轮轻量选项的实现缺陷（后续已修正）

- 上一轮增加的 `--lightweight-timing` 只从 benchmark 计划中移除了 `AUDIT_ALL_PAGES`，但 PowerShell 封面流程仍无条件执行 `Repair-WordBlankPages` 和保存后的 `Get-WordPageAudit`。因此它并未真正绕过 Word COM 全页审计，不应使用其名称推断实际运行模式。
- 下面“2026-09-29 逐页筛查方案升级”记录了修正后的默认路径、脚本实现和轻量 QA 结果；以该节为准。

## 2026-09-29 逐页筛查方案升级

### 新旧路径

- 旧 Word 路径在封面修复阶段无条件调用 `Repair-WordBlankPages`；该函数先全量 `Get-WordPageAudit`，发现布局变化还会重复扫描。保存后即使已经审计过，又无条件收集第二份完整 `page_audit`。`Get-WordPageAudit` 会逐段跨 COM 读取 `Range.Information(wdActiveEndPageNumber)`、文本、分页符和表格状态，再遍历图片、浮动图形与分节；这才是主要耗时。
- 新常规封面路径由 `combined_document_workflow.py` 发出 `AUDIT_PDF_ALL_PAGES`，不发 `AUDIT_ALL_PAGES`。PowerShell 仍完成封面处理、局部封面检查、保存、单次 Word 页数统计及 PDF 导出，但只有显式 `AUDIT_ALL_PAGES` 或 `REPAIR_BLANK_PAGES` 时才运行全页 Word COM 映射/修复；报告写明 `page_audit_mode`，避免把“未做 Word 全页审计”误记成零空白页。
- `process_batch.py` 已支持对 `AUDIT_PDF_ALL_PAGES` 导出 PDF 后逐页提取正文文本、图片/表单/注释对象信号、页数及书签错误。只在正文为空且没有可解释内容对象的候选页渲染；确认正文空白后，它会添加 `REPAIR_BLANK_PAGES` 并重试，PowerShell 此时升级到完整 Word 审计。也可显式使用 `benchmark.py --full-word-audit` 做深度诊断。

### DOCX 与 PDF 差异检查

- `benchmark.py` 的原文保护签名现在只从封面起始区域排除标题和独立标签；正文中再次出现同名内容仍受保护。它比较正文段落文字、表格行/单元格/合并结构及文字、各 Word 故事部件中的图片引用与图片字节哈希、关系有效性、drawing 数量、未引用媒体数和媒体成员数；同时校验 DOCX ZIP 成员、XML 及图片可读性。封面文字/排版变化允许发生，文件时间戳不参与这些语义比较。
- `fullbook_qa.py` 轻量扫描每份输出 PDF 的页数及每页过滤页眉页脚后的正文字符/图片/表单/注释信号；只渲染空白候选页。传入已验收参考输出时，按正文页文本序列对齐，不依赖页码硬配；除文字差异外，还比较正文页逐字符位置、图片框、矢量图形框和纸张尺寸，几何信号超过阈值时只渲染对应输出页。第一页封面按预期允许变化。benchmark 的候选轮默认以已验收 `baseline-resume05-1` 为 PDF 对照；首轮也会比较，不再只拿候选首轮作参考。
- 几何信号比较能筛查文本/图片/矢量对象位置、数量或页面尺寸变化；它是页级筛选信号，**不构成视觉排版证明**，也不能证明所有字体、颜色或细微渲染差异。最终 Word 视觉效果仍由已完成的 resume05 完整验收担保；新增路径用于常规重复计时和异常页筛查。

### 已执行的轻量验证

- 使用现有输入 `inputs/full_rule_corpus`、候选完整输出 `optimized-resume05-1` 和基线参考 `baseline-resume05-1`，运行新版本 `fullbook_qa.py ... --reference-folder baseline-resume05-1`。没有启动 Word，也没有重跑基线或完整验收。
- 修正几何信号并最终复跑后用时 7.060 秒。两份 DOCX 均可读；正文、表格结构、引用图片内容及关系、drawing 数量、媒体成员保护均通过；Word 与 PDF 页数分别为 40/40 和 13/13；封面标题/标签正确；PDF 共检查 53 页正文信号，确认和未解决空白候选均为 0。与 baseline 对齐的正文页为 39+12 页，文字差异页和几何差异页均为 0；最终这次运行没有渲染页面。首次几何实现曾因 PDF 提取顺序差异误报并生成页面图，修正后这些误报图已清理。
- `tests/test_combined_document_workflow.py` 18 项通过，确认默认计划选择 PDF 筛查而非 Word 全页审计；PowerShell 语法解析、Python 脚本语法检查通过。benchmark 现在按 Word 报告里的 `page_audit_mode` 标实际审计方式，防止不支持新动作的旧 baseline 被误标成轻量；候选计时默认引用验收 baseline PDF。没有执行 Word 全量基准。
- 本次未真实启动一次新轻量 Word 输出；resume05 现有 DOCX/PDF 允许实测结构与 PDF 筛查器，常规计划动作由单测核实，PowerShell 轻量分支由解析检查核实。后续只在确需验证 COM 运行分支时最多做一轮候选轻量输出，不重跑全量五轮。

### 后续计时口径与升级条件

- `benchmark.py` 现以轻量 PDF 筛查为默认重复计时路径，`--full-word-audit` 才启用 Word 全页审计。`total_seconds` 计 Word 处理和 PDF 导出；`lightweight_qa_seconds` 单独计结构保护与 PDF 筛查。旧 resume05 的约 115 秒候选 Word 全页审计和约 4 秒 PDF 导出是旧完整路径计时，不代表新轻量路径耗时；本次 3.171 秒是验证方案处理现有输出的时间，不是 Word 处理提速结果。
- 原 DOCX 结构保护失败、封面检查失败、Word/PDF 页数异常、PDF 正文空白/无法判定、书签错误、参考输出正文或几何信号变化时，该轮标记失败并停止后续计时；空白候选、正文差异或几何差异只渲染对应页面。需要查分页成因或执行空白页修复时，再用 `AUDIT_ALL_PAGES`/`REPAIR_BLANK_PAGES` 完整 Word 审计兜底。benchmark 的总计时只包含 Word 处理与 PDF 导出；新增 DOCX 保护和 PDF 差异筛查单独记为 `lightweight_qa_seconds`。

## 2026-09-29 新轻量流程计时尝试

- 用户请求查看新流程实际耗时后，启动一次候选轻量 benchmark；输入仍为 `full_rule_corpus`，输出目录 `runs/cover-upgrade-20260928/optimized-light-1`。启动前确认没有 `WINWORD` 进程，目录与停止标记均不存在。
- 本次无法创建 Word COM 实例：`New-Object -ComObject Word.Application` 返回 HRESULT `80070520`（“指定的登录会话不存在，可能已终止”）。Word 处理约 1.05 秒后失败，没有生成 Word 报告或 PDF；QA 约 1.08 秒仅记录了缺失输出。因此这些秒数不是流程耗时，不能作为提速结果。没有留下运行中的 Word 进程。
- 失败证据在 `optimized-light-1/stderr.log`、`evidence.json` 和 `production-fullbook-qa.json`。待恢复可用的交互式 Windows 登录会话后，只需重跑一轮候选轻量 benchmark；不需重跑 baseline 或全书审计。

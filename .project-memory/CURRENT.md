# 当前架构导航

状态：Confirmed

| 职责 | 当前权威入口 |
|---|---|
| Windows 双击启动 | `启动EasySoftware.cmd` → `启动EasySoftware.ps1` |
| GUI、确认弹窗、进度展示 | `启动EasySoftware.ps1` |
| CLI 路由和公共进度 | `scripts/unified_workflow.py` |
| 订单资料下载与自动衔接 | `scripts/order_material_downloader_gui.ps1`、`scripts/order_download_core.py` |
| 说明 + 代码 + TXT 统一编排 | `scripts/combined_document_workflow.py` |
| 软件名称本地一致性提示 | `scripts/software_name_consistency.py` |
| 说明 OCR、相关性、正文年份、截图年份 | `scripts/manual_content_stage.py` |
| RapidOCR 批处理桥接 | `scripts/rapid_ocr_bridge.py` |
| 旧日期引擎桥接 | `scripts/date_year_rewriter_bridge.py` |
| 封面修复、Word 保存与说明 PDF | `scripts/word_manual_pipeline_stable.ps1` |
| 说明预审与动作计划 | `scripts/process_batch.py` |
| 登录模块结构识别、图文迁移与编号顺延 | `scripts/manual_login_module.py` + `scripts/manual_login_typography.py`（接入统一页面整理与既有 Word/PDF 管线） |
| 代码中文注释清理 | `scripts/code_comment_cleaner.py` |
| 说明阶段测试 | `tests/test_manual_content_stage.py`、`tests/test_manual_cover_rules.py` |
| 统一入口测试 | `tests/test_combined_document_workflow.py`、`tests/test_unified_document_entry.py` |

行为边界以 `contracts/document-stage.schema.json` 和 `contracts/code-phase.md` 为准。

2026-10-16：正式执行的标题正文 MISMATCH 剪贴板清单标题包含规范化客户名（去掉输入标签末尾件数）；用户反馈的 `easysoftware` DOCX 样本中，该字样是处理后 `word/document.xml` 的首段正文，不是水印对象，输入原稿不存在该段。当前处理脚本没有主动写入该字样的规则，具体来源尚不能从 DOCX 单独确认；反馈样本目录保持只读。

2026-10-16：订单批次的 MISMATCH 剪贴板标题从 `item-xx` 下客户归档目录获取客户名，不再显示批次根目录的时间戳编号；多个客户名以顿号列出。单客户正式输出名使用 `客户名_匹配件数（共总件数）_已处理_时间戳`，移除 `item-xx` 与客户归档两层包装，保留材料内项目目录；多客户批次仍分客户隔离以避免重名覆盖。未运行测试，详见 [HANDOFF](../HANDOFF.md)。

2026-10-16（Confirmed）：登录整理及说明字体对齐已完成。普通文本标题与 Word 标题共同定位，已有登录小节保留结构；新建和已有小节说明文字均对齐占优普通正文格式，存疑转人工，一致文件跳过 Word。当前规则以 [登录契约](contracts/login-module.schema.json) 为准，最新验证/速度见 [字体补充证据](evidence/login-font-20261016.md)，首轮结构验证见 [原证据](evidence/login-module-20261016.md)。

2026-10-16：下载后自动开始处理不再强制启用“一键处理”；主界面的一键选项写入 `title_body_config.json`，自动衔接时按已保存状态选择直接处理或预览确认。本次实现尚未运行测试，细节见 [HANDOFF](../HANDOFF.md)。

2026-10-15 当前状态（Confirmed；证据：提交 `f3622a0`、[HANDOFF](../HANDOFF.md)、`tests/test_combined_document_workflow.py`、`tests/test_progress_reporting.py`）：候选版包含订单下载自动衔接、自定义步骤持久化与重新打开恢复、TXT 检查/主要功能字段修复，以及进度文件占用时继续处理的保护。该提交已推送到 `origin/release/jev-shadow-20260928`。完整单元测试共 321 项，320 通过、1 跳过；真实订单与 Word/PDF 完整批次仍待本版本复验。以下按日期保留的旧状态为历史记录。

2026-10-15 后续修正（未提交）：空心线检测支持 Word VML `o:hr` 横线；图片 OCR 只提示以“公司”结尾的行，并在统一 HTML 报告输出命中截图；代码页数处理以 60 页为目标，不再按原 PDF 62 页补长，报告显示实际页数、末页行数及第 60 页行数。为保护代码正文，不裁切原文；自然排版超过 60 页时保留并报告实际页数。本轮未运行测试或 Word/PDF 批次。

2026-09-28：候选版部署封面提速规则 `3.4.27`。空白页修复只复用同一未变动布局的映射；保存后的 Word/PDF 终检仍独立执行。封面末行按需倒扫并保留异常回退，格式化范围限定连续、非表格、非目录封面块。真实速度和全部场景验收待后续模型完成，见 [部署交接](evidence/cover-speed-deployment-20260928.md)。

2026-09-29：统一文档处理默认只将 Jev `MATCH` 项目复制并处理；`MISMATCH` 与 `NOT_EVALUATED` 留在原输入。新目录名按 `客户名{MATCH件数}件（共{识别项目总数}件）_已处理_{时间戳}` 生成。实现与验证细节见当前 [HANDOFF](../HANDOFF.md)。

## 2026-10-13 后续更新目录（用户指定）

- 后续所有功能更新、修复和配置界面修改，只在 `D:/360MoveData/Users/ban/Desktop/work_about/EasySoftware_候选_审查版_20260928` 实施。此目录为当前唯一更新目标；不要再使用优化开发版作为开发或运行入口。
- 用户计划删除优化开发版；删除前仍需保全其中未推送的提交、未提交改动，并迁移候选版缺失的软件名称一致性模块及相关流程/报告接入。此记录当时尚缺 `match_software_name`；现已按下方候选版补齐记录迁移，旧报告仍不包含此检查。
- 已在候选版同步的本轮内容：Jev 网络有界重试与诊断、可调整并记忆的 Jev 通过线。迁移其他改动时保留候选版已有功能和配置，不整目录覆盖。

## 2026-10-13 候选版补齐名称一致性检查

- 以后唯一更新目录为本候选审查版。旧开发版、旧验证材料已由用户移入 `../EasySoftware_项目备份/`；旧路径不再是运行或修改目标。名称模块从备份旧版复用，更新仅落在候选版。
- 已接入 `match_software_name`：自定义步骤“匹配检测”内可独立勾选“软件名称一致性检查（本地）”；默认完整处理、仅检测和说明书预设包含。只选此项时不调用 Jev；与主题评分同时选中时复用正文抽取，评分通过或缓存命中仍运行名称检查。
- HTML/JSON 报告增加软件名称一致性结果；疑似名称用红色区域提示正式名称、正文名称、原句及段落位置，GUI 预检/处理完成状态和日志同步提示。名称检查仅报警，不改评分、失败计数或材料筛选。正文、表格、章节标题包含在范围内，图片内名称不检查。未选择或无法读取的材料不会被记为已检查。
- 保留候选版原有一键处理、订单下载、MATCH 筛选复制、输出件数命名、封面/空白页改进、现有 Jev 模式及通过线设置。网络重试与阈值记忆已在候选版；本次未调整当前 70 分模板默认值。
- 旧报告不会自动更新，需重新生成报告才显示此检查。未改写用户原材料及此前结果。
- 本轮完成 Python 静态语法检查及 diff 格式检查；未运行测试、GUI 或真实批次。迁移前快照位于 `../EasySoftware_项目备份/20261013_candidate_name_update/before/`。旧开发版移动导致候选版 Git 路径失效，已修复关联；Git 主库仍保存在备份中的旧开发版，保留该备份以保全历史。

## 2026-10-13 名称报警精简（v2）

- 候选版 `software_name_consistency.py` 使用 `software-name-consistency-v2`。完整正式名称及显式别名在提取候选之前保护，避免包含“平台/系统”的全称被拆开；过滤代称、操作句、运行环境、第三方引用和明确模块语境。
- 只有明确软件自称/定义、早期独立软件名称或带版本的独立产品名称等证据形成醒目报警。名称相近的简称/名称变体及证据不足的名称进入 `review_findings` / `review_count`，报告使用默认折叠的“待确认名称”，不计入 `warning_count`，不触发 GUI 冲突提示。正式名称未被自动更改，待确认项也不是已认可别名。
- 报告和 GUI 现有报警计数继续读取 `findings` / `warning_count`，只包含明确证据；疑似项状态为 `PENDING_REVIEW`。判断更保守，含糊自称可能留在待确认或未识别，图片内名称仍未检查。
- 保持复用已提取正文，规则和文档内保护模式预编译，未新增 Jev/OCR/Word 请求或修改语义缓存身份；模型缓存命中时名称规则仍按最新版本重算。材料筛选、Jev 阈值、修改及失败计数不变。
- 在真实反馈材料的只读副本生成本地预览：17 件原报告 63 条提示变为 1 条醒目报警、2 条折叠待确认，保留“伊甸之扉•永恒的誓约游戏软件”；名称规则本身合计 61.577 毫秒（不含启动和 DOCX 抽取），模型请求 0。最初管道样本仍检出“瑶池阆苑地下管网数字孪生可视化平台”。这些结果不代表独立留出集准确率。
- 源材料与用户旧报告未改写；新本地预览、比较记录和回退快照在 `../EasySoftware_项目备份/20261013_name_filter_v2/`。本轮未新增/运行测试套件，也未运行完整 Word/PDF 批次。

## 2026-10-16 图片 OCR 公司名称清除

- 新增独立步骤 `manual_company_text`，初始勾选；旧版已保存的步骤设置升级时默认加入此项，之后取消并保存会持续关闭。扫描 DOCX 全部内嵌图片，识别行末为“公司”（允许尾部标点）的 OCR 框，在输出副本中修复，并在统一 HTML 报告保存前后图对比及命中原文。
- 复用 `date_year_rewriter_bridge.py` 与批处理 OCR 引擎；公司清除/图片日期扫描阶段的一次图片 OCR 供日期识别和公司定位共同使用，命中框合并后每图只做一次 inpaint。既有抽样相关性 OCR 在同时选择时仍单独预检。预览不扫描全部图片，源材料保持不变。
- 既有 `company_ocr_alerts` 仍只覆盖抽样图片且仅提示，不修改；这一项选择才执行图片清除。修复质量需人工复核。

## 2026-10-15 图片 OCR“公司”字样红色提示

- 候选审查版统一报告复用说明书主题相关性/视觉复核阶段已经识别的抽样图片 OCR 文本；仅当 OCR 行末为“公司”（允许末尾标点）时，报告顶部附近显示醒目红色警告，并逐项列出项目、图片序号/文件名、命中次数和 OCR 原文行。JSON 同步保存 `company_ocr_alerts`。
- 仍只检查既有 OCR 样本，不增加 OCR 图片数量、模型或视觉调用，不启动 Word，不写入材料；没有选择截图 OCR/视觉检查时会明确写未执行。提醒只供人工复核，不自动通过、拦截或修改项目。
- OCR 为抽样识别，报告明确警告范围仅限实际抽取的样本；未抽取的图片不能据此断言没有“公司”。既有报告不会自动重写，新生成报告才包含此提示。
- 候选版 `scripts/combined_document_workflow.py` 已通过 Python 静态语法解析。未运行 OCR 批次或测试套件。修改前快照见 `../EasySoftware_项目备份/20261015_company_ocr/before/`。

## 2026-10-15 说明书分节垂直对齐规则

- 处理输出若只有一个 Word 节，不得将该节设置为垂直居中；封面独占第一节时才允许第一节居中，所有正文节保持顶端对齐。原因：封面格式化可能在分节边界缺失时把唯一正文节一起居中。
- 实现与审计位于 scripts/word_manual_pipeline_stable.ps1、scripts/repair_manual_structure.ps1 和 scripts/process_batch.py。

# EasySoftware 接手入口

状态：可接续。唯一工作目录为本候选审查版。历史提交 `f3622a0` 已推送；最近工作进度、验证边界与远端状态以 [HANDOFF 当前状态与接续路线](../HANDOFF.md) 为准。

## 阅读顺序

1. [REBUILD.md](REBUILD.md)：启动、处理顺序和验证命令。
2. [SPEC.md](SPEC.md)：材料、持久化及进度状态边界。
3. [contracts/document-stage.schema.json](contracts/document-stage.schema.json)、[contracts/login-module.schema.json](contracts/login-module.schema.json) 与 [contracts/code-phase.md](contracts/code-phase.md)：执行契约。
4. [CURRENT.md](CURRENT.md)：当前模块入口和最近变更。
5. [HANDOFF.md](../HANDOFF.md)：故障诊断、提交状态与真实验证边界。
6. [gaps.md](gaps.md)：尚未确认的扩展规则和性能问题。

## 当前行为

状态：Confirmed；已提交行为以提交 `f3622a0` 为历史基线，后续登录行为以登录契约与下述证据为准。

- 主界面与 CLI 复用统一材料流程；可按项选择匹配检测、说明书、代码和 TXT，自动补入必需依赖。自定义步骤及 Jev 通过线保存在本版本配置，重新打开设置可恢复。
- 新增图片 OCR 公司名称清除：`manual_company_text` 初始勾选，取消后跳过；旧版步骤设置只在首次迁移时加入新项。行为以 [document-stage.schema.json](contracts/document-stage.schema.json) 为准。
- 订单下载校验完成后可按用户勾选自动衔接候选审查。TXT 已支持名称/字段/主题检查，以及仅限主要功能字段的修复；更广的规则尚未确定。
- 原始材料保持不变；DOCX 是编辑源，PDF 由当前 DOCX 导出。Word COM 串行。进度文件被占用不能仅因状态显示中断材料处理。
- “整理登录模块”同时识别 Word 与普通文本标题；已有登录小节不重复新增。新建和已有小节的说明文字按占优普通正文格式局部对齐，存疑转人工。一致文件跳过 Word。

## 验证边界

本轮最新回归、独立副本 Word/PDF、保全/幂等及速度对照统一见 [字体与速度验证](evidence/login-font-20261016.md)，结构首轮见 [登录整理验证](evidence/login-module-20261016.md)。此前 321 项是历史提交基线。页码审计测试需要 `reportlab`，项目便携 Python 未预装。旧主窗口需退出重启加载修复。真实订单下载、外部模型和所有阶段组合的完整批次仍未在本轮重验。

本记忆包缺少 manifest 及若干重建资源，仅为接手导航；缺口见 [gaps.md](gaps.md)，不要宣称完整可复现。

历史 10 件材料验证和封面提速部署见 [corrected-real-batch-20260923.md](evidence/corrected-real-batch-20260923.md) 与 [cover-speed-deployment-20260928.md](evidence/cover-speed-deployment-20260928.md)，其结果不代表当前提交的完整验收。

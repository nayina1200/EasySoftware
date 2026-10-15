# EasySoftware 接手入口

状态：可接续。当前候选版提交 `f3622a0` 已推送到 `origin/release/jev-shadow-20260928`；代码与自动化测试已验证，真实订单和完整 Word/PDF 批次仍需本版本复验。

## 阅读顺序

1. [REBUILD.md](REBUILD.md)：启动、处理顺序和验证命令。
2. [SPEC.md](SPEC.md)：材料、持久化及进度状态边界。
3. [contracts/document-stage.schema.json](contracts/document-stage.schema.json) 与 [contracts/code-phase.md](contracts/code-phase.md)：执行契约。
4. [CURRENT.md](CURRENT.md)：当前模块入口和最近变更。
5. [HANDOFF.md](../HANDOFF.md)：故障诊断、提交状态与真实验证边界。
6. [gaps.md](gaps.md)：尚未确认的扩展规则和性能问题。

## 当前行为

状态：Confirmed；证据见上述代码入口、契约、测试和提交 `f3622a0`。

- 主界面与 CLI 复用统一材料流程；可按项选择匹配检测、说明书、代码和 TXT，自动补入必需依赖。自定义步骤及 Jev 通过线保存在本版本配置，重新打开设置可恢复。
- 订单下载校验完成后可按用户勾选自动衔接候选审查。TXT 已支持名称/字段/主题检查，以及仅限主要功能字段的修复；更广的规则尚未确定。
- 原始材料保持不变；DOCX 是编辑源，PDF 由当前 DOCX 导出。Word COM 串行。进度文件被占用不能仅因状态显示中断材料处理。

## 验证边界

提交前完整测试共 321 项，320 通过、1 跳过；页码审计测试需要 `reportlab`，项目便携 Python 未预装。旧主窗口需退出重启加载修复。真实订单下载、外部模型和 Word/PDF 全批次未随该提交重新运行。

历史 10 件材料验证和封面提速部署见 [corrected-real-batch-20260923.md](evidence/corrected-real-batch-20260923.md) 与 [cover-speed-deployment-20260928.md](evidence/cover-speed-deployment-20260928.md)，其结果不代表当前提交的完整验收。

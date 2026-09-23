# 系统稳定规格

## 目标

状态：Confirmed  
证据：[README.md](../README.md)、[开发维护规则.md](../开发维护规则.md)

EasySoftware 是 Windows 绿色便携程序，用于批量审计和处理软件著作权材料。公共入口同时服务 GUI 和命令行，输入材料、解压副本、修改副本、报告和最终交付必须保持清晰隔离。

## 参与者与外部边界

| 对象 | 状态 | 约束 | 证据 |
|---|---|---|---|
| 操作用户 | Confirmed | 通过 GUI 选择压缩包或文件夹并审核预览 | `README.md` |
| Microsoft Word | Confirmed | DOCX 修改后的 PDF 导出首选；全局串行 | `开发维护规则.md` 第 9、11 节 |
| LibreOffice | Confirmed | Word 不可用时的 PDF 导出兜底 | `README.md` |
| 7-Zip / Rar.exe | Confirmed | 解压及最终归档工具 | `README.md` |

## 不变量

状态：Confirmed  
证据：[开发维护规则.md](../开发维护规则.md)

- 原始输入不得被原地修改。
- PDF 不直接编辑；先修改 DOCX，再生成并验证 PDF。
- 有歧义的内容宁可保留并报告，不得冒险删除。
- 真实进度统一写入 `progress.json`；后续期次不得另建跑马灯式进度系统。
- 最终归档只在全部处理期完成后执行。

## 启动与持久化

状态：Confirmed  
证据：[启动EasySoftware.cmd](../启动EasySoftware.cmd)、[启动EasySoftware.ps1](../启动EasySoftware.ps1)、[scripts/unified_workflow.py](../scripts/unified_workflow.py)

- Windows 双击入口为 `启动EasySoftware.cmd`，由其调用 PowerShell GUI。
- 统一工作流入口为 `scripts/unified_workflow.py`。
- 运行状态和审计结果写入工作目录；项目源码自身不依赖数据库或后台服务。

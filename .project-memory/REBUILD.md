# 重建与接续

## 环境

- Windows 10/11；使用项目自带 `python\python.exe`。
- Word 封面修复和 PDF 导出需要 Microsoft Word COM；所有 Word 调用必须持有统一全局锁并串行。
- OCR 默认复用相邻目录 `图片日期智能处理工具_用户版_20260903` 的 RapidOCR/PP-OCRv6；不可用时降级 Windows OCR。

## 启动

双击 `启动EasySoftware.cmd`，选择“处理说明和代码”。命令行可用：

```powershell
.\python\python.exe .\scripts\unified_workflow.py <材料目录或压缩包> --process-documents
.\python\python.exe .\scripts\unified_workflow.py <材料目录或压缩包> --apply-document-processing
```

第一条仅预览，第二条在同一时间戳输出副本中正式执行。

## 流程顺序

1. 解压外层并递归展开内层压缩包。
2. 完成所有说明 OCR 预检；先决定项目拦截和批量熔断，再允许日期写入。
3. 若选中 `manual_company_text`，复用日期图片桥接及批处理 OCR 结果扫描说明书全部内嵌图，对行末“公司”文字框合并执行一次 inpaint，并把原图/处理图写入统一报告素材；只修改输出副本。抽样相关性 OCR 如另选，仍按既有流程执行。
4. 处理正文年份和截图年份。
5. 若选择 manual_login_module，先在既有 Word 锁与备份保护下完成登录结构整理/说明字体对齐，再复用同一 Word 会话执行所选封面等页面修复和说明 PDF 导出；最后复检登录结构、字体及 Word/PDF 页数。仅选登录整理且无需修改时跳过 Word。契约见 [login-module.schema.json](contracts/login-module.schema.json)。
6. 清理代码中文注释并串行导出代码 PDF。
7. 按所选项检查 TXT 名称、字段与主题；仅在选中 `txt_repair` 时修复主要功能字段，保留原编码和其他字段。

## 验收

```powershell
python -m pip install -r tests/requirements.txt
python -m unittest discover -s tests -q
.\python\python.exe -m py_compile .\scripts\combined_document_workflow.py .\scripts\manual_content_stage.py
```

页码审计测试使用 `reportlab` 生成合成 PDF，便携 Python 未预装此测试依赖；已安装依赖的工作区 Python 最新验证见 [字体与速度证据](evidence/login-font-20261016.md)，不要把历史 321 项基线当作本轮最新结果。布局改动必须以 Word/PDF 渲染首两页复核，不能只看结构化 QA。对输出还应检查 DOCX ZIP、PDF 可读性、TXT 编码与非目标字段、临时文件和原始压缩包哈希。

## 登录整理的定向复验

状态：Confirmed；命令入口见 scripts/unified_workflow.py，行为及既有结果见登录契约和上述证据。

```powershell
python -m unittest discover -s tests -p "test_manual_login*.py" -v
python -m unittest discover -s tests -p "test_combined_document_workflow.py" -v
.\python\python.exe .\scripts\unified_workflow.py <独立材料副本目录> --process-documents --document-stages manual_login_module
.\python\python.exe .\scripts\unified_workflow.py <独立材料副本目录> --apply-document-processing --document-stages manual_login_module
```

保持正常资格门禁；先预览，再对独立副本执行。检查新增小节的编号和图文顺序、已有小节无新增、说明文字格式、原文及图片保全、PDF 页数和登录页排版、二次运行无修改。字体参考不足应报告人工，不能猜测；本机样例速度不能外推为大文档保证。

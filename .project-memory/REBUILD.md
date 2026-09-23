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
3. 处理正文年份和截图年份。
4. 串行执行 Word 封面修复和说明 PDF。
5. 清理代码中文注释并串行导出代码 PDF。
6. TXT 仅登记，不读写。

## 验收

```powershell
.\python\python.exe -m unittest discover -s .\tests -q
.\python\python.exe -m py_compile .\scripts\combined_document_workflow.py .\scripts\manual_content_stage.py
```

布局改动必须以 Word/PDF 渲染首两页复核，不能只看结构化 QA。对输出还应检查 DOCX ZIP、PDF 可读性、TXT 哈希、临时文件和原始压缩包哈希。

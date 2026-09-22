# EasySoftware 3.5.0

中国软件著作权材料批量审计、修复、两轮审核与交付打包程序。

> 开发或修改程序前，必须先阅读并遵循 [开发维护规则](开发维护规则.md)。该文档记录当前已确认的材料规则、审核流程和验收标准。

## 双击启动

双击 `EasySoftware.lnk` 或 `启动EasySoftware.cmd`：

1. 点击“选择压缩包”，直接选择一个 RAR/ZIP（仅处理该文件）；也可点击“选择文件夹”处理整批目录。
2. 点击“开始处理”，程序递归展开内嵌压缩包并打开首轮图片审核页；只勾选图片不合格、需要退回的项目。
3. 未勾选项目进入自动修复，程序只修改 DOCX，再由 Word 导出对应 PDF。
4. 浏览器打开最终审核页，显示 TXT 前后对比、说明书封面前两页、代码 PDF 首尾页和每项修复结果。
5. 在窗口填写最终通过项目序号（如 `1,3-5` 或 `all`），点击“打包通过项目”。程序仅交付通过项目，并保持 RAR→RAR、ZIP→ZIP。
6. 如果文件夹同时包含 ZIP 和 RAR，GUI 会询问本次交付格式。

原始归档和解压审计副本保持不变；只有修改版会被修复。选择单个压缩包时，输出目录默认为同级的 `压缩包名_解压版` 和 `压缩包名_修改版`；选择批次目录时使用 `解压版` 和 `修改版`。缓存与统一审核页位于系统临时目录。

> 最终交付只支持 ZIP 或 RAR。7z 可以作为批次中的嵌套输入被展开，但不能作为顶层交付格式。

## 环境依赖

- Windows 10/11、Microsoft Word（说明书和代码 DOCX 修复、PDF 导出）
- Python 3.10+，并安装 `lxml`、`Pillow`、`pypdf`
- 7-Zip（RAR/7z 解压；程序也会查找常见的 AMD/NVIDIA 附带 7z）
- Poppler `pdftoppm` 或可用的 `pypdfium2`（PDF 审核图）
- WinRAR `Rar.exe`（RAR 交付打包）

程序会自动寻找带 Python 文档依赖的兼容运行时；缺失依赖时会返回机器可读错误。

## 命令行

```powershell
python scripts\unified_workflow.py "C:\材料\批次.rar"
python scripts\unified_workflow.py "C:\材料\批次.rar" --reject-images "2,5"
python scripts\unified_workflow.py "C:\材料\批次文件夹" --reject-images none
python scripts\unified_workflow.py "C:\材料\批次文件夹" --approve "1,3-5" --format rar
```

### 独立代码中文注释清理

该阶段不进入 TXT、说明书、终审或打包流程。默认先生成执行前预览，确认后再处理：

```powershell
python scripts\unified_workflow.py "C:\材料\批次.rar" --clean-code-comments
python scripts\unified_workflow.py "C:\材料\批次.rar" --apply-comment-cleanup
```

也可在主窗口选择 RAR、文件夹或已解压的材料目录后，点击“清理代码中文注释”。输入为压缩包时只接受 RAR；输入为文件夹时优先处理其中递归找到的 RAR，若无 RAR 但文件夹内已有“代码.docx/代码.pdf”则直接处理该已解压材料（无需解压）。程序只删除“代码.docx”中含汉字的真实注释，保护字符串、模板字符串、正则、HTML 页面文字和 Python 三引号内容。不确定项保留在文档中并列入报告。输出为同级的“原名_已处理_时间戳”目录，原 RAR 不移动、不修改；本阶段不重新压缩。清理后先修改 DOCX，再用 Word 优先、LibreOffice 兜底重新生成同名 PDF，只有新 PDF 验证有效后才替换输出副本中的旧 PDF。

任一 DOCX 处理失败或 PDF 无法重新导出时，批次状态为 `CODE_COMMENT_CLEANUP_PARTIAL`并返回非零退出码；GUI 会显示警告而不会报“完成”。单个文件或 RAR 的失败不会中断其他文件，总报告会保留每项失败原因。

预览和正式执行都显示全批次真实进度：已处理/待处理件数、当前 RAR 和项目、阶段、失败数、已用时间及 ETA。PDF 导出期间仍保持确定型进度条。

## 翻新失败业务助手

主窗口右上角的“翻新失败业务助手”用于处理已登录业务系统中的翻新批次：

1. 先在已登录的 Edge 或 Chrome 页面获取当前 Bearer 令牌，粘贴到助手的“Edge/Chrome 令牌”框；令牌只传给本次子进程，不写入材料目录和日志。
2. 点击“读取并预览”，助手按操作人和可选提交日期读取“进行中”“部分完成”批次。
3. “进行中”只列出已暂停订单，保留给人工进入“流程”处理；“部分完成”只提交处于代码生成步骤且具备订单编号的失败订单恢复。
4. 确认预览结果后取消“只预览”，点击“开始正式处理”。程序按批次顺序提交，每个批次单独记录结果。

每次预览或正式处理都会在 `%LOCALAPPDATA%\EasySoftware\refurbish-reports` 生成一份中文 HTML 报告，并保留同名 JSON 数据副本。报告使用绿色、蓝色和橙色直接标识已提交恢复、预览发现和待人工处理项目，含批次、候选订单、提交结果与失败原因，不包含登录令牌；助手完成后可点击“打开本次报告”。

助手默认记住令牌，保存在 `%LOCALAPPDATA%\EasySoftware\refurbish-settings.json`。文件使用当前 Windows 用户的 DPAPI 加密，其他 Windows 用户不能直接读取；需要移除时点击“清除已保存令牌”。

业务接口使用 `http://36.212.60.8:10000/prod-api`。浏览器令牌桥接采用人工粘贴，Edge 未开放时直接在 Chrome 登录并获取令牌即可，令牌过期时重新获取。

第一次运行成功后返回 `NEEDS_IMAGE_REVIEW`；填写图片未通过项目序号后返回 `NEEDS_UNIFIED_REVIEW`。两次审核页都带有可继续执行的 `next` 命令。只有最终打包成功返回退出码 0；审核等待或错误状态返回退出码 2，调用方应以 JSON 的 `status` 字段判断。

## 自检与测试

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\启动EasySoftware.ps1 -SelfTest
python -m unittest discover -s tests -v
```

公共入口只有 `scripts/unified_workflow.py`。旧文件 `scripts/workflow.py` 会自动转发到该入口；其他低层脚本仅用于诊断，不应作为正常处理入口。

## 项目结构与开发

```text
easysoftware/
├── 启动EasySoftware.ps1 / .cmd / .lnk   GUI 入口（选择压缩包/文件夹、开始处理、打包）
├── SKILL.md                           程序化工作流的技能说明（与 GUI 同一套流程）
├── 开发维护规则.md                      开发前必读：材料规则、处理流程、验收标准、文件安全
├── README.md
├── scripts/
│   ├── unified_workflow.py             唯一公共命令行入口（GUI 也调用它）
│   ├── workflow.py                     旧入口，自动转发到 unified_workflow.py
│   ├── process_batch.py                批处理主逻辑：解压、审核、修复、验证、审核页
│   ├── extract_archives.py             递归展开 ZIP/RAR/7z
│   ├── word_manual_pipeline_stable.ps1 说明书 DOCX 修复 + Word 导出 PDF（单会话、带备份回滚）
│   ├── repair_code_headers.ps1         代码 DOCX 页眉/页码标准化 + PDF 替换（审批范围内）
│   ├── package_approved_modified.py    按审批项目打包 ZIP/RAR + 完整性校验
│   ├── refurbish_workflow.py / _gui.ps1 翻新失败业务助手
│   ├── enforce_cover_typography.py / normalize_manual_covers.ps1 / qa_manual_covers.py  封面排版
│   ├── audit_packages.py / audit_cover_word.ps1 / windows_ocr.ps1                      审核辅助
│   └── run_word_item.py / export_*pdf*.ps1                                             单项目/导出诊断
├── references/rules.md                 材料规则（与开发维护规则.md 配套）
├── tests/                              unittest 测试（test_archives / test_word_item / test_workflow_safety）
└── 临时目录（系统 %TEMP%）：缓存、统一审核页
```

关键开发约定（完整约束见 `开发维护规则.md`）：

- **DOCX 是唯一编辑源**：修复 PDF 前必须先修 DOCX，再由 Word 导出 PDF；禁止直接在 PDF 上写页眉/页码。
- **只有 `修改版` 可变**：输入压缩包与 `解压版` 永远保持原样。
- **缓存由 SHA-256 + 规则版本决定**：修改修复规则时必须提升规则版本或使缓存失效。
- **不要手动拼接低层脚本**：正常流程只走 `unified_workflow.py`；低层脚本用于诊断失败或写测试。
- 修改后运行 `python -m unittest discover -s tests -v`；涉及流程/报告/打包时跑完整测试集。

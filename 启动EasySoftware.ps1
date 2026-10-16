param([switch]$SelfTest, [string]$InputPath = "", [switch]$AutoStart, [string]$DocumentStages = "", [switch]$IncludeTopicMismatch, [switch]$OrderDownloads)

$ErrorActionPreference = "Stop"
$projectRoot = $PSScriptRoot
$launchDocumentStages = $DocumentStages
$entry = Join-Path $projectRoot "scripts\unified_workflow.py"
$env:PYTHONIOENCODING = "utf-8"

function Assert-NoReparseComponents([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { throw "拒绝空隔离路径。" }
    $fullPath = [IO.Path]::GetFullPath($Path)
    $cursor = [IO.Path]::GetPathRoot($fullPath)
    $relativePath = $fullPath.Substring($cursor.Length)
    $segments = @($relativePath.Split([char[]]@([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar), [StringSplitOptions]::RemoveEmptyEntries))
    foreach ($segment in $segments) {
        $cursor = Join-Path $cursor $segment
        if (Test-Path -LiteralPath $cursor) {
            $item = Get-Item -LiteralPath $cursor -Force
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "优化开发版隔离路径包含链接或重解析点：$cursor"
            }
        }
    }
}

$script:isOptimizationDev = Test-Path -LiteralPath (Join-Path $projectRoot ".optimization-dev")
if ($script:isOptimizationDev) {
    $script:validationRoot = [IO.Path]::GetFullPath((Join-Path (Split-Path -Parent $projectRoot) "EasySoftware_优化验证材料"))
    Assert-NoReparseComponents $script:validationRoot
    $script:validationRunsRoot = Join-Path $script:validationRoot "runs"
    $script:devPaths = @{
        Temp = Join-Path $script:validationRoot "temp"
        PyCache = Join-Path $script:validationRoot "cache\python"
        YoloCache = Join-Path $script:validationRoot "cache\yolo"
        UserData = Join-Path $script:validationRoot "state\user-data"
        DateTool = Join-Path $script:validationRoot "dependencies\date_ocr"
        Llama = Join-Path $script:validationRoot "dependencies\llama"
        VisionConfig = Join-Path $script:validationRoot "config\vision_config.json"
    }
    $requiredDevPaths = @(
        (Join-Path $projectRoot "python\python.exe"),
        (Join-Path $script:devPaths.DateTool "_internal\python.exe"),
        (Join-Path $script:devPaths.DateTool "app.py"),
        (Join-Path $projectRoot "scripts\rapid_ocr_bridge.py"),
        (Join-Path $projectRoot "scripts\date_year_rewriter_bridge.py"),
        (Join-Path $script:devPaths.Llama "llama-server.exe"),
        (Join-Path $script:validationRoot "dependencies\model\QwenPaw-Flash-2B-Q4_K_M.gguf"),
        $script:devPaths.VisionConfig
    )
    foreach ($required in $requiredDevPaths) {
        Assert-NoReparseComponents $required
        if (-not (Test-Path -LiteralPath $required -PathType Leaf)) { throw "优化开发版隔离依赖不完整；为避免能力静默降级，已阻止启动：$required" }
    }
    foreach ($directory in @($script:devPaths.Temp, $script:devPaths.PyCache, $script:devPaths.YoloCache, $script:devPaths.UserData, $script:validationRunsRoot)) {
        Assert-NoReparseComponents $directory
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    $env:TEMP = $script:devPaths.Temp
    $env:TMP = $script:devPaths.Temp
    $env:PYTHONPYCACHEPREFIX = $script:devPaths.PyCache
    $env:YOLO_CONFIG_DIR = $script:devPaths.YoloCache
    $env:EASYSOFTWARE_USER_DATA_ROOT = $script:devPaths.UserData
    $env:EASYSOFTWARE_DATE_TOOL_ROOT = $script:devPaths.DateTool
    $env:EASYSOFTWARE_LLAMA_HOME = $script:devPaths.Llama
    $env:EASYSOFTWARE_VISION_CONFIG = $script:devPaths.VisionConfig
    $env:EASYSOFTWARE_VALIDATION_ROOT = $script:validationRoot
    $env:EASYSOFTWARE_VALIDATION_RUNS_ROOT = $script:validationRunsRoot
    $env:EASYSOFTWARE_DEVELOPMENT_READ_ONLY = "1"
}

# 便携版增强：把捆绑的工具目录加入 PATH，使 find_7z / find_rar / poppler 能找到
foreach ($sub in @("tools\7z", "tools\rar")) {
    $toolDir = Join-Path $projectRoot $sub
    if (Test-Path -LiteralPath $toolDir) {
        $env:PATH = $toolDir + ";" + $env:PATH
    }
}

function Find-Python {
    # 便携版优先使用捆绑的 Python 运行时
    $bundledPython = Join-Path $projectRoot "python\python.exe"
    $moduleCheck = if ($script:isOptimizationDev) {
        "import lxml, PIL, pypdf, pypdfium2, pdfminer"
    } else {
        "import lxml, PIL, pypdf, pypdfium2"
    }
    if (Test-Path -LiteralPath $bundledPython -PathType Leaf) {
        try {
            & $bundledPython -c $moduleCheck 2>$null
            if ($LASTEXITCODE -eq 0) { return $bundledPython }
        } catch {}
    }
    if ($script:isOptimizationDev) {
        $isolatedPython = [string]$env:EASYSOFTWARE_PYTHON
        if (-not [string]::IsNullOrWhiteSpace($isolatedPython) -and (Test-Path -LiteralPath $isolatedPython -PathType Leaf)) {
            try {
                & $isolatedPython -c $moduleCheck 2>$null
                if ($LASTEXITCODE -eq 0) { return $isolatedPython }
            } catch {}
        }
        return $null
    }
    $candidates = @()
    $command = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($null -ne $command) { $candidates += $command.Source }
    $command = Get-Command python -ErrorAction SilentlyContinue
    if ($null -ne $command) { $candidates += $command.Source }
    $profileRoot = [Environment]::GetFolderPath([Environment+SpecialFolder]::UserProfile)
    $bundled = Join-Path $profileRoot '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
    if (Test-Path -LiteralPath $bundled) { $candidates += $bundled }
    foreach ($profile in Get-ChildItem -LiteralPath 'C:\Users' -Directory -ErrorAction SilentlyContinue) {
        $bundled = Join-Path $profile.FullName '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
        if (Test-Path -LiteralPath $bundled) { $candidates += $bundled }
    }
    foreach ($candidate in ($candidates | Select-Object -Unique)) {
        try {
            & $candidate -c "import lxml, PIL, pypdf" 2>$null
            if ($LASTEXITCODE -eq 0) { return $candidate }
        } catch {}
    }
    return $null
}

function Quote-Argument([string]$Value) {
    return '"' + ($Value -replace '"', '\"') + '"'
}

# 可作为输入的全部压缩包类型。必须与 scripts/archive_types.py 的 ARCHIVE_SUFFIXES 保持一致。
# 交付（打包输出）仍然只有 RAR/ZIP，其余类型只作为输入读取。
$script:ArchiveExtensions = @(
    '.tar.gz', '.tar.bz2', '.tar.xz', '.tar.lz4',
    '.7z', '.rar', '.rarx', '.zip', '.jar', '.iso', '.cab',
    '.tar', '.gz', '.tgz', '.bz2', '.tbz', '.tbz2', '.xz', '.txz',
    '.lz4', '.lzma', '.lz', '.zst', '.zstd', '.z', '.pax', '.zpaq',
    '.ace', '.arj', '.lzh', '.lha', '.wim', '.msi', '.deb', '.rpm',
    '.dmg', '.hfs', '.pet', '.vhd', '.vhdx', '.gem'
)

function Test-ArchivePath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return $false }
    $item = Get-Item -LiteralPath $Path
    if ($item.PSIsContainer) { return $false }
    $name = $item.Name.ToLowerInvariant()
    foreach ($ext in $script:ArchiveExtensions) {
        if ($name.EndsWith($ext)) { return $true }
    }
    return $false
}

function Test-ArchiveName([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $leaf = ([IO.Path]::GetFileName($Value.Trim())).ToLowerInvariant()
    foreach ($ext in $script:ArchiveExtensions) {
        if ($leaf.EndsWith($ext)) { return $true }
    }
    return $false
}

function Assert-OptimizationPath([string]$Path, [switch]$Material) {
    if (-not $script:isOptimizationDev) { return }
    if ([string]::IsNullOrWhiteSpace($Path)) { throw "优化开发版隔离拒绝空路径。" }
    $target = [IO.Path]::GetFullPath($Path)
    if ($Material) {
        Assert-NoReparseComponents $target
        return
    }
    $allowedRoot = $script:validationRoot
    $allowedRoot = [IO.Path]::GetFullPath($allowedRoot).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    $prefix = $allowedRoot + [IO.Path]::DirectorySeparatorChar
    if (-not $target.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "优化开发版工作目录必须位于验证隔离目录内：$target"
    }
    $relative = $target.Substring($prefix.Length)
    $segments = @($relative.Split([char[]]@([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar), [StringSplitOptions]::RemoveEmptyEntries))
    $workspaceAllowed = $false
    foreach ($workspaceName in @("temp", "cache", "reports", "state")) {
        $workspaceRoot = [IO.Path]::GetFullPath((Join-Path $script:validationRoot $workspaceName)).TrimEnd([IO.Path]::DirectorySeparatorChar)
        if ($target.StartsWith($workspaceRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { $workspaceAllowed = $true }
    }
    if (-not $workspaceAllowed) { throw "优化开发版工作目录必须位于专用 temp/cache/reports/state 目录：$target" }
    $rootCursor = $allowedRoot
    while (-not [string]::IsNullOrWhiteSpace($rootCursor)) {
        if (Test-Path -LiteralPath $rootCursor) {
            $rootItem = Get-Item -LiteralPath $rootCursor -Force
            if (($rootItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "优化开发版隔离根或其父目录是链接路径：$rootCursor"
            }
        }
        $parentCursor = Split-Path -Parent $rootCursor
        if ([string]::Equals($parentCursor, $rootCursor, [StringComparison]::OrdinalIgnoreCase)) { break }
        $rootCursor = $parentCursor
    }
    $cursor = $allowedRoot
    $separators = [char[]]@([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    foreach ($segment in $segments) {
        if ([string]::IsNullOrWhiteSpace($segment)) { continue }
        $cursor = Join-Path $cursor $segment
        if (Test-Path -LiteralPath $cursor) {
            $item = Get-Item -LiteralPath $cursor -Force
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "优化开发版隔离不接受指向其他位置的链接路径：$cursor"
            }
        }
    }
}

function Prepare-ArchiveWorkspace([System.IO.FileInfo]$Archive) {
    Assert-OptimizationPath $Archive.FullName -Material
    while ($Archive.Directory.Name -ieq $Archive.BaseName -and $Archive.Directory.Parent.Name -ieq $Archive.BaseName) {
        $target = Join-Path $Archive.Directory.Parent.FullName $Archive.Name
        Assert-OptimizationPath $target -Material
        Assert-OptimizationPath $Archive.Directory.FullName -Material
        if (Test-Path -LiteralPath $target) {
            $existing = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
            $incoming = (Get-FileHash -LiteralPath $Archive.FullName -Algorithm SHA256).Hash
            if ($existing -ne $incoming) { throw "同名文件夹内已存在不同的压缩包：$target" }
            Remove-Item -LiteralPath $Archive.FullName -Force
        } else {
            Move-Item -LiteralPath $Archive.FullName -Destination $target
        }
        $emptyFolder = $Archive.Directory.FullName
        if (-not (Get-ChildItem -LiteralPath $emptyFolder -Force | Select-Object -First 1)) { Remove-Item -LiteralPath $emptyFolder -Force }
        $Archive = Get-Item -LiteralPath $target
    }
    $folder = Join-Path $Archive.DirectoryName $Archive.BaseName
    if ($Archive.Directory.Name -ieq $Archive.BaseName) { return $Archive.FullName }
    Assert-OptimizationPath $folder -Material
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    $target = Join-Path $folder $Archive.Name
    Assert-OptimizationPath $target -Material
    if (Test-Path -LiteralPath $target) {
        $existing = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
        $incoming = (Get-FileHash -LiteralPath $Archive.FullName -Algorithm SHA256).Hash
        if ($existing -ne $incoming) { throw "同名文件夹内已存在不同的压缩包：$target" }
        Remove-Item -LiteralPath $Archive.FullName -Force
    } else {
        Move-Item -LiteralPath $Archive.FullName -Destination $target
    }
    return $target
}

$python = Find-Python
if ($script:isOptimizationDev) {
    $datePython = Join-Path $script:devPaths.DateTool "_internal\python.exe"
    & $datePython -c "import cv2, numpy, rapidocr, gradio, PIL, docx, ultralytics" 2>$null
    if ($LASTEXITCODE -ne 0) { throw "日期 OCR 运行时无法加载必要引擎；为避免能力静默降级，已阻止启动。" }
}
if ($SelfTest) {
    Write-Host "EasySoftware GUI launcher self-test"
    Write-Host "Project root: $projectRoot"
    Write-Host "Python: $python"
    Write-Host "Entry: $entry"
    if ([string]::IsNullOrWhiteSpace($python) -or -not (Test-Path -LiteralPath $entry -PathType Leaf)) { exit 2 }
    if ($script:isOptimizationDev) {
        Assert-OptimizationPath (Join-Path $script:validationRunsRoot "__isolation_selftest\input") -Material
        Assert-OptimizationPath (Join-Path (Split-Path -Parent $projectRoot) "formal-orders") -Material
        $workspaceBlocked = $false
        try { Assert-OptimizationPath (Join-Path (Split-Path -Parent $projectRoot) "formal-output") }
        catch { $workspaceBlocked = $true }
        if (-not $workspaceBlocked) { throw "隔离自检失败：验证区外工作目录未被阻止。" }
    }
    & $python $entry --help
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    Write-Host "Self-test OK"
    exit 0
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

if ([string]::IsNullOrWhiteSpace($python)) {
    [System.Windows.Forms.MessageBox]::Show("未找到带 lxml、Pillow、pypdf 依赖的 Python 运行时。", "EasySoftware", "OK", "Error") | Out-Null
    exit 2
}
if (-not (Test-Path -LiteralPath $entry -PathType Leaf)) {
    [System.Windows.Forms.MessageBox]::Show("Main program not found: $entry", "EasySoftware", "OK", "Error") | Out-Null
    exit 2
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "EasySoftware - Software Copyright Materials"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(980, 720)
$form.MinimumSize = New-Object System.Drawing.Size(820, 600)
$form.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)

$title = New-Object System.Windows.Forms.Label
$title.Text = "EasySoftware 3.5"
$title.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 18, [System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(22, 16)
$title.AutoSize = $true
$form.Controls.Add($title)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = "选择一个压缩包（RAR / ZIP / 7z 等），或选择一个已解压的材料文件夹。"
$subtitle.ForeColor = [System.Drawing.Color]::DimGray
$subtitle.Location = New-Object System.Drawing.Point(25, 55)
$subtitle.AutoSize = $true
$form.Controls.Add($subtitle)

$visionSettingsButton = New-Object System.Windows.Forms.Button
$visionSettingsButton.Text = "视觉模型配置"
$visionSettingsButton.Location = New-Object System.Drawing.Point(810, 20)
$visionSettingsButton.Size = New-Object System.Drawing.Size(130, 32)
$visionSettingsButton.Anchor = "Top,Right"
$form.Controls.Add($visionSettingsButton)

$inputBox = New-Object System.Windows.Forms.TextBox
$inputBox.Location = New-Object System.Drawing.Point(26, 88)
$inputBox.Size = New-Object System.Drawing.Size(650, 30)
$inputBox.Anchor = "Top,Left,Right"
$form.Controls.Add($inputBox)

$archiveButton = New-Object System.Windows.Forms.Button
$archiveButton.Text = "选择压缩包"
$archiveButton.Location = New-Object System.Drawing.Point(690, 85)
$archiveButton.Size = New-Object System.Drawing.Size(120, 34)
$archiveButton.Anchor = "Top,Right"
$form.Controls.Add($archiveButton)

$folderButton = New-Object System.Windows.Forms.Button
$folderButton.Text = "选择文件夹"
$folderButton.Location = New-Object System.Drawing.Point(820, 85)
$folderButton.Size = New-Object System.Drawing.Size(120, 34)
$folderButton.Anchor = "Top,Right"
$form.Controls.Add($folderButton)

$startButton = New-Object System.Windows.Forms.Button
$startButton.Text = "开始处理"
$startButton.Location = New-Object System.Drawing.Point(26, 135)
$startButton.Size = New-Object System.Drawing.Size(130, 38)
$startButton.BackColor = [System.Drawing.Color]::FromArgb(22, 93, 255)
$startButton.ForeColor = [System.Drawing.Color]::White
$startButton.FlatStyle = "Flat"
$form.Controls.Add($startButton)

$fromSecondButton = New-Object System.Windows.Forms.Button
$fromSecondButton.Text = "从第二轮开始"
$fromSecondButton.Location = New-Object System.Drawing.Point(166, 135)
$fromSecondButton.Size = New-Object System.Drawing.Size(150, 38)
$form.Controls.Add($fromSecondButton)

$continueExtractedButton = New-Object System.Windows.Forms.Button
$continueExtractedButton.Text = "从解压版继续"
$continueExtractedButton.Location = New-Object System.Drawing.Point(326, 135)
$continueExtractedButton.Size = New-Object System.Drawing.Size(140, 38)
$form.Controls.Add($continueExtractedButton)

$reviewButton = New-Object System.Windows.Forms.Button
$reviewButton.Text = "打开统一审核页"
$reviewButton.Location = New-Object System.Drawing.Point(476, 135)
$reviewButton.Size = New-Object System.Drawing.Size(140, 38)
$reviewButton.Enabled = $false
$form.Controls.Add($reviewButton)

$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Text = "修改版 Word 导出 PDF"
$exportButton.Location = New-Object System.Drawing.Point(626, 135)
$exportButton.Size = New-Object System.Drawing.Size(160, 38)
$exportButton.Enabled = $false
$form.Controls.Add($exportButton)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Ready"
$statusLabel.Location = New-Object System.Drawing.Point(800, 177)
$statusLabel.Size = New-Object System.Drawing.Size(160, 26)
$statusLabel.Anchor = "Top,Left,Right"
$form.Controls.Add($statusLabel)

$refurbishButton = New-Object System.Windows.Forms.Button
$refurbishButton.Text = "翻新失败业务助手"
$refurbishButton.Location = New-Object System.Drawing.Point(780, 591)
$refurbishButton.Size = New-Object System.Drawing.Size(160, 38)
$refurbishButton.Anchor = "Bottom,Right"
$refurbishButton.BackColor = [System.Drawing.Color]::FromArgb(8, 116, 67)
$refurbishButton.ForeColor = [System.Drawing.Color]::White
$refurbishButton.FlatStyle = "Flat"
$form.Controls.Add($refurbishButton)

$orderDownloadButton = New-Object System.Windows.Forms.Button
$orderDownloadButton.Text = "下载订单资料"
$orderDownloadButton.Location = New-Object System.Drawing.Point(610, 591)
$orderDownloadButton.Size = New-Object System.Drawing.Size(160, 38)
$orderDownloadButton.Anchor = "Bottom,Right"
$orderDownloadButton.BackColor = [System.Drawing.Color]::FromArgb(35, 109, 173)
$orderDownloadButton.ForeColor = [System.Drawing.Color]::White
$orderDownloadButton.FlatStyle = "Flat"
$form.Controls.Add($orderDownloadButton)

$codeCommentButton = New-Object System.Windows.Forms.Button
$codeCommentButton.Text = "处理说明和代码"
$codeCommentButton.Location = New-Object System.Drawing.Point(796, 135)
$codeCommentButton.Size = New-Object System.Drawing.Size(160, 38)
$codeCommentButton.Anchor = "Top,Right"
$codeCommentButton.BackColor = [System.Drawing.Color]::FromArgb(154, 91, 0)
$codeCommentButton.ForeColor = [System.Drawing.Color]::White
$codeCommentButton.FlatStyle = "Flat"
$form.Controls.Add($codeCommentButton)

$documentStageLabel = New-Object System.Windows.Forms.Label
$documentStageLabel.Text = "处理说明和代码流程："
$documentStageLabel.Location = New-Object System.Drawing.Point(26, 179)
$documentStageLabel.AutoSize = $true
$documentStageLabel.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$form.Controls.Add($documentStageLabel)

$workflowOptionsButton = New-Object System.Windows.Forms.Button
$workflowOptionsButton.Text = "自定义步骤…"
$workflowOptionsButton.Location = New-Object System.Drawing.Point(190, 174)
$workflowOptionsButton.Size = New-Object System.Drawing.Size(125, 30)
$form.Controls.Add($workflowOptionsButton)

$stageSelectionHint = New-Object System.Windows.Forms.Label
$stageSelectionHint.Text = "完整处理"
$stageSelectionHint.ForeColor = [System.Drawing.Color]::DimGray
$stageSelectionHint.Location = New-Object System.Drawing.Point(327, 200)
$stageSelectionHint.Size = New-Object System.Drawing.Size(610, 18)
$stageSelectionHint.Anchor = "Top,Left,Right"
$stageSelectionHint.AutoSize = $false
$stageSelectionHint.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8.5)
$form.Controls.Add($stageSelectionHint)

$oneClickProcessingCheck = New-Object System.Windows.Forms.CheckBox
$oneClickProcessingCheck.Text = "一键处理（跳过预览直接生成结果）"
$oneClickProcessingCheck.Location = New-Object System.Drawing.Point(327, 176)
$oneClickProcessingCheck.Size = New-Object System.Drawing.Size(210, 27)
$oneClickProcessingCheck.Checked = $false
$oneClickProcessingCheck.Anchor = "Top,Left"
$oneClickProcessingCheck.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8.5)
$oneClickProcessingCheck.Add_CheckedChanged({
    $toolTip = if ($oneClickProcessingCheck.Checked) {
        "点击处理说明和代码后直接正式处理并生成结果，中间不打开预览。"
    } else {
        "先生成预览；确认后点击确认处理并生成结果。"
    }
    $documentWorkflowTips.SetToolTip($oneClickProcessingCheck, $toolTip)
})
$documentWorkflowTips = New-Object System.Windows.Forms.ToolTip
$documentWorkflowTips.SetToolTip($oneClickProcessingCheck, "先生成预览；确认后点击确认处理并生成结果。")
$form.Controls.Add($oneClickProcessingCheck)

$skipTopicMismatchCheck = New-Object System.Windows.Forms.CheckBox
$skipTopicMismatchCheck.Text = "跳过标题正文不匹配项目（默认）"
$skipTopicMismatchCheck.Location = New-Object System.Drawing.Point(540, 176)
$skipTopicMismatchCheck.Size = New-Object System.Drawing.Size(250, 27)
$skipTopicMismatchCheck.Checked = $true
$skipTopicMismatchCheck.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8.5)
$skipTopicMismatchCheck.Anchor = "Top,Right"
$form.Controls.Add($skipTopicMismatchCheck)

$workflowGuide = New-Object System.Windows.Forms.Label
$workflowGuide.Text = "工作流：1 选择材料  >  2 首轮看图并微调 Word  >  3 二次处理  >  4 终审  >  5 导出 PDF  >  6 压缩交付"
$workflowGuide.ForeColor = [System.Drawing.Color]::FromArgb(49, 82, 138)
$workflowGuide.BackColor = [System.Drawing.Color]::FromArgb(238, 244, 255)
$workflowGuide.Location = New-Object System.Drawing.Point(26, 220)
$workflowGuide.Size = New-Object System.Drawing.Size(710, 18)
$workflowGuide.Anchor = "Top,Left,Right"
$workflowGuide.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8.5)
$form.Controls.Add($workflowGuide)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(26, 242)
$progress.Size = New-Object System.Drawing.Size(914, 22)
$progress.Anchor = "Top,Left,Right"
$progress.Style = "Continuous"
$progress.Maximum = 100
$progress.Value = 0
$progress.ForeColor = [System.Drawing.Color]::FromArgb(22, 93, 255)
$form.Controls.Add($progress)

$progressDetail = New-Object System.Windows.Forms.Label
$progressDetail.Text = "等待开始"
$progressDetail.ForeColor = [System.Drawing.Color]::DimGray
$progressDetail.Location = New-Object System.Drawing.Point(26, 268)
$progressDetail.Size = New-Object System.Drawing.Size(914, 24)
$progressDetail.Anchor = "Top,Left,Right"
$form.Controls.Add($progressDetail)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(26, 296)
$logBox.Size = New-Object System.Drawing.Size(914, 279)
$logBox.Anchor = "Top,Bottom,Left,Right"
$logBox.Multiline = $true
$logBox.ReadOnly = $true
$logBox.ScrollBars = "Both"
$logBox.WordWrap = $false
$logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($logBox)

$approvalLabel = New-Object System.Windows.Forms.Label
$approvalLabel.Text = "通过的项目序号："
$approvalLabel.Location = New-Object System.Drawing.Point(26, 598)
$approvalLabel.AutoSize = $true
$approvalLabel.Anchor = "Bottom,Left"
$form.Controls.Add($approvalLabel)

$approvalBox = New-Object System.Windows.Forms.TextBox
$approvalBox.Location = New-Object System.Drawing.Point(220, 594)
$approvalBox.Size = New-Object System.Drawing.Size(220, 30)
$approvalBox.Anchor = "Bottom,Left"
$approvalBox.Enabled = $false
$form.Controls.Add($approvalBox)

$recheckButton = New-Object System.Windows.Forms.Button
$recheckButton.Text = "重新验收修改版"
$recheckButton.Location = New-Object System.Drawing.Point(450, 591)
$recheckButton.Size = New-Object System.Drawing.Size(145, 36)
$recheckButton.Anchor = "Bottom,Left"
$recheckButton.Enabled = $false
$form.Controls.Add($recheckButton)

$packageButton = New-Object System.Windows.Forms.Button
$packageButton.Text = "打包通过项目"
$packageButton.Location = New-Object System.Drawing.Point(605, 591)
$packageButton.Size = New-Object System.Drawing.Size(165, 36)
$packageButton.Anchor = "Bottom,Left"
$packageButton.Enabled = $false
$form.Controls.Add($packageButton)

$hint = New-Object System.Windows.Forms.Label
$hint.Text = "例如：1,3-5 或 all"
$hint.ForeColor = [System.Drawing.Color]::DimGray
$hint.Location = New-Object System.Drawing.Point(780, 635)
$hint.AutoSize = $true
$hint.Anchor = "Bottom,Left"
$form.Controls.Add($hint)

$script:process = $null
$script:stdoutPath = $null
$script:stderrPath = $null
$script:batchRoot = $null
$script:reviewPath = $null
$script:operation = $null
$script:archiveFormat = "auto"
$script:workDir = $null
$script:reviewStage = $null
$script:fromExtracted = $false
$script:exportWorkDir = $null
$script:workflowOptionMetadata = [ordered]@{
    match_title_body = @{ Group = "匹配度检测"; Label = "标题与说明书正文主题匹配"; Tip = "逐件输出 0–100 主题相似度与判定依据；默认资格筛选会运行。" }
    match_software_name = @{ Group = "匹配度检测"; Label = "软件名称一致性检查（本地）"; Tip = "检查正文、表格和章节标题中的其他软件名称；本地运行，不调用模型，只提示人工复核。" }
    match_manual_ocr = @{ Group = "匹配度检测"; Label = "截图 OCR 与软件主题匹配"; Tip = "抽样识别说明书界面图并检查与项目主题的相关性。" }
    match_manual_vision = @{ Group = "匹配度检测"; Label = "视觉模型复核截图"; Tip = "只作为相关/不相关/无法判断的人工复核证据；自动依赖 OCR 候选筛选。" }
    match_code = @{ Group = "匹配度检测"; Label = "代码 DOCX/PDF 配套检查"; Tip = "检查同名配套文件和代码文档中的中文注释候选。" }
    match_txt = @{ Group = "匹配度检测"; Label = "TXT 名称、字段与主题检查"; Tip = "检查编码、字段、软件名称、主要功能及主题冲突。" }
    manual_body_years = @{ Group = "说明书内容处理"; Label = "检查/修正说明书正文年份"; Tip = "按既有规则处理高置信时效语境中的正文年份。" }
    manual_image_dates = @{ Group = "说明书内容处理"; Label = "检查/修正说明书图片日期"; Tip = "只更新早于目标年份的有效日期年份；自动运行 OCR 相关性预检。" }
    manual_company_text = @{ Group = "说明书内容处理"; Label = "清除说明书图片中的公司名称（OCR）"; Tip = '扫描说明书内嵌图片，清除 OCR 识别行末以公司结尾的文字；仅修改输出副本，并在报告中提供原图与处理图对比。' }
    manual_ai_hollow_lines = @{ Group = "说明书页面整理"; Label = "检测空心线"; Tip = "检查 DOCX 中参考图所示的横向细长空心矩形；报告会单列说明书名称，并提示将该说明书发给其他人。只检测、不修改文档。" }
    manual_login_module = @{ Group = "说明书页面整理"; Label = "整理登录模块"; Tip = "同时识别 Word 标题和普通文本标题；在第二个正文一级标题下新增登录页小节，保留原小节并顺延编号。无法可靠定位的文档报人工处理，自动导出配套 PDF。" }
    manual_cover = @{ Group = "说明书页面整理"; Label = "整理说明书封面"; Tip = "检查并整理封面及其与正文的分页边界。" }
    manual_header = @{ Group = "说明书页面整理"; Label = "整理说明书页眉"; Tip = "单独处理页眉内容，不自动补页码。" }
    manual_page_numbers = @{ Group = "说明书页面整理"; Label = "整理说明书页码"; Tip = "单独检查和补齐动态页码，不自动整理页眉。" }
    manual_pdf = @{ Group = "说明书页面整理"; Label = "导出说明书 PDF"; Tip = "在所选说明书修改全部完成后导出配套 PDF。" }
    code_cleanup = @{ Group = "代码处理"; Label = "清理代码中文注释"; Tip = "仅清理识别出的中文注释。" }
    code_pdf = @{ Group = "代码处理"; Label = "导出代码 PDF"; Tip = "代码 DOCX 清理后重新导出同名 PDF。" }
    txt_repair = @{ Group = "TXT 处理"; Label = "修复 TXT 主要功能字段"; Tip = "保留原编码和其他字段，只清理主要功能字段中的既定格式问题。" }
}
$script:workflowOptionGroups = [ordered]@{
    "匹配度检测" = @("match_title_body", "match_software_name", "match_manual_ocr", "match_manual_vision", "match_code", "match_txt")
    "说明书内容处理" = @("manual_body_years", "manual_image_dates", "manual_company_text")
    "说明书页面整理" = @("manual_ai_hollow_lines", "manual_login_module", "manual_cover", "manual_header", "manual_page_numbers", "manual_pdf")
    "代码处理" = @("code_cleanup", "code_pdf")
    "TXT 处理" = @("txt_repair")
}
$script:allWorkflowStages = @()
foreach ($stageId in $script:workflowOptionMetadata.Keys) {
    $script:allWorkflowStages += [string]$stageId
}
$script:documentStages = @($script:allWorkflowStages)
$script:hasUnsavedLaunchStages = -not [string]::IsNullOrWhiteSpace($launchDocumentStages)
if (-not [string]::IsNullOrWhiteSpace($launchDocumentStages)) {
    # 兼容旧启动窗口已传出的空格分隔参数；新窗口统一传逗号分隔列表。
    $requestedStages = @($launchDocumentStages -split '[,\s]+' | Where-Object { $_ })
    $unknownStages = @($requestedStages | Where-Object { -not $script:workflowOptionMetadata.Contains($_) })
    if ($unknownStages.Count -gt 0) { throw "未知材料处理步骤：$($unknownStages -join ",")" }
    $script:documentStages = @($script:allWorkflowStages | Where-Object { $requestedStages -contains $_ })
}
if ($AutoStart) {
    if ($IncludeTopicMismatch) { $skipTopicMismatchCheck.Checked = $false }
}

function Get-SelectedDocumentStages {
    return [string]::Join(",", [string[]]$script:documentStages)
}

function Get-DocumentStageSummary($Stages) {
    if ($null -eq $Stages) {
        $selected = @()
    } elseif ($Stages -is [string]) {
        $selected = @($Stages -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    } else {
        $selected = @($Stages | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
    }
    $matchCount = @($selected | Where-Object { $_ -like "match_*" }).Count
    $manualCount = @($selected | Where-Object { $_ -like "manual_*" }).Count
    $codeCount = @($selected | Where-Object { $_ -like "code_*" }).Count
    $txtCount = @($selected | Where-Object { $_ -like "txt_*" }).Count
    return "匹配检测 $matchCount 项；说明书 $manualCount 项；代码 $codeCount 项；TXT $txtCount 项"
}

function Update-DocumentWorkflowSummary {
    $stageSelectionHint.Text = Get-DocumentStageSummary (Get-SelectedDocumentStages)
}

function New-WorkflowOptionCheckBox([System.Windows.Forms.Control]$Parent, [string]$OptionId, [int]$Y, [hashtable]$Controls, [System.Windows.Forms.ToolTip]$Tips) {
    $metadata = $script:workflowOptionMetadata[$OptionId]
    $check = New-Object System.Windows.Forms.CheckBox
    $check.Text = [string]$metadata.Label
    $check.Location = New-Object System.Drawing.Point(16, $Y)
    $check.Size = New-Object System.Drawing.Size(640, 25)
    $check.Checked = $script:documentStages -contains $OptionId
    $check.Anchor = "Top,Left,Right"
    $Tips.SetToolTip($check, [string]$metadata.Tip)
    $Parent.Controls.Add($check)
    $Controls[$OptionId] = $check
}

function Read-JevSettings {
    $configPath = Join-Path $projectRoot "title_body_config.json"
    $sourcePath = $configPath
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        $sourcePath = Join-Path $projectRoot "title_body_config.example.json"
        if (-not (Test-Path -LiteralPath $sourcePath)) {
            $sourcePath = Join-Path $projectRoot "scripts\title_body_config.example.json"
        }
    }
    $config = Get-Content -LiteralPath $sourcePath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($null -eq $config -or $config -isnot [System.Management.Automation.PSCustomObject]) {
        throw "Jev 配置必须是 JSON 对象，请检查：$sourcePath"
    }
    $threshold = [decimal]70
    if ($null -ne $config.PSObject.Properties["score_threshold"]) {
        $rawThreshold = $config.score_threshold
        $text = [Convert]::ToString($rawThreshold, [Globalization.CultureInfo]::InvariantCulture)
        if ($rawThreshold -is [bool] -or -not [decimal]::TryParse($text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$threshold) -or $threshold -lt 0 -or $threshold -gt 100) {
            throw "Jev 通过线必须是 0～100 的数值，请检查：$sourcePath"
        }
    }
    return [PSCustomObject]@{ Path = $configPath; Config = $config; Threshold = $threshold }
}

function Write-JevSettings($Settings) {
    $json = $Settings.Config | ConvertTo-Json -Depth 100
    $tempPath = $Settings.Path + "." + [Guid]::NewGuid().ToString("N") + ".tmp"
    $backupPath = $Settings.Path + "." + [Guid]::NewGuid().ToString("N") + ".bak"
    try {
        [IO.File]::WriteAllText($tempPath, $json + [Environment]::NewLine, (New-Object System.Text.UTF8Encoding($false)))
        if ([IO.File]::Exists($Settings.Path)) {
            [IO.File]::Replace($tempPath, $Settings.Path, $backupPath)
            try { [IO.File]::Delete($backupPath) } catch {}
        } else {
            [IO.File]::Move($tempPath, $Settings.Path)
        }
    } finally {
        if ([IO.File]::Exists($tempPath)) { [IO.File]::Delete($tempPath) }
    }
}

function Save-OneClickPreference([bool]$Enabled) {
    # 单独更新此偏好，避免覆盖同时保存的评分线和步骤配置。
    $settings = Read-JevSettings
    $settings.Config | Add-Member -MemberType NoteProperty -Name one_click_processing -Value $Enabled -Force
    Write-JevSettings $settings
}

function Save-WorkflowSettings([decimal]$Threshold, [string[]]$Stages) {
    if ($Threshold -lt 0 -or $Threshold -gt 100) { throw "Jev 通过线必须在 0～100 之间。" }
    # 保存时重新读取配置，保留评分模式、请求参数及其他设置。
    $settings = Read-JevSettings
    $settings.Config | Add-Member -MemberType NoteProperty -Name score_threshold -Value ([double]$Threshold) -Force
    $settings.Config | Add-Member -MemberType NoteProperty -Name document_stages -Value ([string[]]@($Stages)) -Force
    $settings.Config | Add-Member -MemberType NoteProperty -Name document_stages_version -Value 2 -Force
    Write-JevSettings $settings
}

function Get-SavedWorkflowStages($Settings) {
    $savedStages = $Settings.Config.PSObject.Properties["document_stages"]
    if ($null -eq $savedStages -or $savedStages.Value -isnot [array]) { return @() }
    $knownStages = @($script:allWorkflowStages | Where-Object { $savedStages.Value -contains $_ })
    $versionProperty = $Settings.Config.PSObject.Properties["document_stages_version"]
    $version = if ($null -ne $versionProperty) { [int]$versionProperty.Value } else { 0 }
    if ($version -lt 2) {
        if ($knownStages -notcontains "manual_company_text") { $knownStages += "manual_company_text" }
        $Settings.Config | Add-Member -MemberType NoteProperty -Name document_stages -Value ([string[]]$knownStages) -Force
        $Settings.Config | Add-Member -MemberType NoteProperty -Name document_stages_version -Value 2 -Force
        try { Write-JevSettings $Settings } catch { }
    }
    return [string[]]$knownStages
}

try {
    $savedOneClick = (Read-JevSettings).Config.PSObject.Properties["one_click_processing"]
    if ($null -ne $savedOneClick -and $savedOneClick.Value -is [bool]) {
        $oneClickProcessingCheck.Checked = [bool]$savedOneClick.Value
    }
} catch {
    # 配置读取失败时沿用默认的预览确认流程。
}
$oneClickProcessingCheck.Add_CheckedChanged({
    try {
        Save-OneClickPreference $oneClickProcessingCheck.Checked
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show($form, ("一键处理偏好未能保存：" + $_.Exception.Message), "设置保存失败", "OK", "Error")
    }
})

if ([string]::IsNullOrWhiteSpace($launchDocumentStages)) {
    try {
        $knownStages = @(Get-SavedWorkflowStages (Read-JevSettings))
        if ($knownStages.Count -gt 0) {
            $script:documentStages = $knownStages
        }
    } catch {
        # 配置错误由自定义流程弹窗提示；主窗口先沿用完整处理默认项。
    }
}

function Show-DocumentWorkflowOptions {
    try {
        $jevSettings = Read-JevSettings
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show($form, ("无法读取 Jev 通过线：" + $_.Exception.Message), "配置读取失败", "OK", "Error")
        return
    }
    if (-not $script:hasUnsavedLaunchStages) {
        $knownStages = @(Get-SavedWorkflowStages $jevSettings)
        if ($knownStages.Count -gt 0) { $script:documentStages = $knownStages }
    }
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "处理说明和代码 - 自定义流程"
    $dialog.StartPosition = "CenterParent"
    $dialog.Size = New-Object System.Drawing.Size(720, 680)
    $dialog.MinimumSize = New-Object System.Drawing.Size(680, 620)
    $dialog.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $note = New-Object System.Windows.Forms.Label
    $note.Text = "可多选。系统按固定顺序执行；依赖步骤会自动加入并在预览报告列出。"
    $note.Location = New-Object System.Drawing.Point(14, 12)
    $note.Size = New-Object System.Drawing.Size(675, 24)
    $dialog.Controls.Add($note)

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Location = New-Object System.Drawing.Point(12, 42)
    $tabs.Size = New-Object System.Drawing.Size(680, 500)
    $tabs.Anchor = "Top,Left,Right,Bottom"
    $dialog.Controls.Add($tabs)
    $controls = @{}
    $tips = New-Object System.Windows.Forms.ToolTip
    $tips.AutoPopDelay = 12000
    $tips.InitialDelay = 350
    $tips.ReshowDelay = 100

    $matchTab = New-Object System.Windows.Forms.TabPage
    $matchTab.Text = "匹配检测"
    $matchPanel = New-Object System.Windows.Forms.Panel
    $matchPanel.Dock = "Fill"
    $matchPanel.AutoScroll = $true
    $matchTab.Controls.Add($matchPanel)
    $tabs.TabPages.Add($matchTab)
    $y = 18
    foreach ($id in $script:workflowOptionGroups["匹配度检测"]) {
        New-WorkflowOptionCheckBox $matchPanel $id $y $controls $tips
        $y += 42
    }

    $thresholdLabel = New-Object System.Windows.Forms.Label
    $thresholdLabel.Text = "Jev 通过线："
    $thresholdLabel.Location = New-Object System.Drawing.Point(18, ($y + 8))
    $thresholdLabel.Size = New-Object System.Drawing.Size(125, 26)
    $matchPanel.Controls.Add($thresholdLabel)
    $thresholdInput = New-Object System.Windows.Forms.NumericUpDown
    $thresholdInput.Minimum = 0
    $thresholdInput.Maximum = 100
    $thresholdInput.DecimalPlaces = 2
    $thresholdInput.Increment = 1
    $thresholdInput.Value = $jevSettings.Threshold
    $thresholdInput.Location = New-Object System.Drawing.Point(150, ($y + 6))
    $thresholdInput.Size = New-Object System.Drawing.Size(100, 28)
    $matchPanel.Controls.Add($thresholdInput)
    $thresholdNote = New-Object System.Windows.Forms.Label
    $thresholdNote.Text = "0～100 分；达到通过线即通过。保存后下次打开仍使用此值。"
    $thresholdNote.Location = New-Object System.Drawing.Point(18, ($y + 42))
    $thresholdNote.Size = New-Object System.Drawing.Size(615, 40)
    $thresholdNote.Anchor = "Top,Left,Right"
    $matchPanel.Controls.Add($thresholdNote)

    $manualTab = New-Object System.Windows.Forms.TabPage
    $manualTab.Text = "说明书"
    $manualPanel = New-Object System.Windows.Forms.Panel
    $manualPanel.Dock = "Fill"
    $manualPanel.AutoScroll = $true
    $manualTab.Controls.Add($manualPanel)
    $tabs.TabPages.Add($manualTab)
    $y = 14
    foreach ($groupName in @("说明书内容处理", "说明书页面整理")) {
        $ids = @($script:workflowOptionGroups[$groupName])
        $box = New-Object System.Windows.Forms.GroupBox
        $box.Text = $groupName
        $box.Location = New-Object System.Drawing.Point(12, $y)
        $box.Size = New-Object System.Drawing.Size(630, (35 + 42 * $ids.Count))
        $box.Anchor = "Top,Left,Right"
        $manualPanel.Controls.Add($box)
        for ($index = 0; $index -lt $ids.Count; $index++) {
            New-WorkflowOptionCheckBox $box $ids[$index] (22 + 42 * $index) $controls $tips
        }
        $y += $box.Height + 10
    }

    $otherTab = New-Object System.Windows.Forms.TabPage
    $otherTab.Text = "代码与 TXT"
    $otherPanel = New-Object System.Windows.Forms.Panel
    $otherPanel.Dock = "Fill"
    $otherPanel.AutoScroll = $true
    $otherTab.Controls.Add($otherPanel)
    $tabs.TabPages.Add($otherTab)
    $y = 14
    foreach ($groupName in @("代码处理", "TXT 处理")) {
        $ids = @($script:workflowOptionGroups[$groupName])
        $box = New-Object System.Windows.Forms.GroupBox
        $box.Text = $groupName
        $box.Location = New-Object System.Drawing.Point(12, $y)
        $box.Size = New-Object System.Drawing.Size(630, (35 + 42 * $ids.Count))
        $box.Anchor = "Top,Left,Right"
        $otherPanel.Controls.Add($box)
        for ($index = 0; $index -lt $ids.Count; $index++) {
            New-WorkflowOptionCheckBox $box $ids[$index] (22 + 42 * $index) $controls $tips
        }
        $y += $box.Height + 10
    }

    $completeButton = New-Object System.Windows.Forms.Button
    $completeButton.Text = "完整处理"
    $completeButton.Location = New-Object System.Drawing.Point(12, 552)
    $completeButton.Size = New-Object System.Drawing.Size(95, 30)
    $dialog.Controls.Add($completeButton)
    $checksButton = New-Object System.Windows.Forms.Button
    $checksButton.Text = "只做检测"
    $checksButton.Location = New-Object System.Drawing.Point(114, 552)
    $checksButton.Size = New-Object System.Drawing.Size(95, 30)
    $dialog.Controls.Add($checksButton)
    $manualButton = New-Object System.Windows.Forms.Button
    $manualButton.Text = "说明书"
    $manualButton.Location = New-Object System.Drawing.Point(216, 552)
    $manualButton.Size = New-Object System.Drawing.Size(85, 30)
    $dialog.Controls.Add($manualButton)
    $codeButton = New-Object System.Windows.Forms.Button
    $codeButton.Text = "代码"
    $codeButton.Location = New-Object System.Drawing.Point(308, 552)
    $codeButton.Size = New-Object System.Drawing.Size(75, 30)
    $dialog.Controls.Add($codeButton)
    $txtButton = New-Object System.Windows.Forms.Button
    $txtButton.Text = "TXT"
    $txtButton.Location = New-Object System.Drawing.Point(390, 552)
    $txtButton.Size = New-Object System.Drawing.Size(75, 30)
    $dialog.Controls.Add($txtButton)
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "保存设置"
    $okButton.Location = New-Object System.Drawing.Point(486, 552)
    $okButton.Size = New-Object System.Drawing.Size(95, 30)
    $dialog.Controls.Add($okButton)
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "取消"
    $cancelButton.Location = New-Object System.Drawing.Point(588, 552)
    $cancelButton.Size = New-Object System.Drawing.Size(80, 30)
    $dialog.Controls.Add($cancelButton)

    # GetNewClosure 的事件脚本块有独立的 script 作用域；先在此处固定步骤清单。
    $allStages = [string[]]@($script:allWorkflowStages)
    $checkStages = [string[]]@($allStages | Where-Object { $_ -like "match_*" })
    $manualStages = [string[]]@($allStages | Where-Object { $_ -in @("match_title_body", "match_software_name", "match_manual_ocr", "match_manual_vision") -or $_ -like "manual_*" })
    $setPreset = {
        param([string[]]$Ids)
        foreach ($key in $controls.Keys) { $controls[$key].Checked = $Ids -contains $key }
    }
    $completeButton.Add_Click(({ & $setPreset $allStages }).GetNewClosure())
    $checksButton.Add_Click(({ & $setPreset $checkStages }).GetNewClosure())
    $manualButton.Add_Click(({ & $setPreset $manualStages }).GetNewClosure())
    $codeButton.Add_Click(({ & $setPreset @("match_code", "code_cleanup", "code_pdf") }).GetNewClosure())
    $txtButton.Add_Click(({ & $setPreset @("match_txt", "txt_repair") }).GetNewClosure())
    $okButton.Add_Click(({
        try {
            [void]$dialog.ValidateChildren()
            $selectedStages = [string[]]@($allStages | Where-Object { $controls[$_].Checked })
            if ($selectedStages.Count -eq 0) { throw "请至少勾选一个处理步骤。" }
            Save-WorkflowSettings $thresholdInput.Value $selectedStages
            $dialog.Tag = [string[]]$selectedStages
        } catch {
            [void][System.Windows.Forms.MessageBox]::Show($dialog, ("自定义流程未保存：" + $_.Exception.Message), "保存失败", "OK", "Error")
            return
        }
        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dialog.Close()
    }).GetNewClosure())
    $cancelButton.Add_Click(({ $dialog.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $dialog.Close() }).GetNewClosure())
    $dialog.AcceptButton = $okButton
    $dialog.CancelButton = $cancelButton
    if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:documentStages = [string[]]$dialog.Tag
        $script:hasUnsavedLaunchStages = $false
        Update-DocumentWorkflowSummary
    }
    $tips.Dispose()
    $dialog.Dispose()
}

Update-DocumentWorkflowSummary

function Append-Log([string]$Text) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return }
    $logBox.AppendText($Text.TrimEnd() + [Environment]::NewLine)
    $logBox.SelectionStart = $logBox.TextLength
    $logBox.ScrollToCaret()
}

function Set-WorkflowGuide([string]$Stage) {
    switch ($Stage) {
        "images" {
            $workflowGuide.Text = "工作流：1 选择材料  >  【2 首轮看图并微调 Word】  >  3 二次处理  >  4 终审  >  5 导出 PDF  >  6 压缩交付"
            $workflowGuide.ForeColor = [System.Drawing.Color]::FromArgb(49, 82, 138)
        }
        "processing" {
            $workflowGuide.Text = "工作流：1 选择材料  >  2 首轮看图并微调 Word  >  【3 正在二次处理】  >  4 终审  >  5 导出 PDF  >  6 压缩交付"
            $workflowGuide.ForeColor = [System.Drawing.Color]::FromArgb(154, 91, 0)
        }
        "final" {
            $workflowGuide.Text = "工作流：1 选择材料  >  2 首轮图审  >  3 二次处理  >  【4 终审：复验或直接导出】  >  5 导出 PDF  >  6 压缩交付"
            $workflowGuide.ForeColor = [System.Drawing.Color]::FromArgb(8, 116, 67)
        }
        "export" {
            $workflowGuide.Text = "工作流：1 选择材料  >  2 首轮图审  >  3 二次处理  >  4 终审  >  【5 导出修改版 Word 为 PDF】  >  6 压缩交付"
            $workflowGuide.ForeColor = [System.Drawing.Color]::FromArgb(154, 91, 0)
        }
        "package" {
            $workflowGuide.Text = "工作流：1 选择材料  >  2 首轮图审  >  3 二次处理  >  4 终审  >  5 导出 PDF  >  【6 压缩交付】"
            $workflowGuide.ForeColor = [System.Drawing.Color]::FromArgb(8, 116, 67)
        }
        default {
            $workflowGuide.Text = "工作流：1 选择材料  >  2 首轮看图并微调 Word  >  3 二次处理  >  4 终审  >  5 导出 PDF  >  6 压缩交付"
            $workflowGuide.ForeColor = [System.Drawing.Color]::FromArgb(49, 82, 138)
        }
    }
}

function Set-Busy([bool]$Busy, [string]$Status) {
    $startButton.Enabled = -not $Busy
    $fromSecondButton.Enabled = -not $Busy
    $continueExtractedButton.Enabled = -not $Busy
    $archiveButton.Enabled = -not $Busy
    $folderButton.Enabled = -not $Busy
    $refurbishButton.Enabled = -not $Busy
    $orderDownloadButton.Enabled = -not $Busy
    $codeCommentButton.Enabled = -not $Busy
    $oneClickProcessingCheck.Enabled = -not $Busy
    $stageChecksEnabled = (-not $Busy) -and ($script:reviewStage -ne "documents")
    $workflowOptionsButton.Enabled = $stageChecksEnabled
    $skipTopicMismatchCheck.Enabled = (-not $Busy) -and ($script:reviewStage -ne "documents")
    $visionSettingsButton.Enabled = -not $Busy
    $exportButton.Enabled = (-not $Busy) -and (-not [string]::IsNullOrWhiteSpace($inputBox.Text))
    $packageButton.Enabled = (-not $Busy) -and ($null -ne $script:reviewPath)
    $recheckButton.Enabled = (-not $Busy) -and ($script:reviewStage -eq "final")
    if ($Busy) {
        $progress.Style = "Continuous"
        $progress.MarqueeAnimationSpeed = 0
        $progress.Maximum = 100
        $progress.Value = 0
        $progressDetail.Text = "正在准备：$Status"
    } else {
        $progress.Style = "Continuous"
        $progress.MarqueeAnimationSpeed = 0
        $progress.Maximum = 100
        $progress.Value = 100
    }
    $statusLabel.Text = $Status
    if (-not $Busy -and $script:reviewStage -in @("documents", "documents-done")) {
        $progressDetail.Text = $Status
    }
}

function Update-ExportAvailability {
    $value = $inputBox.Text.Trim()
    $exportButton.Enabled = ($null -eq $script:process -or $script:process.HasExited) -and (-not [string]::IsNullOrWhiteSpace($value)) -and (Test-Path -LiteralPath $value)
}

function Get-MainWorkDir {
    # Derive the processing work directory from the current batch root so that
    # packaging / re-verification never pick up a "-export" directory that the
    # PDF export step uses for its report.
    $bytes = [Text.Encoding]::UTF8.GetBytes($script:batchRoot.ToLowerInvariant())
    $hash = [Security.Cryptography.SHA1]::Create().ComputeHash($bytes)
    $key = ((-join ($hash | ForEach-Object { $_.ToString("x2") })).Substring(0, 16))
    $suffix = if ($script:fromExtracted) { "-from-extracted" } else { "" }
    return Join-Path ([IO.Path]::GetTempPath()) ("codex-softcopyright\" + $key + $suffix)
}

function Resolve-ModifiedDirectory([string]$BatchRoot, [string]$Source = "") {
    # 定位修改版目录：优先返回真实存在的目录，不把已解压目录命名为
    # “解压版”作为前提；名称约定只作为最后的兜底猜测。
    $candidates = @()
    if ($Source -and (Test-Path -LiteralPath $Source -PathType Container)) {
        $item = Get-Item -LiteralPath $Source
        if ($item.Name.EndsWith("修改版")) { $candidates += $item.FullName }
        $name = $item.Name
        if ($name.EndsWith("_解压版")) { $candidates += (Join-Path $item.Parent.FullName ($name.Substring(0, $name.Length - 4) + "_修改版")) }
        elseif ($name.EndsWith("解压版")) { $candidates += (Join-Path $item.Parent.FullName ($name.Substring(0, $name.Length - 3) + "修改版")) }
        $candidates += (Join-Path $item.Parent.FullName ($item.BaseName + "_修改版"))
        $candidates += (Join-Path $item.FullName "修改版")
    }
    $candidates += (Join-Path $BatchRoot "修改版")
    $candidates += (Join-Path (Split-Path -Parent $BatchRoot) (([IO.Path]::GetFileName($BatchRoot.TrimEnd('\', '/'))) + "_修改版"))
    $candidates += (Join-Path (Split-Path -Parent $BatchRoot) "修改版")
    $ordered = @($candidates | Where-Object { $_ } | Select-Object -Unique)
    foreach ($candidate in $ordered) {
        if (Test-Path -LiteralPath $candidate -PathType Container -ErrorAction SilentlyContinue) {
            if (Get-ChildItem -LiteralPath $candidate -Force -ErrorAction SilentlyContinue | Select-Object -First 1) { return $candidate }
        }
    }
    return ($ordered | Select-Object -First 1)
}

function Get-ActiveModifiedDirectory {
    $manifestPath = if ($null -ne $script:workDir) { Join-Path $script:workDir "unified_review.json" } else { $null }
    if ($manifestPath -and (Test-Path -LiteralPath $manifestPath)) {
        try {
            $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
            if (-not [string]::IsNullOrWhiteSpace([string]$manifest.modified) -and (Test-Path -LiteralPath ([string]$manifest.modified))) {
                return [string]$manifest.modified
            }
        } catch { }
    }
    $source = Get-Item -LiteralPath $script:batchRoot -ErrorAction SilentlyContinue
    if ($null -eq $source) { return $null }
    if (-not $source.PSIsContainer) { return (Join-Path $source.DirectoryName ($source.BaseName + "_修改版")) }
    $batchRoot = if ($script:fromExtracted) { $source.Parent.FullName } else { $source.FullName }
    return (Resolve-ModifiedDirectory $batchRoot $source.FullName)
}

function Confirm-PdfFreshness {
    $modified = Get-ActiveModifiedDirectory
    if ([string]::IsNullOrWhiteSpace($modified) -or -not (Test-Path -LiteralPath $modified -PathType Container)) { return $true }
    $stale = @()
    foreach ($docx in Get-ChildItem -LiteralPath $modified -Recurse -File -Filter "*.docx" -ErrorAction SilentlyContinue) {
        $pdf = [IO.Path]::ChangeExtension($docx.FullName, ".pdf")
        if (-not (Test-Path -LiteralPath $pdf) -or ((Get-Item -LiteralPath $pdf).LastWriteTimeUtc -lt $docx.LastWriteTimeUtc)) {
            $stale += (Join-Path $docx.Directory.Name $docx.Name)
        }
    }
    if ($stale.Count -eq 0) { return $true }
    $shown = ($stale | Select-Object -First 6) -join "`n"
    $more = if ($stale.Count -gt 6) { "`n另有 $($stale.Count - 6) 个文件。" } else { "" }
    $message = "发现 $($stale.Count) 个修改版 Word 尚未导出为最新 PDF：`n`n$shown$more`n`n可先点击直接导出 PDF；仍可选择继续压缩。"
    return ([System.Windows.Forms.MessageBox]::Show($message, "PDF 导出提醒", "YesNo", "Warning") -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Read-Result([string]$Text) {
    $result = $null
    foreach ($line in @($Text -split "`r?`n" | Select-Object -Last 30)) {
        try {
            $candidate = $line | ConvertFrom-Json -ErrorAction Stop
            if ($null -ne $candidate.status) { $result = $candidate }
        } catch { }
    }
    return $result
}

function Start-Operation([string[]]$Arguments, [string]$Operation) {
    if ($null -ne $script:process -and -not $script:process.HasExited) {
        Append-Log "已有任务正在运行，忽略重复启动。"
        return
    }
    if ($script:isOptimizationDev) {
        try {
            if ($Arguments.Count -lt 1) { throw "优化开发版隔离拒绝缺少材料路径的操作。" }
            Assert-OptimizationPath $Arguments[0] -Material
            for ($index = 0; $index -lt ($Arguments.Count - 1); $index++) {
                switch ($Arguments[$index]) {
                    "--work-dir" { Assert-OptimizationPath $Arguments[$index + 1] }
                    "--extracted" { Assert-OptimizationPath $Arguments[$index + 1] -Material }
                    "--modified" { Assert-OptimizationPath $Arguments[$index + 1] -Material }
                    "--input-archive" { Assert-OptimizationPath $Arguments[$index + 1] -Material }
                }
            }
        } catch {
            Append-Log $_.Exception.Message
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "优化开发版隔离", "OK", "Warning") | Out-Null
            return
        }
    }
    $script:stdoutPath = [IO.Path]::GetTempFileName()
    $script:stderrPath = [IO.Path]::GetTempFileName()
    $argumentText = (@(Quote-Argument $entry) + @($Arguments | ForEach-Object { Quote-Argument $_ })) -join " "
    Append-Log ("> " + $python + " " + $argumentText)
    $script:operation = $Operation
    $script:process = Start-Process -FilePath $python -ArgumentList $argumentText -PassThru -WindowStyle Hidden -RedirectStandardOutput $script:stdoutPath -RedirectStandardError $script:stderrPath
    Set-Busy $true $Operation
    $timer.Start()
}

function Read-ProgressState([string]$Path) {
    $share = [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete
    $stream = [IO.FileStream]::new($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, $share)
    try {
        $reader = [IO.StreamReader]::new($stream, [Text.Encoding]::UTF8, $true)
        try { return ($reader.ReadToEnd() | ConvertFrom-Json -ErrorAction Stop) }
        finally { $reader.Dispose() }
    } finally { $stream.Dispose() }
}

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 750
$timer.Add_Tick({
    if ($null -ne $script:workDir) {
        $progressPath = Join-Path $script:workDir "progress.json"
        if (Test-Path -LiteralPath $progressPath) {
            try {
                $state = Read-ProgressState $progressPath
                $elapsed = [math]::Round([double]$state.elapsed_seconds, 0)
                $project = [string]$state.project_name
                if ($state.total -gt 0) {
                    $moduleTotal = [int]$state.total
                    $planName = if ([string]$state.current_module -eq "说明书处理") { "repair_plan_selected.json" } elseif ([string]$state.current_module -eq "代码处理") { "code_repair_plan_selected.json" } else { "" }
                    if ($planName) {
                        $planPath = Join-Path $script:workDir (Join-Path "audit" $planName)
                        if (Test-Path -LiteralPath $planPath) {
                            try {
                                $plan = Get-Content -LiteralPath $planPath -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
                                $planned = @($plan.packages).Count
                                if ($planned -gt 0) { $moduleTotal = $planned }
                            } catch { }
                        }
                    }
                    $current = [math]::Min([int]$state.current, $moduleTotal)
                    $percent = [math]::Floor(($current * 100) / $moduleTotal)
                    $active = ($state.current_project_active -eq $true)
                    $displayPhase = switch ([string]$state.phase) {
                        "WORD_PDF_EXPORT" { "Word 导出 PDF" }
                        "EXPORT_COMPLETE" { "Word 导出完成" }
                        "PREPARING" { "准备导出" }
                        default { [string]$state.phase }
                    }
                    # 项目正在处理不等于总数未知；总数已知时始终显示确定型进度。
                    $progress.Style = "Continuous"
                    $progress.MarqueeAnimationSpeed = 0
                    $progress.Maximum = $moduleTotal
                    $progress.Value = $current
                    $displayIndex = [math]::Min($current + 1, $moduleTotal)
                    $statusLabel.Text = if ($active) { "正在处理第 $displayIndex / $moduleTotal 个：$project" } else { "$percent%  $displayPhase（$current / $moduleTotal）" }
                    $done = if ($null -ne $state.completed_projects) { [int]$state.completed_projects } else { $current }
                    $done = [math]::Min($done, $moduleTotal)
                    $left = [math]::Max(0, $moduleTotal - $done)
                    $eta = if ($null -ne $state.estimated_remaining_seconds) { " · 预计剩余 $([math]::Round([double]$state.estimated_remaining_seconds,0)) 秒" } else { "" }
                    $activity = if ($active) { "正在处理：$project" } else { "最近完成：$project" }
                    $failed = if ($null -ne $state.failures) { [int]$state.failures } else { 0 }
                    $archive = if (-not [string]::IsNullOrWhiteSpace([string]$state.archive_name)) { [string]$state.archive_name } else { "-" }
                    $selectionDetail = if ($script:documentStages -and $script:operation -match "所选材料流程") { " · " + (Get-DocumentStageSummary $script:documentStages) } else { "" }
                    $progressDetail.Text = "已处理 $done 件 · 待处理 $left 件 · 失败 $failed 件 · 压缩包：$archive · $activity · 阶段：$displayPhase$eta · 已用 $elapsed 秒$selectionDetail"
                } else {
                    $progress.Style = "Marquee"
                    $progress.MarqueeAnimationSpeed = 30
                    $statusLabel.Text = $displayPhase
                    $selectionDetail = if ($script:documentStages -and $script:operation -match "所选材料流程") { " · " + (Get-DocumentStageSummary $script:documentStages) } else { "" }
                    $progressDetail.Text = "正在准备项目清单 · 模块：$displayPhase · 已用 $elapsed 秒$selectionDetail"
                }
            } catch { }
        }
    }
    if ($null -eq $script:process -or -not $script:process.HasExited) { return }
    $timer.Stop()
    $script:process = $null
    $output = if (Test-Path $script:stdoutPath) { Get-Content -LiteralPath $script:stdoutPath -Raw -Encoding UTF8 } else { "" }
    $errors = if (Test-Path $script:stderrPath) { Get-Content -LiteralPath $script:stderrPath -Raw -Encoding UTF8 } else { "" }
    Append-Log $output
    Append-Log $errors
    $result = Read-Result ($output + "`n" + $errors)
    if ($null -eq $result) {
        Set-Busy $false "Failed: no readable status"
        $orderDownloadButton.BringToFront()
        [System.Windows.Forms.MessageBox]::Show("The process did not return a readable status. See the log.", "EasySoftware", "OK", "Error") | Out-Null
        return
    }
    if ($result.status -eq "NEEDS_IMAGE_REVIEW") {
        $script:reviewPath = [string]$result.review
        $script:reviewStage = "images"
        $reviewButton.Text = "打开首轮图片审核"
        $reviewButton.Enabled = Test-Path -LiteralPath $script:reviewPath
        $approvalBox.Enabled = $true
        $approvalLabel.Text = "图片未通过项目序号："
        $approvalBox.Text = "none"
        $packageButton.Text = "继续处理其余项目"
        $exportButton.Text = "修改版 Word 导出 PDF"
        $recheckButton.Enabled = $false
        $packageButton.Enabled = $true
        $packageButton.BringToFront()
        Set-WorkflowGuide "images"
        Set-Busy $false "等待图片审核：只填不通过项目"
        if ($reviewButton.Enabled) { Start-Process -FilePath $script:reviewPath }
    } elseif ($result.status -eq "NEEDS_CODE_COMMENT_REVIEW") {
        $script:reviewPath = [string]$result.report
        $script:reviewStage = "code-comments"
        $reviewButton.Text = "打开注释清理预览"
        $reviewButton.Enabled = Test-Path -LiteralPath $script:reviewPath
        $approvalBox.Enabled = $false
        $approvalLabel.Text = "代码注释清理："
        $approvalBox.Text = ""
        $packageButton.Text = "确认并执行清理"
        $packageButton.Enabled = $true
        $packageButton.BringToFront()
        Set-Busy $false ("预览完成：预计删除 " + [string]$result.deletions + " 处")
        if ($reviewButton.Enabled) { Start-Process -FilePath $script:reviewPath }
    } elseif ($result.status -eq "NEEDS_DOCUMENT_PROCESSING_REVIEW") {
        $script:reviewPath = [string]$result.report
        $script:reviewStage = "documents"
        $reviewButton.Text = "打开材料处理预览"
        $reviewButton.Enabled = Test-Path -LiteralPath $script:reviewPath
        $approvalBox.Enabled = $false
        $approvalLabel.Text = "材料处理流程："
        $approvalBox.Text = ""
        $packageButton.Text = "确认处理并生成结果"
        $packageButton.Enabled = $true
        $packageButton.BringToFront()
        $nameWarning = ""
        if ($result.PSObject.Properties["name_warning_count"] -and [int]$result.name_warning_count -gt 0) {
            $nameWarning = "；软件名称疑似不一致 " + [string]$result.name_warning_count + " 处，请复核报告"
            Append-Log ("软件名称不一致，请复核：" + (@($result.name_warning_names) -join "、"))
        }
        Set-Busy $false ("预检完成：" + (Get-DocumentStageSummary $script:documentStages) + $nameWarning)
        if ($reviewButton.Enabled) { Start-Process -FilePath $script:reviewPath }
    } elseif ($result.status -in @("DOCUMENT_PROCESSING_OK", "DOCUMENT_PROCESSING_PARTIAL")) {
        $script:reviewPath = [string]$result.report
        $script:reviewStage = "documents-done"
        $reviewButton.Text = "打开材料处理报告"
        $reviewButton.Enabled = Test-Path -LiteralPath $script:reviewPath
        $approvalBox.Enabled = $false
        $orderDownloadButton.BringToFront()
        $caption = if ($result.status -eq "DOCUMENT_PROCESSING_OK") { "材料统一处理完成" } else { "材料处理部分完成，请复核异常" }
        $nameWarning = ""
        if ($result.PSObject.Properties["name_warning_count"] -and [int]$result.name_warning_count -gt 0) {
            $nameWarning = "；软件名称疑似不一致 " + [string]$result.name_warning_count + " 处，请复核报告"
            Append-Log ("软件名称不一致，请复核：" + (@($result.name_warning_names) -join "、"))
        }
        Set-Busy $false ($caption + "：失败/待处理 " + [string]$result.failures + " 项；" + (Get-DocumentStageSummary $script:documentStages) + $nameWarning)
        $packageButton.Enabled = $false
        Append-Log ($caption + "；失败/待处理 " + [string]$result.failures + " 项")
        Append-Log ("输出目录：" + [string]$result.output)
        Append-Log ("处理报告：" + [string]$result.report)
        if ($reviewButton.Enabled) {
            try { Start-Process -FilePath $script:reviewPath -ErrorAction Stop }
            catch { Append-Log ("自动打开报告失败，可点击报告按钮：" + $_.Exception.Message) }
        } else {
            [System.Windows.Forms.MessageBox]::Show(("处理结束，但报告文件不存在：`n" + [string]$result.report), "EasySoftware", "OK", "Warning") | Out-Null
        }
    } elseif ($result.status -eq "CODE_COMMENT_CLEANUP_OK") {
        $script:reviewPath = [string]$result.report
        $script:reviewStage = "code-comments-done"
        $reviewButton.Text = "打开注释清理报告"
        $reviewButton.Enabled = Test-Path -LiteralPath $script:reviewPath
        $packageButton.Enabled = $false
        $approvalBox.Enabled = $false
        Set-Busy $false ("清理完成：删除 " + [string]$result.deletions + " 处")
        [System.Windows.Forms.MessageBox]::Show(("代码中文注释清理完成。`n`n输出目录：`n" + (@($result.outputs) -join "`n") + "`n`n报告：`n" + [string]$result.report), "EasySoftware", "OK", "Information") | Out-Null
    } elseif ($result.status -eq "CODE_COMMENT_CLEANUP_PARTIAL") {
        $script:reviewPath = [string]$result.report
        $script:reviewStage = "code-comments-partial"
        $reviewButton.Text = "打开注释清理异常报告"
        $reviewButton.Enabled = Test-Path -LiteralPath $script:reviewPath
        $packageButton.Enabled = $false
        $approvalBox.Enabled = $false
        Set-Busy $false ("部分完成：有 " + [string]$result.failures + " 个文件未完成")
        [System.Windows.Forms.MessageBox]::Show(("代码注释清理未全部完成，部分 PDF 导出或文件处理失败。旧 PDF 已保留，请打开报告复核。`n`n报告：`n" + [string]$result.report), "EasySoftware", "OK", "Warning") | Out-Null
    } elseif ($result.status -eq "NEEDS_UNIFIED_REVIEW") {
        $script:reviewPath = [string]$result.review
        $script:reviewStage = "final"
        $reviewButton.Text = "打开最终审核页"
        $reviewButton.Enabled = Test-Path -LiteralPath $script:reviewPath
        $approvalBox.Enabled = $true
        $approvalLabel.Text = "最终通过项目序号："
        $approvalBox.Text = ""
        $exportButton.Text = "直接导出 PDF"
        $packageButton.Text = "打包通过项目"
        Set-WorkflowGuide "final"
        $packageButton.Enabled = $true
        $packageButton.BringToFront()
        Set-Busy $false "终审已生成：可重新验收修改版或直接导出 PDF"
        if ($reviewButton.Enabled) { Start-Process -FilePath $script:reviewPath }
    } elseif ($result.status -eq "NEEDS_ARCHIVE_FORMAT") {
        Set-Busy $false "请选择交付格式"
        $choice = [System.Windows.Forms.MessageBox]::Show("批次同时包含 ZIP 和 RAR。`n`n选择【是】生成 ZIP，选择【否】生成 RAR。", "选择交付格式", "YesNoCancel", "Question")
        if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) { $script:archiveFormat = "zip" }
        elseif ($choice -eq [System.Windows.Forms.DialogResult]::No) { $script:archiveFormat = "rar" }
        else { return }
        $approved = $approvalBox.Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($approved)) {
            $arguments = @($script:batchRoot, "--work-dir", $script:workDir, "--approve", $approved, "--format", $script:archiveFormat)
            if ($script:fromExtracted) { $arguments += "--from-extracted" }
            Set-WorkflowGuide "package"
            Start-Operation -Arguments $arguments -Operation "正在打包通过项目..."
        }
    } elseif ($result.status -eq "OK") {
        Set-WorkflowGuide "package"
        Set-Busy $false "Completed"
        $message = "Archive created:`n" + [string]$result.archive + "`n`nSHA-256:`n" + [string]$result.sha256
        [System.Windows.Forms.MessageBox]::Show($message, "EasySoftware", "OK", "Information") | Out-Null
    } elseif ($result.status -eq "EXPORT_OK") {
        $exportButton.Text = "再次导出 PDF"
        Set-WorkflowGuide "package"
        Set-Busy $false "修改版 Word 已导出并覆盖 PDF"
        [System.Windows.Forms.MessageBox]::Show(("已导出：{0} 个文件`n已跳过：{1} 个文件`n失败：{2} 个文件`n`n报告：{3}" -f $result.exported, $result.skipped, $result.failed, (Join-Path ([string]$result.work_dir) "modified_pdf_export_report.json")), "EasySoftware", "OK", "Information") | Out-Null
    } else {
        $orderDownloadButton.BringToFront()
        Set-Busy $false ("Status: " + [string]$result.status)
        $message = [string]$result.status
        if ($null -ne $result.error) { $message += "`n`n" + [string]$result.error }
        if ($null -ne $result.work_dir) { $message += "`n`nWork directory:`n" + [string]$result.work_dir }
        [System.Windows.Forms.MessageBox]::Show($message, "EasySoftware", "OK", "Warning") | Out-Null
    }
})

$archiveButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "选择软著材料压缩包"
    $dialog.Filter = "软著压缩包|*.rar;*.zip;*.7z;*.tar;*.gz;*.tgz;*.tar.gz;*.bz2;*.xz;*.lz4;*.zst|RAR (*.rar)|*.rar|ZIP (*.zip)|*.zip|7z (*.7z)|*.7z|所有文件 (*.*)|*.*"
    $dialog.CheckFileExists = $true
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $inputBox.Text = $dialog.FileName
        $statusLabel.Text = "已选择压缩包，仅处理该文件。"
        Update-ExportAvailability
    }
})

$folderButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "选择软著材料批次文件夹"
    $dialog.ShowNewFolderButton = $false
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $inputBox.Text = $dialog.SelectedPath
        $statusLabel.Text = "已选择批次文件夹。"
        Update-ExportAvailability
    }
})

$inputBox.Add_TextChanged({ Update-ExportAvailability })

$workflowOptionsButton.Add_Click({ Show-DocumentWorkflowOptions })

$visionSettingsButton.Add_Click({
    $configPath = if ($script:isOptimizationDev) { $env:EASYSOFTWARE_VISION_CONFIG } else { Join-Path $projectRoot "vision_config.json" }
    if (-not (Test-Path -LiteralPath $configPath)) {
        $examplePath = Join-Path $projectRoot "vision_config.example.json"
        if (-not (Test-Path -LiteralPath $examplePath)) {
            [System.Windows.Forms.MessageBox]::Show("找不到视觉模型配置模板：$examplePath", "EasySoftware", "OK", "Error") | Out-Null
            return
        }
        Copy-Item -LiteralPath $examplePath -Destination $configPath
    }
    Start-Process -FilePath "notepad.exe" -ArgumentList (Quote-Argument $configPath)
})

function Test-OrderDownloadBatchPath([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return $false }
    $batch = Get-Item -LiteralPath $Path
    if (-not $batch.Name.StartsWith("订单批量_", [StringComparison]::Ordinal)) { return $false }
    foreach ($item in Get-ChildItem -LiteralPath $batch.FullName -Directory) {
        if ($item.Name -notmatch '^item-\d+$') { continue }
        if (Get-ChildItem -LiteralPath $item.FullName -Recurse -File -Filter "*.zip" | Select-Object -First 1) { return $true }
    }
    return $false
}

$codeCommentButton.Add_Click({
    $value = $inputBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($value) -or -not (Test-Path -LiteralPath $value)) {
        [System.Windows.Forms.MessageBox]::Show("请先选择压缩包，或已解压的软著材料文件夹。", "EasySoftware", "OK", "Warning") | Out-Null
        return
    }
    $selectedStages = Get-SelectedDocumentStages
    if ([string]::IsNullOrWhiteSpace($selectedStages)) {
        [System.Windows.Forms.MessageBox]::Show("请至少勾选一个材料处理流程。", "EasySoftware", "OK", "Warning") | Out-Null
        return
    }
    if ($script:documentStages.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("请至少勾选一个检测或处理步骤。", "EasySoftware", "OK", "Warning") | Out-Null
        return
    }
    $script:batchRoot = (Get-Item -LiteralPath $value).FullName
    $bytes = [Text.Encoding]::UTF8.GetBytes($script:batchRoot.ToLowerInvariant())
    $hash = [Security.Cryptography.SHA1]::Create().ComputeHash($bytes)
    $key = ((-join ($hash | ForEach-Object { $_.ToString("x2") })).Substring(0, 16))
    $script:workDir = Join-Path ([IO.Path]::GetTempPath()) ("codex-softcopyright\" + $key + "-documents")
    $script:reviewPath = $null
    $script:reviewStage = $null
    $reviewButton.Enabled = $false
    $packageButton.Enabled = $false
    $approvalBox.Enabled = $false
    $logBox.Clear()
    Append-Log ("材料统一处理预览：" + $script:batchRoot)
    Append-Log (Get-DocumentStageSummary $script:documentStages)
    if ($oneClickProcessingCheck.Checked) {
        $arguments = @($script:batchRoot, "--work-dir", $script:workDir, "--apply-document-processing", "--document-stages", $selectedStages)
        if (-not $skipTopicMismatchCheck.Checked) { $arguments += "--include-topic-mismatch" }
        if ($OrderDownloads -or (Test-OrderDownloadBatchPath $value)) { $arguments += "--expand-order-downloads" }
        Append-Log "一键处理已启用：跳过预览，直接执行所选流程并生成结果。"
        Set-WorkflowGuide "processing"
        Start-Operation -Arguments $arguments -Operation "正在一键处理材料并生成结果..."
    } else {
        $arguments = @($script:batchRoot, "--work-dir", $script:workDir, "--process-documents", "--document-stages", $selectedStages)
        if (-not $skipTopicMismatchCheck.Checked) { $arguments += "--include-topic-mismatch" }
        if ($OrderDownloads -or (Test-OrderDownloadBatchPath $value)) { $arguments += "--expand-order-downloads" }
        Start-Operation -Arguments $arguments -Operation "正在预检所选材料流程..."
    }
})

$exportButton.Add_Click({
    $value = $inputBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($value) -or -not (Test-Path -LiteralPath $value)) {
        [System.Windows.Forms.MessageBox]::Show("请先选择要导出 PDF 的文件夹或压缩包。", "EasySoftware", "OK", "Warning") | Out-Null
        return
    }
    $item = Get-Item -LiteralPath $value
    if (Test-ArchivePath $value) {
        try { $value = Prepare-ArchiveWorkspace $item; $inputBox.Text = $value } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "EasySoftware", "OK", "Error") | Out-Null; return }
    }
    $selected = Get-Item -LiteralPath $value
    # 导出只跟随本次选择的目录，不要求目录名叫“解压版”或“修改版”：程序按
    # 目录内容推导导出位置。这里不改 $script:batchRoot 与 $script:workDir，
    # 打包与复验仍沿用上一次流程的批次根目录；导出报告写入独立的 -export 目录。
    $exportBatchRoot = $selected.FullName
    $bytes = [Text.Encoding]::UTF8.GetBytes($exportBatchRoot.ToLowerInvariant())
    $hash = [Security.Cryptography.SHA1]::Create().ComputeHash($bytes)
    $key = ((-join ($hash | ForEach-Object { $_.ToString("x2") })).Substring(0, 16))
    $script:exportWorkDir = Join-Path ([IO.Path]::GetTempPath()) ("codex-softcopyright\" + $key + "-export")
    $arguments = @($exportBatchRoot, "--work-dir", $script:exportWorkDir, "--export-modified-pdfs")
    if ($script:fromExtracted) { $arguments += "--from-extracted" }
    Set-WorkflowGuide "export"
    Start-Operation -Arguments $arguments -Operation "正在将修改版 Word 导出并覆盖 PDF..."
})

$startButton.Add_Click({
    $value = $inputBox.Text.Trim()
    if (Test-ArchiveName $value) {
        $nested = Join-Path (Join-Path (Split-Path -Parent $value) ([IO.Path]::GetFileNameWithoutExtension($value))) ([IO.Path]::GetFileName($value))
        if (Test-Path -LiteralPath $nested) { $value = $nested; $inputBox.Text = $nested }
    }
    if ([string]::IsNullOrWhiteSpace($value) -or -not (Test-Path -LiteralPath $value)) {
        [System.Windows.Forms.MessageBox]::Show("Select an archive or batch folder first.", "EasySoftware", "OK", "Warning") | Out-Null
        return
    }
    $item = Get-Item -LiteralPath $value
    if (Test-ArchivePath $value) {
        try {
            $prepared = Prepare-ArchiveWorkspace $item
            $item = Get-Item -LiteralPath $prepared
            $inputBox.Text = $item.FullName
            Append-Log ("Archive moved into project folder: " + $item.DirectoryName)
        } catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "EasySoftware", "OK", "Error") | Out-Null
            return
        }
    }
    $script:batchRoot = $item.FullName
    $bytes = [Text.Encoding]::UTF8.GetBytes($script:batchRoot.ToLowerInvariant())
    $hash = [Security.Cryptography.SHA1]::Create().ComputeHash($bytes)
    $key = ((-join ($hash | ForEach-Object { $_.ToString("x2") })).Substring(0, 16))
    $script:workDir = Join-Path ([IO.Path]::GetTempPath()) ("codex-softcopyright\" + $key)
    $script:reviewPath = $null
    $script:reviewStage = $null
    $script:archiveFormat = "auto"
    $script:fromExtracted = $false
    $reviewButton.Enabled = $false
    $recheckButton.Enabled = $false
    $approvalBox.Enabled = $false
    $packageButton.Enabled = $false
    $exportButton.Text = "修改版 Word 导出 PDF"
    Set-WorkflowGuide "default"
    $logBox.Clear()
    Append-Log ("Selected input: " + $value)
    if (-not $item.PSIsContainer) { Append-Log "Only the selected archive will be processed." }
    Start-Operation -Arguments @($script:batchRoot, "--work-dir", $script:workDir) -Operation "正在提取并生成首轮图片审核..."
})

$fromSecondButton.Add_Click({
    $value = $inputBox.Text.Trim()
    if (Test-ArchiveName $value) {
        $nested = Join-Path (Join-Path (Split-Path -Parent $value) ([IO.Path]::GetFileNameWithoutExtension($value))) ([IO.Path]::GetFileName($value))
        if (Test-Path -LiteralPath $nested) { $value = $nested; $inputBox.Text = $nested }
    }
    if ([string]::IsNullOrWhiteSpace($value) -or -not (Test-Path -LiteralPath $value)) {
        [System.Windows.Forms.MessageBox]::Show("Select an archive or batch folder first.", "EasySoftware", "OK", "Warning") | Out-Null
        return
    }
    $item = Get-Item -LiteralPath $value
    if (Test-ArchivePath $value) {
        try {
            $prepared = Prepare-ArchiveWorkspace $item
            $item = Get-Item -LiteralPath $prepared
            $inputBox.Text = $item.FullName
            Append-Log ("Archive moved into project folder: " + $item.DirectoryName)
        } catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "EasySoftware", "OK", "Error") | Out-Null
            return
        }
    }
    $isExtracted = $item.PSIsContainer -and $item.Name.EndsWith("解压版")
    if (-not $isExtracted) {
        # The selected batch may already carry an extracted 解压版 directory
        # (from an earlier run or supplied directly). When the normal work
        # directory never recorded that extraction, use the 解压版 itself as the
        # input so the second round does not fail with "解压版已存在但缺少
        # extraction.json". If it was recorded, keep the normal flow and reuse
        # the existing caches.
        $candidateExtracted = if ($item.PSIsContainer) {
            Join-Path $item.FullName "解压版"
        } else {
            Join-Path $item.DirectoryName ($item.BaseName + "_解压版")
        }
        if ((Test-Path -LiteralPath $candidateExtracted) -and (Get-ChildItem -LiteralPath $candidateExtracted -Force | Select-Object -First 1)) {
            $mainBytes = [Text.Encoding]::UTF8.GetBytes($item.FullName.ToLowerInvariant())
            $mainHash = [Security.Cryptography.SHA1]::Create().ComputeHash($mainBytes)
            $mainKey = ((-join ($mainHash | ForEach-Object { $_.ToString("x2") })).Substring(0, 16))
            $mainWorkDir = Join-Path ([IO.Path]::GetTempPath()) ("codex-softcopyright\" + $mainKey)
            $hasRecordedExtraction = Test-Path -LiteralPath (Join-Path $mainWorkDir "extraction.json")
            if (-not $hasRecordedExtraction) {
                $isExtracted = $true
                $item = Get-Item -LiteralPath $candidateExtracted
                $value = $item.FullName
                Append-Log ("已检测到解压版目录，改用该目录继续：" + $value)
            } else {
                Append-Log "检测到已记录的解压版，继续使用原工作目录缓存。"
            }
        }
    }
    $script:batchRoot = $item.FullName
    $bytes = [Text.Encoding]::UTF8.GetBytes($script:batchRoot.ToLowerInvariant())
    $hash = [Security.Cryptography.SHA1]::Create().ComputeHash($bytes)
    $key = ((-join ($hash | ForEach-Object { $_.ToString("x2") })).Substring(0, 16))
    $suffix = if ($isExtracted) { "-from-extracted" } else { "" }
    $script:workDir = Join-Path ([IO.Path]::GetTempPath()) ("codex-softcopyright\" + $key + $suffix)
    $script:reviewPath = $null
    $script:reviewStage = $null
    $script:archiveFormat = "auto"
    $script:fromExtracted = $isExtracted
    $reviewButton.Enabled = $false
    $recheckButton.Enabled = $false
    $approvalBox.Enabled = $false
    $packageButton.Enabled = $false
    $exportButton.Text = "修改版 Word 导出 PDF"
    Set-WorkflowGuide "processing"
    $logBox.Clear()
    Append-Log ("Selected input: " + $value)
    Append-Log "已跳过首轮图片审核，直接开始二次处理及后续。"
    $arguments = @($script:batchRoot, "--work-dir", $script:workDir, "--reject-images", "none")
    if ($isExtracted) { $arguments += "--from-extracted" }
    Start-Operation -Arguments $arguments -Operation "正在跳过首轮图审，直接二次处理..."
})

$continueExtractedButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "选择已解压材料目录（名称通常以解压版结尾）"
    $dialog.ShowNewFolderButton = $false
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $selected = Get-Item -LiteralPath $dialog.SelectedPath
    if (-not $selected.PSIsContainer -or -not (Get-ChildItem -LiteralPath $selected.FullName -Force | Select-Object -First 1)) {
        [System.Windows.Forms.MessageBox]::Show("请选择非空的已解压材料文件夹。", "EasySoftware", "OK", "Warning") | Out-Null
        return
    }
    $script:batchRoot = $selected.FullName
    $bytes = [Text.Encoding]::UTF8.GetBytes($script:batchRoot.ToLowerInvariant())
    $hash = [Security.Cryptography.SHA1]::Create().ComputeHash($bytes)
    $key = ((-join ($hash | ForEach-Object { $_.ToString("x2") })).Substring(0, 16))
    $script:workDir = Join-Path ([IO.Path]::GetTempPath()) ("codex-softcopyright\" + $key + "-from-extracted")
    $script:reviewPath = $null
    $script:reviewStage = $null
    $script:archiveFormat = "auto"
    $script:fromExtracted = $true
    $reviewButton.Enabled = $false
    $recheckButton.Enabled = $false
    $approvalBox.Enabled = $false
    $packageButton.Enabled = $false
    $exportButton.Text = "修改版 Word 导出 PDF"
    Set-WorkflowGuide "default"
    $inputBox.Text = $selected.FullName
    $logBox.Clear()
    Append-Log ("Selected extracted input: " + $selected.FullName)
    Start-Operation -Arguments @($script:batchRoot, "--from-extracted", "--work-dir", $script:workDir) -Operation "正在从解压版生成首轮图片审核..."
})

$reviewButton.Add_Click({
    if ($null -ne $script:reviewPath -and (Test-Path -LiteralPath $script:reviewPath)) { Start-Process -FilePath $script:reviewPath }
})

$refurbishButton.Add_Click({
    $helper = Join-Path $projectRoot "scripts\refurbish_workflow_gui.ps1"
    if (-not (Test-Path -LiteralPath $helper)) {
        [System.Windows.Forms.MessageBox]::Show("翻新失败业务助手文件不存在：$helper", "EasySoftware", "OK", "Error") | Out-Null
        return
    }
    Start-Process -FilePath "powershell.exe" -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $helper) -WindowStyle Normal
})

$orderDownloadButton.Add_Click({
    $helper = Join-Path $projectRoot "scripts\order_material_downloader_gui.ps1"
    if (-not (Test-Path -LiteralPath $helper)) {
        [System.Windows.Forms.MessageBox]::Show("订单资料下载助手文件不存在：$helper", "EasySoftware", "OK", "Error") | Out-Null
        return
    }
    $selectedStages = Get-SelectedDocumentStages
    if ([string]::IsNullOrWhiteSpace($selectedStages)) {
        [System.Windows.Forms.MessageBox]::Show("请先在自定义步骤中勾选至少一项。", "EasySoftware", "OK", "Warning") | Out-Null
        return
    }
    $arguments = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", (Quote-Argument $helper), "-DocumentStages", (Quote-Argument $selectedStages))
    if (-not $skipTopicMismatchCheck.Checked) { $arguments += "-IncludeTopicMismatch" }
    Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -WindowStyle Normal
})

$recheckButton.Add_Click({
    if ($script:reviewStage -ne "final" -or [string]::IsNullOrWhiteSpace($script:batchRoot)) { return }
    $script:workDir = Get-MainWorkDir
    $arguments = @($script:batchRoot, "--work-dir", $script:workDir, "--recheck-modified")
    if ($script:fromExtracted) { $arguments += "--from-extracted" }
    Set-WorkflowGuide "processing"
    Start-Operation -Arguments $arguments -Operation "正在重新验收修改版..."
})

$packageButton.Add_Click({
    if ($script:reviewStage -eq "documents") {
        $packageButton.Enabled = $false
        $orderDownloadButton.BringToFront()
        Append-Log ("确认处理，" + (Get-DocumentStageSummary $script:documentStages))
        $arguments = @($script:batchRoot, "--work-dir", $script:workDir, "--apply-document-processing", "--document-stages", (Get-SelectedDocumentStages))
        if (-not $skipTopicMismatchCheck.Checked) { $arguments += "--include-topic-mismatch" }
        Start-Operation -Arguments $arguments -Operation "正在处理所选材料流程..."
        return
    }
    if ($script:reviewStage -eq "code-comments") {
        $packageButton.Enabled = $false
        Start-Operation -Arguments @($script:batchRoot, "--work-dir", $script:workDir, "--apply-comment-cleanup") -Operation "正在清理注释并重新导出代码 PDF..."
        return
    }
    $approved = $approvalBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($approved)) {
        [System.Windows.Forms.MessageBox]::Show("首轮图片审核未勾选任何不通过项目时，请填写 none。", "EasySoftware", "OK", "Warning") | Out-Null
        return
    }
    $script:workDir = Get-MainWorkDir
    if ($script:reviewStage -eq "images") {
        $approvalBox.Enabled = $false
        $packageButton.Enabled = $false
        $arguments = @($script:batchRoot, "--work-dir", $script:workDir)
        if ($script:fromExtracted) { $arguments += "--from-extracted" }
        $arguments += @("--reject-images", $approved)
        Set-WorkflowGuide "processing"
        Start-Operation -Arguments $arguments -Operation "正在修复图片审核通过的项目..."
        return
    }
    if (-not (Confirm-PdfFreshness)) { return }
    $arguments = @($script:batchRoot, "--work-dir", $script:workDir, "--approve", $approved)
    if ($script:fromExtracted) { $arguments += "--from-extracted" }
    if ($script:archiveFormat -ne "auto") { $arguments += @("--format", $script:archiveFormat) }
    Set-WorkflowGuide "package"
    Start-Operation -Arguments $arguments -Operation "正在打包通过项目..."
})

$form.Add_FormClosing({
    if ($null -ne $script:process -and -not $script:process.HasExited) {
        $answer = [System.Windows.Forms.MessageBox]::Show("Processing is still running. Close the application?", "EasySoftware", "YesNo", "Warning")
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { $_.Cancel = $true }
    }
})

if (-not [string]::IsNullOrWhiteSpace($InputPath)) {
    try {
        $resolvedInputPath = [System.IO.Path]::GetFullPath($InputPath)
        if (-not (Test-Path -LiteralPath $resolvedInputPath)) {
            [System.Windows.Forms.MessageBox]::Show("传入的订单资料不存在：$resolvedInputPath", "EasySoftware", "OK", "Error") | Out-Null
        }
        else {
            $inputBox.Text = $resolvedInputPath
            Update-ExportAvailability
            if ($AutoStart) {
                if ((Test-ArchivePath $resolvedInputPath) -or (Test-Path -LiteralPath $resolvedInputPath -PathType Container)) {
                    $script:autoStartAfterShown = $true
                    $statusLabel.Text = "订单资料已载入；窗口显示后将自动开始处理。"
                }
                else {
                    $statusLabel.Text = "路径已载入；自动处理仅支持有效的压缩包，请检查后手动开始。"
                    [System.Windows.Forms.MessageBox]::Show("自动处理未启动：输入路径不是有效的压缩包文件。", "EasySoftware", "OK", "Warning") | Out-Null
                }
            }
            else {
                $statusLabel.Text = "已载入订单资料；检查路径后点击开始处理。"
            }
        }
    }
    catch {
        $script:autoStartAfterShown = $false
        [System.Windows.Forms.MessageBox]::Show("订单资料路径无效：$($_.Exception.Message)", "EasySoftware", "OK", "Error") | Out-Null
    }
}
elseif ($AutoStart) {
    [System.Windows.Forms.MessageBox]::Show("未提供订单资料路径，已取消自动处理。", "EasySoftware", "OK", "Warning") | Out-Null
}

$form.Add_Shown({
    if (-not $script:autoStartAfterShown) { return }
    $script:autoStartAfterShown = $false
    $form.BeginInvoke([System.Windows.Forms.MethodInvoker]{
        $candidatePath = $inputBox.Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($candidatePath) -and (Test-Path -LiteralPath $candidatePath) -and ((Test-ArchivePath $candidatePath) -or (Test-Path -LiteralPath $candidatePath -PathType Container))) {
            $codeCommentButton.PerformClick()
        }
    }) | Out-Null
}.GetNewClosure())

[void]$form.ShowDialog()

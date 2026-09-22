param([switch]$SelfTest)

$ErrorActionPreference = "Stop"
$projectRoot = $PSScriptRoot
$entry = Join-Path $projectRoot "scripts\unified_workflow.py"
$env:PYTHONIOENCODING = "utf-8"

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
    if (Test-Path -LiteralPath $bundledPython) {
        try {
            & $bundledPython -c "import lxml, PIL, pypdf, pypdfium2" 2>$null
            if ($LASTEXITCODE -eq 0) { return $bundledPython }
        } catch {}
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

function Prepare-ArchiveWorkspace([System.IO.FileInfo]$Archive) {
    while ($Archive.Directory.Name -ieq $Archive.BaseName -and $Archive.Directory.Parent.Name -ieq $Archive.BaseName) {
        $target = Join-Path $Archive.Directory.Parent.FullName $Archive.Name
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
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    $target = Join-Path $folder $Archive.Name
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
if ($SelfTest) {
    Write-Host "EasySoftware GUI launcher self-test"
    Write-Host "Project root: $projectRoot"
    Write-Host "Python: $python"
    Write-Host "Entry: $entry"
    if ([string]::IsNullOrWhiteSpace($python) -or -not (Test-Path -LiteralPath $entry -PathType Leaf)) { exit 2 }
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
$subtitle.Text = "选择一个 RAR / ZIP 压缩包，或选择一个批次文件夹。"
$subtitle.ForeColor = [System.Drawing.Color]::DimGray
$subtitle.Location = New-Object System.Drawing.Point(25, 55)
$subtitle.AutoSize = $true
$form.Controls.Add($subtitle)

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
$statusLabel.Location = New-Object System.Drawing.Point(750, 177)
$statusLabel.Size = New-Object System.Drawing.Size(190, 26)
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

$codeCommentButton = New-Object System.Windows.Forms.Button
$codeCommentButton.Text = "清理代码中文注释"
$codeCommentButton.Location = New-Object System.Drawing.Point(796, 135)
$codeCommentButton.Size = New-Object System.Drawing.Size(160, 38)
$codeCommentButton.Anchor = "Top,Right"
$codeCommentButton.BackColor = [System.Drawing.Color]::FromArgb(154, 91, 0)
$codeCommentButton.ForeColor = [System.Drawing.Color]::White
$codeCommentButton.FlatStyle = "Flat"
$form.Controls.Add($codeCommentButton)

$workflowGuide = New-Object System.Windows.Forms.Label
$workflowGuide.Text = "工作流：1 选择材料  >  2 首轮看图并微调 Word  >  3 二次处理  >  4 终审  >  5 导出 PDF  >  6 压缩交付"
$workflowGuide.ForeColor = [System.Drawing.Color]::FromArgb(49, 82, 138)
$workflowGuide.BackColor = [System.Drawing.Color]::FromArgb(238, 244, 255)
$workflowGuide.Location = New-Object System.Drawing.Point(26, 176)
$workflowGuide.Size = New-Object System.Drawing.Size(710, 18)
$workflowGuide.Anchor = "Top,Left,Right"
$workflowGuide.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8.5)
$form.Controls.Add($workflowGuide)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(26, 198)
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
$progressDetail.Location = New-Object System.Drawing.Point(26, 224)
$progressDetail.Size = New-Object System.Drawing.Size(914, 24)
$progressDetail.Anchor = "Top,Left,Right"
$form.Controls.Add($progressDetail)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(26, 252)
$logBox.Size = New-Object System.Drawing.Size(914, 323)
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
    $codeCommentButton.Enabled = -not $Busy
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
    if ($script:fromExtracted) {
        $name = $source.Name
        if ($name.EndsWith("_解压版")) { return (Join-Path $source.Parent.FullName ($name.Substring(0, $name.Length - 4) + "_修改版")) }
        if ($name.EndsWith("解压版")) { return (Join-Path $source.Parent.FullName ($name.Substring(0, $name.Length - 3) + "修改版")) }
        return (Join-Path $source.Parent.FullName "修改版")
    }
    if (-not $source.PSIsContainer) { return (Join-Path $source.DirectoryName ($source.BaseName + "_修改版")) }
    if ($source.Name.EndsWith("修改版")) { return $source.FullName }
    return (Join-Path $source.FullName "修改版")
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
    $script:stdoutPath = [IO.Path]::GetTempFileName()
    $script:stderrPath = [IO.Path]::GetTempFileName()
    $argumentText = (@(Quote-Argument $entry) + @($Arguments | ForEach-Object { Quote-Argument $_ })) -join " "
    Append-Log ("> " + $python + " " + $argumentText)
    $script:operation = $Operation
    $script:process = Start-Process -FilePath $python -ArgumentList $argumentText -PassThru -WindowStyle Hidden -RedirectStandardOutput $script:stdoutPath -RedirectStandardError $script:stderrPath
    Set-Busy $true $Operation
    $timer.Start()
}

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 750
$timer.Add_Tick({
    if ($null -ne $script:workDir) {
        $progressPath = Join-Path $script:workDir "progress.json"
        if (Test-Path -LiteralPath $progressPath) {
            try {
                $state = Get-Content -LiteralPath $progressPath -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
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
                    $progressDetail.Text = "已处理 $done 件 · 待处理 $left 件 · 失败 $failed 件 · RAR：$archive · $activity · 阶段：$displayPhase$eta · 已用 $elapsed 秒"
                } else {
                    $progress.Style = "Marquee"
                    $progress.MarqueeAnimationSpeed = 30
                    $statusLabel.Text = $displayPhase
                    $progressDetail.Text = "正在准备项目清单 · 模块：$displayPhase · 已用 $elapsed 秒"
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
        Set-Busy $false ("预览完成：预计删除 " + [string]$result.deletions + " 处")
        if ($reviewButton.Enabled) { Start-Process -FilePath $script:reviewPath }
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
    $dialog.Filter = "软著压缩包 (*.rar;*.zip)|*.rar;*.zip|RAR (*.rar)|*.rar|ZIP (*.zip)|*.zip|所有文件 (*.*)|*.*"
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

$codeCommentButton.Add_Click({
    $value = $inputBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($value) -or -not (Test-Path -LiteralPath $value)) {
        [System.Windows.Forms.MessageBox]::Show("请先选择 RAR 压缩包、包含 RAR 的文件夹，或已解压的代码材料文件夹。", "EasySoftware", "OK", "Warning") | Out-Null
        return
    }
    $script:batchRoot = (Get-Item -LiteralPath $value).FullName
    $bytes = [Text.Encoding]::UTF8.GetBytes($script:batchRoot.ToLowerInvariant())
    $hash = [Security.Cryptography.SHA1]::Create().ComputeHash($bytes)
    $key = ((-join ($hash | ForEach-Object { $_.ToString("x2") })).Substring(0, 16))
    $script:workDir = Join-Path ([IO.Path]::GetTempPath()) ("codex-softcopyright\" + $key + "-code-comments")
    $script:reviewPath = $null
    $script:reviewStage = $null
    $reviewButton.Enabled = $false
    $packageButton.Enabled = $false
    $approvalBox.Enabled = $false
    $logBox.Clear()
    Append-Log ("代码注释清理预览：" + $script:batchRoot)
    Start-Operation -Arguments @($script:batchRoot, "--work-dir", $script:workDir, "--clean-code-comments") -Operation "正在扫描代码中文注释..."
})

$exportButton.Add_Click({
    $value = $inputBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($value) -or -not (Test-Path -LiteralPath $value)) {
        [System.Windows.Forms.MessageBox]::Show("请先选择原批次文件夹或压缩包。", "EasySoftware", "OK", "Warning") | Out-Null
        return
    }
    $item = Get-Item -LiteralPath $value
    if (-not $item.PSIsContainer -and $item.Extension.ToLowerInvariant() -in @('.zip', '.rar')) {
        try { $value = Prepare-ArchiveWorkspace $item; $inputBox.Text = $value } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "EasySoftware", "OK", "Error") | Out-Null; return }
    }
    $selected = Get-Item -LiteralPath $value
    # Use the batch root from the last processing step so the export matches the
    # same modified directory. Never overwrite $script:batchRoot or $script:workDir
    # here: packaging and re-verification rely on them. The export report goes to
    # a separate "-export" directory instead.
    if ([string]::IsNullOrWhiteSpace($script:batchRoot)) {
        $script:batchRoot = if ($selected.PSIsContainer -and $selected.Name.EndsWith("修改版")) { $selected.Parent.FullName } else { $selected.FullName }
    }
    $bytes = [Text.Encoding]::UTF8.GetBytes($script:batchRoot.ToLowerInvariant())
    $hash = [Security.Cryptography.SHA1]::Create().ComputeHash($bytes)
    $key = ((-join ($hash | ForEach-Object { $_.ToString("x2") })).Substring(0, 16))
    $script:exportWorkDir = Join-Path ([IO.Path]::GetTempPath()) ("codex-softcopyright\" + $key + "-export")
    $arguments = @($script:batchRoot, "--work-dir", $script:exportWorkDir, "--export-modified-pdfs")
    if ($script:fromExtracted) { $arguments += "--from-extracted" }
    Set-WorkflowGuide "export"
    Start-Operation -Arguments $arguments -Operation "正在将修改版 Word 导出并覆盖 PDF..."
})

$startButton.Add_Click({
    $value = $inputBox.Text.Trim()
    if (-not (Test-Path -LiteralPath $value) -and [IO.Path]::GetExtension($value).ToLowerInvariant() -in @('.zip', '.rar')) {
        $nested = Join-Path (Join-Path (Split-Path -Parent $value) ([IO.Path]::GetFileNameWithoutExtension($value))) ([IO.Path]::GetFileName($value))
        if (Test-Path -LiteralPath $nested) { $value = $nested; $inputBox.Text = $nested }
    }
    if ([string]::IsNullOrWhiteSpace($value) -or -not (Test-Path -LiteralPath $value)) {
        [System.Windows.Forms.MessageBox]::Show("Select an archive or batch folder first.", "EasySoftware", "OK", "Warning") | Out-Null
        return
    }
    $item = Get-Item -LiteralPath $value
    if (-not $item.PSIsContainer -and $item.Extension.ToLowerInvariant() -in @('.zip', '.rar')) {
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
    if (-not (Test-Path -LiteralPath $value) -and [IO.Path]::GetExtension($value).ToLowerInvariant() -in @('.zip', '.rar')) {
        $nested = Join-Path (Join-Path (Split-Path -Parent $value) ([IO.Path]::GetFileNameWithoutExtension($value))) ([IO.Path]::GetFileName($value))
        if (Test-Path -LiteralPath $nested) { $value = $nested; $inputBox.Text = $nested }
    }
    if ([string]::IsNullOrWhiteSpace($value) -or -not (Test-Path -LiteralPath $value)) {
        [System.Windows.Forms.MessageBox]::Show("Select an archive or batch folder first.", "EasySoftware", "OK", "Warning") | Out-Null
        return
    }
    $item = Get-Item -LiteralPath $value
    if (-not $item.PSIsContainer -and $item.Extension.ToLowerInvariant() -in @('.zip', '.rar')) {
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

$recheckButton.Add_Click({
    if ($script:reviewStage -ne "final" -or [string]::IsNullOrWhiteSpace($script:batchRoot)) { return }
    $script:workDir = Get-MainWorkDir
    $arguments = @($script:batchRoot, "--work-dir", $script:workDir, "--recheck-modified")
    if ($script:fromExtracted) { $arguments += "--from-extracted" }
    Set-WorkflowGuide "processing"
    Start-Operation -Arguments $arguments -Operation "正在重新验收修改版..."
})

$packageButton.Add_Click({
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

[void]$form.ShowDialog()

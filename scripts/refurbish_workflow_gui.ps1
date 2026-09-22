$ErrorActionPreference = "Stop"
$scriptRoot = Split-Path -Parent $PSScriptRoot
$entry = Join-Path $PSScriptRoot "refurbish_workflow.py"
$env:PYTHONIOENCODING = "utf-8"

Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class RefurbishConsoleWindow {
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr handle, int command);
}
"@

$consoleWindow = [RefurbishConsoleWindow]::GetConsoleWindow()
if ($consoleWindow -ne [IntPtr]::Zero) {
    [RefurbishConsoleWindow]::ShowWindow($consoleWindow, 0) | Out-Null
}

function Find-Python {
    $candidates = @()
    foreach ($name in @("python.exe", "python")) {
        $command = Get-Command $name -ErrorAction SilentlyContinue
        if ($null -ne $command) { $candidates += $command.Source }
    }
    $profileRoot = [Environment]::GetFolderPath([Environment+SpecialFolder]::UserProfile)
    $bundled = Join-Path $profileRoot '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
    if (Test-Path -LiteralPath $bundled) { $candidates += $bundled }
    foreach ($candidate in ($candidates | Select-Object -Unique)) {
        try {
            & $candidate -c "import urllib.request" 2>$null
            if ($LASTEXITCODE -eq 0) { return $candidate }
        } catch {}
    }
    return $null
}

function Quote-Argument([string]$Value) { return '"' + ($Value -replace '"', '\"') + '"' }

$python = Find-Python
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Security
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = "翻新失败业务助手"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(900, 680)
$form.MinimumSize = New-Object System.Drawing.Size(820, 600)
$form.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)

$title = New-Object System.Windows.Forms.Label
$title.Text = "翻新失败业务助手"
$title.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 18, [System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(24, 18)
$title.AutoSize = $true
$form.Controls.Add($title)

$guide = New-Object System.Windows.Forms.Label
$guide.Text = "工作流：1 筛选 admin 批次  >  2 进行中查看暂停项  >  3 部分完成恢复代码  >  4 逐批完成并记录"
$guide.Location = New-Object System.Drawing.Point(26, 58)
$guide.Size = New-Object System.Drawing.Size(830, 28)
$guide.BackColor = [System.Drawing.Color]::FromArgb(238, 244, 255)
$guide.ForeColor = [System.Drawing.Color]::FromArgb(49, 82, 138)
$guide.Padding = New-Object System.Windows.Forms.Padding(6, 4, 6, 2)
$form.Controls.Add($guide)

function Add-Label([string]$Text, [int]$X, [int]$Y, [int]$Width = 90) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.Size = New-Object System.Drawing.Size($Width, 28)
    $form.Controls.Add($label) | Out-Null
}

Add-Label "操作人" 26 108 70
$operatorBox = New-Object System.Windows.Forms.TextBox
$operatorBox.Text = "admin"
$operatorBox.Location = New-Object System.Drawing.Point(96, 104)
$operatorBox.Size = New-Object System.Drawing.Size(160, 30)
$form.Controls.Add($operatorBox)

Add-Label "开始日期" 280 108 75
$beginBox = New-Object System.Windows.Forms.TextBox
$beginBox.Location = New-Object System.Drawing.Point(355, 104)
$beginBox.Size = New-Object System.Drawing.Size(135, 30)
$form.Controls.Add($beginBox)

Add-Label "结束日期" 510 108 75
$endBox = New-Object System.Windows.Forms.TextBox
$endBox.Location = New-Object System.Drawing.Point(585, 104)
$endBox.Size = New-Object System.Drawing.Size(135, 30)
$form.Controls.Add($endBox)

$dateHint = New-Object System.Windows.Forms.Label
$dateHint.Text = "日期可留空，格式 YYYY-MM-DD"
$dateHint.ForeColor = [System.Drawing.Color]::DimGray
$dateHint.Location = New-Object System.Drawing.Point(730, 109)
$dateHint.AutoSize = $true
$form.Controls.Add($dateHint)

Add-Label "Edge/Chrome 令牌" 26 148 100
$tokenBox = New-Object System.Windows.Forms.TextBox
$tokenBox.Location = New-Object System.Drawing.Point(126, 144)
$tokenBox.Size = New-Object System.Drawing.Size(594, 30)
$tokenBox.PasswordChar = '*'
$form.Controls.Add($tokenBox)

$tokenHint = New-Object System.Windows.Forms.Label
$tokenHint.Text = "从已登录的 Edge 或 Chrome 页面获取；仅本机使用"
$tokenHint.ForeColor = [System.Drawing.Color]::DimGray
$tokenHint.Location = New-Object System.Drawing.Point(730, 149)
$tokenHint.AutoSize = $true
$form.Controls.Add($tokenHint)

$rememberTokenBox = New-Object System.Windows.Forms.CheckBox
$rememberTokenBox.Text = "记住令牌（当前 Windows 用户）"
$rememberTokenBox.Location = New-Object System.Drawing.Point(26, 184)
$rememberTokenBox.Size = New-Object System.Drawing.Size(250, 28)
$rememberTokenBox.Checked = $true
$form.Controls.Add($rememberTokenBox)

$clearTokenButton = New-Object System.Windows.Forms.Button
$clearTokenButton.Text = "清除已保存令牌"
$clearTokenButton.Location = New-Object System.Drawing.Point(280, 181)
$clearTokenButton.Size = New-Object System.Drawing.Size(130, 32)
$form.Controls.Add($clearTokenButton)

$previewBox = New-Object System.Windows.Forms.CheckBox
$previewBox.Text = "只预览，不提交接口（建议先预览一次）"
$previewBox.Location = New-Object System.Drawing.Point(430, 184)
$previewBox.Size = New-Object System.Drawing.Size(300, 28)
$previewBox.Checked = $true
$form.Controls.Add($previewBox)

$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = "读取并预览"
$runButton.Location = New-Object System.Drawing.Point(26, 218)
$runButton.Size = New-Object System.Drawing.Size(160, 36)
$runButton.BackColor = [System.Drawing.Color]::FromArgb(22, 93, 255)
$runButton.ForeColor = [System.Drawing.Color]::White
$runButton.FlatStyle = "Flat"
$form.Controls.Add($runButton)

$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = "关闭"
$closeButton.Location = New-Object System.Drawing.Point(200, 218)
$closeButton.Size = New-Object System.Drawing.Size(100, 36)
$form.Controls.Add($closeButton)

$openReportButton = New-Object System.Windows.Forms.Button
$openReportButton.Text = "打开本次报告"
$openReportButton.Location = New-Object System.Drawing.Point(315, 218)
$openReportButton.Size = New-Object System.Drawing.Size(130, 36)
$openReportButton.Enabled = $false
$form.Controls.Add($openReportButton)

$status = New-Object System.Windows.Forms.Label
$status.Text = "等待开始"
$status.Location = New-Object System.Drawing.Point(465, 224)
$status.Size = New-Object System.Drawing.Size(390, 28)
$form.Controls.Add($status)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(26, 265)
$progress.Size = New-Object System.Drawing.Size(830, 20)
$progress.Style = "Continuous"
$progress.Value = 0
$form.Controls.Add($progress)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(26, 297)
$logBox.Size = New-Object System.Drawing.Size(830, 308)
$logBox.Multiline = $true
$logBox.ReadOnly = $true
$logBox.ScrollBars = "Both"
$logBox.WordWrap = $true
$logBox.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$form.Controls.Add($logBox)

$script:process = $null
$script:stdoutPath = $null
$script:stderrPath = $null
$script:reportPath = $null
$script:settingsDirectory = Join-Path $env:LOCALAPPDATA "EasySoftware"
$script:settingsPath = Join-Path $script:settingsDirectory "refurbish-settings.json"

function Read-SavedToken {
    try {
        if (-not (Test-Path -LiteralPath $script:settingsPath)) { return "" }
        $settings = Get-Content -LiteralPath $script:settingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ([string]::IsNullOrWhiteSpace([string]$settings.token_protected)) { return "" }
        $protected = [Convert]::FromBase64String([string]$settings.token_protected)
        $plain = [Security.Cryptography.ProtectedData]::Unprotect($protected, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)
        return [Text.Encoding]::UTF8.GetString($plain)
    } catch {
        return ""
    }
}

function Save-Token([string]$Token) {
    New-Item -ItemType Directory -Path $script:settingsDirectory -Force | Out-Null
    $plain = [Text.Encoding]::UTF8.GetBytes($Token)
    $protected = [Security.Cryptography.ProtectedData]::Protect($plain, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)
    $settings = @{ token_protected = [Convert]::ToBase64String($protected); saved_at = (Get-Date -Format "s") } | ConvertTo-Json
    [IO.File]::WriteAllText($script:settingsPath, $settings, [Text.UTF8Encoding]::new($false))
}

function Clear-SavedToken {
    if (Test-Path -LiteralPath $script:settingsPath) {
        Remove-Item -LiteralPath $script:settingsPath -Force
    }
}

function Append-Log([string]$Text) {
    if (-not [string]::IsNullOrWhiteSpace($Text)) { $logBox.AppendText($Text.TrimEnd() + [Environment]::NewLine) }
}

function Start-Workflow {
    if ($null -ne $script:process -and -not $script:process.HasExited) { return }
    if ($null -eq $python) { [System.Windows.Forms.MessageBox]::Show("未找到可用 Python。", "翻新失败业务助手", "OK", "Error") | Out-Null; return }
    if ([string]::IsNullOrWhiteSpace($tokenBox.Text)) { [System.Windows.Forms.MessageBox]::Show("请填入 Edge 或 Chrome 登录令牌。", "翻新失败业务助手", "OK", "Warning") | Out-Null; return }
    if ($rememberTokenBox.Checked) { Save-Token $tokenBox.Text.Trim() } else { Clear-SavedToken }
    $arguments = @("$entry", "--operator", $operatorBox.Text.Trim(), "--base-url", "http://36.212.60.8:10000")
    if (-not [string]::IsNullOrWhiteSpace($beginBox.Text)) { $arguments += @("--begin", $beginBox.Text.Trim()) }
    if (-not [string]::IsNullOrWhiteSpace($endBox.Text)) { $arguments += @("--end", $endBox.Text.Trim()) }
    if ($previewBox.Checked) { $arguments += "--dry-run" }
    $reportDirectory = Join-Path $env:LOCALAPPDATA "EasySoftware\refurbish-reports"
    New-Item -ItemType Directory -Path $reportDirectory -Force | Out-Null
    $script:reportPath = Join-Path $reportDirectory ("refurbish-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".html")
    $arguments += @("--report", $script:reportPath)
    $argumentText = ($arguments | ForEach-Object { Quote-Argument $_ }) -join " "
    $script:stdoutPath = [IO.Path]::GetTempFileName()
    $script:stderrPath = [IO.Path]::GetTempFileName()
    $previousToken = $env:EASYSOFTWARE_REFURBISH_TOKEN
    try {
        $env:EASYSOFTWARE_REFURBISH_TOKEN = $tokenBox.Text.Trim()
        $script:process = Start-Process -FilePath $python -ArgumentList $argumentText -PassThru -WindowStyle Hidden -RedirectStandardOutput $script:stdoutPath -RedirectStandardError $script:stderrPath
    } finally {
        $env:EASYSOFTWARE_REFURBISH_TOKEN = $previousToken
    }
    $runButton.Enabled = $false
    $closeButton.Enabled = $false
    $previewBox.Enabled = $false
    $rememberTokenBox.Enabled = $false
    $clearTokenButton.Enabled = $false
    $openReportButton.Enabled = $false
    $progress.Style = "Marquee"
    $status.Text = if ($previewBox.Checked) { "正在读取批次..." } else { "正在顺序处理..." }
    $timer.Start()
}

function Show-Report($result) {
    $logBox.Clear()
    if ($result.status -eq "OK") {
        Append-Log ("操作人：" + [string]$result.operator)
        Append-Log ("批次结果：发现 " + @($result.batches).Count + " 个相关批次")
        Append-Log ("暂停订单：" + [string]$result.paused_orders + " 个（需进入流程的复杂去重暂未自动执行）")
        Append-Log ("代码恢复候选：" + [string]$result.resume_candidates + " 个")
        if (-not $result.dry_run) { Append-Log ("已提交恢复请求：" + [string]$result.submitted_orders + " 个；系统确认恢复：" + [string]$result.resumed_orders + " 个；后端跳过：" + [string]$result.skipped_orders + " 个") }
        foreach ($item in @($result.batches)) {
            $action = switch ([string]$item.action) { "RESUME_CODE_GEN" { "已提交代码生成恢复" } "PLAN_ONLY" { "预览：将恢复代码生成" } default { "暂停项待人工流程" } }
            Append-Log (([string]$item.batch_no) + "：" + $action + "，订单 " + [string]$item.order_count + " 个")
            foreach ($order in @($item.orders)) { Append-Log ("  - " + [string]$order.name) }
        }
        if ($result.report_path) { Append-Log ("本次报告：" + [string]$result.report_path) }
    } else {
        Append-Log (([string]$result.status) + "：" + [string]$result.error)
    }
}

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 500
$timer.Add_Tick({
    if ($null -eq $script:process -or -not $script:process.HasExited) { return }
    $timer.Stop()
    $output = if (Test-Path -LiteralPath $script:stdoutPath) { Get-Content -LiteralPath $script:stdoutPath -Raw -Encoding UTF8 } else { "" }
    $errors = if (Test-Path -LiteralPath $script:stderrPath) { Get-Content -LiteralPath $script:stderrPath -Raw -Encoding UTF8 } else { "" }
    $result = $null
    foreach ($line in @($output -split "`r?`n" | Select-Object -Last 20)) { try { $candidate = $line | ConvertFrom-Json -ErrorAction Stop; if ($candidate.status) { $result = $candidate } } catch {} }
    $progress.Style = "Continuous"
    $progress.Value = 100
    $runButton.Enabled = $true
    $closeButton.Enabled = $true
    $previewBox.Enabled = $true
    $rememberTokenBox.Enabled = $true
    $clearTokenButton.Enabled = $true
    $script:process = $null
    if ($result) { Show-Report $result; $openReportButton.Enabled = -not [string]::IsNullOrWhiteSpace([string]$result.report_path) -and (Test-Path -LiteralPath ([string]$result.report_path)); $status.Text = "处理完成" } else { Append-Log $output; Append-Log $errors; $status.Text = "未返回结果" }
})

$runButton.Add_Click({ Start-Workflow })
$previewBox.Add_CheckedChanged({ $runButton.Text = if ($previewBox.Checked) { "读取并预览" } else { "开始正式处理" } })
$clearTokenButton.Add_Click({ Clear-SavedToken; $tokenBox.Clear(); $rememberTokenBox.Checked = $false; Append-Log "已清除本机保存的令牌" })
$openReportButton.Add_Click({ if (-not [string]::IsNullOrWhiteSpace($script:reportPath) -and (Test-Path -LiteralPath $script:reportPath)) { Start-Process -FilePath $script:reportPath } })
$closeButton.Add_Click({ $form.Close() })
$tokenBox.Text = Read-SavedToken
$form.Add_FormClosing({ if ($null -ne $script:process -and -not $script:process.HasExited) { $_.Cancel = $true } })
[void]$form.ShowDialog()

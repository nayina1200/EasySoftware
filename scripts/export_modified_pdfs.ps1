param(
    [Parameter(Mandatory = $true)][string]$BaseDir,
    [Parameter(Mandatory = $true)][string]$ReportPath,
    [string]$ProgressPath = "",
    [int]$ItemTimeoutSeconds = 90
)

$ErrorActionPreference = "Stop"
$itemScript = Join-Path $PSScriptRoot "export_modified_pdf_item.ps1"
$started = [DateTime]::UtcNow
$base = (Resolve-Path -LiteralPath $BaseDir).Path
$results = @()
$docxFiles = @(Get-ChildItem -LiteralPath $base -Recurse -File -Filter "*.docx" |
    Where-Object { $_.Name -notlike "~$*" -and $_.FullName -notmatch "\\(?:backup|\.easysoftware-)" } |
    Sort-Object FullName)

function Write-Utf8([string]$Path, [string]$Text) {
    $parent = [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($Path))
    if ($parent) { [IO.Directory]::CreateDirectory($parent) | Out-Null }
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($Path), $Text, (New-Object Text.UTF8Encoding($false)))
}

function Write-ProgressState([int]$Current, [string]$Name, [string]$Status = "RUNNING", [string]$ErrorText = "") {
    if ([string]::IsNullOrWhiteSpace($ProgressPath)) { return }
    $elapsed = ([DateTime]::UtcNow - $started).TotalSeconds
    $estimate = if ($Current -gt 0) { ($elapsed / $Current) * [Math]::Max(0, $docxFiles.Count - $Current) } else { $null }
    $state = [ordered]@{
        version = "3.4.25"
        status = $Status
        phase = "WORD_PDF_EXPORT"
        current = $Current
        total = $docxFiles.Count
        project_name = $Name
        elapsed_seconds = [Math]::Round($elapsed, 1)
        completed_projects = $Current
        remaining_projects = [Math]::Max(0, $docxFiles.Count - $Current)
        estimated_remaining_seconds = if ($null -ne $estimate) { [Math]::Round($estimate, 1) } else { $null }
        current_module = "WORD_PDF_EXPORT"
        current_project_active = ($Status -eq "RUNNING")
        failures = @($results | Where-Object status -eq "FAILED").Count
        updated_at = [DateTime]::Now.ToString("o")
    }
    if ($ErrorText) { $state.error = $ErrorText }
    Write-Utf8 $ProgressPath ($state | ConvertTo-Json -Depth 5)
}

function Get-WordProcessIds {
    return @(Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
}

function Stop-ChildWord([System.Diagnostics.Process]$Child, [string]$PidPath, [int[]]$FallbackWordPids = @()) {
    $wordPid = 0
    try { $wordPid = [int](Get-Content -LiteralPath $PidPath -Raw -ErrorAction Stop).Trim() } catch {}
    $wordPids = @($FallbackWordPids)
    if ($wordPid -gt 0) { $wordPids += $wordPid }
    foreach ($targetProcessId in @($wordPids | Select-Object -Unique)) {
        Start-Process -FilePath "taskkill.exe" -ArgumentList @("/PID", $targetProcessId, "/T", "/F") -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue | Out-Null
    }
    if ($null -ne $Child -and -not $Child.HasExited) {
        Start-Process -FilePath "taskkill.exe" -ArgumentList @("/PID", $Child.Id, "/T", "/F") -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue | Out-Null
    }
}

function Quote-ProcessArgument([string]$Value) {
    return '"' + ($Value -replace '"', '\"') + '"'
}

$fatalError = ""
try {
    Write-ProgressState 0 "PREPARING"
    for ($index = 0; $index -lt $docxFiles.Count; $index++) {
        $docx = $docxFiles[$index]
        $pdf = [IO.Path]::ChangeExtension($docx.FullName, ".pdf")
        $temporaryPdf = $pdf + ".easysoftware-export-" + [Guid]::NewGuid().ToString("N") + ".tmp"
        $itemDir = Join-Path ([IO.Path]::GetTempPath()) ("easysoftware-export-" + [Guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Path $itemDir -Force | Out-Null
        $itemResult = Join-Path $itemDir "result.json"
        $pidPath = Join-Path $itemDir "word.pid"
        $itemStarted = [DateTime]::UtcNow
        $child = $null
        $wordPidsBefore = Get-WordProcessIds
        Write-ProgressState $index $docx.BaseName
        try {
            if ((Test-Path -LiteralPath $pdf -PathType Leaf) -and ((Get-Item -LiteralPath $pdf).LastWriteTimeUtc -ge $docx.LastWriteTimeUtc)) {
                $results += [pscustomobject]@{ name = $docx.BaseName; docx = $docx.FullName; pdf = $pdf; status = "SKIPPED"; reason = "PDF_IS_CURRENT"; seconds = 0 }
                Write-ProgressState ($index + 1) $docx.BaseName
                continue
            }
            $shell = (Get-Command powershell.exe -ErrorAction Stop).Source
            $arguments = (@(
                "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $itemScript,
                "-DocxPath", $docx.FullName, "-TemporaryPdfPath", $temporaryPdf,
                "-ResultPath", $itemResult, "-PidPath", $pidPath
            ) | ForEach-Object { Quote-ProcessArgument ([string]$_) }) -join " "
            $child = Start-Process -FilePath $shell -ArgumentList $arguments -PassThru -WindowStyle Hidden
            if (-not $child.WaitForExit($ItemTimeoutSeconds * 1000)) {
                $message = "SINGLE_DOCUMENT_TIMEOUT"
                $newWordPids = @(Get-WordProcessIds | Where-Object { $_ -notin $wordPidsBefore })
                Stop-ChildWord $child $pidPath $newWordPids
                $results += [pscustomobject]@{ name = $docx.BaseName; docx = $docx.FullName; pdf = $pdf; status = "FAILED"; error = $message; seconds = $ItemTimeoutSeconds }
                Write-ProgressState ($index + 1) $docx.BaseName "RUNNING" $message
            } else {
                $row = if (Test-Path -LiteralPath $itemResult) { Get-Content -LiteralPath $itemResult -Raw -Encoding UTF8 | ConvertFrom-Json } else { $null }
                if ($null -ne $row -and [string]$row.status -eq "OK" -and (Test-Path -LiteralPath $temporaryPdf -PathType Leaf)) {
                    Move-Item -LiteralPath $temporaryPdf -Destination $pdf -Force
                    $results += [pscustomobject]@{ name = $docx.BaseName; docx = $docx.FullName; pdf = $pdf; status = "OK"; seconds = [Math]::Round(([DateTime]::UtcNow - $itemStarted).TotalSeconds, 2) }
                } else {
                    $errorText = if ($null -ne $row) { [string]$row.error } else { "CHILD_PROCESS_NO_REPORT" }
                    $results += [pscustomobject]@{ name = $docx.BaseName; docx = $docx.FullName; pdf = $pdf; status = "FAILED"; error = $errorText; seconds = [Math]::Round(([DateTime]::UtcNow - $itemStarted).TotalSeconds, 2) }
                }
            }
        } catch {
            if ($null -ne $child) {
                $newWordPids = @(Get-WordProcessIds | Where-Object { $_ -notin $wordPidsBefore })
                Stop-ChildWord $child $pidPath $newWordPids
            }
            if (Test-Path -LiteralPath $temporaryPdf) { Remove-Item -LiteralPath $temporaryPdf -Force -ErrorAction SilentlyContinue }
            $results += [pscustomobject]@{ name = $docx.BaseName; docx = $docx.FullName; pdf = $pdf; status = "FAILED"; error = $_.Exception.Message; seconds = [Math]::Round(([DateTime]::UtcNow - $itemStarted).TotalSeconds, 2) }
        } finally {
            Remove-Item -LiteralPath $itemDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        Write-ProgressState ($index + 1) $docx.BaseName
    }
} catch {
    $fatalError = $_.Exception.Message
    Write-ProgressState $results.Count "EXPORT_ERROR" "FAILED" $_.Exception.Message
} finally {
    $failedCount = @($results | Where-Object status -eq "FAILED").Count
    $finalStatus = if ($fatalError -or $failedCount -gt 0) { "EXPORT_FAILED" } else { "EXPORT_OK" }
    $report = [pscustomobject]@{
        version = "3.4.25"
        status = $finalStatus
        base = $base
        exported = @($results | Where-Object status -eq "OK").Count
        skipped = @($results | Where-Object status -eq "SKIPPED").Count
        failed = $failedCount
        files = $results
        seconds = [Math]::Round(([DateTime]::UtcNow - $started).TotalSeconds, 2)
    }
    Write-Utf8 $ReportPath ($report | ConvertTo-Json -Depth 8)
    Write-ProgressState $docxFiles.Count "EXPORT_COMPLETE" $(if ($finalStatus -eq "EXPORT_OK") { "DONE" } else { "FAILED" }) $fatalError
}

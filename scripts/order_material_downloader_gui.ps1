param([switch]$SelfTest)

$ErrorActionPreference = "Stop"
$scriptRoot = Split-Path -Parent $PSScriptRoot
$corePath = Join-Path $PSScriptRoot "order_download_core.py"
$mainGuiPath = Join-Path $scriptRoot "启动EasySoftware.ps1"
$targetUrl = "http://36.212.60.8:10000/material/batchRefurbish"
$debugPort = 9222
$debugEndpoint = "http://127.0.0.1:$debugPort"
$profilePath = Join-Path $env:LOCALAPPDATA "EasySoftware\OrderMaterialBrowser"
$settingsDirectory = Join-Path $env:LOCALAPPDATA "EasySoftware"
$settingsPath = Join-Path $settingsDirectory "order-download-settings.json"
$script:socket = $null
$script:requestId = 0
$script:cdpEvents = @()
$script:python = $null
$script:searchRows = @()
$script:finalDownloadPath = $null

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

function Find-Python {
    $candidates = @((Join-Path $scriptRoot "python\python.exe"))
    foreach ($name in @("python.exe", "python")) {
        $command = Get-Command $name -ErrorAction SilentlyContinue
        if ($null -ne $command) { $candidates += $command.Source }
    }
    foreach ($candidate in ($candidates | Select-Object -Unique)) {
        if (-not (Test-Path -LiteralPath $candidate)) { continue }
        try {
            & $candidate -c "import zipfile" 2>$null
            if ($LASTEXITCODE -eq 0) { return $candidate }
        } catch {}
    }
    return $null
}

function Quote-Argument([string]$Value) {
    return '"' + ($Value -replace '"', '\"') + '"'
}

function Get-DedicatedChromeProcessIds {
    $expectedProfile = [IO.Path]::GetFullPath($profilePath).TrimEnd('\')
    try {
        foreach ($process in @(Get-CimInstance Win32_Process -Filter "Name = 'chrome.exe'" -ErrorAction Stop)) {
            $commandLine = [string]$process.CommandLine
            if (-not [regex]::IsMatch($commandLine, "(?i)(?:^|\s)--remote-debugging-port=$debugPort(?:\s|$)")) { continue }
            $profileMatch = [regex]::Match($commandLine, '(?i)(?:^|\s)--user-data-dir=(?:"([^"]+)"|([^\s]+))')
            if (-not $profileMatch.Success) { continue }
            $candidateProfile = if ($profileMatch.Groups[1].Success) { $profileMatch.Groups[1].Value } else { $profileMatch.Groups[2].Value }
            if ([string]::Equals([IO.Path]::GetFullPath($candidateProfile).TrimEnd('\'), $expectedProfile, [StringComparison]::OrdinalIgnoreCase)) {
                [int]$process.ProcessId
            }
        }
    } catch { return @() }
}

function Assert-DedicatedChromeListener {
    $processIds = @(Get-DedicatedChromeProcessIds)
    if ($processIds.Count -eq 0) {
        throw "未找到使用订单资料专用配置目录和调试端口的 Chrome。请通过助手启动专用窗口。"
    }
    try {
        $listeners = @(Get-NetTCPConnection -State Listen -LocalPort $debugPort -ErrorAction Stop)
    } catch {
        throw "无法验证本机调试端口的监听进程；为避免连接到其他程序，已停止。"
    }
    if ($listeners.Count -ne 1 -or $listeners[0].LocalAddress -ne "127.0.0.1" -or $processIds -notcontains [int]$listeners[0].OwningProcess) {
        throw "调试端口没有由专用 Chrome 绑定到本机回环地址；为避免连接到其他程序，已停止。"
    }
}

function Assert-LoopbackWebSocketTarget([string]$WebSocketUrl) {
    try { $uri = [Uri]$WebSocketUrl } catch { throw "浏览器返回了无效的调试地址，已停止。" }
    $isLoopback = $false
    if ($uri.Host -eq "localhost") { $isLoopback = $true }
    else {
        $address = $null
        if ([Net.IPAddress]::TryParse($uri.Host, [ref]$address)) { $isLoopback = [Net.IPAddress]::IsLoopback($address) }
    }
    if ($uri.Scheme -ne "ws" -or -not $isLoopback -or $uri.Port -ne $debugPort -or $uri.UserInfo -or $uri.AbsolutePath -notmatch '^/devtools/page/[A-Za-z0-9_-]+$') {
        throw "浏览器调试地址不属于预期的本机页面端点，已停止。"
    }
    return $uri
}

function Convert-CdpEventToSafeRecord($Event) {
    $method = [string]$Event.method
    $params = $Event.params
    $requestId = [string]$params.requestId
    switch ($method) {
        "Network.requestWillBeSent" {
            return [pscustomobject]@{ method = $method; requestId = $requestId; url = [string]$params.request.url; httpMethod = [string]$params.request.method }
        }
        "Network.responseReceived" {
            $status = $null
            if ($null -ne $params.response.status) { $status = [int]$params.response.status }
            return [pscustomobject]@{ method = $method; requestId = $requestId; status = $status }
        }
        "Network.loadingFinished" {
            return [pscustomobject]@{ method = $method; requestId = $requestId }
        }
        "Network.loadingFailed" {
            return [pscustomobject]@{ method = $method; requestId = $requestId }
        }
        default { return $null }
    }
}

function Assert-BatchSearchRequest([string]$Method, [string]$Url, [string]$QueryField, [string]$QueryValue) {
    if ($Method -ne "GET" -or $QueryField -notin @("batchNo", "orderName")) { throw "批次搜索请求不符合已确认的只读查询格式。" }
    try { $uri = [Uri]$Url } catch { throw "批次搜索请求地址无效。" }
    if ($uri.Scheme -ne "http" -or $uri.Host -ne "36.212.60.8" -or $uri.Port -ne 10000 -or $uri.AbsolutePath -notmatch '^/(?:prod-api/)?doc/refurbish/batch/groups$') {
        throw "未捕获到预期的批次列表查询请求。"
    }
    $values = @()
    foreach ($part in $uri.Query.TrimStart('?').Split('&')) {
        if ([string]::IsNullOrWhiteSpace($part)) { continue }
        $pair = $part.Split('=', 2)
        $key = [Uri]::UnescapeDataString($pair[0].Replace('+', ' '))
        if ($key -ne $QueryField) { continue }
        $value = if ($pair.Count -gt 1) { [Uri]::UnescapeDataString($pair[1].Replace('+', ' ')) } else { "" }
        $values += $value
    }
    if ($values.Count -ne 1 -or $values[0] -cne $QueryValue) { throw "批次查询参数与本次输入不一致，未执行打包。" }
}

function Get-UniqueBatchCandidate([int]$TotalRows, [object[]]$Rows, [string]$Fingerprint, [string]$BeforeFingerprint, [string]$QueryField, [string]$QueryValue) {
    if ($TotalRows -ne 1 -or $Rows.Count -ne 1) { throw "搜索结果并非唯一批次，未执行打包。" }
    if ([string]::IsNullOrWhiteSpace($Fingerprint)) { throw "搜索结果缺少可复核数据，未执行打包。" }
    $row = $Rows[0]
    if ([string]::IsNullOrWhiteSpace([string]$row.batchNo) -or [string]::IsNullOrWhiteSpace([string]$row.customer) -or [int]$row.packs -ne 1 -or [int]$row.completed -lt 1) {
        throw "唯一批次的客户、批次号、完成数量或打包控件无法确认，未执行打包。"
    }
    if ($QueryField -eq "batchNo" -and [string]$row.batchNo -cne $QueryValue) { throw "批次号结果与本次精确查询不一致，未执行打包。" }
    if ($QueryField -eq "orderName" -and -not [string]::IsNullOrWhiteSpace($BeforeFingerprint) -and $Fingerprint -ceq $BeforeFingerprint) {
        throw "订单名称搜索后页面结果未发生变化，无法证明当前唯一批次属于本次搜索，未执行打包。请改用批次号精确搜索。"
    }
    return $row
}

function Get-DebugTargets {
    try {
        return @(Invoke-RestMethod -Uri "$debugEndpoint/json/list" -Method Get -TimeoutSec 3)
    } catch {
        return @()
    }
}

function Find-ExpectedTarget {
    $matches = @()
    foreach ($target in (Get-DebugTargets)) {
        if ($target.type -ne "page" -or [string]::IsNullOrWhiteSpace($target.webSocketDebuggerUrl)) { continue }
        try {
            $uri = [Uri]$target.url
            if ($uri.Scheme -eq "http" -and $uri.Host -eq "36.212.60.8" -and $uri.Port -eq 10000 -and $uri.AbsolutePath.TrimEnd("/") -eq "/material/batchRefurbish") {
                $matches += $target
            }
        } catch {}
    }
    if ($matches.Count -ne 1) {
        if ($matches.Count -gt 1) { throw "调试端口上存在多个目标网页，无法安全确定唯一页面。请关闭多余的目标标签后重试。" }
        throw "未找到目标批次页面。请先打开专用浏览器并登录后，再点「连接/检查登录」。"
    }
    return $matches[0]
}

function Connect-Cdp {
    Assert-DedicatedChromeListener
    $target = Find-ExpectedTarget
    $webSocketUri = Assert-LoopbackWebSocketTarget ([string]$target.webSocketDebuggerUrl)
    if ($null -ne $script:socket) {
        try { $script:socket.Dispose() } catch {}
    }
    $script:socket = New-Object System.Net.WebSockets.ClientWebSocket
    $cancel = New-Object System.Threading.CancellationTokenSource
    $cancel.CancelAfter(10000)
    try {
        $script:socket.ConnectAsync($webSocketUri, $cancel.Token).GetAwaiter().GetResult()
    } finally {
        $cancel.Dispose()
    }
    [void](Send-CdpCommand "Page.enable" @{})
    [void](Send-CdpCommand "Runtime.enable" @{})
    [void](Send-CdpCommand "Network.enable" @{})
    return $target
}

function Send-CdpCommand([string]$Method, [hashtable]$Parameters) {
    if ($null -eq $script:socket -or $script:socket.State -ne [System.Net.WebSockets.WebSocketState]::Open) {
        throw "浏览器调试连接已断开。请重新连接后再试。"
    }
    $script:requestId++
    $id = $script:requestId
    $payload = @{ id = $id; method = $Method; params = $Parameters } | ConvertTo-Json -Depth 20 -Compress
    $bytes = [Text.Encoding]::UTF8.GetBytes($payload)
    $segment = New-Object 'System.ArraySegment[byte]' (,$bytes)
    $cancel = New-Object System.Threading.CancellationTokenSource
    $cancel.CancelAfter(15000)
    try {
        $script:socket.SendAsync($segment, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $cancel.Token).GetAwaiter().GetResult()
        while ($true) {
            $stream = New-Object System.IO.MemoryStream
            try {
                do {
                    $buffer = New-Object byte[] 65536
                    $receiveSegment = New-Object 'System.ArraySegment[byte]' (,$buffer)
                    $received = $script:socket.ReceiveAsync($receiveSegment, $cancel.Token).GetAwaiter().GetResult()
                    if ($received.MessageType -eq [System.Net.WebSockets.WebSocketMessageType]::Close) {
                        throw "浏览器调试连接已关闭。"
                    }
                    $stream.Write($buffer, 0, $received.Count)
                } while (-not $received.EndOfMessage)
                $response = [Text.Encoding]::UTF8.GetString($stream.ToArray()) | ConvertFrom-Json
            } finally {
                $stream.Dispose()
            }
            if ($null -eq $response.id) {
                $safeEvent = Convert-CdpEventToSafeRecord $response
                if ($null -ne $safeEvent) { $script:cdpEvents += $safeEvent }
                continue
            }
            if ($response.id -ne $id) { continue }
            if ($null -ne $response.error) { throw "浏览器拒绝操作：$($response.error.message)" }
            return $response.result
        }
    } finally {
        $cancel.Dispose()
    }
}

function Update-BatchSearchRequestState([hashtable]$Requests, [object[]]$Events, [string]$QueryField, [string]$QueryValue) {
    foreach ($event in $Events) {
        $requestId = [string]$event.requestId
        if ($event.method -eq "Network.requestWillBeSent") {
            try { $requestUri = [Uri]$event.url } catch { continue }
            if ($requestUri.Scheme -eq "http" -and $requestUri.Host -eq "36.212.60.8" -and $requestUri.Port -eq 10000 -and $requestUri.AbsolutePath -match '^/(?:prod-api/)?doc/refurbish/batch/groups$') {
                if ([string]::IsNullOrWhiteSpace($requestId)) { throw "批次搜索请求缺少请求标识，未执行打包。" }
                Assert-BatchSearchRequest $event.httpMethod $event.url $QueryField $QueryValue
                if ($Requests.Count -gt 0 -and -not $Requests.ContainsKey($requestId)) { throw "搜索触发了多个批次列表请求，无法确认唯一结果。" }
                $Requests[$requestId] = @{ status = $null; finished = $false; method = [string]$event.httpMethod; url = [string]$event.url }
            }
        } elseif ($Requests.ContainsKey($requestId) -and $event.method -eq "Network.responseReceived") {
            $Requests[$requestId].status = $event.status
        } elseif ($Requests.ContainsKey($requestId) -and $event.method -eq "Network.loadingFailed") {
            throw "批次搜索请求失败，未执行打包。"
        } elseif ($Requests.ContainsKey($requestId) -and $event.method -eq "Network.loadingFinished") {
            $Requests[$requestId].finished = $true
        }
    }
    foreach ($request in $Requests.Values) {
        if ($request.finished) {
            if ($null -eq $request.status -or $request.status -lt 200 -or $request.status -ge 300) { throw "批次搜索接口未成功返回，未执行打包。" }
            return $request
        }
    }
    return $null
}

function Wait-ForBatchListResponse([string]$QueryField, [string]$QueryValue) {
    $requests = @{}
    $deadline = [DateTime]::UtcNow.AddSeconds(30)
    while ([DateTime]::UtcNow -lt $deadline) {
        $completedRequest = Update-BatchSearchRequestState $requests @($script:cdpEvents) $QueryField $QueryValue
        $script:cdpEvents = @()
        if ($null -ne $completedRequest) { return $completedRequest }
        [void](Send-CdpCommand "Runtime.evaluate" @{ expression = "void 0"; returnByValue = $true })
        Start-Sleep -Milliseconds 150
        [System.Windows.Forms.Application]::DoEvents()
    }
    throw "未观察到批次列表搜索完成，未执行打包。"
}

function Invoke-PageExpression([string]$Expression) {
    $result = Send-CdpCommand "Runtime.evaluate" @{
        expression = $Expression
        returnByValue = $true
        awaitPromise = $true
        userGesture = $true
    }
    if ($null -ne $result.exceptionDetails) { throw "页面检查或操作未能完成。" }
    if ($null -eq $result.result -or $null -eq $result.result.value) { throw "页面未返回可识别的结果。" }
    return $result.result.value
}

function Get-PageEvidence {
    $expression = @'
(() => {
  const visible = e => !!(e && e.getClientRects().length && getComputedStyle(e).visibility !== "hidden");
  const inputs = [...document.querySelectorAll("input")].filter(visible);
  const exact = value => inputs.filter(e => (e.getAttribute("placeholder") || "").trim() === value).length;
  const searchButtons = [...document.querySelectorAll("button,a,[role=button]")]
    .filter(e => visible(e) && (e.innerText || "").trim() === "搜索").length;
  const headerRows = [...document.querySelectorAll(".el-table__header-wrapper tr")].filter(visible).map(row => ({
    row,
    cells: [...row.querySelectorAll(":scope > th")].filter(cell => !cell.classList.contains("gutter")).map(e => (e.innerText || "").trim())
  })).filter(item => ["批次号", "客户", "总数", "已完成"].every(name => item.cells.includes(name)));
  const totalRows = [...document.body.innerText.matchAll(/共\s*(\d+)\s*条/g)].map(match => Number(match[1]));
  const header = headerRows.length === 1 ? headerRows[0] : null;
  const table = header ? header.row.closest(".el-table") : null;
  const bodyTable = table ? table.querySelector(".el-table__body-wrapper table.el-table__body") : null;
  const rows = bodyTable ? [...bodyTable.querySelectorAll("tbody > tr")].filter(row => visible(row) && row.closest(".el-table") === table).flatMap(row => {
    const cells = [...row.querySelectorAll(":scope > td")].map(e => (e.innerText || "").trim());
    return cells.length === header.cells.length ? [cells] : [];
  }) : [];
  return {
    href: location.href,
    title: document.title,
    batchInputs: exact("批次号"),
    orderInputs: exact("订单名/编号"),
    searchButtons,
    matchingHeaders: headerRows.length,
    resultFingerprint: JSON.stringify({total: totalRows, rows})
  };
})()
'@
    return Invoke-PageExpression $expression
}

function Assert-ExpectedPage($Evidence) {
    $uri = [Uri]$Evidence.href
    if ($uri.Scheme -ne "http" -or $uri.Host -ne "36.212.60.8" -or $uri.Port -ne 10000 -or $uri.AbsolutePath.TrimEnd("/") -ne "/material/batchRefurbish") {
        throw "当前页面不是已确认的批次管理页面。请人工登录后回到批次翻新页面，再重新连接。"
    }
    if ($Evidence.batchInputs -ne 1 -or $Evidence.orderInputs -ne 1 -or $Evidence.searchButtons -ne 1 -or $Evidence.matchingHeaders -ne 1) {
        throw "页面控件与已确认的批次管理页面结构不一致。为避免误操作，已停止。"
    }
}

function Load-SavedDestination {
    try {
        if (Test-Path -LiteralPath $settingsPath) {
            $settings = Get-Content -LiteralPath $settingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if (-not [string]::IsNullOrWhiteSpace($settings.destination) -and (Test-Path -LiteralPath $settings.destination -PathType Container)) {
                return [IO.Path]::GetFullPath($settings.destination)
            }
        }
    } catch {}
    return Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)) "EasySoftware订单资料"
}

function Save-Destination([string]$Path) {
    New-Item -ItemType Directory -Path $settingsDirectory -Force | Out-Null
    @{ destination = [IO.Path]::GetFullPath($Path) } | ConvertTo-Json -Compress | Set-Content -LiteralPath $settingsPath -Encoding UTF8
}

function Set-Busy([bool]$Busy) {
    foreach ($control in @($openBrowserButton, $connectButton, $searchButton, $browseButton, $destinationBox, $searchKind, $queryBox, $batchList, $handoffButton, $autoProcessCheck)) {
        if ($null -ne $control) { $control.Enabled = -not $Busy }
    }
    $searchButton.Enabled = (-not $Busy) -and ($null -ne $script:python)
    $handoffButton.Enabled = (-not $Busy) -and (-not [string]::IsNullOrWhiteSpace($script:finalDownloadPath))
}

function Show-Error([string]$Message) {
    $statusLabel.Text = "操作停止：" + $Message
    [System.Windows.Forms.MessageBox]::Show($Message, "订单资料下载", "OK", "Warning") | Out-Null
}

if ($SelfTest) {
    $goodSearchUrl = "http://36.212.60.8:10000/prod-api/doc/refurbish/batch/groups?pageNum=1&orderName=%E5%90%88%E6%B3%95%E6%A0%B7%E4%BE%8B"
    Assert-BatchSearchRequest "GET" $goodSearchUrl "orderName" "合法样例"
    foreach ($case in @(
        @{ method = "POST"; url = $goodSearchUrl; field = "orderName"; value = "合法样例" },
        @{ method = "GET"; url = $goodSearchUrl; field = "orderName"; value = "其他样例" },
        @{ method = "GET"; url = "http://example.test:10000/doc/refurbish/batch/groups?orderName=%E5%90%88%E6%B3%95%E6%A0%B7%E4%BE%8B"; field = "orderName"; value = "合法样例" },
        @{ method = "GET"; url = "http://36.212.60.8:10000/doc/refurbish/list?orderName=%E5%90%88%E6%B3%95%E6%A0%B7%E4%BE%8B"; field = "orderName"; value = "合法样例" }
    )) {
        $rejected = $false
        try { Assert-BatchSearchRequest $case.method $case.url $case.field $case.value } catch { $rejected = $true }
        if (-not $rejected) { throw "只读搜索策略自检失败：接受了无效请求。" }
    }

    $safeTarget = Assert-LoopbackWebSocketTarget "ws://127.0.0.1:9222/devtools/page/abc-123"
    foreach ($url in @(
        "ws://192.0.2.1:9222/devtools/page/abc-123",
        "ws://127.0.0.1:9333/devtools/page/abc-123",
        "wss://127.0.0.1:9222/devtools/page/abc-123",
        "ws://127.0.0.1:9222/devtools/browser/abc-123"
    )) {
        $rejected = $false
        try { [void](Assert-LoopbackWebSocketTarget $url) } catch { $rejected = $true }
        if (-not $rejected) { throw "本机 CDP 目标自检失败：接受了非专用端点。" }
    }
    if ($safeTarget.Port -ne $debugPort) { throw "本机 CDP 目标自检失败。" }

    $mockEvent = [pscustomobject]@{
        method = "Network.requestWillBeSent"
        params = [pscustomobject]@{
            requestId = "mock-1"
            request = [pscustomobject]@{
                url = $goodSearchUrl
                method = "GET"
                headers = @{ Cookie = "must-not-be-copied"; Authorization = "must-not-be-copied" }
            }
        }
    }
    $safeEvent = Convert-CdpEventToSafeRecord $mockEvent
    if (($safeEvent.PSObject.Properties.Name -contains "headers") -or ($safeEvent.PSObject.Properties.Name -contains "Cookie") -or $safeEvent.httpMethod -ne "GET") {
        throw "CDP 事件白名单自检失败。"
    }

    $mockResponse = [pscustomobject]@{ method = "Network.responseReceived"; params = [pscustomobject]@{ requestId = "mock-1"; response = [pscustomobject]@{ status = 200; headers = @{ "Set-Cookie" = "must-not-be-copied" } } } }
    $mockFinished = [pscustomobject]@{ method = "Network.loadingFinished"; params = [pscustomobject]@{ requestId = "mock-1" } }
    $safeEvents = @(
        $safeEvent,
        (Convert-CdpEventToSafeRecord $mockResponse),
        (Convert-CdpEventToSafeRecord $mockFinished)
    )
    $completedRequest = Update-BatchSearchRequestState @{} $safeEvents "orderName" "合法样例"
    if ($null -eq $completedRequest -or $completedRequest.status -ne 200) { throw "批次搜索事件流自检失败：成功响应未被识别。" }
    foreach ($badResponse in @(
        [pscustomobject]@{ method = "Network.responseReceived"; params = [pscustomobject]@{ requestId = "mock-1"; response = [pscustomobject]@{ status = 500 } } },
        [pscustomobject]@{ method = "Network.responseReceived"; params = [pscustomobject]@{ requestId = "mock-1"; response = [pscustomobject]@{} } }
    )) {
        $badEvents = @($safeEvent, (Convert-CdpEventToSafeRecord $badResponse), (Convert-CdpEventToSafeRecord $mockFinished))
        $rejected = $false
        try { [void](Update-BatchSearchRequestState @{} $badEvents "orderName" "合法样例") } catch { $rejected = $true }
        if (-not $rejected) { throw "批次搜索事件流自检失败：接受了非成功或缺失状态码的响应。" }
    }

    $validRow = [pscustomobject]@{ batchNo = "fp-test"; customer = "模拟客户"; completed = 1; total = 1; packs = 1 }
    [void](Get-UniqueBatchCandidate 1 @($validRow) "fingerprint" "fingerprint" "batchNo" "fp-test")
    foreach ($case in @(
        @{ total = 2; rows = @($validRow, $validRow); fingerprint = "fingerprint"; field = "batchNo"; value = "fp-test" },
        @{ total = 1; rows = @(); fingerprint = "fingerprint"; field = "batchNo"; value = "fp-test" },
        @{ total = 1; rows = @($validRow); fingerprint = ""; field = "batchNo"; value = "fp-test" },
        @{ total = 1; rows = @($validRow); fingerprint = "fingerprint"; field = "batchNo"; value = "other-batch" },
        @{ total = 1; rows = @($validRow); fingerprint = "same-page"; before = "same-page"; field = "orderName"; value = "客户订单" }
    )) {
        $rejected = $false
        try { [void](Get-UniqueBatchCandidate $case.total $case.rows $case.fingerprint $case.before $case.field $case.value) } catch { $rejected = $true }
        if (-not $rejected) { throw "唯一批次校验自检失败：接受了多条、缺失或与精确批次号查询不符的结果。" }
    }
    Write-Host "Order download policy self-test passed."
    exit 0
}

$script:python = Find-Python
$form = New-Object System.Windows.Forms.Form
$form.Text = "订单资料下载前置工作流"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(940, 690)
$form.MinimumSize = New-Object System.Drawing.Size(900, 640)
$form.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)

$title = New-Object System.Windows.Forms.Label
$title.Text = "订单资料下载前置工作流"
$title.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 17, [System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(22, 16)
$title.AutoSize = $true
$form.Controls.Add($title)

$guide = New-Object System.Windows.Forms.Label
$guide.Text = "人工登录专用浏览器 → 搜索并确认唯一批次 → 下载校验 ZIP → 按需自动启动候选审查"
$guide.Location = New-Object System.Drawing.Point(24, 56)
$guide.Size = New-Object System.Drawing.Size(865, 34)
$guide.BackColor = [System.Drawing.Color]::FromArgb(238, 244, 255)
$guide.ForeColor = [System.Drawing.Color]::FromArgb(49, 82, 138)
$guide.Padding = New-Object System.Windows.Forms.Padding(6, 7, 6, 2)
$form.Controls.Add($guide)

$destinationLabel = New-Object System.Windows.Forms.Label
$destinationLabel.Text = "下载位置"
$destinationLabel.Location = New-Object System.Drawing.Point(25, 108)
$destinationLabel.Size = New-Object System.Drawing.Size(80, 28)
$form.Controls.Add($destinationLabel)

$destinationBox = New-Object System.Windows.Forms.TextBox
$destinationBox.Location = New-Object System.Drawing.Point(105, 104)
$destinationBox.Size = New-Object System.Drawing.Size(650, 30)
$destinationBox.Text = Load-SavedDestination
$form.Controls.Add($destinationBox)

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "选择目录"
$browseButton.Location = New-Object System.Drawing.Point(765, 102)
$browseButton.Size = New-Object System.Drawing.Size(120, 34)
$form.Controls.Add($browseButton)

$openBrowserButton = New-Object System.Windows.Forms.Button
$openBrowserButton.Text = "打开人工登录窗口"
$openBrowserButton.Location = New-Object System.Drawing.Point(26, 151)
$openBrowserButton.Size = New-Object System.Drawing.Size(180, 36)
$form.Controls.Add($openBrowserButton)

$connectButton = New-Object System.Windows.Forms.Button
$connectButton.Text = "连接 / 检查登录"
$connectButton.Location = New-Object System.Drawing.Point(216, 151)
$connectButton.Size = New-Object System.Drawing.Size(180, 36)
$connectButton.BackColor = [System.Drawing.Color]::FromArgb(22, 93, 255)
$connectButton.ForeColor = [System.Drawing.Color]::White
$connectButton.FlatStyle = "Flat"
$form.Controls.Add($connectButton)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "等待连接。登录过程由你在专用浏览器窗口中手动完成。"
$statusLabel.Location = New-Object System.Drawing.Point(410, 157)
$statusLabel.Size = New-Object System.Drawing.Size(480, 28)
$form.Controls.Add($statusLabel)

$searchKind = New-Object System.Windows.Forms.ComboBox
$searchKind.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$searchKind.Location = New-Object System.Drawing.Point(26, 207)
$searchKind.Size = New-Object System.Drawing.Size(160, 30)
[void]$searchKind.Items.Add("订单名称/编号")
[void]$searchKind.Items.Add("批次号")
$searchKind.SelectedIndex = 0
$form.Controls.Add($searchKind)

$queryBox = New-Object System.Windows.Forms.TextBox
$queryBox.Location = New-Object System.Drawing.Point(196, 207)
$queryBox.Size = New-Object System.Drawing.Size(500, 30)
$queryBox.Add_KeyDown({ if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $searchButton.PerformClick() } })
$form.Controls.Add($queryBox)

$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Text = "搜索并下载"
$searchButton.Location = New-Object System.Drawing.Point(708, 204)
$searchButton.Size = New-Object System.Drawing.Size(175, 36)
$form.Controls.Add($searchButton)

$batchList = New-Object System.Windows.Forms.ListView
$batchList.Location = New-Object System.Drawing.Point(26, 254)
$batchList.Size = New-Object System.Drawing.Size(858, 240)
$batchList.View = [System.Windows.Forms.View]::Details
$batchList.FullRowSelect = $true
$batchList.GridLines = $true
$batchList.MultiSelect = $false
[void]$batchList.Columns.Add("批次号", 250)
[void]$batchList.Columns.Add("客户", 240)
[void]$batchList.Columns.Add("已完成 / 总数", 140)
[void]$batchList.Columns.Add("可下载", 110)
$form.Controls.Add($batchList)

$handoffButton = New-Object System.Windows.Forms.Button
$handoffButton.Text = "交给候选审查"
$handoffButton.Location = New-Object System.Drawing.Point(26, 510)
$handoffButton.Size = New-Object System.Drawing.Size(180, 38)
$handoffButton.Enabled = $false
$form.Controls.Add($handoffButton)

$openFolderButton = New-Object System.Windows.Forms.Button
$openFolderButton.Text = "打开下载文件夹"
$openFolderButton.Location = New-Object System.Drawing.Point(218, 510)
$openFolderButton.Size = New-Object System.Drawing.Size(180, 38)
$form.Controls.Add($openFolderButton)

$autoProcessCheck = New-Object System.Windows.Forms.CheckBox
$autoProcessCheck.Text = "下载校验后自动开始候选审查处理"
$autoProcessCheck.Location = New-Object System.Drawing.Point(420, 514)
$autoProcessCheck.Size = New-Object System.Drawing.Size(440, 30)
$autoProcessCheck.Checked = $false
$form.Controls.Add($autoProcessCheck)

$footnote = New-Object System.Windows.Forms.Label
$footnote.Text = "一次点击先搜索；只有完整结果唯一且行内控件可确认时才会打包一次。无结果、多结果或页面结构异常时停止。勾选后会启动原候选审查流程，仍按流程等待人工审核。"
$footnote.ForeColor = [System.Drawing.Color]::DimGray
$footnote.Location = New-Object System.Drawing.Point(28, 566)
$footnote.Size = New-Object System.Drawing.Size(850, 48)
$form.Controls.Add($footnote)

$browseButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "选择下载订单资料的位置"
    $dialog.ShowNewFolderButton = $true
    if (Test-Path -LiteralPath $destinationBox.Text -PathType Container) { $dialog.SelectedPath = $destinationBox.Text }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $destinationBox.Text = $dialog.SelectedPath
        try { Save-Destination $dialog.SelectedPath } catch { Show-Error $_.Exception.Message }
    }
})

$openBrowserButton.Add_Click({
    try {
        $dedicatedProcessIds = @(Get-DedicatedChromeProcessIds)
        $targets = @()
        if ($dedicatedProcessIds.Count -gt 0) {
            Assert-DedicatedChromeListener
            $targets = @(Get-DebugTargets)
        }
        $existing = @($targets | Where-Object { $_.type -eq "page" -and $_.url -like "http://36.212.60.8:10000/material/batchRefurbish*" })
        if ($existing.Count -eq 1) {
            $statusLabel.Text = "专用窗口已打开。请在窗口中人工登录，随后点击「连接 / 检查登录」。"
            return
        }
        if ($existing.Count -gt 1) { throw "专用浏览器中存在多个目标页面。请关闭多余标签后重试。" }
        $chromeCandidates = @()
        $chromeCommand = Get-Command "chrome.exe" -ErrorAction SilentlyContinue
        if ($null -ne $chromeCommand) { $chromeCandidates += $chromeCommand.Source }
        foreach ($base in @(${env:ProgramFiles}, ${env:ProgramFiles(x86)}, $env:LOCALAPPDATA)) {
            if (-not [string]::IsNullOrWhiteSpace($base)) { $chromeCandidates += (Join-Path $base "Google\Chrome\Application\chrome.exe") }
        }
        $chrome = $chromeCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
        if (-not $chrome) { throw "找不到 Chrome。请安装 Chrome，或先用已配置的专用窗口启动调试端口 9222。" }
        New-Item -ItemType Directory -Path $profilePath -Force | Out-Null
        Start-Process -FilePath $chrome -ArgumentList @("--remote-debugging-port=$debugPort", "--remote-debugging-address=127.0.0.1", "--user-data-dir=$(Quote-Argument $profilePath)", (Quote-Argument $targetUrl)) -WindowStyle Normal | Out-Null
        $statusLabel.Text = "已打开独立浏览器。请完成手动登录，再点击「连接 / 检查登录」。"
    } catch { Show-Error $_.Exception.Message }
})

$connectButton.Add_Click({
    try {
        $target = Connect-Cdp
        $evidence = Get-PageEvidence
        try {
            Assert-ExpectedPage $evidence
            $statusLabel.Text = "已连接目标批次页面；登录态和筛选控件可用。"
        } catch {
            $statusLabel.Text = "已连上浏览器，请先在窗口内完成登录并回到批次页面。"
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "需要人工登录", "OK", "Information") | Out-Null
        }
    } catch { Show-Error $_.Exception.Message }
})

$searchButton.Add_Click({
    try {
        if (-not [string]::IsNullOrWhiteSpace([string]$script:pendingDownloadStaging)) {
            throw "上一笔打包的下载状态仍未确认。请先检查暂存目录中的下载是否完成：$script:pendingDownloadStaging。为避免串批，当前会话不会再次触发打包。"
        }
        if ([string]::IsNullOrWhiteSpace($queryBox.Text)) { throw "请输入订单名称、编号或批次号。" }
        if ($null -eq $script:socket -or $script:socket.State -ne [System.Net.WebSockets.WebSocketState]::Open) { throw "请先连接并检查登录状态。" }
        if ($null -eq $script:python) { throw "找不到可用的 Python 运行时，无法校验下载 ZIP。" }
        $evidence = Get-PageEvidence
        Assert-ExpectedPage $evidence
        $readRows = @'
(() => {
  const visible = e => !!(e && e.getClientRects().length && getComputedStyle(e).visibility !== "hidden");
  const totals = [...document.body.innerText.matchAll(/共\s*(\d+)\s*条/g)].map(match => Number(match[1]));
  const headerRows = [...document.querySelectorAll(".el-table__header-wrapper tr")].filter(visible).map(row => {
    const cells = [...row.querySelectorAll(":scope > th")].filter(cell => !cell.classList.contains("gutter")).map(e => (e.innerText || "").trim());
    return {row, cells};
  }).filter(item => ["批次号", "客户", "总数", "已完成"].every(name => item.cells.includes(name)));
  if (headerRows.length !== 1 || totals.length !== 1) return {known:false, rows:[], totalRows:null};
  const headers = headerRows[0].cells;
  const table = headerRows[0].row.closest(".el-table");
  const bodyTable = table && table.querySelector(".el-table__body-wrapper table.el-table__body");
  if (!bodyTable) return {known:false, rows:[], totalRows:null};
  const indexes = Object.fromEntries(["批次号", "客户", "总数", "已完成"].map(name => [name, headers.indexOf(name)]));
  const dataRows = [...bodyTable.querySelectorAll("tbody > tr")].filter(row => visible(row) && row.closest(".el-table") === table);
  const actionRows = [...table.querySelectorAll(".el-table__fixed-right .el-table__fixed-body-wrapper table.el-table__body tbody > tr")].filter(visible);
  const rows = dataRows.flatMap((row, rowIndex) => {
    const cells = [...row.querySelectorAll(":scope > td")].map(e => (e.innerText || "").trim());
    if (cells.length !== headers.length || cells.length === 0) return [];
    const completed = Number.parseInt(cells[indexes["已完成"]].replace(/[^0-9]/g, ""), 10);
    const total = Number.parseInt(cells[indexes["总数"]].replace(/[^0-9]/g, ""), 10);
    const batchNo = cells[indexes["批次号"]];
    const customer = cells[indexes["客户"]];
    if (!batchNo || !customer || !Number.isFinite(completed) || !Number.isFinite(total)) return [];
    const actionRow = actionRows[rowIndex] || row;
    const packs = [...actionRow.querySelectorAll("button,a,[role=button]")].filter(e => visible(e) && (e.innerText || "").trim() === "打包").length;
    return [{batchNo, customer, completed, total, packs}];
  });
  const fingerprintRows = dataRows.flatMap(row => {
    const cells = [...row.querySelectorAll(":scope > td")].map(e => (e.innerText || "").trim());
    return cells.length === headers.length ? [cells] : [];
  });
  return {known:true, rows, totalRows:totals[0], fingerprint:JSON.stringify({total:totals,rows:fingerprintRows})};
})()
'@
        $beforeSearch = Invoke-PageExpression $readRows
        if (-not $beforeSearch.known) { throw "搜索前无法识别当前批次表格，未提交搜索。" }
        $beforeFingerprint = [string]$beforeSearch.fingerprint
        Set-Busy $true
        $script:finalDownloadPath = $null
        $handoffButton.Enabled = $false
        $batchList.Items.Clear()
        $statusLabel.Text = "正在提交搜索条件…"
        $placeholder = if ($searchKind.SelectedIndex -eq 1) { "批次号" } else { "订单名/编号" }
        $queryField = if ($searchKind.SelectedIndex -eq 1) { "batchNo" } else { "orderName" }
        $queryValue = $queryBox.Text.Trim()
        $queryJson = ConvertTo-Json -InputObject $queryBox.Text.Trim() -Compress
        $placeholderJson = ConvertTo-Json -InputObject $placeholder -Compress
        $expression = @"
(() => {
  const visible = e => !!(e && e.getClientRects().length && getComputedStyle(e).visibility !== "hidden");
  const fields = [...document.querySelectorAll("input")].filter(e => visible(e) && (e.getAttribute("placeholder") || "").trim() === $placeholderJson);
  const buttons = [...document.querySelectorAll("button,a,[role=button]")].filter(e => visible(e) && (e.innerText || "").trim() === "搜索");
  if (fields.length !== 1 || buttons.length !== 1) return {ok:false};
  const field = fields[0];
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value").set;
  setter.call(field, $queryJson);
  field.dispatchEvent(new Event("input", {bubbles:true}));
  field.dispatchEvent(new Event("change", {bubbles:true}));
  buttons[0].click();
  return {ok:true};
})()
"@
        $script:cdpEvents = @()
        $submitted = Invoke-PageExpression $expression
        if (-not $submitted.ok) { throw "搜索控件与已确认页面结构不一致，已停止。" }
        $statusLabel.Text = "搜索请求已提交，正在确认接口完成和完整结果数…"
        $form.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
        $searchResponse = Wait-ForBatchListResponse $queryField $queryValue
        Start-Sleep -Milliseconds 1800
        $result = Invoke-PageExpression $readRows
        if (-not $result.known) { throw "无法唯一识别结果表格结构，为避免误点已停止。" }
        $script:searchRows = @($result.rows)
        $batchList.Items.Clear()
        foreach ($row in $script:searchRows) {
            $item = New-Object System.Windows.Forms.ListViewItem([string]$row.batchNo)
            [void]$item.SubItems.Add([string]$row.customer)
            [void]$item.SubItems.Add("$($row.completed) / $($row.total)")
            [void]$item.SubItems.Add($(if ($row.packs -eq 1) { "可打包" } else { "控件异常" }))
            $item.Tag = $row
            [void]$batchList.Items.Add($item)
        }
        if ([int]$result.totalRows -ne $batchList.Items.Count) {
            $statusLabel.Text = "页面总记录数与当前可见行数不一致，未点击打包。请缩小搜索范围后重试。"
            return
        }
        if ([int]$result.totalRows -eq 0) {
            $statusLabel.Text = "搜索结果为零，没有执行任何打包操作。"
            return
        }
        if ([int]$result.totalRows -ne 1 -or $batchList.Items.Count -ne 1) {
            $statusLabel.Text = "搜索命中 $($result.totalRows) 条；为避免批次歧义，没有执行打包。请缩小搜索范围后重试。"
            return
        }
        $selectedRow = Get-UniqueBatchCandidate ([int]$result.totalRows) @($script:searchRows) ([string]$result.fingerprint) $beforeFingerprint $queryField $queryValue
        Invoke-SelectedBatchDownload $selectedRow
    } catch { Show-Error $_.Exception.Message }
    finally { Set-Busy $false }
})

function Invoke-SelectedBatchDownload($selectedRow) {
    $staging = $null
    $packClicked = $false
    $downloadFinalized = $false
    $preserveStaging = $false
    try {
        $destination = $destinationBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($destination)) { throw "请先设置下载位置。" }
        Save-Destination $destination
        $evidence = Get-PageEvidence
        Assert-ExpectedPage $evidence
        $staging = Join-Path ([IO.Path]::GetTempPath()) ("EasySoftware-OrderDownload-" + [Guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Path $staging -Force | Out-Null
        [void](Send-CdpCommand "Page.setDownloadBehavior" @{ behavior = "allow"; downloadPath = $staging })
        $rowJson = @{
            batchNo = [string]$selectedRow.batchNo
            customer = [string]$selectedRow.customer
            completed = [int]$selectedRow.completed
            total = [int]$selectedRow.total
        } | ConvertTo-Json -Depth 5 -Compress
        $clickExpression = @"
(() => {
  const expected = $rowJson;
  const visible = e => !!(e && e.getClientRects().length && getComputedStyle(e).visibility !== "hidden");
  const totals = [...document.body.innerText.matchAll(/共\s*(\d+)\s*条/g)].map(match => Number(match[1]));
  const headerRows = [...document.querySelectorAll(".el-table__header-wrapper tr")].filter(visible).map(row => ({row,cells:[...row.querySelectorAll(":scope > th")].filter(cell => !cell.classList.contains("gutter")).map(e => (e.innerText || "").trim())}))
    .filter(item => ["批次号", "客户", "总数", "已完成"].every(name => item.cells.includes(name)));
  if (headerRows.length !== 1 || totals.length !== 1 || totals[0] !== 1) return {ok:false};
  const table = headerRows[0].row.closest(".el-table");
  const bodyTable = table && table.querySelector(".el-table__body-wrapper table.el-table__body");
  if (!bodyTable) return {ok:false};
  const headers = [headerRows[0].cells];
  const indexes = Object.fromEntries(["批次号", "客户", "总数", "已完成"].map(name => [name, headers[0].indexOf(name)]));
  const rows = [...bodyTable.querySelectorAll("tbody > tr")].filter(row => visible(row) && row.closest(".el-table") === table && row.querySelectorAll(":scope > td").length === headers[0].length);
  if (rows.length !== 1) return {ok:false};
  const matches = rows.filter(row => {
    const cells = [...row.querySelectorAll(":scope > td")].map(e => (e.innerText || "").trim());
    if (cells.length !== headers[0].length) return false;
    const completed = Number.parseInt(cells[indexes["已完成"]].replace(/[^0-9]/g, ""), 10);
    const total = Number.parseInt(cells[indexes["总数"]].replace(/[^0-9]/g, ""), 10);
    return cells[indexes["批次号"]] === expected.batchNo && cells[indexes["客户"]] === expected.customer && completed === expected.completed && total === expected.total;
  });
  if (matches.length !== 1) return {ok:false};
  const actionRows = [...table.querySelectorAll(".el-table__fixed-right .el-table__fixed-body-wrapper table.el-table__body tbody > tr")].filter(visible);
  const actionRow = actionRows[rows.indexOf(matches[0])] || matches[0];
  const packs = [...actionRow.querySelectorAll("button,a,[role=button]")].filter(e => visible(e) && (e.innerText || "").trim() === "打包");
  if (packs.length !== 1) return {ok:false};
  packs[0].click();
  return {ok:true};
})()
"@
        # Mark the request as potentially dispatched before the CDP round trip.
        # The click may reach the page even if its acknowledgement is lost.
        $packClicked = $true
        $clicked = Invoke-PageExpression $clickExpression
        if (-not $clicked.ok) {
            # The page expression returns {ok:false} before invoking click().
            $packClicked = $false
            throw "批次行或打包按钮与搜索时不一致，已停止，没有触发下载。"
        }
        $statusLabel.Text = "已验证唯一批次并点击其打包一次，正在等待并检查 ZIP…"
        $form.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
        $downloadStartedAt = [DateTime]::UtcNow
        $nextProgressAt = $downloadStartedAt.AddSeconds(10)
        $deadline = $downloadStartedAt.AddMinutes(15)
        $stableName = $null
        $stableLength = -1L
        $stablePasses = 0
        $source = $null
        while ([DateTime]::UtcNow -lt $deadline) {
            Start-Sleep -Milliseconds 700
            [System.Windows.Forms.Application]::DoEvents()
            if ([DateTime]::UtcNow -ge $nextProgressAt) {
                $elapsedSeconds = [int]([DateTime]::UtcNow - $downloadStartedAt).TotalSeconds
                $statusLabel.Text = "批次资料仍在打包或下载（已等待 $elapsedSeconds 秒）；完成后会自动校验 ZIP。"
                $form.Refresh()
                $nextProgressAt = [DateTime]::UtcNow.AddSeconds(10)
            }
            $files = @(Get-ChildItem -LiteralPath $staging -File -Force -ErrorAction SilentlyContinue)
            $partials = @($files | Where-Object { $_.Name.EndsWith(".crdownload", [StringComparison]::OrdinalIgnoreCase) })
            $complete = @($files | Where-Object { -not $_.Name.EndsWith(".crdownload", [StringComparison]::OrdinalIgnoreCase) })
            if ($complete.Count -gt 1) { throw "浏览器生成多个下载文件，无法确定目标 ZIP。" }
            if ($complete.Count -eq 1 -and $partials.Count -eq 0) {
                if ($complete[0].Name -eq $stableName -and $complete[0].Length -eq $stableLength) { $stablePasses++ }
                else { $stableName = $complete[0].Name; $stableLength = $complete[0].Length; $stablePasses = 0 }
                if ($stablePasses -ge 2) { $source = $complete[0].FullName; break }
            }
        }
        if ($null -eq $source) { throw "等待浏览器下载超时，或下载文件未稳定完成。" }
        $coreResultText = & $script:python $corePath "finalize" "--source" $source "--destination" $destination "--customer" ([string]$selectedRow.customer)
        if ($LASTEXITCODE -ne 0) {
            try { $coreResult = ($coreResultText -join "`n") | ConvertFrom-Json; throw [string]$coreResult.error }
            catch { if ($_.Exception.Message) { throw $_.Exception.Message }; throw "ZIP 检查或安全命名失败。" }
        }
        $coreResult = ($coreResultText -join "`n") | ConvertFrom-Json
        if ($coreResult.status -ne "OK" -or [string]::IsNullOrWhiteSpace($coreResult.path)) { throw "ZIP 校验或归档命名未返回成功结果。" }
        $script:finalDownloadPath = [string]$coreResult.path
        $downloadFinalized = $true
        $script:pendingDownloadStaging = $null
        $handoffButton.Enabled = $true
        $statusLabel.Text = "下载并校验完成：$($coreResult.filename)（ZIP 内 $($coreResult.count) 个文件）。"
        if ($autoProcessCheck.Checked) {
            Invoke-CandidateReviewHandoff -AutoStart
        } else {
            $statusLabel.Text += "可点击「交给候选审查」手动启动。"
        }
    } catch {
        if ($packClicked -and -not $downloadFinalized) {
            $preserveStaging = $true
            $script:pendingDownloadStaging = $staging
            throw "$($_.Exception.Message)`n下载状态未能确认，已保留暂存目录：$staging。请先核实此目录，不要再次触发打包。"
        }
        throw
    }
    finally {
        if (-not $preserveStaging) {
            if ($null -ne $script:socket -and $script:socket.State -eq [System.Net.WebSockets.WebSocketState]::Open) {
                try { [void](Send-CdpCommand "Page.setDownloadBehavior" @{ behavior = "default" }) } catch {}
            }
            if ($null -ne $staging -and (Test-Path -LiteralPath $staging)) {
                Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

function Invoke-CandidateReviewHandoff([switch]$AutoStart) {
    if ([string]::IsNullOrWhiteSpace($script:finalDownloadPath) -or -not (Test-Path -LiteralPath $script:finalDownloadPath -PathType Leaf)) { throw "没有已校验的下载文件可交给候选审查。" }
    if (-not (Test-Path -LiteralPath $mainGuiPath -PathType Leaf)) { throw "候选审查主界面不存在：$mainGuiPath" }
    $arguments = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", (Quote-Argument $mainGuiPath), "-InputPath", (Quote-Argument $script:finalDownloadPath))
    if ($AutoStart) { $arguments += "-AutoStart" }
    Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -WindowStyle Normal | Out-Null
    if ($AutoStart) {
        $handoffButton.Enabled = $false
        $statusLabel.Text = "已自动启动候选审查处理；流程会在需要人工审核时停下。"
    } else {
        $statusLabel.Text = "已打开候选审查界面并预填 ZIP 路径；尚未启动材料处理。"
    }
}

$handoffButton.Add_Click({
    try {
        Invoke-CandidateReviewHandoff
    } catch { Show-Error $_.Exception.Message }
})

$openFolderButton.Add_Click({
    try {
        $path = $destinationBox.Text.Trim()
        if (-not (Test-Path -LiteralPath $path -PathType Container)) { throw "下载目录还不存在。" }
        Start-Process -FilePath "explorer.exe" -ArgumentList (Quote-Argument $path) | Out-Null
    } catch { Show-Error $_.Exception.Message }
})

$form.Add_FormClosed({
    if ($null -ne $script:socket) {
        try {
            if ($script:socket.State -eq [System.Net.WebSockets.WebSocketState]::Open) {
                $script:socket.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, "closed", [System.Threading.CancellationToken]::None).GetAwaiter().GetResult()
            }
        } catch {}
        try { $script:socket.Dispose() } catch {}
    }
})

$form.Add_Shown({
    if ($null -eq $script:python) {
        $searchButton.Enabled = $false
        $statusLabel.Text = "未找到 Python 运行时，无法安全校验下载文件。"
    }
})

[void]$form.ShowDialog()

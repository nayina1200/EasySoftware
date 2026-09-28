$script:CoverRuleDataPath = Join-Path $PSScriptRoot "cover_rules.json"
$script:CoverRuleData = Get-Content -LiteralPath $script:CoverRuleDataPath -Raw -Encoding UTF8 | ConvertFrom-Json

function Normalize-CoverText([string]$Text) {
    $value = if ($null -eq $Text) { "" } else { [regex]::Replace($Text, '\s+', '') }
    foreach ($pair in $script:CoverRuleData.wrappers) {
        $opening = [string]$pair[0]
        $closing = [string]$pair[1]
        if ($value.StartsWith($opening, [StringComparison]::Ordinal) -and
            $value.EndsWith($closing, [StringComparison]::Ordinal) -and
            $value.Length -ge ($opening.Length + $closing.Length)) {
            $value = $value.Substring($opening.Length, $value.Length - $opening.Length - $closing.Length)
            break
        }
    }
    return $value
}

function Get-CoverClassification([string]$Text, [string]$Title = "") {
    $normalized = Normalize-CoverText $Text
    $result = [ordered]@{ normalized = $normalized; kind = "other"; reason = "未匹配完整软件名称或说明书标签" }
    if ($normalized.Length -eq 0) {
        $result.reason = "内容为空"
        return [pscustomobject]$result
    }

    $versionless = [regex]::Replace($normalized, [string]$script:CoverRuleData.cover_version_pattern, "")
    $expectedTitle = if ([string]::IsNullOrWhiteSpace($Title)) { "" } else {
        [regex]::Replace((Normalize-CoverText $Title), [string]$script:CoverRuleData.cover_version_pattern, "")
    }
    $hadOrdinal = $false
    if ($expectedTitle.Length -gt 0) {
        $candidate = ([regex]::new([string]$script:CoverRuleData.cover_ordinal_pattern)).Replace($versionless, "", 1)
        $hadOrdinal = -not [string]::Equals($candidate, $versionless, [StringComparison]::Ordinal)
        $versionless = $candidate
    }
    if ($expectedTitle.Length -gt 0 -and [string]::Equals($versionless, $expectedTitle, [StringComparison]::Ordinal)) {
        $result.kind = "cover_title"
        $reasons = @()
        if ($hadOrdinal) { $reasons += "封面编号可移除" }
        if (-not [string]::Equals([regex]::Replace($normalized, [string]$script:CoverRuleData.cover_version_pattern, ""), $normalized, [StringComparison]::Ordinal)) { $reasons += "末尾版本号可移除" }
        $result.reason = if ($reasons.Count -gt 0) { "与软件全称完整匹配（$($reasons -join '；')）" } else { "与软件全称完整匹配" }
        return [pscustomobject]$result
    }

    foreach ($term in $script:CoverRuleData.exclude_terms) {
        if ($normalized.Contains([string]$term)) {
            $result.kind = "excluded"
            $result.reason = ('命中排除词“{0}”' -f $term)
            return [pscustomobject]$result
        }
    }
    foreach ($mark in $script:CoverRuleData.sentence_punctuation) {
        if ($normalized.Contains([string]$mark)) {
            $result.reason = "含句末标点或文档编号字符，不是完整标签"
            return [pscustomobject]$result
        }
    }
    if ($normalized.Length -gt [int]$script:CoverRuleData.maximum_label_length) {
        $result.reason = "文本过长，不是完整封面标签"
        return [pscustomobject]$result
    }
    foreach ($label in $script:CoverRuleData.full_labels) {
        if ([string]::Equals($normalized, [string]$label, [StringComparison]::Ordinal)) {
            $result.kind = "manual_label"
            $result.reason = ('完整标签“{0}”' -f $normalized)
            return [pscustomobject]$result
        }
    }

    foreach ($documentType in @($script:CoverRuleData.document_types | Sort-Object { ([string]$_).Length } -Descending)) {
        $type = [string]$documentType
        if (-not $normalized.EndsWith($type, [StringComparison]::Ordinal)) { continue }
        $purpose = $normalized.Substring(0, $normalized.Length - $type.Length)
        foreach ($prefix in @($script:CoverRuleData.allowed_prefixes | Sort-Object { ([string]$_).Length } -Descending)) {
            $prefixValue = [string]$prefix
            if ($purpose.StartsWith($prefixValue, [StringComparison]::Ordinal)) {
                $purpose = $purpose.Substring($prefixValue.Length)
                break
            }
        }
        foreach ($purposeTerm in $script:CoverRuleData.purpose_terms) {
            if ([string]::Equals($purpose, [string]$purposeTerm, [StringComparison]::Ordinal)) {
                $result.kind = "manual_label"
                $result.reason = ('用途词“{0}”与类型词“{1}”完整匹配' -f $purposeTerm, $type)
                return [pscustomobject]$result
            }
        }
    }
    $result.reason = "缺少明确文档用途词"
    return [pscustomobject]$result
}

function Test-ManualLabel([string]$Text) {
    return (Get-CoverClassification $Text).kind -eq "manual_label"
}

function Test-CoverTitle([string]$Text, [string]$Title) {
    return (Get-CoverClassification $Text $Title).kind -eq "cover_title"
}

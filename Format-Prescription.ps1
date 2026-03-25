<#
.FILENAME
    Format-Prescription.ps1
.VERSION
    1.1.2
.DESCRIPTION
    メモ帳での「?」化を徹底的に防ぐため、.NETのDataObjectを使用して
    プレーンテキスト(Unicode)としてクリップボードに強制上書きします。
#>

# 1. クリップボードからテキストを取得
$inputContent = Get-Clipboard -Raw
if ([string]::IsNullOrWhiteSpace($inputContent)) {
    Write-Warning "クリップボードが空です。"
    exit
}

# --- 前処理：制御文字の除去 ---
$inputContent = $inputContent -replace "[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", ""

$lines = $inputContent -split "`r?`n" | ForEach-Object {
    $line = $_.Trim()
    $line = $line -replace "^商品名\s*", ""
    $line = $line -replace "[（\(].*?として[）\)]", ""
    
    if ($line -match "^(処方箋使用期限|--|<)") { return $null }
    if ([string]::IsNullOrWhiteSpace($line)) { return $null }
    
    $line
} | Where-Object { $_ -ne $null }

# ブロック分割処理
$processedText = $lines -join "`r?`n"
$blocks = $processedText -split "(?=処方日[:：])" | Where-Object { $_ -match "処方日" }

$finalOutput = New-Object System.Collections.Generic.List[string]

foreach ($block in $blocks) {
    $blockLines = $block -split "`r?`n" | Where-Object { $_ -notmatch "^処方日" -and ![string]::IsNullOrWhiteSpace($_) }
    if ($blockLines.Count -eq 0) { continue }

    $maxLen = ($blockLines | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum
    $mergedLines = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $blockLines.Count; $i++) {
        $current = $blockLines[$i]
        if ($i -lt $blockLines.Count - 1 -and $current.Length -lt $maxLen) {
            $mergedLines.Add($current + $blockLines[$i+1])
            $i++ 
        } else {
            $mergedLines.Add($current)
        }
    }

    $resultBlock = New-Object System.Collections.Generic.List[string]
    foreach ($line in $mergedLines) {
        $l = $line
        if ($l -match "^分") {
            if ($l -match "\s{2,}") {
                $l = $l -replace "\s{2,}.*$", ""
                $resultBlock.Add($l); break 
            }
            if ($l -match "^分[1１]\s*") {
                $l = $l -replace "^分[1１]\s*", "" -replace "食後", ""
            } elseif ($l -match "^分[3３]") {
                $l = $l -replace "毎食後", ""
            } else {
                $l = $l -replace "食後", ""
            }
        }
        elseif ($l -match "時") {
            if ($l -match "\s{2,}") {
                $l = $l -replace "\s{2,}.*$", ""
                $resultBlock.Add($l); break 
            }
        }
        elseif ($l -match "^外）") {
            $l = $l -replace "^外）", ""
            if ($l -match "\s{2,}") {
                $l = $l -replace "\s{2,}.*$", ""
                $currentIdx = $mergedLines.IndexOf($line)
                if ($currentIdx -lt $mergedLines.Count - 1) {
                    $l += ($mergedLines[($currentIdx + 1)..($mergedLines.Count - 1)] -join "")
                }
                $resultBlock.Add($l); break
            }
        }
        $resultBlock.Add($l)
    }
    $finalOutput.Add(($resultBlock -join "`r?`n"))
}

# 最終整形
$rawResult = ($finalOutput -join "`r?`n") -replace "　", " " -replace "\s{2,}", " "
$finalLines = $rawResult -split "`r?`n" | Where-Object { ![string]::IsNullOrWhiteSpace($_) }
$resultText = $finalLines -join "`r?`n"

# --- クリップボードへの「強制プレーンテキスト」書き込み ---
if (![string]::IsNullOrEmpty($resultText)) {
    try {
        # PowerShell 5.1のSet-Clipboardを使わず、Windows FormsのClipboardクラスを使用
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.Clipboard]::SetText($resultText)
        Write-Host "--- v1.1.2 完了 ---" -ForegroundColor Green
    } catch {
        # 万が一Formsが失敗した場合の予備（これでもダメなら環境依存の可能性が高いです）
        $resultText | clip
    }
}

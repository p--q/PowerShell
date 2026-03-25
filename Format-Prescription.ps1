<#
.FILENAME
    Format-Prescription.ps1
.VERSION
    1.1.1
.DESCRIPTION
    1. 不可視の制御文字や特殊な空白を徹底除去（「?」化対策）
    2. 不要な行の削除、薬品名の結合、用法の正規化
    3. 結果を標準的なテキスト形式でクリップボードへ保存
#>

# 1. クリップボードからテキストを取得
$inputContent = Get-Clipboard -Raw
if ([string]::IsNullOrWhiteSpace($inputContent)) {
    Write-Warning "クリップボードが空です。"
    exit
}

# --- 前処理：不可視の制御文字を削除 ---
# 改行(10,13)とタブ(9)以外の、ASCII制御文字や特殊な文字を置換
$inputContent = $inputContent -replace "[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", ""

$lines = $inputContent -split "`r?`n" | ForEach-Object {
    $line = $_
    # 行頭・行末のトリム（不可視の空白も含む）
    $line = $line.Trim()
    
    # 「商品名」とその後の空白を削除
    $line = $line -replace "^商品名\s*", ""
    # 「（任意の文字として）」を削除
    $line = $line -replace "[（\(].*?として[）\)]", ""
    
    # 削除対象の行
    if ($line -match "^(処方箋使用期限|--|<)") { return $null }
    
    # 文字列が空になった場合も除外
    if ([string]::IsNullOrWhiteSpace($line)) { return $null }
    
    $line
} | Where-Object { $_ -ne $null }

# 再結合して「処方日」ごとにブロック化
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

# 最終結合と空白の正規化
$rawResult = ($finalOutput -join "`r?`n") -replace "　", " " -replace "\s{2,}", " "
$finalLines = $rawResult -split "`r?`n" | Where-Object { ![string]::IsNullOrWhiteSpace($_) }
$resultText = $finalLines -join "`r?`n"

# 4. クリップボードに挿入（確実にテキストとして渡す）
if (![string]::IsNullOrEmpty($resultText)) {
    # メモ帳での「?」化を防ぐため、明示的に文字列としてセット
    Set-Clipboard -Value ([string]$resultText)
    Write-Host "--- 修正版(v1.1.1) 処理完了 ---" -ForegroundColor Green
}

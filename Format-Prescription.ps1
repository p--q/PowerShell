<#
.FILENAME
    Format-Prescription.ps1
.VERSION
    1.1.3
.DESCRIPTION
    ?化対策の最終手段：
    1. 全ての行をフラットなリストとして管理し、空行や制御文字を徹底排除。
    2. 文字列結合を最後の一回に限定し、.NETの機能でクリップボードへ転送。
#>

# 1. クリップボードからテキストを取得
$inputContent = Get-Clipboard -Raw
if ([string]::IsNullOrWhiteSpace($inputContent)) {
    Write-Warning "クリップボードが空です。"
    exit
}

# 制御文字（改行・タブ以外）を即座に消去
$inputContent = $inputContent -replace "[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", ""

# --- 前処理：不要な行の削除とトリム ---
$allLines = New-Object System.Collections.Generic.List[string]
foreach ($line in ($inputContent -split "`r?`n")) {
    $l = $line.Trim()
    if ([string]::IsNullOrWhiteSpace($l)) { continue }
    if ($l -match "^(商品名|処方箋使用期限|--|<)") { continue }
    
    # 商品名が先頭にある場合の削除処理
    $l = $l -replace "^商品名\s*", ""
    # 「（...として）」の削除
    $l = $l -replace "[（\(].*?として[）\)]", ""
    
    if (![string]::IsNullOrWhiteSpace($l)) { $allLines.Add($l) }
}

# --- 処方日ごとのブロック処理と整形 ---
$finalList = New-Object System.Collections.Generic.List[string]
$currentBlock = New-Object System.Collections.Generic.List[string]

foreach ($line in $allLines) {
    if ($line -match "処方日[:：]") {
        # 溜まっていたブロックを処理して最終リストへ
        if ($currentBlock.Count -gt 0) {
            $processed = Process-PrescriptionBlock $currentBlock
            $finalList.AddRange($processed)
            $currentBlock.Clear()
        }
        continue # 「処方日」行自体は結果に含めない
    }
    $currentBlock.Add($line)
}
# 最後のブロックを処理
if ($currentBlock.Count -gt 0) {
    $processed = Process-PrescriptionBlock $currentBlock
    $finalList.AddRange($processed)
}

# --- ブロック整形用関数 ---
function Process-PrescriptionBlock($lines) {
    $res = New-Object System.Collections.Generic.List[string]
    # 行の結合（最短行対策）
    $maxLen = ($lines | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum
    $merged = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $curr = $lines[$i]
        if ($i -lt $lines.Count - 1 -and $curr.Length -lt $maxLen) {
            $merged.Add($curr + $lines[$i+1]); $i++
        } else { $merged.Add($curr) }
    }

    foreach ($line in $merged) {
        $l = $line
        if ($l -match "^分") {
            if ($l -match "\s{2,}") { $res.Add(($l -replace "\s{2,}.*$", "")); break }
            $l = $l -replace "^分[1１]\s*", "" -replace "食後", ""
            $l = $l -replace "毎食後", ""
            $res.Add($l.Trim())
        }
        elseif ($l -match "時") {
            if ($l -match "\s{2,}") { $res.Add(($l -replace "\s{2,}.*$", "")); break }
            $res.Add($l.Trim())
        }
        elseif ($l -match "^外）") {
            $l = $l -replace "^外）", ""
            if ($l -match "\s{2,}") {
                $l = $l -replace "\s{2,}.*$", ""
                # 以降をすべて結合
                $idx = $merged.IndexOf($line)
                if ($idx -lt $merged.Count - 1) { $l += ($merged[($idx+1)..($merged.Count-1)] -join "") }
                $res.Add($l.Trim()); break
            }
            $res.Add($l.Trim())
        }
        else { $res.Add($l.Trim()) }
    }
    return $res
}

# --- 最終出力：空白を1つに集約し、重複行や空行を消してクリップボードへ ---
$outputText = ($finalList | ForEach-Object { 
    ($_ -replace "　", " " -replace "\s{2,}", " ").Trim() 
} | Where-Object { ![string]::IsNullOrWhiteSpace($_) }) -join "`r`n"

if (![string]::IsNullOrEmpty($outputText)) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Clipboard]::SetText($outputText)
    Write-Host "--- v1.1.3 完了 ---" -ForegroundColor Green
}

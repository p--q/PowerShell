<#
.FILENAME
    Format-Prescription.ps1
.VERSION
    1.1.0
.DESCRIPTION
    クリップボード内の処方箋テキストを取得し、以下の整形を行います。
    1. 不要な行（商品名、処方箋使用期限、--、<、注意書き等）の削除
    2. 泣き別れ（改行）した薬品名の自動結合
    3. 用法（分1、分3、外用等）の簡略化と正規化
    4. 連続する空白の半角スペース1つへの集約
    5. 最終結果をクリップボードへ書き戻し
#>

# 1. クリップボードからテキストを取得
$inputContent = Get-Clipboard -Raw
if ([string]::IsNullOrWhiteSpace($inputContent)) {
    Write-Warning "クリップボードが空です。"
    exit
}

# --- 前処理 (全体に対して適用) ---
$lines = $inputContent -split "`r?`n" | ForEach-Object {
    $line = $_
    # 「商品名」とその後の空白を削除
    $line = $line -replace "^商品名\s*", ""
    # 「（任意の文字として）」のパターンを削除（全角半角問わず）
    $line = $line -replace "[（\(].*?として[）\)]", ""
    # 行先頭の空白削除
    $line = $line.TrimStart()
    
    # 削除対象の行（処方箋使用期限、--、< から始まる行）
    if ($line -match "^(処方箋使用期限|--|<)") { return $null }
    
    $line
} | Where-Object { $_ -ne $null }

# 再結合して「処方日」ごとにブロック化
$processedText = $lines -join "`r?`n"
# 「処方日：」または「処方日 :」で分割（肯定先読みでキーワードを残す）
$blocks = $processedText -split "(?=処方日[:：])" | Where-Object { $_ -match "処方日" }

$finalOutput = New-Object System.Collections.Generic.List[string]

# --- 各ブロック（処方日単位）への処理 ---
foreach ($block in $blocks) {
    # 処方日の行を除いた中身を取得
    $blockLines = $block -split "`r?`n" | Where-Object { $_ -notmatch "^処方日" -and ![string]::IsNullOrWhiteSpace($_) }
    if ($blockLines.Count -eq 0) { continue }

    # 行の結合ロジック（最長行より短い行を次行と結合して薬品名の泣き別れを解消）
    $maxLen = ($blockLines | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum
    $mergedLines = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $blockLines.Count; $i++) {
        $current = $blockLines[$i]
        if ($i -lt $blockLines.Count - 1 -and $current.Length -lt $maxLen) {
            $mergedLines.Add($current + $blockLines[$i+1])
            $i++ # 次の行を消費
        } else {
            $mergedLines.Add($current)
        }
    }

    # 各行の個別置換ルール
    $resultBlock = New-Object System.Collections.Generic.List[string]
    foreach ($line in $mergedLines) {
        $l = $line

        # 「分」から始まる行 (全角半角両対応)
        if ($l -match "^分") {
            if ($l -match "\s{2,}") {
                $l = $l -replace "\s{2,}.*$", ""
                $resultBlock.Add($l)
                break # 以降の行を削除
            }
            # 用法の置換ルール
            if ($l -match "^分[1１]\s*") {
                $l = $l -replace "^分[1１]\s*", "" -replace "食後", ""
            } elseif ($l -match "^分[3３]") {
                $l = $l -replace "毎食後", ""
            } else {
                $l = $l -replace "食後", ""
            }
        }
        # 「時」を含む行
        elseif ($l -match "時") {
            if ($l -match "\s{2,}") {
                $l = $l -replace "\s{2,}.*$", ""
                $resultBlock.Add($l)
                break # 以降の行を削除
            }
        }
        # 「外）」から始まる行
        elseif ($l -match "^外）") {
            $l = $l -replace "^外）", ""
            if ($l -match "\s{2,}") {
                $l = $l -replace "\s{2,}.*$", ""
                # 以降の行をすべて結合
                $currentIdx = $mergedLines.IndexOf($line)
                if ($currentIdx -lt $mergedLines.Count - 1) {
                    $l += ($mergedLines[($currentIdx + 1)..($mergedLines.Count - 1)] -join "")
                }
                $resultBlock.Add($l)
                break
            }
        }
        $resultBlock.Add($l)
    }
    $finalOutput.Add(($resultBlock -join "`r?`n"))
}

# 最終結合と空白・文字の正規化
$rawResult = ($finalOutput -join "`r?`n") -replace "　", " " -replace "\s{2,}", " "
$finalLines = $rawResult -split "`r?`n" | Where-Object { ![string]::IsNullOrWhiteSpace($_) }
$resultText = $finalLines -join "`r?`n"

# 4. クリップボードに挿入
if (![string]::IsNullOrEmpty($resultText)) {
    Set-Clipboard -Value $resultText
    Write-Host "--- 処理完了 (v1.1.0) ---" -ForegroundColor Green
    Write-Host "整形したテキストをクリップボードに保存しました。"
} else {
    Write-Warning "整形結果が空になりました。"
}

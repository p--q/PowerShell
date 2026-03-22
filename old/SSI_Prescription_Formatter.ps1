<#
.FILENAME
    SSI_Prescription_Formatter.ps1
.VERSION
    1.0
.DESCRIPTION
    SSI電子カルテの処方引用テキストを、紹介状向けに自動整形するスクリプトです。
    【主な機能】
    1. 英数字のみを半角化（薬品名のカタカナは全角を維持して視認性を確保）
    2. 薬剤名と用量の間の過剰な空白を半角スペース1つに集約
    3. 「1日...」で始まる行の改行を削除し、上の行と連結
    4. 「(〇〇日分)」などの日数表記をカッコごと削除
#>

Add-Type -AssemblyName Microsoft.VisualBasic
Write-Host "--------------------------------------------------"
Write-Host " SSI Prescription Formatter v1.0 "
Write-Host " 状態: 監視中... (コピーすると自動整形します) "
Write-Host " 終了するには、このウィンドウを閉じてください。 "
Write-Host "--------------------------------------------------"

$lastText = ""

while($true) {
    # クリップボードからテキスト取得
    try {
        $text = Get-Clipboard -Raw -ErrorAction SilentlyContinue
    } catch {
        $text = $null
    }
    
    if ($text -and $text -ne $lastText) {
        # 1. 英数字・記号のみを半角に変換（カタカナの全角を維持するため個別処理）
        $chars = $text.ToCharArray()
        for ($i=0; $i -lt $chars.Length; $i++) {
            $val = [int]$chars[$i]
            # 全角英数字の範囲を判定して半角へシフト
            if (($val -ge 0xFF10 -and $val -le 0xFF19) -or  # 数字
                ($val -ge 0xFF21 -and $val -le 0xFF3A) -or  # 大文字英
                ($val -ge 0xFF41 -and $val -le 0xFF5A)) {   # 小文字英
                $chars[$i] = [char]($val - 0xFEE0)
            }
        }
        $work = New-Object String($chars, 0, $chars.Length)

        # 2. 日数表記「(●日分)」を削除（全角・半角両方のカッコに対応）
        $work = $work -replace "[\(（]\d+日分[\)）]", ""

        # 3. 「1日...」の行を上の行と結合（改行と行頭空白を半角スペースに）
        $work = $work -replace "\r?\n\s*(1日)", " $1"

        # 4. 2つ以上続く空白（全角・半角）を半角スペース1つに集約
        $work = $work -replace "[ 　]{2,}", " "

        # クリップボードに書き戻し
        $finalText = $work.Trim()
        if ($finalText) {
            Set-Clipboard -Value $finalText
            $lastText = $finalText
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 整形を実行しました。"
        }
    }
    Start-Sleep -Milliseconds 500
}

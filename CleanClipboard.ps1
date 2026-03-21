<#
.FileName
    CleanClipboard.ps1
.Version
    1.1
.Description
    LibreOffice Calc等からコピーした際の画像データや装飾を破棄し、
    特定の文字列整形（「商品名」の削除、「（として）」の削除）を
    行った上でプレーンテキストとしてクリップボードに戻します。
    SSI電子カルテ等への貼り付け最適化用スクリプトです。
#>

# クリップボードからテキスト形式のみを取得（画像・HTML形式をカット）
$rawText = Get-Clipboard -Format Text

if ($rawText) {
    # 1. 行頭の「商品名」とそれに続く空白を削除
    # 2. 全角/半角括弧で囲まれた「として」を削除
    # (?m) はマルチラインモード（各行の先頭を認識）
    $processedText = $rawText -replace "(?m)^商品名\s*", "" `
                              -replace "[（\(]として[）\)]", ""
    
    # 前後の余白を整えてクリップボードに再セット
    # これによりクリップボード内はプレーンテキストのみの状態になります
    $processedText.Trim() | Set-Clipboard
    
    # 実行確認（ショートカット実行時に最小化設定なら表示されません）
    # Write-Host "Process Complete: Text Cleaned." -ForegroundColor Green
}

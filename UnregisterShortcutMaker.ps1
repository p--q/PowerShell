<#
.FileName
    UnregisterShortcutMaker.ps1
.Version
    1.0
.Description
    「送る」メニューに登録された「ショートカットキーを割り当てて作成」を削除します。
#>

# 1. 「送る (SendTo)」フォルダのパスを取得
$sendToPath = [Environment]::GetFolderPath("SendTo")
$lnkName = "ショートカットキーを割り当てて作成.lnk"
$targetPath = Join-Path $sendToPath $lnkName

# 2. ファイルの存在確認と削除
if (Test-Path $targetPath) {
    try {
        Remove-Item $targetPath -Force
        Write-Host "--- 削除完了 ---" -ForegroundColor Yellow
        Write-Host "「送る」メニューから以下の項目を削除しました:"
        Write-Host " $lnkName"
    } catch {
        Write-Error "削除中にエラーが発生しました: $($_.Exception.Message)"
    }
} else {
    Write-Host "対象のショートカットは見つかりませんでした。" -ForegroundColor Gray
}

Write-Host "`n3秒後にウィンドウを閉じます..."
Start-Sleep -Seconds 3

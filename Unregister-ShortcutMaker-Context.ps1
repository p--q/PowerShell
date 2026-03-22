<#
.FileName
    Unregister-ShortcutMaker-Context.ps1
.Version
    1.0.0
.Description
    Register-ShortcutMaker-Context.ps1 で登録した
    右クリックメニュー（コンテキストメニュー）の項目を削除します。
#>

# 削除対象のレジストリパス
$registryPath = "HKCU:\Software\Classes\Microsoft.PowerShellScript.1\Shell\AssignShortcut"

try {
    Write-Host "--- 右クリックメニューの削除を開始します ---" -ForegroundColor Cyan
    
    # レジストリキーが存在するか確認
    if (Test-Path $registryPath) {
        # キーとその配下（commandキーなど）をすべて削除
        Remove-Item -Path $registryPath -Recurse -Force
        Write-Host "成功: メニュー『キーボードショートカットを割り当て』を削除しました。" -ForegroundColor Green
    } else {
        Write-Host "通知: 削除対象のメニューは見つかりませんでした。" -ForegroundColor Yellow
    }
} catch {
    Write-Host "失敗: レジストリの削除中にエラーが発生しました。" -ForegroundColor Red
    Write-Host "理由: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host "`n完了しました。任意のキーを押すと終了します。"
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

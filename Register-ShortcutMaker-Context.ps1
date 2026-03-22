<#
.FileName
    Register-ShortcutMaker-Context.ps1
.Version
    1.1.1
.Description
    ShortcutMaker.ps1をWindowsの右クリック（コンテキスト）メニューに登録します。
    引数のエラーを回避するため、レジストリの既定値の設定方法を修正しました。
#>

# --- 設定項目 ---
$menuName = "キーボードショートカットを割り当て"
$scriptPath = Join-Path $PSScriptRoot "ShortcutMaker.ps1"

# --- 登録処理 ---
$command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" `"%1`""
$registryPath = "HKCU:\Software\Classes\Microsoft.PowerShellScript.1\Shell\AssignShortcut"

try {
    Write-Host "--- 右クリックメニューへの登録を開始します ---" -ForegroundColor Cyan
    
    # 1. メニュー項目の作成
    if (-not (Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }
    # (既定)値にメニュー名を設定（Set-Itemを使うことで確実に「既定」を書き換えます）
    Set-Item -Path $registryPath -Value $menuName
    
    # 2. 実行コマンドの登録
    $commandPath = Join-Path $registryPath "command"
    if (-not (Test-Path $commandPath)) {
        New-Item -Path $commandPath -Force | Out-Null
    }
    # (既定)値にコマンドライン文字列を設定
    Set-Item -Path $commandPath -Value $command
    
    Write-Host "成功: .ps1専用メニュー『$menuName』を登録しました。" -ForegroundColor Green
} catch {
    Write-Host "失敗: レジストリ登録中にエラーが発生しました。" -ForegroundColor Red
    Write-Host "理由: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host "`n完了しました。任意のキーを押すと終了します。"
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

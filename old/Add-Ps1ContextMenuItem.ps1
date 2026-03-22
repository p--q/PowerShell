<#
.FileName
    Add-Ps1ContextMenuItem.ps1
.Version
    1.0.0
.Description
    .ps1ファイルを右クリックした際のコンテキストメニューに、
    「キーボードショートカットを割り当て」という項目を追加します。
    このメニューはPowerShellスクリプト（.ps1）に対してのみ表示されます。
#>

# 表示されるメニュー名
$menuName = "キーボードショートカットを割り当て"

# 呼び出す本体スクリプトのパス（このスクリプトと同じフォルダにある ShortcutMaker.ps1 を指定）
$scriptPath = Join-Path $PSScriptRoot "ShortcutMaker.ps1"

# 実行するコマンドの組み立て
# -NoProfile: プロファイルを読み込まず高速起動
# -ExecutionPolicy Bypass: 実行ポリシーを一時的に回避
# -File: 実行するスクリプトを指定
# "%1": 右クリックしたファイルのフルパスを引数として渡す
$command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" `"%1`""

# レジストリの登録先（現在のユーザー設定内: .ps1ファイルのシェル拡張）
$registryPath = "HKCU:\Software\Classes\Microsoft.PowerShellScript.1\Shell\AssignShortcut"

try {
    Write-Host "--- メニュー登録処理を開始します ---" -ForegroundColor Cyan
    
    # 1. メニュー項目の作成
    if (-not (Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }
    New-ItemProperty -Path $registryPath -Name "" -Value $menuName -PropertyType String -Force | Out-Null
    
    # 2. 実行コマンドの登録
    $commandPath = Join-Path $registryPath "command"
    if (-not (Test-Path $commandPath)) {
        New-Item -Path $commandPath -Force | Out-Null
    }
    New-ItemProperty -Path $commandPath -Name "" -Value $command -PropertyType String -Force | Out-Null
    
    Write-Host "成功: .ps1専用メニュー『$menuName』を登録しました。" -ForegroundColor Green
    Write-Host "今後は .ps1 ファイルを右クリックするだけで実行可能です。"
} catch {
    Write-Host "失敗: レジストリへの登録中にエラーが発生しました。" -ForegroundColor Red
    Write-Host "理由: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host "`n任意のキーを押すと終了します。"
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

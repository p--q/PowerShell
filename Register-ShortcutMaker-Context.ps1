<#
.FileName
    Register-ShortcutMaker-Context.ps1
.Version
    1.1.0
.Description
    ShortcutMaker.ps1をWindowsの右クリック（コンテキスト）メニューに登録します。
    この登録を行うと、.ps1ファイルを右クリックした際にのみ
    「キーボードショートカットを割り当て」という項目が表示されるようになります。
#>

# --- 設定項目 ---
# 右クリックメニューに表示される名前
$menuName = "キーボードショートカットを割り当て"

# 呼び出す本体スクリプトのパス（このスクリプトと同じフォルダにある ShortcutMaker.ps1 を指定）
$scriptPath = Join-Path $PSScriptRoot "ShortcutMaker.ps1"

# --- 登録処理 ---
# 実行コマンドの組み立て
# "%1" は右クリックした対象ファイルのパスを受け取るための特殊な記法です
$command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" `"%1`""

# レジストリの登録先（.ps1ファイルに関連付けられたシェル拡張）
$registryPath = "HKCU:\Software\Classes\Microsoft.PowerShellScript.1\Shell\AssignShortcut"

try {
    Write-Host "--- 右クリックメニューへの登録を開始します ---" -ForegroundColor Cyan
    
    # 1. メニュー項目の作成（キーの作成と表示名の設定）
    if (-not (Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }
    # (既定) 値にメニュー名を設定
    New-ItemProperty -Path $registryPath -Name "" -Value $menuName -PropertyType String -Force | Out-Null
    
    # 2. 実行コマンドの登録（実際の動作を定義）
    $commandPath = Join-Path $registryPath "command"
    if (-not (Test-Path $commandPath)) {
        New-Item -Path $commandPath -Force | Out-Null
    }
    # (既定) 値にコマンドライン文字列を設定
    New-ItemProperty -Path $commandPath -Name "" -Value $command -PropertyType String -Force | Out-Null
    
    Write-Host "成功: .ps1専用メニュー『$menuName』を登録しました。" -ForegroundColor Green
    Write-Host "確認: .ps1ファイルを右クリックしてメニューが表示されるか見てください。"
} catch {
    Write-Host "失敗: レジストリ登録中にエラーが発生しました。" -ForegroundColor Red
    Write-Host "理由: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host "`n完了しました。任意のキーを押すとウィンドウを閉じます。"
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

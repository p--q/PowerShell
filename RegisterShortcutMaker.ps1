<#
.FileName
    RegisterShortcutMaker.ps1
.Version
    1.0
.Description
    ShortcutMaker.ps1 を Windows の「送る」メニューに登録します。
    このファイル自体を実行することで設定が完了します。
#>

# 1. ShortcutMaker.ps1 の場所を特定（このスクリプトと同じフォルダにあると想定）
$scriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$targetScript = Join-Path $scriptDir "ShortcutMaker.ps1"

if (-not (Test-Path $targetScript)) {
    Write-Host "エラー: ShortcutMaker.ps1 が見つかりません。" -ForegroundColor Red
    Write-Host "このスクリプトと同じフォルダに ShortcutMaker.ps1 を置いてから実行してください。"
    Pause
    exit
}

# 2. 「送る (SendTo)」フォルダのパスを取得
$sendToPath = [Environment]::GetFolderPath("SendTo")

# 3. ショートカットの作成設定
$shell = New-Object -ComObject WScript.Shell
$lnkPath = Join-Path $sendToPath "ショートカットキーを割り当てて作成.lnk"
$shortcut = $shell.CreateShortcut($lnkPath)

# 4. 実行コマンドの構成
# 実行実感を出すため、通常のウィンドウ(1)で起動するように設定
$shortcut.TargetPath = "powershell.exe"
$shortcut.Arguments = "-ExecutionPolicy Bypass -File ""$targetScript"" -SelectedPath"
$shortcut.IconLocation = "powershell.exe, 0"
$shortcut.Description = "選択したps1ファイルにショートカットキーを割り当てます"
$shortcut.Save()

Write-Host "--- 登録完了 ---" -ForegroundColor Cyan
Write-Host "「送る」メニューに登録しました: $lnkPath"
Write-Host "これで、他のps1ファイルを右クリックからショートカット化できます。"
Start-Sleep -Seconds 3

<#
.FileName
    ShortcutMaker.ps1
.Version
    1.1
.Description
    エクスプローラーで選択したPS1ファイルを「送る」から受け取り、
    同じフォルダに実行用ショートカットを作成します。
    実行時に黒い画面（コンソール）が表示される設定にしています。
#>

# 引数（送るメニューから渡されたファイルパス）を取得
param([string]$SelectedPath)

# ファイルが選択されていない、またはPS1でない場合のチェック
if ([string]::IsNullOrEmpty($SelectedPath) -or ($SelectedPath -notlike "*.ps1")) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("PowerShellスクリプト(.ps1)を右クリックして選択してください。")
    exit
}

# 1. ショートカットキーの入力を取得
Add-Type -AssemblyName Microsoft.VisualBasic
$prompt = "割り当てるショートカットキーを入力してください。`n(例: Ctrl+Alt+C, Ctrl+Alt+S)"
$shortcutKey = [Microsoft.VisualBasic.Interaction]::InputBox($prompt, "ショートカットキー設定", "Ctrl+Alt+C")

if ([string]::IsNullOrWhiteSpace($shortcutKey)) { exit }

# 2. ショートカット作成の準備
$shell = New-Object -ComObject WScript.Shell
$file = Get-Item $SelectedPath
$shortcutPath = Join-Path $file.DirectoryName "$($file.BaseName).lnk"
$shortcut = $shell.CreateShortcut($shortcutPath)

# 3. 実行コマンドの設定
# -NoExit を付けると、処理が終わっても黒い画面が閉じずに残ります（動作確認用）
# 実行実感を出すため、通常のウィンドウ表示（WindowStyle 1）に設定
$shortcut.TargetPath = "powershell.exe"
$shortcut.Arguments = "-ExecutionPolicy Bypass -File ""$($file.FullName)"""
$shortcut.WorkingDirectory = $file.DirectoryName
$shortcut.Hotkey = $shortcutKey
$shortcut.WindowStyle = 1 # 通常のウィンドウで表示
$shortcut.Description = "Shortcut for $($file.Name)"

# 4. 保存と完了通知
try {
    $shortcut.Save()
    Write-Host "--- 設定完了 ---" -ForegroundColor Cyan
    Write-Host "作成先: $shortcutPath"
    Write-Host "割り当てキー: $shortcutKey"
    Write-Host "確認のため5秒後に閉じます..."
    Start-Sleep -Seconds 5
} catch {
    Write-Error "ショートカットの保存に失敗しました: $($_.Exception.Message)"
    Pause
}

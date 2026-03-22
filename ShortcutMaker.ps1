<#
.FileName
    ShortcutMaker.ps1
.Version
    1.2.0
.Description
    実行すると、対象ファイルへのショートカットをデスクトップに作成し、
    左手で届く範囲の空きショートカットキー候補を提示・割り当てます。
#>

param([string]$targetPath)

if (-not $targetPath) {
    Write-Host "エラー: 対象ファイルが指定されていません。" -ForegroundColor Red
    pause ; exit
}

$shell = New-Object -ComObject WScript.Shell
$desktopPath = [Environment]::GetFolderPath("Desktop")
$fileName = [System.IO.Path]::GetFileNameWithoutExtension($targetPath)
$shortcutPath = Join-Path $desktopPath "$fileName.lnk"

# --- 1. 左手で届く「Ctrl+Alt+」の候補リスト ---
# 先生の可動域（右はRFVまで、上は数字キーの下）に基づいた厳選リスト
$candidates = @("S", "X", "Z", "A", "Q", "W", "E", "R", "D", "F", "C", "V", "1", "2", "3", "4", "5")

# --- 2. デスクトップ上の既存ショートカットとの重複チェック ---
$usedKeys = Get-ChildItem "$desktopPath\*.lnk" | ForEach-Object {
    $lnk = $shell.CreateShortcut($_.FullName)
    if ($lnk.Hotkey) { $lnk.Hotkey.Replace("Ctrl+Alt+", "") }
}

$availableCandidates = $candidates | Where-Object { $_ -notin $usedKeys }

# --- 3. ユーザー選択ダイアログ ---
$message = "割り当てるキーを選択してください（Ctrl+Alt + [選択キー]）`n`n※現在のデスクトップで未使用の候補を表示しています。"
$title = "ショートカットキーの割り当て"

# 選択肢の作成
$choices = foreach ($c in $availableCandidates) {
    "&" + $c  # &を付けるとキーボードで選択可能に
}
$default = 0
$choice = $Host.UI.PromptForChoice($title, $message, [System.Management.Automation.Host.ChoiceDescription[]]$choices, $default)

$selectedKey = $availableCandidates[$choice]
$fullHotkey = "Ctrl+Alt+$selectedKey"

# --- 4. ショートカット作成と保存 ---
try {
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = "powershell.exe"
    $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$targetPath`""
    $shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName($targetPath)
    $shortcut.Hotkey = $fullHotkey
    $shortcut.IconLocation = "powershell.exe,0"
    $shortcut.Description = "Shortcut created by ShortcutMaker"
    $shortcut.Save()

    Write-Host "`n成功！" -ForegroundColor Green
    Write-Host "作成先: $shortcutPath"
    Write-Host "割り当てキー: $fullHotkey" -ForegroundColor Cyan
} catch {
    Write-Host "エラー: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n3秒後に閉じます..."
Start-Sleep -Seconds 3

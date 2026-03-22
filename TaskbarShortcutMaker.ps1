<#
.FileName
    TaskbarShortcutMaker.ps1
.Version
    1.0.0
.Description
    .ps1ファイルをタスクバーにピン留め可能な形式のショートカットとして作成します。
    作成後、エクスプローラーでショートカットを表示するので、手動でタスクバーへドラッグしてください。
#>

param([string]$targetPath)

if (-not $targetPath) {
    Write-Host "エラー: 対象のps1ファイルを指定してください。" -ForegroundColor Red
    Start-Sleep -Seconds 3 ; exit
}

$shell = New-Object -ComObject WScript.Shell
$desktopPath = [Environment]::GetFolderPath("Desktop")
$fileName = [System.IO.Path]::GetFileNameWithoutExtension($targetPath)
$shortcutPath = Join-Path $desktopPath "$fileName.lnk"

try {
    $shortcut = $shell.CreateShortcut($shortcutPath)
    
    # ポイント: 直接ps1を指定せず、powershell.exeの引数として渡すことでピン留め可能になります
    $shortcut.TargetPath = "powershell.exe"
    $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$targetPath`""
    
    $shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName($targetPath)
    $shortcut.Description = "Taskbar Pin-able Shortcut for $fileName"
    
    # アイコンを分かりやすくPowerShellの青いアイコンに設定（インデックス0）
    $shortcut.IconLocation = "powershell.exe,0"
    
    $shortcut.Save()

    Write-Host "--- ショートカットの作成に成功しました ---" -ForegroundColor Green
    Write-Host "作成先: $shortcutPath"
    Write-Host "`n【重要：最後のステップ】" -ForegroundColor Cyan
    Write-Host "デスクトップに作成された「$fileName」のショートカットを"
    Write-Host "マウスでタスクバーへドラッグ＆ドロップしてください。"
    
    # 作成したファイルをエクスプローラーで選択状態で表示する
    explorer.exe /select, "$shortcutPath"

} catch {
    Write-Host "エラー: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n5秒後に終了します..."
Start-Sleep -Seconds 5

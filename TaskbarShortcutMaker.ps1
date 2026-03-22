<#
.FileName
    TaskbarShortcutMaker.ps1
.Version
    1.1.0
.Description
    .ps1ファイルをタスクバーにピン留め可能な形式で作成します。
    引数がない場合はファイル選択ダイアログを表示します。
#>

param([string]$targetPath)

# --- 1. 引数がない場合にファイル選択ダイアログを表示 ---
if (-not $targetPath) {
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $FileBrowser.Title = "タスクバーに登録したい .ps1 ファイルを選択してください"
    $FileBrowser.Filter = "PowerShell スクリプト (*.ps1)|*.ps1|すべてのファイル (*.*)|*.*"
    $FileBrowser.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
    
    if ($FileBrowser.ShowDialog() -eq "OK") {
        $targetPath = $FileBrowser.FileName
    } else {
        Write-Host "キャンセルされました。" -ForegroundColor Yellow
        Start-Sleep -Seconds 2 ; exit
    }
}

# --- 2. ショートカット作成の準備 ---
$shell = New-Object -ComObject WScript.Shell
$desktopPath = [Environment]::GetFolderPath("Desktop")
$fileName = [System.IO.Path]::GetFileNameWithoutExtension($targetPath)
$shortcutPath = Join-Path $desktopPath "$fileName.lnk"

try {
    $shortcut = $shell.CreateShortcut($shortcutPath)
    
    # ピン留め可能にするための特殊設定
    $shortcut.TargetPath = "powershell.exe"
    $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$targetPath`""
    
    $shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName($targetPath)
    $shortcut.Description = "Taskbar Shortcut for $fileName"
    
    # アイコン設定（PowerShellの青いロゴ）
    $shortcut.IconLocation = "powershell.exe,0"
    
    $shortcut.Save()

    Write-Host "--- 作成完了 ---" -ForegroundColor Green
    Write-Host "作成されたショートカット: $fileName.lnk"
    Write-Host "`n【次の操作】" -ForegroundColor Cyan
    Write-Host "デスクトップにあるこのファイルを、タスクバーへ直接ドラッグしてください。"
    
    # 作成したファイルをエクスプローラーで選択状態で表示
    explorer.exe /select, "$shortcutPath"

} catch {
    Write-Host "エラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n5秒後に終了します..."
Start-Sleep -Seconds 5

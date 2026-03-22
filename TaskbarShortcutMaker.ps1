<#
.FileName
    TaskbarShortcutMaker.ps1
.Version
    1.1.1
.Description
    実行中のスクリプトと同じフォルダをデフォルトで開き、
    選択したファイルをタスクバー登録用ショートカットとして作成します。
#>

param([string]$targetPath)

# --- 1. 引数がない場合にファイル選択ダイアログを表示 ---
if (-not $targetPath) {
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $FileBrowser.Title = "タスクバーに登録したい .ps1 ファイルを選択してください"
    $FileBrowser.Filter = "PowerShell スクリプト (*.ps1)|*.ps1|すべてのファイル (*.*)|*.*"
    
    # 【修正箇所】このスクリプトがあるフォルダを初期ディレクトリに設定
    $FileBrowser.InitialDirectory = $PSScriptRoot
    
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
    
    # タスクバーにピン留め可能にするための設定
    $shortcut.TargetPath = "powershell.exe"
    $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$targetPath`""
    $shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName($targetPath)
    $shortcut.Description = "Taskbar Shortcut for $fileName"
    
    # アイコン（PowerShellの青いロゴ）
    $shortcut.IconLocation = "powershell.exe,0"
    
    $shortcut.Save()

    Write-Host "--- ショートカットの作成に成功しました ---" -ForegroundColor Green
    Write-Host "作成先: $shortcutPath"
    Write-Host "`n【次の操作】" -ForegroundColor Cyan
    Write-Host "デスクトップに作成された「$fileName」のショートカットを"
    Write-Host "マウスでタスクバーへドラッグ＆ドロップしてください。"
    
    # 作成したファイルをエクスプローラーで選択状態で表示する
    explorer.exe /select, "$shortcutPath"

} catch {
    Write-Host "エラー: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n5秒後に終了します..."
Start-Sleep -Seconds 5

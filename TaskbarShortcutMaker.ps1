<#
.FileName
    TaskbarShortcutMaker.ps1
.Version
    1.3.0
.Description
    ショートカット作成時に、Windows標準のアイコン選択ダイアログを表示します。
    好きなアイコンを選んで、タスクバーでの視認性を高めることができます。
#>

param([string]$targetPath)

# --- 1. ファイル選択ダイアログ ---
if (-not $targetPath) {
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $FileBrowser.Title = "タスクバーに登録したい .ps1 ファイルを選択してください"
    $FileBrowser.Filter = "PowerShell スクリプト (*.ps1)|*.ps1|すべてのファイル (*.*)|*.*"
    $FileBrowser.InitialDirectory = $PSScriptRoot
    
    if ($FileBrowser.ShowDialog() -eq "OK") {
        $targetPath = $FileBrowser.FileName
    } else {
        exit 
    }
}

# --- 2. パス設定 ---
$shell = New-Object -ComObject WScript.Shell
$desktopPath = [Environment]::GetFolderPath("Desktop")
$fileName = [System.IO.Path]::GetFileNameWithoutExtension($targetPath)
$shortcutPath = Join-Path $desktopPath "$fileName.lnk"

# --- 3. アイコン選択ダイアログの表示 (ここが新機能) ---
# Windows APIを利用して標準のアイコン選択画面を呼び出します
$iconPath = "shell32.dll" # デフォルトのライブラリ
$iconIndex = 0

$code = @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public class IconPicker {
    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    public static extern int PickIconDlg(IntPtr hwndOwner, StringBuilder lpstrFile, int nMaxFile, ref int lpdwIconIndex);
}
"@
Add-Type -TypeDefinition $code

$sb = New-Object System.Text.StringBuilder($iconPath, 260)
$res = [IconPicker]::PickIconDlg([IntPtr]::Zero, $sb, $sb.Capacity, [ref]$iconIndex)

if ($res -eq 1) {
    $selectedIconPath = $sb.ToString()
    $selectedIconIndex = $iconIndex
} else {
    # キャンセルされた場合はPowerShellの標準アイコン
    $selectedIconPath = "powershell.exe"
    $selectedIconIndex = 0
}

try {
    # --- 4. 重複チェックと作成 ---
    if (Test-Path $shortcutPath) { Remove-Item $shortcutPath -Force }

    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = "powershell.exe"
    $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$targetPath`""
    $shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName($targetPath)
    $shortcut.Description = "Taskbar Shortcut for $fileName"
    
    # 選んだアイコンを適用
    $shortcut.IconLocation = "$selectedIconPath,$selectedIconIndex"
    
    $shortcut.Save()

    Write-Host "--- 作成完了 ---" -ForegroundColor Green
    Write-Host "ショートカット: $fileName.lnk"
    Write-Host "適用アイコン: $selectedIconPath, index $selectedIconIndex" -ForegroundColor Cyan
    Write-Host "`nこのファイルをタスクバーへドラッグしてください。"

} catch {
    Write-Host "エラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
}

Start-Sleep -Seconds 3

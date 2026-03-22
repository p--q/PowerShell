<#
.FileName
    ShortcutMaker.ps1
.Version
    1.3.0
.Description
    対象ファイルをデスクトップへショートカットとして作成します。
    左手可動域内の空きキーを自動判定し、実行時の黒い画面と警告音を抑制する
    プロ仕様の起動オプション（Hidden / Bypass）を付与します。
#>

param([string]$targetPath)

# ターゲットが指定されていない場合の処理
if (-not $targetPath) {
    Write-Host "エラー: 対象ファイルが指定されていません。" -ForegroundColor Red
    Start-Sleep -Seconds 3
    exit
}

$shell = New-Object -ComObject WScript.Shell
$desktopPath = [Environment]::GetFolderPath("Desktop")
$fileName = [System.IO.Path]::GetFileNameWithoutExtension($targetPath)
$shortcutPath = Join-Path $desktopPath "$fileName.lnk"

# --- 1. 左手で届く「Ctrl+Alt+」の候補リスト ---
# 先生の可動域（右はRFVまで、上は数字キーの下）
$candidates = @("S", "X", "Z", "A", "Q", "W", "E", "R", "D", "F", "C", "V", "1", "2", "3", "4", "5")

# --- 2. デスクトップ上の既存ショートカットとの重複チェック ---
$usedKeys = Get-ChildItem "$desktopPath\*.lnk" | ForEach-Object {
    try {
        $lnk = $shell.CreateShortcut($_.FullName)
        if ($lnk.Hotkey) { $lnk.Hotkey.Replace("Ctrl+Alt+", "") }
    } catch { $null }
}

$availableCandidates = $candidates | Where-Object { $_ -notin $usedKeys }

# --- 3. ユーザー選択ダイアログ ---
$message = "割り当てるキーを選択してください（Ctrl+Alt + [選択キー]）`n`n※デスクトップで使用されていない候補のみ表示しています。"
$title = "ショートカットキーの自動割り当て (v1.3.0)"

$choices = foreach ($c in $availableCandidates) {
    "&" + $c  # キーボードの該当キーで即選択可能にする
}
$default = 0
$choice = $Host.UI.PromptForChoice($title, $message, [System.Management.Automation.Host.ChoiceDescription[]]$choices, $default)

$selectedKey = $availableCandidates[$choice]
$fullHotkey = "Ctrl+Alt+$selectedKey"

# --- 4. ショートカット作成と保存（ここが修正のキモです） ---
try {
    $shortcut = $shell.CreateShortcut($shortcutPath)
    
    # 実行プログラムはPowerShell本体
    $shortcut.TargetPath = "powershell.exe"
    
    # 引数の設定: 
    # -NoProfile: 高速起動
    # -ExecutionPolicy Bypass: 警告音の原因となる実行制限を回避
    # -WindowStyle Hidden: 黒い画面を一瞬も見せない
    # -File: ターゲットのスクリプトを指定
    $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$targetPath`""
    
    # 作業ディレクトリをスクリプトのある場所に設定
    $shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName($targetPath)
    
    # ショートカットキーの割り当て
    $shortcut.Hotkey = $fullHotkey
    
    # アイコンはPowerShellのものを使用
    $shortcut.IconLocation = "powershell.exe,0"
    
    $shortcut.Save()

    Write-Host "`n作成成功！" -ForegroundColor Green
    Write-Host "----------------------------------------"
    Write-Host " 作成ファイル : $fileName.lnk"
    Write-Host " 割り当てキー : $fullHotkey" -ForegroundColor Cyan
    Write-Host "----------------------------------------"
    Write-Host "※キーを押しても反応がない場合は、一度デスクトップを"
    Write-Host "  クリックしてから再度お試しください。"
} catch {
    Write-Host "エラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n5秒後に自動で閉じます..."
Start-Sleep -Seconds 5

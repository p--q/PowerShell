<#
.FileName
    RegisterShortcutMaker.ps1
.Version
    1.1
.Description
    ShortcutMaker.ps1 を Windows の「送る」メニューに登録します。
    エラー発生時に画面が閉じないよう修正しました。
#>

try {
    # 1. スクリプト自身の場所を厳密に特定
    $currentScriptPath = $MyInvocation.MyCommand.Path
    if ([string]::IsNullOrEmpty($currentScriptPath)) {
        # 直接貼り付けなどでパスが取れない場合の予備
        $currentScriptPath = Get-Location
    }
    $scriptDir = Split-Path $currentScriptPath -Parent
    $targetScript = Join-Path $scriptDir "ShortcutMaker.ps1"

    Write-Host "確認中のパス: $targetScript" -ForegroundColor Gray

    # 2. ファイルの存在チェック
    if (-not (Test-Path $targetScript)) {
        throw "エラー: 同じフォルダに 'ShortcutMaker.ps1' が見つかりません。`n現在の作業フォルダ: $scriptDir"
    }

    # 3. 「送る (SendTo)」フォルダの取得
    $sendToPath = [Environment]::GetFolderPath("SendTo")
    $lnkPath = Join-Path $sendToPath "ショートカットキーを割り当てて作成.lnk"

    # 4. ショートカットの作成
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($lnkPath)
    $shortcut.TargetPath = "powershell.exe"
    # 引数の構成をより確実に
    $shortcut.Arguments = "-ExecutionPolicy Bypass -File ""$targetScript"" -SelectedPath"
    $shortcut.IconLocation = "powershell.exe, 0"
    $shortcut.Save()

    Write-Host "--- 登録完了 ---" -ForegroundColor Cyan
    Write-Host "「送る」メニューに登録しました。"
    Write-Host "パス: $lnkPath"
}
catch {
    Write-Host "`n！！！ エラーが発生しました ！！！" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
}

Write-Host "`nキーを押すと終了します..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

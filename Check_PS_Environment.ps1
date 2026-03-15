<#
.FILENAME
    Check_PS_Environment.ps1
.VERSION
    1.0
.DESCRIPTION
    PowerShellスクリプトが実行可能な設定になっているか確認します。
#>

Write-Host "--- PowerShell 実行環境チェック ---" -ForegroundColor Cyan

# 現在の実行ポリシーを取得
$policy = Get-ExecutionPolicy

Write-Host "現在の実行ポリシー: " -NoNewline
Write-Host "$policy" -ForegroundColor Yellow

# 判定
if ($policy -eq "Restricted") {
    Write-Host "[NG] スクリプトの実行が禁止されています。" -ForegroundColor Red
    Write-Host "対策: 管理者としてPowerShellを開き、以下のコマンドを実行してください。"
    Write-Host "Set-ExecutionPolicy RemoteSigned -Force" -ForegroundColor Green
}
elseif ($policy -eq "AllSigned" -or $policy -eq "RemoteSigned" -or $policy -eq "Unrestricted" -or $policy -eq "Bypass") {
    Write-Host "[OK] スクリプトを実行可能な設定です。" -ForegroundColor Green
}
else {
    Write-Host "[注意] 特殊な設定 ($policy) です。実行できない可能性があります。" -ForegroundColor Magenta
}

Write-Host "----------------------------------"
Read-Host "Enterキーを押すと閉じます"

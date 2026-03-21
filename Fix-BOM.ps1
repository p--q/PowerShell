# Fix-BOM.ps1
# 共有フォルダ内の ps1 を Windows で読める形式に一括変換
$files = Get-ChildItem -Filter *.ps1
foreach ($file in $files) {
    if ($file.Name -eq $MyInvocation.MyCommand.Name) { continue }
    $content = Get-Content $file.FullName -Raw
    # 強制的に BOM 付き UTF-8 で保存
    [System.IO.File]::WriteAllText($file.FullName, $content, [System.Text.Encoding]::UTF8)
}
Write-Host "Windows用に最適化完了！" -ForegroundColor Cyan
Start-Sleep -Seconds 2

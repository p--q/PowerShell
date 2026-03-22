<#
.Description
    Encodes all *.ps1 files in the current directory to UTF-8 with BOM.
    This version uses only ASCII characters to avoid encoding issues.
#>

$files = Get-ChildItem -Filter *.ps1
foreach ($file in $files) {
    # Skip this script itself
    if ($file.Name -eq $MyInvocation.MyCommand.Name) { continue }
    
    try {
        Write-Host "Processing: $($file.Name)... " -NoNewline
        $content = Get-Content $file.FullName -Raw -Encoding UTF8
        
        # Force save as UTF-8 with BOM
        [System.IO.File]::WriteAllText($file.FullName, $content, [System.Text.Encoding]::UTF8)
        
        Write-Host "Done" -ForegroundColor Green
    } catch {
        Write-Host "Failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`nAll files have been optimized for Windows PowerShell."
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

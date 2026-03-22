<#
.Description
    Registers ShortcutMaker.ps1 to the 'SendTo' menu.
    This version uses explicit environment paths to avoid issues in Shared Folders.
#>

# 1. Define Paths
$sendToPath = [System.IO.Path]::Combine($env:APPDATA, "Microsoft\Windows\SendTo")
$shortcutPath = Join-Path $sendToPath "ShortcutMaker.lnk"
$scriptPath = Join-Path $PSScriptRoot "ShortcutMaker.ps1"

# 2. Check if Script exists
if (-not (Test-Path $scriptPath)) {
    Write-Host "Error: ShortcutMaker.ps1 not found in $PSScriptRoot" -ForegroundColor Red
    pause
    exit
}

try {
    # 3. Create Shortcut using COM object
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    
    # Target: powershell.exe
    $shortcut.TargetPath = "powershell.exe"
    
    # Arguments: Bypass policy and run the script with the right-clicked file as argument
    $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -File ""$scriptPath"""
    
    # Icon (Optional: PowerShell Icon)
    $shortcut.IconLocation = "powershell.exe,0"
    
    $shortcut.Save()
    
    Write-Host "Success! 'ShortcutMaker' added to SendTo menu." -ForegroundColor Green
    Write-Host "Target: $shortcutPath"
} catch {
    Write-Host "Failed to create shortcut: $($_.Exception.Message)" -ForegroundColor Red
}

pause

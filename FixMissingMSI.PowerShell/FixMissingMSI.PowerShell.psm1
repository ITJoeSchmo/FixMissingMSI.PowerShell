# FixMissingMSI.PowerShell.psm1
# Entry point for the module. Loads all public functions from Functions directory.

# Import each function script in the Functions folder
Get-ChildItem -Path "$PSScriptRoot\Functions" -Filter *.ps1 -File | ForEach-Object {
    try {
        . $_.FullName
    } catch {
        Write-Error "Failed to import function file: $($_.FullName). Error: $_"
    }
}
Get-ChildItem -Path "$PSScriptRoot\Functions\extras" -Filter *.ps1 -File | ForEach-Object {
    try {
        . $_.FullName
    } catch {
        Write-Error "Failed to import function file: $($_.FullName). Error: $_"
    }
}


Export-ModuleMember -Function @(
    'Initialize-InstallerCacheFileShare',
    'Invoke-InstallerCacheRepair',
    'Merge-InstallerCacheReports',
    'Update-InstallerCache',
    'Get-InstallerRegistration',
    'Remove-InstallerRegistration'
)

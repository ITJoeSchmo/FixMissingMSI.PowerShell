<#
    Example: Using FixMissingMSI.PowerShell with MECM
    -------------------------------------------------
    This script demonstrates how to:
      1. Initialize the FixMissingMSI workspace
      2. Create MECM scripts for deployment
      3. Run the repair, merge reports, populate cache, and rerun
#>

# --- Requirements ---
# - MECM PowerShell module (ConfigurationManager)
# - Admin rights in MECM
# - Existing fileshare with high level access (to set sharing permissions on subfolders)

# 1. Install and initialize the shared workspace
Install-Module FixMissingMSI.PowerShell
$fileSharePath = Read-Host "Enter the UNC path to the file share (e.g. \\FS01\Software)"
Initialize-InstallerCacheFileShare -FileSharePath $fileSharePath

# Locate module paths
$modulePsd1Path = (Get-Module FixMissingMSI.PowerShell -ListAvailable | Select-Object -First 1 -ExpandProperty Path)
$modulePath     = Split-Path $modulePsd1Path
$functionsPath  = Join-Path $modulePath "Functions"

# 2. Register MECM scripts
# Connect to the MECM site provider first (e.g., cd XYZ:)
$InvokeInstallerCacheRepair = New-CMScript -ScriptName "Invoke-InstallerCacheRepair" -ScriptFile "$functionsPath\Invoke-InstallerCacheRepair.ps1" -Fast
$UpdateInstallerCache       = New-CMScript -ScriptName "Update-InstallerCache"       -ScriptFile "$functionsPath\Update-InstallerCache.ps1"       -Fast

# 3. Run the first repair pass
$repairParams = @{
    FileSharePath = $fileSharePath
    SourcePaths   = "" # can add other paths to source MSI/MSP files from here example: "\\FS01\Software\SQLServer2019\SP3","\\FS01\Software\SQLServer2017\"
    RunFromShare  = $false 
    ReportOnly    = $false # Use ReportOnly = $true for a dry run before actual repair.
}

$targetCollection = Read-Host "Enter MECM collection name"
Invoke-CMScript -ScriptGuid $InvokeInstallerCacheRepair.ScriptGuid -CollectionName $targetCollection -ScriptParameter $repairParams

# 4. Merge reports and populate shared cache
Merge-InstallerCacheReports -FileSharePath $fileSharePath

$targetCollection = Read-Host "Enter MECM collection name to populate shared cache"
$updateParams = @{ FileSharePath = $fileSharePath }
Invoke-CMScript -ScriptGuid $UpdateInstallerCache.ScriptGuid -CollectionName $targetCollection -ScriptParameter $updateParams

# 5. Re-run repair pass using the populated cache
Invoke-CMScript -ScriptGuid $InvokeInstallerCacheRepair.ScriptGuid -CollectionName $targetCollection -ScriptParameter $repairParams

# --- Notes ---
# - The shared cache is demand-driven: only files reported missing are uploaded.
# - Use ReportOnly = $true for a dry run before actual repair.
# - Steps 3–5 can be scheduled or repeated as maintenance.

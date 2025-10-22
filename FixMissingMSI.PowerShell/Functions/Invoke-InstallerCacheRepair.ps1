#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Repairs the local Windows Installer cache using FixMissingMSI, operating from a local install or shared cache.

.DESCRIPTION
    Invokes FixMissingMSI non-interactively via .NET reflection to detect and repair missing or mismatched
    MSI and MSP files in C:\Windows\Installer.

    When -FileSharePath is specified, FixMissingMSI and its cache are sourced from the shared path:
        \\<Server>\<Share>\FixMissingMSI\
    
    When -FileSharePath is omitted, FixMissingMSI is expected to exist locally at:
        $env:TEMP\FixMissingMSI  (or as provided via -LocalWorkPath)
    
    The script:
      1. Loads FixMissingMSI.exe from either the local or shared path.
      2. Runs FixMissingMSI non-interactively by invoking its internal methods via reflection.
      3. Scans the configured setup media paths, installed products, and patch metadata.
      4. Attempts to reconstruct missing installer files using local and/or shared caches.
      5. Exports unresolved entries to a CSV report.
      6. Optionally executes generated FixCommand operations to repopulate C:\Windows\Installer.

    By default, the function stages FixMissingMSI locally before running it.  
    Use -RunFromShare to execute directly from the share (recommended only for trusted, low-latency environments).

.PARAMETER FileSharePath
    UNC path to the share that contains FixMissingMSI and its Cache/Reports folders.  
    Example: \\FS01\Software

    When not provided, the function assumes FixMissingMSI is already installed locally
    (for example, via Install-FixMissingMSI) under $env:TEMP\FixMissingMSI.

.PARAMETER SourcePaths
    One or more setup media paths (local or UNC) to scan for MSI/MSP packages.  
    Defaults to:
      - The shared cache (if -FileSharePath is provided)
      - An empty string if not, which triggers FixMissingMSI to use registry-based LastUsedSource metadata.

.PARAMETER LocalWorkPath
    The local working directory where FixMissingMSI is staged, executed, and where logs/reports are written.  
    Defaults to $env:TEMP\FixMissingMSI.

.PARAMETER RunFromShare
    If specified, runs FixMissingMSI directly from the network share instead of copying binaries locally.  
    This can improve performance in trusted, low-latency environments but increases risk of transient I/O or antivirus interference.

.PARAMETER ReportOnly
    Scans and generates the unresolved CSV report without executing any FixCommand operations.  
    Useful for audit or pre-flight scenarios.

.EXAMPLE
PS> Invoke-InstallerCacheRepair

    Runs FixMissingMSI non-interactively using the local install in $env:TEMP\FixMissingMSI,
    scans for missing MSI/MSP files, attempts repair, and exports the unresolved report locally.

.EXAMPLE
PS> Invoke-InstallerCacheRepair -FileSharePath \\FS01\Software

    Uses the shared cache \\FS01\Software\FixMissingMSI\Cache for repairs,
    exports the unresolved report to \\FS01\Software\FixMissingMSI\Reports,
    and executes generated FixCommands.

.EXAMPLE
PS> Invoke-InstallerCacheRepair -FileSharePath \\FS01\Software -SourcePaths 'D:\Media','\\FS01\Builds\Office'

    Scans the provided media paths in the given order instead of the default shared cache.

.EXAMPLE
PS> Invoke-InstallerCacheRepair -FileSharePath \\FS01\Software -RunFromShare

    Runs FixMissingMSI directly from the network share without copying it locally first.  
    Recommended only in environments with reliable, low-latency access to the share.

.EXAMPLE
PS> Invoke-InstallerCacheRepair -ReportOnly

    Runs FixMissingMSI from the local temp install, performs discovery only,
    and generates the unresolved CSV report without making any file repairs.

.NOTES
    Author: Joey Eckelbarger

    Credits:
        FixMissingMSI is authored and maintained by suyouquan  
        Source: https://github.com/suyouquan/SQLSetupTools/releases/tag/V2.2.1

    Behavior:
        - If -FileSharePath is omitted, assumes FixMissingMSI is present in $LocalWorkPath.
        - If FixMissingMSI.exe is missing locally, the function will throw and recommend running Install-FixMissingMSI first.
        - All repairs are performed locally under C:\Windows\Installer.

    Security:
        The script writes to C:\Windows\Installer when executing FixCommands.
        Ensure it runs with administrative rights (see #Requires -RunAsAdministrator).

    Requirements:
        - PowerShell 5.1+
        - Administrative privileges
        - NTFS permissions to write to $LocalWorkPath
        - (Optional) Read/write permissions on the share when using -FileSharePath
#>
function Invoke-InstallerCacheRepair {
    param(
        [string]$FileSharePath,
        [string[]]$SourcePaths = "",
        [string]$LocalWorkPath = (Join-Path $env:TEMP 'FixMissingMSI'),
        [switch]$RunFromShare,
        [switch]$ReportOnly
    )

    $ErrorActionPreference = 'Stop'

    if($null -eq $FileSharePath -and $(Test-Path -Path (Join-Path $LocalWorkPath "FixMissingMSI.exe")) -eq $false){
        Throw "No FileSharePath specified and FixMissingMSI.exe is not present in $LocalWorkPath, please specify a FileSharePath for FixMissingMSI to copy from or run Install-FixMissingMSI if running without a fileshare."
    }

    # We want to ensure an empty string to be in here to ensure that the for loop runs at least 1x with an emptry string so FixMissingMSI tries to recover using sources from the original installation metadata in registry
    if($sourcePaths -notcontains ""){
        $sourcePaths += ""
    }

    # Compose shared paths (app folder and caches) once. Keep all shared I/O under the app folder.\
    if($FileSharePath){
        $ShareRoot      = $FileSharePath.TrimEnd('\')
        $CacheRoot      = Join-Path $AppFolder 'Cache'
        $ProductsCache  = Join-Path $CacheRoot 'Products'
        $PatchesCache   = Join-Path $CacheRoot 'Patches'
        $ReportsPath    = Join-Path $AppFolder 'Reports'
        $AppFolder      = Join-Path $ShareRoot 'FixMissingMSI'
        $RunningLocally = $false
    } else {
        Write-Warning "Running locally only as FileSharePath parameter was not passed..."
        $RunningLocally = $true # used later to ensure we dont try to write/do actions requiring the share to be defined. 
        $AppFolder      = $LocalWorkPath
        $ReportsPath    = $LocalWorkPath 
    }

    # Prepare local working directory and transcript path first (Start-Transcript requires an existing folder).
    if (-not (Test-Path -LiteralPath $LocalWorkPath)) {
        New-Item -ItemType Directory -Path $LocalWorkPath -Force | Out-Null
    }
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $TranscriptPath = Join-Path $LocalWorkPath "Transcript-Invoke-InstallerCacheRepair-$($env:COMPUTERNAME)-$timestamp.txt"
    Start-Transcript -Path $TranscriptPath | Out-Null

    try {
        # Stage FixMissingMSI locally. Copy only top-level binaries/config files.
        # Why: We need FixMissingMSI.exe and its dependencies, but not Cache\ or Reports\ folders.
        if($RunFromShare -eq $false){
            Get-ChildItem -Path $AppFolder -File | ForEach-Object {
                if((Test-Path -LiteralPath (Join-Path $LocalWorkPath $_.Name)) -eq $false){
                    Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $LocalWorkPath $_.Name) -Force
                }
            }

            $exePath = Join-Path $LocalWorkPath 'FixMissingMSI.exe'
        } else {
            $exePath = Join-Path $AppFolder 'FixMissingMSI.exe'
        }

        Push-Location $LocalWorkPath

        try {
            $serverName = $env:COMPUTERNAME

            $mergedBadRowsWithoutFix = @()
            $mergedFixCommands       = @()

            foreach ($source in $SourcePaths) {
                $sourceLabel = if($source -eq ""){
                    "No specified source; FixMissingMSI will look for local sources from Registry install metadata (as well as the shared cache if available/populated)."
                } else {
                    $source
                }

                # 1) Load the FixMissingMSI assembly
                if (-not (Test-Path -LiteralPath $exePath)) {
                    throw "FixMissingMSI.exe not found at expected path: $exePath"
                }
                $asm = [System.Reflection.Assembly]::LoadFrom($exePath)

                # Create an instance of the main form. The backend expects a form handle; no UI is shown.
                Add-Type -AssemblyName System.Windows.Forms
                $formType = $asm.GetTypes() | Where-Object { $_.Name -eq 'Form1' }
                if (-not $formType) { throw "Form1 type not found in FixMissingMSI assembly." }
                $form = [Activator]::CreateInstance($formType)
                [System.Windows.Forms.Application]::EnableVisualStyles()
                [void]$form.Handle  # Initialize handle to avoid null reference exceptions in backend calls.

                # 2) Access internal types and required fields
                $myDataType = $asm.GetType('FixMissingMSI.myData')
                if (-not $myDataType) { throw "Type FixMissingMSI.myData not found." }

                # Set filters off and clear filter string
                $null = $myDataType.GetField('isFilterOn', [Reflection.BindingFlags]'Static,Public')
                $fldFilterStr = $myDataType.GetField('filterString', [Reflection.BindingFlags]'Static,Public')
                if ($fldFilterStr) { $fldFilterStr.SetValue($null, '') }

                # 3) Point to setup media source directory
                $fldSetupSource = $myDataType.GetField('setupSource', [Reflection.BindingFlags]'Static,Public')
                if ($fldSetupSource) { $fldSetupSource.SetValue($null, $source) }

                # 4) Scan supplied media for MSI/MSP packages
                $scanMedia = $myDataType.GetMethod('ScanSetupMedia', [Reflection.BindingFlags]'Static,Public')
                if ($scanMedia) { $null = $scanMedia.Invoke($null, @()) }

                # 5) Scan installed products and patches
                $scanProducts = $myDataType.GetMethod('ScanProducts', [Reflection.BindingFlags]'Static,Public')
                if ($scanProducts) { $null = $scanProducts.Invoke($null, @()) }

                # 6) Include extra packages from LastUsedSource (mirrors AfterDone behavior)
                $addFromLast = $myDataType.GetMethod('AddMsiMspPackageFromLastUsedSource', [Reflection.BindingFlags]'Static,NonPublic')
                if ($addFromLast) { $null = $addFromLast.Invoke($null, @()) }

                # 7) Generate FixCommand strings for missing/mismatched rows
                $updateFix = $myDataType.GetMethod('UpdateFixCommand', [Reflection.BindingFlags]'Static,Public')
                if ($updateFix) { $null = $updateFix.Invoke($null, @()) }

                # 8) Retrieve rows collection
                $rowsField = $myDataType.GetField('rows', [Reflection.BindingFlags]'Static,Public')
                $rows = if ($rowsField) { $rowsField.GetValue($null) } else { $null }
                if (-not $rows) { Write-Verbose "No rows returned from myData.rows."; continue }

                # 9) Filter to missing or mismatched
                $badRows = $rows | Where-Object { $_.Status -in 'Missing','Mismatched' }

                if($null -eq $badRows){
                    Write-Output "No missing files; exiting"
                    break 
                }

                # 9.5) If FixCommand is empty, try building a COPY command from the shared cache
                if($RunningLocally -eq $false){
                    foreach ($row in ($badRows | Where-Object { -not $_.FixCommand })) {
                        if($row.ProductCode -and $row.PackageCode -and $row.PackageName){
                            $productCandidate = Join-Path $ProductsCache (Join-Path $($row.ProductCode) (Join-Path $($row.PackageCode) $($row.PackageName)))
                            if ((Test-Path -LiteralPath $productCandidate)) {
                                Write-Output "Found missing files in shared cache, populating FixCommand value"
                                $row.FixCommand = "COPY `"$productCandidate`" `"C:\Windows\Installer\$($row.CachedMsiMsp)`""
                                continue
                            }
                        }
    
                        if($row.ProductCode -and $row.PatchCode -and $row.PackageName){
                            $patchCandidate   = Join-Path $PatchesCache (Join-Path $($row.ProductCode) (Join-Path $($row.PatchCode)   $($row.PackageName)))
                            if ((Test-Path -LiteralPath $patchCandidate)) {
                                Write-Output "Found missing files in shared cache, populating FixCommand value"
                                $row.FixCommand = "COPY `"$patchCandidate`" `"C:\Windows\Installer\$($row.CachedMsiMsp)`""
    
                                continue
                            }
                        }
                    }
                }

                $badRowsWithFix    = @($badRows | Where-Object { $_.FixCommand })      | Select-Object *,@{N='Hostname';E={$serverName}}, @{N='SourcePath';E={$source}}, @{N="CompareString";E={"$($_.ProductCode)-$($_.PackageCode)-$($_.PatchCode)-$($_.PackageName)"}}
                $badRowsWithoutFix = @($badRows | Where-Object { -not $_.FixCommand }) | Select-Object *,@{N='Hostname';E={$serverName}}, @{N='SourcePath';E={$source}}, @{N="CompareString";E={"$($_.ProductCode)-$($_.PackageCode)-$($_.PatchCode)-$($_.PackageName)"}}
                
                if($badRowsWithoutFix){
                    $mergedBadRowsWithoutFix += $badRowsWithoutFix | Where-Object {$_.CompareString -notin $mergedBadRowsWithoutFix.CompareString}
                }

                $mergedBadRowsWithoutFix = $mergedBadRowsWithoutFix | Where-Object {$_.CompareString -notin $badRowsWithFix.CompareString}
                

                [array]$mergedFixCommands += $badRowsWithFix | Select-Object -ExpandProperty FixCommand

                $missingCount    = @($badRows | Where-Object { $_.Status -eq 'Missing'   }).Count
                $mismatchedCount = @($badRows | Where-Object { $_.Status -eq 'Mismatched'}).Count
                Write-Output "Source: $sourceLabel"
                Write-Output "Missing: $missingCount`nMismatched: $mismatchedCount`nTo be fixed: $($badRowsWithFix.Count)"
            }

            $mergedFixCommands = $mergedFixCommands | Sort-Object * -Unique

            # 10) Execute fix commands. Copies to C:\Windows\Installer as needed.
            foreach ($fixCommand in $mergedFixCommands) {
                Write-Output "$fixCommand" # logs cmd executed to the transcript
                if($reportOnly){
                    continue # dont run if report only
                }
                & cmd /c $fixCommand
            }
            # Export unresolved rows to central report (per host). Overwrites by host design; adjust if you prefer timestamped files.
            $reportFile = Join-Path $ReportsPath "$serverName.csv"
            $mergedBadRowsWithoutFix | Sort-Object * -Unique | Export-Csv -Path $reportFile -NoTypeInformation -Force
        } finally {
            Pop-Location
        }
    } finally {
        Stop-Transcript | Out-Null
    }
}

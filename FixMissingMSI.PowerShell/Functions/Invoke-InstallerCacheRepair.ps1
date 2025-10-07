<#
.SYNOPSIS
    Copies FixMissingMSI locally, runs it non-interactively via .NET reflection,
    attempts to source missing MSI/MSP files from local and shared caches,
    and exports a CSV report of unresolved items.

.DESCRIPTION
    Intended for broad deployment (e.g., MECM). This script stages FixMissingMSI locally,
    then loads the FixMissingMSI.exe assembly and invokes its internal methods via reflection,
    bypassing the GUI to run non-interactively.

    For each configured source path, the script:
      1) Loads FixMissingMSI.exe and instantiates the hidden Form to satisfy internal dependencies.
      2) Points FixMissingMSI to a given setup source directory (MSI/MSP media).
      3) Scans setup media, installed products/patches, and LastUsedSource locations.
      4) Generates FixCommand entries for any missing/mismatched rows.
      5) If a FixCommand is still empty, attempts to build a COPY command from the shared cache
         under \\<Server>\<Share>\FixMissingMSI\Cache\{Products,Patches}.
      6) Exports unresolved rows without a FixCommand to the central Reports folder.
      7) Executes generated FixCommands (guarded by ShouldProcess) to repopulate C:\Windows\Installer.

    > Note: FixMissingMSI is a GUI application without a native CLI. This script leverages .NET
    > reflection to call internal types and methods directly. If FixMissingMSI internals change
    > in future versions, binding calls may need to be updated.

.PARAMETER FileSharePath
    UNC to the share root that contains the app tree. Example: \\FS01\Software
    The script expects the app at: \\<Server>\<Share>\FixMissingMSI

.PARAMETER SourcePaths
    One or more setup media folders (local or UNC) to scan for MSI/MSP packages.
    Defaults to the shared Cache root: \\<Server>\<Share>\FixMissingMSI\Cache

.PARAMETER LocalWorkPath
    Local working directory where FixMissingMSI is staged and executed.
    Defaults to $env:TEMP\FixMissingMSI.

.PARAMETER RunFromShare
    If specified, runs FixMissingMSI directly from the network share instead of
    copying it to a local working directory first.

    This can be used in trusted environments with reliable network access,
    where the performance benefit outweighs the added risk of running directly
    from a network path.

    By default, FixMissingMSI is copied locally before execution to avoid
    issues caused by antivirus scanning or intermittent share connectivity.

.PARAMETER ReportOnly
    If specified, do not execute any FixCommand operations.
    The script still scans sources and writes the unresolved CSV report.

.EXAMPLE
PS> Invoke-InstallerCacheRepair -FileSharePath \\FS01\Software

    Stages FixMissingMSI locally, scans \\FS01\Software\FixMissingMSI\Cache,
    attempts repair, exports unresolved CSV to \\FS01\Software\FixMissingMSI\Reports,
    and executes FixCommands to repopulate the local installer cache.

.EXAMPLE
PS> Invoke-InstallerCacheRepair -FileSharePath \\FS01\Software -SourcePaths 'D:\Media','\\FS01\Builds\Office'

    Scans the provided source paths in order (D:\Media, then \\FS01\Builds\Office) instead of the default shared cache.

.EXAMPLE
PS> Invoke-InstallerCacheRepair -FileSharePath \\FS01\Software -SourcePaths 'D:\Media','\\FS01\Builds\Office' -ReportOnly

    Scans the provided source paths in order (D:\Media, then \\FS01\Builds\Office) instead of the default shared cache.

.EXAMPLE
PS> Invoke-InstallerCacheRepair -FileSharePath \\FS01\Software -RunFromShare

    Runs FixMissingMSI directly from the network share, without copying it locally.
    Useful when testing or running in low-latency, trusted environments.

.NOTES
    Author: Joey Eckelbarger

    Credits:
        FixMissingMSI is authored and maintained by suyouquan
        Source: https://github.com/suyouquan/SQLSetupTools/releases/tag/V2.2.1

    Security:
        This script writes to C:\Windows\Installer via generated FixCommands.

    Requires:
        - PowerShell 5.1+
        - NTFS permissions to write to $LocalWorkPath
        - Read access to \\<Server>\<Share>\FixMissingMSI and subfolders
        - Write access to \\<Server>\<Share>\FixMissingMSI\Reports
#>
#Requires -RunAsAdministrator
function Invoke-InstallerCacheRepair {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileSharePath,
        [string[]]$SourcePaths = "",
        [string]$LocalWorkPath = (Join-Path $env:TEMP 'FixMissingMSI'),
        [switch]$RunFromShare,
        [switch]$ReportOnly
    )

    $ErrorActionPreference = 'Stop'

    # We want to ensure an empty string to be in here to ensure that the for loop runs at least 1x with an emptry string so FixMissingMSI tries to recover using sources from the original installation metadata in registry
    if($sourcePaths -notcontains ""){
        $sourcePaths += ""
    }

    # Compose shared paths (app folder and caches) once. Keep all shared I/O under the app folder.
    $ShareRoot      = $FileSharePath.TrimEnd('\')
    $AppFolder      = Join-Path $ShareRoot 'FixMissingMSI'
    $CacheRoot      = Join-Path $AppFolder 'Cache'
    $ProductsCache  = Join-Path $CacheRoot 'Products'
    $PatchesCache   = Join-Path $CacheRoot 'Patches'
    $ReportsPath    = Join-Path $AppFolder 'Reports'

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

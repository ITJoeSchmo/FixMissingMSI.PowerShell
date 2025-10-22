<#
.SYNOPSIS
    Downloads and installs the FixMissingMSI tool into a temporary working directory.

.DESCRIPTION
    Retrieves the FixMissingMSI utility from a specified URI (defaulting to the latest public GitHub release)
    and installs it to the local temporary path.

    It performs the following steps:
    1. Removes any previous FixMissingMSI artifacts from the target temp folder.
    2. Downloads the FixMissingMSI .zip archive from the provided URI.
    3. Expands the archive into a local subfolder under $Path\FixMissingMSI.

    The installation occurs entirely in the local temp directory (e.g. $env:TEMP),
    and does not modify system-wide locations or registry keys.

.PARAMETER FixMissingMsiUri
    URI to the FixMissingMSI zip archive to download. Defaults to the current public release:
    https://github.com/suyouquan/SQLSetupTools/releases/download/V2.2.1/FixMissingMSI_V2.2.1_For_NET45.zip

.PARAMETER Path
    Destination directory for the downloaded and extracted files.
    Defaults to the user's temporary directory ($env:TEMP).

.EXAMPLE
PS> Install-FixMissingMSI

    Downloads FixMissingMSI to the current user's temp directory, replacing any previous version,
    and expands it to $env:TEMP\FixMissingMSI.

.EXAMPLE
PS> Install-FixMissingMSI -Path 'C:\Temp\Tools'

    Installs FixMissingMSI into C:\Temp\Tools\FixMissingMSI.

.NOTES
    Author: Joey Eckelbarger
#>
function Install-FixMissingMSI {
    param(
        [uri]$FixMissingMsiUri = 'https://github.com/suyouquan/SQLSetupTools/releases/download/V2.2.1/FixMissingMSI_V2.2.1_For_NET45.zip',
        [string]$Path = $env:TEMP
    )

    $ErrorActionPreference = 'Stop'
    $ProgressPreference = 'SilentlyContinue'  # Progress UI slows iwr download speed noticeably.

    $ZipPath   = Join-Path $Path 'FixMissingMSI.zip'
    $AppFolder = Join-Path $Path 'FixMissingMSI'

    # Ensure TLS 1.2 on older hosts (e.g., Server 2016) to avoid protocol negotiation failures.
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


    # Clean previous temp artifacts to avoid mixing versions.
    # Separate guard for destructive actions to make -Confirm meaningful here.
    if (Test-Path -LiteralPath $ZipPath) {
        Remove-Item -LiteralPath $ZipPath -Force
    }
    if (Test-Path -LiteralPath $AppFolder) {
        Remove-Item -LiteralPath $AppFolder -Recurse -Force
    }

    # Download  tool.
    Invoke-WebRequest -Uri $FixMissingMsiUri -UseBasicParsing -OutFile $ZipPath

    # Unblock and expand. MOTW can block execution in some environments. I don't think iwr downloads get tagged with MOTW but just to be 100% I included it.
    Unblock-File -LiteralPath $ZipPath
    Expand-Archive -LiteralPath $ZipPath -DestinationPath $AppFolder -Force

    Write-Output "Downloaded and installed FixMissingMSI: $AppFolder"
}
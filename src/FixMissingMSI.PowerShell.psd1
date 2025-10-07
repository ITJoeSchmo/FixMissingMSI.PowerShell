@{
    # Script module or binary module file associated with this manifest
    RootModule        = 'FixMissingMSI.PowerShell.psm1'

    # Version of this module.
    ModuleVersion     = '1.0.0'

    # ID used uniquely identify this module
    GUID              = '69ffbf20-83d2-4eb5-88b1-9ce47ddcd7eb'

    # Author of this module
    Author            = 'Joey Eckelbarger'

    # Company or vendor of this module
    CompanyName       = 'Joey Eckelbarger'

    # Copyright statement for this module
    Copyright         = '(c) 2025 Joey Eckelbarger. Licensed under MIT License.'

    # Description of the functionality provided by this module
    Description       = 'PowerShell module for detecting, reporting, and repairing missing Windows Installer (MSI/MSP) cache files. Automates the FixMissingMSI utility for non-interactive, scalable recovery using a shared cache model.'

    # Minimum PowerShell version supported
    PowerShellVersion = '5.1'

    # Functions to export from this module
    FunctionsToExport = @(
        'Initialize-InstallerCacheFileShare',
        'Invoke-InstallerCacheRepair',
        'Merge-InstallerCacheReports',
        'Update-InstallerCache',
        'Get-InstallerRegistration',
        'Remove-InstallerRegistration'
    )

    # Cmdlets to export from this module
    CmdletsToExport   = @()

    # Variables to export from this module
    VariablesToExport = '*'

    # Aliases to export from this module
    AliasesToExport   = @()

    # Private data to pass to the module
    PrivateData = @{
        PSData = @{
            # Tags for PowerShell Gallery / discoverability
            Tags = @('MSI','MSP','InstallerCache','WindowsInstaller','Automation','FixMissingMSI')

            # License and Project URLs
            LicenseUri = 'https://opensource.org/licenses/MIT'
            ProjectUri = 'https://github.com/ITJoeSchmo/FixMissingMSI.PowerShell'
            IconUri    = ''

            # Release notes
            ReleaseNotes = @"
Initial release of FixMissingMSI.PowerShell (v1.0.0)

Highlights:
- Automates FixMissingMSI through PowerShell for non-interactive execution.
- Enables centralized detection and repair of missing MSI/MSP cache files.
- Supports shared, demand-driven cache population across systems.
- Adds advanced recovery helpers for inspecting and removing broken MSI registrations.
- Designed for easy integration with MECM, Azure Arc, Intune, or standalone PowerShell execution.
"@
        }
    }
}



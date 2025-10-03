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
    Copyright         = ''

    # Description of the functionality provided by this module
    Description       = 'Automates recovery of missing Windows Installer cache (MSI/MSP) files at scale, leveraging FixMissingMSI for discovery and a shared cache model for repair.'

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
            ProjectUri = 'https://github.com/yourusername/FixMissingMSI.PowerShell'
            IconUri    = ''

            # Release notes
            ReleaseNotes = 'Initial release of FixMissingMSI.PowerShell module. Provides functions to initialize a fileshare, run FixMissingMSI non-interactively, merge reports, and populate the fileshare installer cache.'
        }
    }
}

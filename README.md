# FixMissingMSI.PowerShell

When Windows Installer cache files (`C:\Windows\Installer\*.msi` / `*.msp`) are missing, updates and uninstalls can break for products like SQL Server, Exchange, Azure Arc Agent, and others. The original [**FixMissingMSI**](https://github.com/suyouquan/SQLSetupTools) utility helps identify and restore these files, but it is a GUI-only tool.

**FixMissingMSI.PowerShell** turns that manual process into a repeatable, non-interactive workflow. It lets you scan systems at scale, merge reports, and repopulate missing files from a shared cache — all orchestrated through PowerShell and easily integrated with centralized management tools.

---

## Why This Exists

Windows Installer cache corruption is a common problem that prevents applications from updating or uninstalling cleanly. Typical scenarios include:

- Disk cleanup scripts that delete files from `C:\Windows\Installer` without validating against both Products and Patches registry keys.
- Missing MSI files breaking updates for SQL Server, Exchange Server, Azure Arc Agent, Microsoft Edge, and other applications.
- Manual repair processes requiring sourcing of the original installer media for every affected server.

**FixMissingMSI** is a GUI-only troubleshooting tool. On its own it:
- Cannot be run non-interactively.
- Cannot scale across many servers.
- Cannot share repair data between machines.

This project solves those limitations by automating FixMissingMSI with PowerShell.

---

## Features

- Non-interactive execution of FixMissingMSI via .NET Reflection
- Centralized reporting in `.CSV` format of missing MSI/MSP files
- Shared cache support with demand-driven population
- Integration with deployment tools (MECM, Intune, Ansible, Azure Arc)
- Safe ACL design for file share permissions
- Advanced recovery helpers for MSI registration scrubbing

---

## Installation

Clone this repository and import the module:

```powershell
git clone https://github.com/ITJoeSchmo/FixMissingMSI.PowerShell
Import-Module ./src/FixMissingMSI.PowerShell.psd1
````

(Planned: publish to PowerShell Gallery -> `Install-Module FixMissingMSI.PowerShell`)

---

## Exported Functions

### Core Cache Automation

| Function                             | Purpose                                                              |
| ------------------------------------ | -------------------------------------------------------------------- |
| `Initialize-InstallerCacheFileShare` | Prepares a file share with the tool and reports, applies scoped ACLs |
| `Invoke-InstallerCacheRepair`        | Non-interactively scans each system, outputs per-host CSV report     |
| `Merge-InstallerCacheReports`        | Merges per-host CSV reports into a deduplicated summary              |
| `Update-InstallerCache`              | Populates the shared cache with only the required MSI/MSP files      |

### Registration Helpers

| Function                       | Purpose                                                                   |
| ------------------------------ | ------------------------------------------------------------------------- |
| `Get-InstallerRegistration`    | Lists MSI-registered products from the Windows Installer registry         |
| `Remove-InstallerRegistration` | Forcefully removes MSI registrations when repair/uninstall isn’t possible |

---

## Quickstart

```powershell
# Initialize cache share
Initialize-InstallerCacheFileShare -FileSharePath "\\FS01\"

# Scan a system
Invoke-InstallerCacheRepair -FileSharePath "\\FS01\" -SourcePaths '\\FS01\SQL Server 2019\setup\'

# Merge reports
Merge-InstallerCacheReports -FileSharePath "\\FS01\"

# Populate shared cache
Update-InstallerCache -FileSharePath "\\FS01\"

# Explore registered MSI products
Get-InstallerRegistration -Filter {$_.DisplayName -like "*SQL*"}

# Scrub problematic registration (last resort to enable installing over the bad installation)
Remove-InstallerRegistration -Filter { $_.DisplayName -like "Azure Connected Machine Agent*" -and $_.DisplayVersion -eq "1.56.03167" }
```

---

## Example Workflows

See [examples/](examples/) for:

* **Standalone machine**: detect and repair cache issues on a single host
* **MECM deployment**: run steps across entire collections and centralize results
* **Azure Arc or Intune**: push scans to hybrid or cloud-managed hosts

---

## Common Errors

This project automates recovery for errors such as:

* `The cached MSI file 'C:\Windows\Installer\xxxxxx.msi' is missing`
* `SQL Server Setup has encountered the following error: The cached MSI file is missing`
* `Missing MSI or MSP files prevent service pack or cumulative update installation`

---

## Security Posture

* The tool share is **read/execute only** for Domain Computers
* Cache and reports are **writable as scoped**
* Servers cannot write to the tool folder itself

---

## How the Non-Interactive Execution Works

FixMissingMSI was designed as a WinForms GUI with no CLI support.
This module bypasses the UI by loading the EXE via .NET Reflection and invoking internal methods directly.

---

## Credits

* FixMissingMSI is authored and maintained by [suyouquan (Simon Su @ Microsoft)](https://github.com/suyouquan/SQLSetupTools)
* Installer registration helper functions adapted from Microsoft’s [Program Install and Uninstall Troubleshooter](https://support.microsoft.com/en-us/topic/fix-problems-that-block-programs-from-being-installed-or-removed-cca7d1b6-65a9-3d98-426b-e9f927e1eb4d)

---

## Roadmap

* Publish module to PowerShell Gallery
* Add orchestration examples
* Optional read-only discovery mode for audit-only runs
* Evaluate replacing FixMissingMSI dependency with native implementation (Update-InstallerCache doesn't rely on FixMissingMSI so there is some groundwork laid)

---

## License

MIT License © 2025 Joey Eckelbarger

```

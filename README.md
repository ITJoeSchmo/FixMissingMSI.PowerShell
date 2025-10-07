# FixMissingMSI.PowerShell - Windows Installer Cache Recovery Automation

[![PowerShell Gallery](https://img.shields.io/powershellgallery/v/FixMissingMSI.PowerShell.svg)](https://www.powershellgallery.com/packages/FixMissingMSI.PowerShell)
[![Downloads](https://img.shields.io/powershellgallery/dt/FixMissingMSI.PowerShell.svg)](https://www.powershellgallery.com/packages/FixMissingMSI.PowerShell)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

---

When Windows Installer cache files (`C:\Windows\Installer\*.msi` / `*.msp`) are missing, updates and uninstalls can fail for products like SQL Server, Exchange, Azure Arc Agent, and others.  
The original [**FixMissingMSI**](https://github.com/suyouquan/SQLSetupTools) utility identifies and restores these files but is GUI-only.

**FixMissingMSI.PowerShell** turns that manual process into a repeatable, non-interactive workflow.  
It lets you scan systems at scale, merge reports, and repopulate missing files from a shared cache - all orchestrated through PowerShell.

---

## Features
- Non-interactive execution of FixMissingMSI via .NET Reflection  
- Centralized reporting of missing MSI/MSP files  
- Shared cache support with demand-driven population  
- Secure ACL design for file share permissions  
- Advanced recovery helpers for MSI registration scrubbing  

---

## Installation

From the PowerShell Gallery:
```powershell
Install-Module FixMissingMSI.PowerShell
Import-Module FixMissingMSI.PowerShell
```

Or from source:

```powershell
git clone https://github.com/ITJoeSchmo/FixMissingMSI.PowerShell
Import-Module ./src/FixMissingMSI.PowerShell.psd1
```

---

## Exported Functions

| Function                             | Purpose                                                             |
| ------------------------------------ | ------------------------------------------------------------------- |
| `Initialize-InstallerCacheFileShare` | Prepare a file share with the tool and reports, apply scoped ACLs   |
| `Invoke-InstallerCacheRepair`        | Run FixMissingMSI non-interactively and output per-host CSV reports |
| `Merge-InstallerCacheReports`        | Merge host reports into a deduplicated summary                      |
| `Update-InstallerCache`              | Populate the shared cache with only the required MSI/MSP files      |
> `Update-InstallerCache` copies only the specific MSI/MSP files that appear as missing in any server’s report -- it does *not* mirror every installer file from each system’s installer cache
> This keeps the cache lean and focused on real recovery needs.
### Key Parameters: `Invoke-InstallerCacheRepair`

| Parameter | Type | Description |
|------------|------|-------------|
| **FileSharePath** | `String` *(Required)* | UNC path where the shared FixMissingMSI directory exists (e.g. `\\FS01\Software` ). |
| **SourcePaths** | `String[]` | One or more local or UNC paths containing MSI/MSP to scan as source for recovering missing files. By default, it also checks the shared cache (e.g. `\\FS01\Software\FixMissingMSI\Cache\{Products, Patches}`). |
| **LocalWorkPath** | `String` | Local directory where FixMissingMSI is staged and executed locally. Defaults to `$env:TEMP\FixMissingMSI`. |
| **RunFromShare** | `Switch` | Runs FixMissingMSI directly from the network share instead of copying locally. |
| **ReportOnly** | `Switch` | Performs discovery and reporting only -- no 'FixCommands' are executed, but still logs all FixCommands to the transcript and exports unresolved entries to CSV.. Useful for audit-only scans. |

> For full syntax, parameters, examples of other functions run:
> ```powershell
> Get-Help <function> -Full
> ```

### Extra

| Function                       | Purpose                                                                                          |
| ------------------------------ | ------------------------------------------------------------------------------------------------ |
| `Get-InstallerRegistration`    | Lists MSI-registered products from the Windows Installer registry.                               |
| `Remove-InstallerRegistration` | Forcefully removes or scrubs MSI registrations when standard uninstall or repair isn’t possible. |

> **⚠️ Warning**
> `Remove-InstallerRegistration` is an advanced recovery function adapted from Microsoft’s *Program Install and Uninstall Troubleshooter* (`MSIMATSFN.ps1`).
> It should be used **only after standard uninstall or repair methods have failed**.
>
> When invoked, the function first attempts a normal uninstallation by calling
> `msiexec /x <ProductCode>` to remove program files and registry data through Windows Installer.
>
> If the installation is broken and cannot be removed by MSI:
>
> * The function runs the **Rapid Product Removal (RPR)** phase, which scrubs Windows Installer registration, cached metadata, and uninstall entries so that the product can be reinstalled.
> * When run with **`-DeepClean`**, it additionally performs Microsoft’s **LPR** phase, which locates and purges related files, shims, ARP entries, and orphaned registry keys.
>
> During either cleanup phase, the function automatically:
>
> * **Backs up files and registry data** related to the target product under
>   `C:\MATS\<ProductCode>\`
> * Generates a **PowerShell restore script (`RestoreYourFilesAndRegistry.ps1`)** in the same directory, which can be used to revert any changes if needed.
>
> These backups are created before any deletion occurs, allowing administrators to inspect or restore the removed data manually if necessary.
>
> This command is intended strictly as a *last-resort recovery tool* to enable clean reinstallation of affected software.
> Always verify backups or create a system restore point before use.

---

## Quickstart

```powershell
# Initialize cache share
Initialize-InstallerCacheFileShare -FileSharePath "\\FS01\InstallerCache"

# Scan a system and try to source missing files from "\\FS01\SQLServer2019\setup\"
Invoke-InstallerCacheRepair -FileSharePath "\\FS01\InstallerCache" -SourcePaths "\\FS01\SQLServer2019\setup\"

# Merge reports
Merge-InstallerCacheReports -FileSharePath "\\FS01\InstallerCache"

# Populate shared cache
Update-InstallerCache -FileSharePath "\\FS01\InstallerCache"
```

> Full example workflows are provided under [examples/](examples/).

---

## Common Errors

This project automates recovery for:

* `The cached MSI file 'C:\Windows\Installer\xxxxxx.msi' is missing`
* `SQL Server Setup has encountered the following error: The cached MSI file is missing`
* `Missing MSI/MSP files prevent service pack or cumulative update installation`

---

## Security Posture

* The application folder is **read/execute only** for Domain Computers
* Cache and reports folders are **writable as scoped**
* Servers cannot write to the tool folder itself

---

## How the Non-Interactive Execution Works

FixMissingMSI is a WinForms GUI with no CLI support.
This module loads the EXE via .NET Reflection, instantiates its form classes to initialize internal state, and invokes internal methods directly.

---

## Credits

* **FixMissingMSI** authored by [suyouquan (Simon Su @ Microsoft)](https://github.com/suyouquan/SQLSetupTools)
* Installer-registration helper functions adapted from Microsoft’s [Program Install and Uninstall Troubleshooter](https://support.microsoft.com/en-us/topic/fix-problems-that-block-programs-from-being-installed-or-removed-cca7d1b6-65a9-3d98-426b-e9f927e1eb4d)

---

## Roadmap

* Add orchestration examples for Ansible and Azure Arc
* Evaluate replacing FixMissingMSI dependency with native implementation (`Update-InstallerCache` already functions independently)

---

## License

MIT License © 2025 Joey Eckelbarger

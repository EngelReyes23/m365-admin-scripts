# Microsoft 365 Administration Scripts

Interactive PowerShell tools for troubleshooting, analyzing, and maintaining Microsoft 365, SharePoint Online, OneDrive for Business, and synchronized local content.

These are personal administration tools, not official Microsoft products. Review the code, permissions, target scope, and generated reports before using destructive operations in production.

The tools share a consistent bilingual terminal interface. Each `.ps1` is distributed as a standalone single file: `Common/`, `tools/`, `tests/`, and repository resources are development-only and are not required at runtime.

## Tools

| Script | Description |
| --- | --- |
| `M365-Legacy-UserId-Cleaner.ps1` | Finds and repairs legacy User ID mismatches in SharePoint and OneDrive `UserInfoList` entries. Supports simulation mode, temporary administrative elevation, restoration, diagnostics, and reports. |
| `M365-Permissions-Scope-Manager.ps1` | Analyzes unique permission scopes in SharePoint and OneDrive libraries and can reset permission inheritance in live mode. |
| `M365-Recycle-Bin-and-Hold-Cleaner.ps1` | Reviews and cleans SharePoint and OneDrive recycle bins and Preservation Hold Library content. Supports simulation, site administration assistance, and cleanup reporting. |
| `M365-Version-History-Cleaner.ps1` | Reviews and cleans historical file versions in SharePoint Online and OneDrive for Business, with simulation, revalidation, and audit reporting. |
| `M365-Tenant-Inventory.ps1` | Creates a read-only inventory of SharePoint sites and OneDrive for Business, including status, owners, activity dates, and storage usage/quota. |
| `M365-Inactive-Site-Finder.ps1` | Finds SharePoint sites and OneDrive for Business locations whose last content modification exceeds a configurable inactivity threshold. It is read-only. |
| `M365-Site-Storage-Analyzer.ps1` | Analyzes SharePoint and OneDrive storage usage against configurable warning and critical thresholds and shows priority findings before exporting detailed CSV data. |
| `OneDrive-Path-Analyzer.ps1` | Scans synchronized OneDrive and SharePoint content locally for long Windows, cloud-relative, synchronization, and individual-name paths. It does not require Microsoft 365 authentication. |
| `OneDrive-Path-Analyzer.zsh` | Legacy macOS/Zsh path analyzer, retained separately from the PowerShell suite. |

## Design goals

The administration tools are designed around safe and repeatable workflows:

- English and Spanish interactive interfaces.
- Simulation or preview before destructive actions whenever supported.
- Revalidation immediately before changes are applied.
- CSV reports and local audit information.
- Persistent local configuration where appropriate.
- Modern PnP / Microsoft Entra ID authentication.
- Existing PnP application Client ID validation or guided new-app registration where implemented.
- PnP.PowerShell requirement checks with explicit install/update workflows.
- Throttling-aware retries for large request sequences where supported.

## Requirements

### Microsoft 365 scripts

- PowerShell 7.4 or later (`pwsh`, PowerShell Core).
- `PnP.PowerShell` 3.2 or later. The scripts detect it and can offer a CurrentUser installation/update.
- An account with the administrative role required by the selected operation.
- Site Collection Administrator permissions may be required for site-level operations. Depending on the script and operation, temporary elevation can be requested and removed afterward.

### OneDrive Path Analyzer

- Windows PowerShell 5.1 or PowerShell 7+.
- No Microsoft 365 authentication or PnP module is required for its local workflow.

### macOS Zsh analyzer

- Zsh and standard system utilities for `OneDrive-Path-Analyzer.zsh`.

## Usage

Download the individual `.ps1` you need and run it directly from PowerShell. No other repository file is required:

```powershell
./M365-Legacy-UserId-Cleaner.ps1
./M365-Permissions-Scope-Manager.ps1
./M365-Recycle-Bin-and-Hold-Cleaner.ps1
./M365-Version-History-Cleaner.ps1
./M365-Tenant-Inventory.ps1
./M365-Inactive-Site-Finder.ps1
./M365-Site-Storage-Analyzer.ps1
./OneDrive-Path-Analyzer.ps1
```

The Microsoft 365 tools guide you through tenant, application, target, permissions, authentication, and operation settings. The local path analyzer detects synchronized locations and lets you select locations or enter a manual path.

For the Zsh tool:

```zsh
chmod +x ./OneDrive-Path-Analyzer.zsh
./OneDrive-Path-Analyzer.zsh
```

## Authentication

The Microsoft 365 tools use modern PnP / Microsoft Entra ID authentication. Depending on the script and local configuration, the workflow can validate an existing Client ID or guide registration of a new PnP application. Authentication and permission requirements should always be reviewed before running against a production tenant.

## Reports and configuration

- Microsoft 365 scripts generate audit and operation reports locally.
- `OneDrive-Path-Analyzer.ps1` creates CSV reports on the Desktop when warning or critical paths are found.
- Each applicable TUI provides an option to open the reports folder.
- Working context can be cleared without deleting existing files or reports.
- Configuration and recovery state are optional local files created in the user's application-data location; they are not required repository resources.

## Safety model

1. Select and validate the target.
2. Analyze the current state.
3. Review the summary and generated report.
4. Use simulation/preview when available.
5. Revalidate the target immediately before modification.
6. Explicitly confirm the live operation.

Do not bypass these checks when working with production data. These tools can modify permissions, legacy user entries, file versions, recycle-bin items, or retained content. Test them in a controlled environment first.

## Repository scope

This repository contains current SharePoint Online and OneDrive for Business tooling. Older Exchange Online / Microsoft 365 scripts are preserved separately in [`Microsoft_365_Scripts`](https://github.com/EngelReyes23/Microsoft_365_Scripts) as historical reference.

## Disclaimer

Use these tools at your own risk. Microsoft 365 behavior, APIs, modules, authentication requirements, retention configuration, and service limits can change. Validate current Microsoft documentation and test in a controlled environment before production use.

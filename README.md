# Microsoft 365 Administration Scripts

A collection of interactive administration and troubleshooting tools for **SharePoint Online** and **OneDrive for Business**, built primarily with PowerShell and PnP.PowerShell.

The repository focuses on practical administrative scenarios such as version-history cleanup, legacy user remediation, permission-scope analysis, recycle-bin and preservation-hold cleanup, and OneDrive/SharePoint path analysis.

> These are personal administration tools, not official Microsoft products. Review the code, permissions, target scope, and generated reports before using any destructive operation in production.

## Tools

| Script | Purpose |
| --- | --- |
| `OneDrive-Path-Analyzer.zsh` | Analyzes long OneDrive/SharePoint paths on macOS and generates CSV findings. |
| `OneDrive-Path-Analyzer.ps1` | Analyzes long OneDrive/SharePoint paths on Windows. |
| `M365-Version-History-Cleaner.ps1` | Analyzes and safely cleans historical file versions in SharePoint Online and OneDrive for Business. |
| `M365-Legacy-UserId-Cleaner.ps1` | Detects and repairs legacy user entries with User ID mismatches in `UserInfoList`. |
| `M365-Recycle-Bin-and-Hold-Cleaner.ps1` | Manages SharePoint/OneDrive recycle-bin content and the Preservation Hold Library. |
| `M365-Permissions-Scope-Manager.ps1` | Analyzes unique permission scopes and can reset permission inheritance. |

## Design Goals

The PowerShell tools are designed around safe and repeatable administration workflows:

- English and Spanish interactive interfaces.
- Simulation / preview before destructive actions whenever possible.
- Revalidation immediately before changes are applied.
- CSV reports and local audit information.
- Persistent local configuration where appropriate.
- Modern PnP / Microsoft Entra ID authentication.
- Support for an existing PnP application Client ID or guided creation of a new PnP application where implemented.
- PnP.PowerShell requirement checks with install/update workflows where implemented.
- Throttling-aware operations, including retry behavior for transient Microsoft 365 responses in tools that perform large request sequences.

## Requirements

### Windows

- PowerShell 7.4 or later.
- `PnP.PowerShell` 3.2 or later for Microsoft 365 administration scripts.
- An account with sufficient tenant permissions and, depending on the operation, Site Collection Administrator permissions.

### macOS

- Zsh and standard system utilities for `OneDrive-Path-Analyzer.zsh`.

## Usage

Clone the repository and run the desired tool from PowerShell 7.4+:

```powershell
git clone https://github.com/EngelReyes23/m365-admin-scripts.git
cd m365-admin-scripts

./M365-Version-History-Cleaner.ps1
./M365-Legacy-UserId-Cleaner.ps1
./M365-Recycle-Bin-and-Hold-Cleaner.ps1
./M365-Permissions-Scope-Manager.ps1
./OneDrive-Path-Analyzer.ps1
```

On macOS:

```zsh
chmod +x ./OneDrive-Path-Analyzer.zsh
./OneDrive-Path-Analyzer.zsh
```

The interactive tools ask for the display language at startup. Select English or Spanish for the current run.

## Authentication

The Microsoft 365 tools use modern PnP / Microsoft Entra ID authentication.

Depending on the script and current local configuration, the workflow can validate a previously registered Client ID or guide the user through registering a PnP application. Authentication and permission requirements should always be reviewed before running a tool against a production tenant.

## Safety Model

Several tools can modify Microsoft 365 content or configuration. The general workflow is intentionally conservative:

1. Select and validate the target.
2. Analyze the current state.
3. Review the summary and generated report.
4. Run in simulation mode when available.
5. Revalidate the target immediately before modification.
6. Explicitly confirm the live operation.

Do not bypass these checks when working with production data.

## Reports and Configuration

The Microsoft 365 scripts store their configuration and reports locally in the corresponding user directory. Path analyzers generate CSV files containing detected findings, while administration tools can generate analysis and audit reports for review.

## Repository Scope

This repository contains current SharePoint Online and OneDrive for Business tooling. Older Exchange Online / Microsoft 365 scripts from an earlier stage of my work are preserved separately in [`Microsoft_365_Scripts`](https://github.com/EngelReyes23/Microsoft_365_Scripts) as historical reference.

## Disclaimer

Use these tools at your own risk. Microsoft 365 behavior, APIs, modules, authentication requirements, retention configuration, and service limits can change over time. Validate the current Microsoft documentation and test in a controlled environment before using the tools in production.

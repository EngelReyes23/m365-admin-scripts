# Microsoft 365 Administration Scripts

A collection of interactive tools for troubleshooting, analyzing, and maintaining OneDrive for Business and SharePoint Online.

## Herramientas

| Script | Purpose |
| --- | --- |
| `OneDrive-Path-Analyzer.zsh` | Analyzes long OneDrive/SharePoint paths on macOS and generates a CSV with warnings or critical findings. Requires Zsh. |
| `OneDrive-Path-Analyzer.ps1` | Analyzes long OneDrive/SharePoint paths on Windows. |
| `M365-Version-History-Cleaner.ps1` | Analyzes and cleans historical file versions in SharePoint Online and OneDrive for Business. Includes simulation, revalidation, and auditing. |
| `M365-Legacy-UserId-Cleaner.ps1` | Detects and repairs legacy user entries with User ID mismatches in `UserInfoList`. |
| `M365-Recycle-Bin-and-Hold-Cleaner.ps1` | Manages the SharePoint/OneDrive recycle bins and the Preservation Hold Library. |
| `M365-Permissions-Scope-Manager.ps1` | Counts unique permission scopes and can reset permission inheritance. |

## Requirements

- Windows: PowerShell 7.4 or later.
- macOS: Zsh and standard system utilities.
- Microsoft 365 tools: `PnP.PowerShell` module 3.2 or later.
- An account with sufficient tenant permissions and, depending on the operation, Site Collection Administrator permissions.

## Usage

On Windows, open PowerShell 7.4+ in this folder and run the desired script:

```powershell
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

Destructive tools start in simulation mode whenever possible. Always review the target, summary, and reports before confirming a real operation.

## Reports and Configuration

The Microsoft 365 scripts store configuration and reports locally in the corresponding user directory. The path analyzers generate CSV files containing detected findings.

## Warning

These tools can modify permissions, legacy users, versions, recycle bin items, or retained content. Use simulation mode first and validate the results in a controlled environment before running them in production.

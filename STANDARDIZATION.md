# Suite standardization — implementation checkpoint

## Scope and status

Eight PowerShell scripts exist in this checkout, not nine. All eight were read before changes. The Zsh analyzer is outside this PowerShell task. Existing README changes and the three previously untracked reporters were preserved. No commit/push, module installation, sign-in, application registration or Microsoft 365 operation was performed during development.

This is a **partial implementation**, not certification that every requested standard has been completed. The shared core and reporter adapters are implemented and tested. The remaining work is listed explicitly below.

## Single-file distribution verification

All eight published `.ps1` files are standalone. `Common/`, `tools/` and `tests/` are development infrastructure only. Shared functions and reporter translations are embedded in every generated script; none of the published scripts dot-sources them or reads their files at runtime.

Each script was copied by itself to a distinct empty temporary directory. Against that isolated copy, the test parsed it, loaded and exercised its offline helper functions, then launched the script directly and selected its normal Exit path. The isolated directory still contained exactly one file afterward. Proxy endpoints were set to a closed local port as an additional guard; no PnP, sign-in, registration or Microsoft 365 operation was reached. Path Analyzer also completed this isolated startup/exit test under Windows PowerShell 5.1.

| Script | Standalone | Repository dependency found | Change required | Parser | Isolated test |
| --- | --- | --- | --- | --- | --- |
| M365-Inactive-Site-Finder.ps1 | Yes | None | Embedded core/catalog retained | Pass | Pass |
| M365-Legacy-UserId-Cleaner.ps1 | Yes | None; redirected-console issue found | Guarded final `Clear-Host` | Pass | Pass |
| M365-Permissions-Scope-Manager.ps1 | Yes | None | Embedded core retained | Pass | Pass |
| M365-Recycle-Bin-and-Hold-Cleaner.ps1 | Yes | None | Embedded core retained | Pass | Pass |
| M365-Site-Storage-Analyzer.ps1 | Yes | None | Embedded core/catalog retained | Pass | Pass |
| M365-Tenant-Inventory.ps1 | Yes | None | Embedded core/catalog retained | Pass | Pass |
| M365-Version-History-Cleaner.ps1 | Yes | None; redirected-console issue found | Canonical `Clear-AppScreen` now skips `Clear-Host` when redirected | Pass | Pass |
| OneDrive-Path-Analyzer.ps1 | Yes | None | Embedded helpers retained; no PnP call during local workflow | Pass (7.x and 5.1) | Pass (7.x and 5.1) |

Configuration JSON and recovery-state JSON files are optional runtime state created beneath the user's configured application-data location. They are not bundled resources and their absence is a supported first-run condition. CSV files are outputs, not inputs. `PnP.PowerShell` is the only dynamically imported external dependency; its bootstrap is embedded. Path Analyzer does not invoke that bootstrap.

**Verified conclusion: 8/8 scripts can be downloaded individually and executed without any other repository file.**

## Canonical decisions derived from the audit

| Component | Existing reference | Implemented decision |
| --- | --- | --- |
| Header/layout | Version History, Legacy, Recycle, Path | Adaptive width capped at 84, fallback 72; same header, section hierarchy and 24-character field labels |
| Status | Three tenant reporters | ASCII `[INFO]`, `[OK]`, `[WARNING]`, `[ERROR]`; native ConsoleColor |
| Normal menus | Mature numeric menus | `[1]`…`[0]`, consistent spacing; adapters retain each caller's original return contract |
| Multiple selection | Permissions selector | Preserved; not replaced with a normal-menu implementation |
| Module checks | Recycle error handling | One shared bootstrap: compatible installed module, explicit installation consent, CurrentUser, revalidation, import, failure propagation |
| Application validation | Legacy | Three reporters now test authentication plus tenant enumeration before replacing saved Client ID; cleanup in finally |
| Destructive operations | Version preview and revalidation | Business logic preserved; common localized typed-confirmation helper added for Legacy/Recycle |
| Localization | Complete Spanish/English pairs in reporters | Tool-specific embedded catalogs + common keyed lookup; existing pairs retained, interpolation mechanically checked |
| Distribution | Standalone original scripts | Development sources are synchronized into each `.ps1`; no runtime dependency on Common/ or tools/ |

The core deliberately uses ASCII separators instead of requiring box-drawing support. Native output also avoids emitting explicit ANSI escapes into redirected output. This does not yet normalize every legacy table, selector or custom screen.

## Changes and regression notes by script

All scripts received the shared header/fields/status/menu/language/pause adapters. Original explicit versions remain unchanged; no release version was invented.

| Script | Additional changes | Business behavior preserved / residual risk |
| --- | --- | --- |
| M365-Inactive-Site-Finder.ps1 | 123 paired catalog keys; robust PnP/app bootstrap; configurable admin URL; precise system-site classification; English day suffix fixed | Inactivity cutoff, UNKNOWN handling, calculations and CSV commands preserved; real last-modified data needs tenant testing |
| M365-Site-Storage-Analyzer.ps1 | 132 paired catalog keys; same bootstrap, validation and URL changes | Quota thresholds, rounding and classification preserved; real quota/unknown cases need tenant testing |
| M365-Tenant-Inventory.ps1 | 135 paired catalog keys; same bootstrap, validation and URL changes; stack trace restricted to debug | Record construction and CSV commands preserved; actual enumeration/permissions need tenant testing |
| M365-Permissions-Scope-Manager.ps1 | Shared PnP; technical endpoints removed from discovery; target kind checked; pending temporary grant blocks target replacement; explicit registration confirmation; safe registration-result extraction | Scope counts, REST reset operations, filters and dry-run preserved; temporary-admin baseline and some context changes still need repair |
| M365-Version-History-Cleaner.ps1 | Shared PnP; precise discovery; validates target kind before connecting; null-safe application Client ID extraction; debug-only stack trace | Preview, retention, version analysis, per-file revalidation, deletion retry and CSV commands preserved |
| M365-Recycle-Bin-and-Hold-Cleaner.ps1 | Shared PnP; rejects technical roots; consistent localized destructive phrases; cancelling first stage stops Both; safe registration-result extraction | First-to-second-stage move, second-stage deletion, Preservation Hold behavior, limits and simulation preserved |
| M365-Legacy-UserId-Cleaner.ps1 | Shared PnP; REPAIR/REPARAR checks the displayed language; pending recovery ledger blocks a new run; safe registration-result extraction | SID matching, ambiguity guards and removal logic preserved; original fast-path admin-baseline risk remains important |
| OneDrive-Path-Analyzer.ps1 | Shared UI; explicit PowerShell 5.1 minimum; debug-only stack trace | Local-only; thresholds, path ownership, nested-root handling, recursion, empty-report behavior and CSV commands preserved; no PnP requirement at runtime |

## Bugs corrected

- Failed PnP imports no longer return success; missing/outdated modules offer installation from the actual operation flow.
- Compatibility is checked against module PowerShell requirements; installation is re-enumerated before import.
- Normal menu adapters accept both dictionaries and objects, including missing optional descriptions.
- Language menu numbering is identical and Enter honors the supplied default.
- English yes/no prompts no longer display Spanish S/N hints.
- Three reporters retain the previous Client ID if app validation fails. Temporary test connections close on both success and failure.
- Registration extraction accepts scalar/array/dictionary results and never treats a directory object `Id` as the Application (client) ID.
- Reporter system detection no longer excludes customer sites merely because they are locked or contain substrings such as `search-project` or `appcatalogue`.
- OneDrive technical host roots are not personal OneDrive targets; admin and malformed endpoints are rejected by the common classifier.
- Reporters can retain an explicit SharePoint admin URL associated with the configured tenant, including renamed SharePoint host prefixes.
- Legacy and Recycle typed confirmations match their displayed language. Recycle Both no longer continues after first-stage cancellation/failure returning null.
- A new Legacy repair run cannot overwrite an existing recovery ledger.
- When `M365_PNP_TOOLKIT_HOME` is set, the three reporters and Permissions now save tool-specific settings filenames. Legacy shared `settings.json` is read as a fallback, not overwritten by the new per-tool saves. Existing report paths remain unchanged.

## Dependencies, settings and hardcodes

- M365 tools retain PowerShell 7.4+ and PnP.PowerShell 3.2+ requirements. Bootstrap also checks Core edition. A fresh install uses the tested minimum; an explicit update action remains available and requires consent.
- Path Analyzer retains Windows PowerShell 5.1 / PowerShell 7 compatibility and local Windows registry/filesystem discovery. Its embedded PnP helper is not invoked.
- No UI packages, Pester dependency, secrets, certificates or tokens were introduced.
- Existing tenant/Client ID settings remain reusable per tool. A cross-tool tenant-aware authentication profile is **not implemented yet**.
- Fixed protocol/SharePoint service suffixes and local OS environment/registry locations are legitimate constants. Sample tenant names in help and synthetic GUIDs/URLs in offline tests are examples, not configured customer values. No actual customer tenant, identity or Client ID was found by the targeted hardcode scan.
- Cloud-specific suffix classification alone does not establish sovereign-cloud authentication support; that remains unverified.
- Report names, delimiters, encoding choices and CSV commands were preserved. Path Analyzer intentionally retains culture-specific delimiter and language-dependent columns. No new operational log format was imposed.

The registration adapter for the reporters explicitly requests delegated SharePoint AllSites.FullControl and Graph User.Read, with a warning before registration. PnP otherwise has broader documented default delegated permissions. See [PnP registration documentation](https://pnp.github.io/powershell/cmdlets/Register-PnPEntraIDAppForInteractiveLogin.html). Interactive connection accepts Tenant in the current implementation; it was not removed based on incomplete syntax examples. See [PnP ConnectOnline source](https://github.com/pnp/powershell/blob/dev/src/Commands/Base/ConnectOnline.cs).

## Validation performed

- Parser validation after each propagation; final eight-script parser pass under PowerShell 7.6.6.
- Separate parser pass for Path Analyzer using actual Windows PowerShell 5.1.26100.8875.
- Isolated single-file startup/exit for all eight scripts, plus actual Windows PowerShell 5.1 isolated startup/exit for Path Analyzer.
- AST scan rejects dot-sourcing, `$PSScriptRoot`, repository-directory references, local modules/resources, dynamic expression loading and CSV/CLIXML inputs in every isolated copy.
- Offline tests load function ASTs only, never script entrypoints.
- Common and reporter translation-key parity, format rendering, language selection/default, yes/no default, typed confirmations/cancellation, dictionary/object menus and original return contracts.
- Valid/nonzero/malformed GUIDs; null/object/array/string registration results; misleading object IDs.
- SharePoint/customer paths, admin host, OneDrive personal path, technical root, spoofed suffix, non-HTTPS and userinfo URLs.
- Mocked missing/installed module, installation accepted/declined, import failure and installation failure; no actual modules installed.
- Three reporter app flows: existing valid/invalid app, new app, cancellation before registration, old ID preservation and connection cleanup.
- Exact embedded-core equivalence across eight scripts; no duplicate functions or removed original functions.
- Baseline comparison: 20 selected business functions unchanged and all original Export-Csv command expressions unchanged. This is targeted regression evidence, not proof of every execution path.

Run the offline checks in PowerShell 7.4+:

```powershell
$scripts = (Get-ChildItem -File *.ps1).Name
./tests/Test-Toolkit.ps1 -ScriptName $scripts
./tests/Test-SuiteIntegrity.ps1
# Optional: compare with your pre-change copy of all eight scripts:
./tests/Test-SuiteIntegrity.ps1 -BaselineDirectory '<baseline folder>'
```

For maintenance, edit Common/M365Toolkit.Core.ps1 or Common/M365Toolkit.Reporter.ps1, then synchronize one tool and validate it before advancing:

```powershell
./tools/Sync-ToolkitCore.ps1 -ScriptName M365-Inactive-Site-Finder.ps1
./tests/Test-Toolkit.ps1 -ScriptName M365-Inactive-Site-Finder.ps1
```

Reporter catalog sources are in Common/Locales/. The synchronization command embeds them automatically. End users need only the resulting `.ps1`.

## Remaining work — do not treat as completed

1. Replace the five mature tools' legacy word/phrase translation system with complete keyed catalogs. Mixed-language dynamic messages and data translation outside the new field renderer can still occur. Complete translations and verify every screen in both languages.
2. Finish canonical custom tables, single-item selectors, multi-select fallbacks, context summaries and final screens. The shared core does not mean every legacy screen is visually identical yet.
3. Unify the remaining four M365 authentication/app-registration adapters, validate candidates before every save, bind reused connections to tenant/client, clarify requested permissions, and finish renamed-tenant support there.
4. **Important safety issue:** preserve pre-existing Site Collection Admin grants in Legacy/Permissions/Recycle. Legacy fast processing currently does not establish the baseline. Recovery state also needs tenant/client binding and stronger corrupted-state handling. Until corrected and tested, use a disposable tenant and inspect admin cleanup manually; do not consider production destructive execution certified.
5. Finish guards for all tenant/client/context changes while a temporary grant is pending; the new target-change and new-run guards are not exhaustive.
6. Implement a validated, nonsecret, tenant-keyed shared app profile if desired; existing per-tool settings are retained for now.
7. Complete exhaustive static command/variable/path checks and mocked operation-level regression cases, including all retry, preview, cleanup and interrupted-session paths. Decide carefully where SupportsShouldProcess/-WhatIf can be added without altering preview semantics.
8. Perform real sign-in/consent/module-version, SharePoint/OneDrive, terminal-width and recovery tests in an isolated tenant. No Microsoft 365 integration test has been performed in this work.

No existing preview/simulation was removed. In particular, Legacy's existing DryRun does **not** simulate administrative grants; this intentional existing distinction remains visible in its help and must not be mistaken for a fully read-only mode.

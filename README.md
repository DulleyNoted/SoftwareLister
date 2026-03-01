# SoftwareLister

A WPF-based Windows tool for inventorying installed software, drivers, and services.
Built for PowerShell 5.1. No external dependencies or installation required.

---

## Required Files

The following files must all be in the **same folder**:

| File | Purpose |
|---|---|
| `SoftwareLister.ps1` | Main script — launch this one |
| `config.json` | Configuration for data sources, columns, and custom app names |
| `ReportEngine.ps1` | Export/report functions (dot-sourced automatically by the main script) |

Everything else in the folder (`_*.ps1`, `v1\`, `files.zip`, etc.) is development tooling or archived versions and is not needed to run the tool.

---

## Requirements

- **Windows 10 or 11**
- **PowerShell 5.1** (ships with Windows — no install needed)
- **Run as Administrator** — required for the Drivers and Services tabs (driver install, service start/stop, ACL analysis). The Software tab works without elevation, but some registry paths may be skipped.

---

## How to Run

Open a PowerShell terminal in the folder and run:

```powershell
.\SoftwareLister.ps1
```

To enable verbose debug output in the terminal while the app is open:

```powershell
.\SoftwareLister.ps1 -DebugMode
```

Debug mode prints a timestamped log of every driver property read, service registry check, ACL query, and WMI call to the terminal window.

> **Tip:** Right-click `SoftwareLister.ps1` in File Explorer and choose **Run with PowerShell** to launch without a terminal. Use the terminal method if you want the `-DebugMode` output or to see any error messages.

---

## The Three Tabs

### Software

Lists all installed applications detected from:

- **Registry** — traditional Win32 installers (Add/Remove Programs)
- **AppX** — Microsoft Store and UWP packages
- **Winget** — Windows Package Manager (disabled by default; enable in `config.json`)

**Controls:**
- Search box — filters by name, publisher, version, or install path
- Source / Architecture dropdowns — narrow results
- **Refresh** — re-scans all enabled data sources
- **Export** — saves the current filtered list as CSV, JSON, or TXT

**Comparison:** Load a previously exported file and compare it against the current scan to see what was added or removed.

---

### Drivers

Lists all PnP devices and their associated driver details (version, provider, date, INF path).

Click **Refresh** to load. Data is read incrementally — the status bar at the bottom shows progress.

**Columns:** Device name, class, status, present/absent, driver version, provider, date, INF path.

**Filters:** Search box, class dropdown, status dropdown, present/absent toggle.

**Actions (require Administrator):**
- **Backup Drivers** — exports all third-party driver packages to a folder you choose (uses `Export-WindowsDriver`). This can take several minutes; the UI stays responsive and a notification appears when complete.
- **Import Driver (.inf)** — installs a single INF file via `pnputil /add-driver /install`.
- **Import Driver Folder** — recursively finds all `.inf` files in a folder and installs them all.

---

### Services

Lists all Windows services with optional security analysis.

Click **Refresh** to load. Three checkboxes control what gets analyzed (uncheck any to speed up the load):

| Checkbox | What it does | Speed impact |
|---|---|---|
| Startup type (WMI) | Resolves exact startup type (e.g. "Automatic (Delayed)") via a single WMI bulk query | Adds ~1-3 s at start |
| Service control permissions | Runs `sc.exe sdshow` per service to check if non-admin accounts can change the service config | Slow if many services |
| Executable ACL analysis | Runs `Get-Acl` on each service's executable to check for non-admin write access | Moderate |

**Columns:** Name, display name, state, startup type, log on as, executable path, service control risk, exe write risk.

**Risk indicators:**
- `! Weak` (Service Control column) — a non-administrator account has `SERVICE_CHANGE_CONFIG`, `WRITE_DAC`, or `WRITE_OWNER` rights, which could allow privilege escalation.
- `! Risky` (Exe Write Risk column) — a non-administrator account has write or take-ownership rights on the service executable.
- `OK` / `N/A` — no issue detected, or the check was skipped.

**Filters:** Search box, state, startup type, and risk dropdowns.

**Actions (require Administrator):** Select a service row, then use the **Start / Stop / Restart** buttons. A confirmation prompt appears before any action is taken.

---

## Configuration (config.json)

| Setting | Description |
|---|---|
| `dataSources.registry.enabled` | Include traditional Win32 programs |
| `dataSources.appx.enabled` | Include Microsoft Store / UWP apps |
| `dataSources.winget.enabled` | Include winget packages (requires winget) |
| `properties.<name>.enabled` | Show or hide individual columns in the Software tab |
| `display.showSystemComponents` | Include hidden system components in the Software list |
| `display.showFrameworks` | Include runtime/framework AppX packages |
| `display.excludePatterns` | Array of name substrings to exclude from Software results |
| `export.defaultFormat` | Default export format: `csv`, `json`, or `txt` |
| `customNames` | Human-readable name overrides for AppX packages and registry entries |

Changes made through the **Settings** tab inside the app are written back to `config.json` automatically.

---

## Execution Policy

If PowerShell blocks the script with an execution policy error, run this once in an elevated terminal:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

Or bypass it for a single launch:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\SoftwareLister.ps1
```

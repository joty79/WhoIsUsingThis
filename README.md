<p align="center">
  <img src="https://img.shields.io/badge/Platform-Windows_10%2F11-0078D6?style=for-the-badge&logo=windows&logoColor=white" alt="Platform">
  <img src="https://img.shields.io/badge/PowerShell-7%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white" alt="PowerShell">
  <img src="https://img.shields.io/badge/Sysinternals-Handle.exe-FF6F00?style=for-the-badge&logo=microsoft&logoColor=white" alt="Handle.exe">
</p>

<h1 align="center">🔍 WhoIsUsingThis</h1>

<p align="center">
  <b>Right-click any file or folder to instantly see what's locking it — and kill the process or fix permissions on the spot</b><br>
  <sub>Handle scan · DLL module detection · ACL analysis · Hard link diagnostics — all from one context menu click</sub>
</p>

---

## ✨ What's Inside

| # | Tool | Description |
|:-:|------|-------------|
| 🔍 | **[Lock Scanner](#-lock-scanner)** | 4-method deep scan to find every process or permission blocking a file/folder |
| 👻 | **[Silent Launcher](#-silent-launcher)** | Zero-flash VBS wrapper — opens in Windows Terminal with full emoji support |
| 📦 | **[Installer](#-installation)** | One-command setup with context menu registration and auto-updates from GitHub |

---

## � Lock Scanner

> Right-click any file or folder, instantly see every process, module, permission, and system lock blocking it — then choose to kill, fix, or skip, all from one interactive scan.

### The Problem

- You try to delete/rename a file and get **"The action can't be completed because the file is open in another program"** — but *which* program?
- `handle.exe` alone doesn't catch **DLL locks** (loaded modules) or **ACL restrictions** (TrustedInstaller ownership)
- Windows system folders under `WinSxS` have invisible locks from **CBS transactions**, **hard links**, and **reparse points**
- Critical system paths need **extra protection** from accidental `takeown` operations that could brick the OS

### The Solution

A 4-method progressive scan that runs from a single context-menu click:

```
┌─────────────────────────────────────────────────────────────┐
│           WHOISUSINGTHIS — 4-METHOD SCAN                    │
│                                                             │
│  1️⃣  Handle.exe Scan (Standard file locks)                  │
│     └─ Sysinternals handle.exe — finds open file handles    │
│     └─ Groups by PID, separates apps vs terminals           │
│     └─ Kill all / Select individually / Skip                │
│                                                             │
│  2️⃣  Module / Deep Scan                                     │
│     ├─ FILE:   tasklist /m — finds loaded DLL modules       │
│     └─ FOLDER: Process.Modules scan — deep memory scan      │
│                for any process using files inside folder     │
│                                                             │
│  3️⃣  ACL & Ownership Analysis                               │
│     └─ Owner check (TrustedInstaller/SYSTEM/Admin)          │
│     └─ Delete permission check for current user             │
│     └─ Subfolder ACL scan (streaming, Owner-only, capped)   │
│     └─ CBS InFlight / WinSxS Temp detection                 │
│                                                             │
│  4️⃣  Direct Access Test & Deep Diagnostics                  │
│     └─ Real file open test (ReadWrite + exclusive lock)     │
│     └─ Hard Links detection (fsutil hardlink list)          │
│     └─ CBS pending.xml / reboot.xml check                   │
│     └─ Reparse Point / Junction / Symlink detection         │
│                                                             │
│  ╔══════════════════════════════════════════════════════╗    │
│  ║  🧪 Diagnosis + 💡 Suggested Solutions               ║    │
│  ║  🛡️ Auto-takeown offer (with CRITICAL PATH guard)    ║    │
│  ╚══════════════════════════════════════════════════════╝    │
└─────────────────────────────────────────────────────────────┘
```

Every method runs automatically in sequence. At the end, the tool provides a **tailored diagnosis** with specific solutions, and offers to fix permissions automatically — with extra safety guards for critical system paths.

### Process Kill Options

At each scan stage, when locking processes are found:

```
[A] Kill ALL        — terminates all non-terminal processes
[S] Skip            — move on to next scan stage
[C] Choose one-by-one — decide per process (Y/N)
```

Terminal processes (`pwsh`, `cmd`, `wt`, `conhost`) are **auto-skipped** when using `[A]` to prevent killing your own session. Use `[C]` to explicitly target them if needed.

### Critical System Path Protection

When targeting paths like `C:\Windows\WinSxS`, `C:\Windows\System32`, or other critical directories, the auto-fix prompt requires typing **`CONFIRM`** (case-sensitive) instead of just `Y`:

```
🚫🚫🚫 CRITICAL SYSTEM PATH DETECTED 🚫🚫�
Αυτός ο φάκελος είναι κρίσιμος για τη λειτουργία των Windows!

Γράψε CONFIRM (κεφαλαία) για συνέχεια  |  Enter = Ακύρωση
```

### Usage

**From context menu** — *Right-click any file or folder → System Tools → Who is using this 🔎?*

**From terminal:**

```powershell
# Scan a specific file
pwsh -NoProfile -ExecutionPolicy Bypass -File .\WhoIsUsingThis.ps1 -targetPath "C:\path\to\locked_file.dll"

# Scan a folder (deep module scan + subfolder ACL scan)
pwsh -NoProfile -ExecutionPolicy Bypass -File .\WhoIsUsingThis.ps1 -targetPath "C:\Windows\WinSxS\Temp"
```

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-targetPath` | `string` | *(required)* | Full path to the file or folder to scan |

### Scan Methods Detail

| Method | Scope | What It Catches | Tool Used |
|--------|-------|-----------------|-----------|
| **1. Handle Scan** | File & Folder | Open file handles — "file in use" errors | `handle.exe` (Sysinternals) |
| **2a. Module Scan** | File only | Loaded DLLs — process loaded the file as a module | `tasklist /m` |
| **2b. Deep Scan** | Folder only | Any process using files inside the folder in memory | `Process.Modules` (.NET) |
| **3. ACL Analysis** | File & Folder | Ownership, permissions, TrustedInstaller locks, CBS InFlight | `Get-Acl`, SID comparison |
| **4. Direct Test** | File & Folder | Real access test, hard links, CBS pending, reparse points | `File.Open`, `fsutil`, registry |

### Diagnostic Codes

| Code | Meaning | Suggested Fix |
|------|---------|---------------|
| `ACCESS_DENIED` | Permissions or WRP block | `takeown` + `icacls` (auto-offered) |
| `SHARING_VIOLATION` | Another process has the file open | Restart Explorer or close indexing apps |
| `LOCK_VIOLATION` | Kernel-level file lock | Reboot required |
| `HARD_LINKS` | WinSxS component store links | `DISM /StartComponentCleanup` |
| `CBS_PENDING` | Interrupted Windows Update | Reboot first, then `DISM /RestoreHealth` |
| `REPARSE_POINT` | Symlink or junction | Investigate target manually |

---

## 👻 Silent Launcher

> A VBS wrapper that launches the scan tool in Windows Terminal with full emoji support — zero flash, proper elevation chain.

### The Problem

- Direct PowerShell launch from registry causes a **visible blue console flash**
- Windows Terminal supports **emoji icons** (✅ ❌ 🔒) but classic console doesn't
- The launch chain needs to: detect WT → elevate to Admin → open the right console

### The Solution

`WhoIsUsingThis.vbs` starts the chain hidden, then the PowerShell script auto-detects the best console:

```
┌──────────────────────────────────────────────────────────┐
│  Registry Click                                          │
│    wscript.exe "WhoIsUsingThis.vbs" "%1"                 │
│    │                                                     │
│    └─→ powershell.exe -WindowStyle Hidden (flag set)     │
│          │                                               │
│          ├─ WT available? → Re-launch in Windows Terminal │
│          │   └─ Full emoji icon set 🎯✅❌🔒             │
│          │                                               │
│          ├─ Safe Mode? → PS5 with Nerd Font icons         │
│          │   └─ Fallback icon set (CaskaydiaCove NFM)    │
│          │                                               │
│          └─ Elevate to Admin if needed                    │
└──────────────────────────────────────────────────────────┘
```

All paths resolved from `WScript.ScriptFullName` — no hardcoded paths.

### Dual Icon Set

| Environment | Icon Style | Example |
|-------------|-----------|---------|
| **Windows Terminal** | Native Unicode emoji | ✅ ❌ 🔒 🎯 💀 ⚠️ |
| **Classic Console / Safe Mode** | Nerd Font glyphs (CaskaydiaCove NFM) | Fallback box-drawing icons |

---

## � Installation

### Quick Setup

```powershell
# Install (registers context menu + bundles handle.exe)
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -Action Install

# Update from GitHub
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -Action Update

# Uninstall (removes registry entries + installed files)
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -Action Uninstall -Force
```

### Requirements

| Requirement | Details |
|-------------|---------|
| **OS** | Windows 10 / 11 |
| **PowerShell** | pwsh 7+ preferred; falls back to Windows PowerShell 5.1 in Safe Mode |
| **handle.exe** | Bundled in `assets\bin\` — no separate download needed |
| **Privileges** | Admin — auto-requested via UAC |
| **Terminal** | Windows Terminal recommended (emoji support); works in classic console too |

### What the Installer Does

- Copies runtime files to `%LOCALAPPDATA%\WhoIsUsingThisContext\`
- Bundles `assets\bin\handle.exe` and `assets\icons\WhoIsUsingThis.ico`
- Registers context menu entries under the shared **System Tools** submenu:
  - `*\shell\SystemTools\shell\WhoIsUsingThis` — files
  - `Directory\shell\SystemTools\shell\WhoIsUsingThis` — folders
  - `Directory\Background\shell\...` + `DesktopBackground\Shell\...` — backgrounds
- Cleans up legacy `WhoIsUsingThis` and `CheckLocks` keys from previous versions
- Adds uninstall entry to Programs & Features

---

## � Project Structure

```
WhoIsUsingThis/
├── WhoIsUsingThis.ps1             # Main scanner (4-method lock detection)
├── WhoIsUsingThis.vbs             # Zero-flash VBS launcher
├── Install.ps1                    # Installer/updater/uninstaller
├── WhoIsUsingThis.reg             # Static registry sample (manual import)
├── WhoIsUsingThis - Remove.reg    # Registry cleanup (manual removal)
├── assets/
│   ├── bin/
│   │   └── handle.exe             # Sysinternals Handle (bundled)
│   └── icons/
│       └── WhoIsUsingThis.ico     # Context menu icon
├── .gitignore                     # Excludes logs, state
├── PROJECT_RULES.md               # Decision log and project guardrails
├── INSTALLER_PLAN.md              # Original installer design notes
└── README.md                      # You are here
```

---

## 🧠 Technical Notes

<details>
<summary><b>Why 4 separate scan methods instead of just handle.exe?</b></summary>

`handle.exe` only detects **open file handles** — the classic "file in use" scenario. But files can also be blocked by **loaded DLL modules** (invisible to handle.exe), **ACL restrictions** (TrustedInstaller ownership), **hard links** in the WinSxS component store, and **CBS transactions** from interrupted Windows Updates. Each method catches a different category of lock that the others miss. Running all four gives you the complete picture.

</details>

<details>
<summary><b>Why SID comparison instead of NTAccount for ACL scanning?</b></summary>

The subfolder ACL scan originally used `Get-Acl` which resolves owner names via `NTAccount`. This caused **`IdentityNotMappedException`** false positives on orphaned SIDs from deleted users or migrated machines. Comparing raw SID strings (`S-1-5-18` for SYSTEM, `S-1-5-80-*` for NT SERVICE) is both faster and immune to name resolution failures.

</details>

<details>
<summary><b>How does the ACL scan stay fast on large directories?</b></summary>

Instead of `Get-ChildItem | Get-Acl` (which creates PowerShell objects for every item), the scan uses **streaming enumeration** via `[System.IO.Directory]::EnumerateFileSystemEntries()` and requests **Owner-only** access control sections. This skips DACL/SACL/Group parsing entirely. The scan is also **capped at 500 items** — enough for diagnostic purposes without multi-minute waits on directories with thousands of files.

</details>

<details>
<summary><b>Why does the script re-launch itself in Windows Terminal?</b></summary>

Windows Terminal supports **Unicode emoji** natively, which the script uses extensively for visual feedback (✅ ❌ � 🎯 💀). The classic `conhost.exe` console has limited Unicode support and can't render multi-codepoint emoji. The script detects `$env:WT_SESSION` — if it's not set and `wt.exe` exists, it re-launches itself inside WT with the `--title "Handle Scan"` flag for a clean tab label.

</details>

<details>
<summary><b>What happens in Safe Mode?</b></summary>

In Safe Mode, Windows Terminal may not be available due to MSIX package restrictions. The script detects Safe Mode via `HKLM:\SYSTEM\CurrentControlSet\Control\SafeBoot\Option` and falls back to classic console with **Nerd Font glyphs** from CaskaydiaCove NFM instead of emoji. It also uses `powershell.exe` (PS5) if `pwsh.exe` (PS7) isn't available.

</details>

---

<p align="center">
  <sub>Built for Windows 10/11 · 4-method deep lock detection · Critical path protection · Zero-flash context menu integration</sub>
</p>

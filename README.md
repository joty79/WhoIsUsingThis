# WhoIsUsingThis

Tool για Windows context menu που εντοπίζει ποιο process ή ποιο permission layer εμποδίζει πρόσβαση/διαγραφή σε file ή folder.

## 🔵 Τι προσφέρει

- 🔸 Explorer context menu entry: `Who is using this 🔎?` σε files και folders.
- 🔸 Multi-layer lock detection:
1. `Handle.exe` scan για standard file handles.
2. `tasklist /m` module scan για file locks μέσω loaded DLL/modules.
3. Deep process scan για folder-related opens.
- 🔸 ACL/ownership diagnostics για `TrustedInstaller`, `SYSTEM`, deny ACLs, pending Windows Update states.
- 🔸 Interactive επιλογές για safe process terminate (με εμφανή exclusion λογική για terminal processes).
- 🔸 Προαιρετικό auto-fix flow με `takeown` + `icacls` όταν το πρόβλημα είναι permissions/ownership.

## 🔵 Πώς δουλεύει

1. 🔸 Ο χρήστης επιλέγει `Who is using this 🔎?` από context menu.
2. 🔸 Το registry command καλεί `WhoIsUsingThis.vbs`.
3. 🔸 Το `WhoIsUsingThis.vbs` εκκινεί το `WhoIsUsingThis.ps1` hidden και περνά το target path.
4. 🔸 Το script κάνει resolve το `handle.exe` με σειρά προτεραιότητας:
1. `assets\bin\handle.exe`
2. `handle.exe` δίπλα στο script
3. `handle.exe` από `PATH`
5. 🔸 Εκτελούνται scans + diagnostics και εμφανίζεται interactive remediation flow.

## 🔵 Installer

Το project περιλαμβάνει profile-based installer (`Install.ps1`) με actions:

- `Install`
- `Update`
- `Uninstall`
- `OpenInstallDirectory`
- `OpenInstallLogs`

### 🔸 Γρήγορη χρήση

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1
```

### 🔸 CLI examples

```powershell
# Install (GitHub-first)
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -Action Install -GitHubRef master -Force

# Update
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -Action Update -GitHubRef master -Force

# Uninstall
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -Action Uninstall -Force
```

## 🔵 Απαιτήσεις

- Windows 10/11
- PowerShell 7+
- `wscript.exe`
- `tasklist.exe`, `takeown.exe`, `icacls.exe`
- `handle.exe` (bundled στο `assets\bin\handle.exe`)
- Για private GitHub package install: authenticated `gh` login (`gh auth status`)

## 🔵 Δομή Repo

- `WhoIsUsingThis.ps1`: main runtime logic/scans/diagnostics.
- `WhoIsUsingThis.vbs`: hidden launcher wrapper.
- `WhoIsUsingThis.reg`: standalone registry entries (manual import path).
- `Install.ps1`: installer/update/uninstall flow.
- `assets\bin\handle.exe`: bundled dependency.
- `assets\icons\WhoIsUsingThis.ico`: bundled icon asset.
- `PROJECT_RULES.md`: project memory/guardrails/decisions.

## 🔵 Troubleshooting

- ⚠️ Αν το context menu δεν φαίνεται άμεσα μετά από install/update, κάνε Explorer restart.
- ⚠️ Αν βλέπεις παλιά verb behavior, τρέξε `Install.ps1 -Action Update` για cleanup + rewrite.
- 💡 Installer logs: `%LOCALAPPDATA%\WhoIsUsingThisContext\logs\installer.log`

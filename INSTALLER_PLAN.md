# WhoIsUsingThis Installer Rollout (Action-Ready)

## Summary
- Implement installer for `WhoIsUsingThis` with Robocopy-style workflow quality.
- Introduce reusable `Template + Profile` installer system in a dedicated core repo.

## Public Interfaces / Contracts
1. `Install.ps1` actions:
- `Install`
- `Update`
- `Uninstall`
- `OpenInstallDirectory`
- `OpenInstallLogs`
2. Profile contract:
- tool identity
- install path suffix
- GitHub source settings
- required/deploy files
- registry cleanup/write/verify specs
- uninstall metadata
3. Defaults:
- `InstallPath`: `%LOCALAPPDATA%\WhoIsUsingThisContext`
- `GitHubRef`: `master`
- package source policy: GitHub-first with safe fallback

## Execution Steps
1. Save this plan as project markdown.
2. Create dedicated `InstallerCore` repo.
3. Add `WhoIsUsingThis` profile.
4. Generate `WhoIsUsingThis\Install.ps1` from template + profile.
5. Bundle required assets:
- `assets\bin\handle.exe`
- `assets\icons\WhoIsUsingThis.ico`
6. Update runtime dependency resolution in `WhoIsUsingThis.ps1` for bundled `handle.exe`.
7. Ensure installer registry flow:
- cleanup `WhoIsUsingThis` and `CheckLocks` keys in HKCU/HKCR
- create only `WhoIsUsingThis` active keys
8. Append decision memory to `PROJECT_RULES.md`.

## Test Cases / Scenarios
1. Parser validation for modified `.ps1` files.
2. Fresh install flow from GitHub `master`.
3. Re-run update for idempotent behavior.
4. Uninstall flow removes menu keys and uninstall entry.
5. Registry read-back verify for command values.
6. Context-menu runtime smoke on file/folder.
7. Missing `handle.exe` path behavior.

## Assumptions / Defaults
1. `joty79/WhoIsUsingThis` keeps branch `master`.
2. Installer dependencies are bundled in repository assets.
3. Explorer restart policy remains installer-controlled with opt-out switch.

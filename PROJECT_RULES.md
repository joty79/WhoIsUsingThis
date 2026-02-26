# PROJECT_RULES - WhoIsUsingThis

## Scope
- Repo: `d:\Users\joty79\scripts\WhoIsUsingThis`
- Focus: context-menu integration and launch chain (`.reg` -> `.vbs` -> `.ps1`).

## Guardrails
- Keep `WhoIsUsingThis.ps1` as the main logic entrypoint.
- Keep `WhoIsUsingThis.vbs` as hidden launcher used by registry verbs.
- In `.reg`, prefer `HKCU\Software\Classes\*\shell` and `HKCU\Software\Classes\Directory\shell`.
- Use targeted legacy cleanup only for known old verb keys; avoid broad key deletion.

## Decision Log

### Entry - 2026-02-26
- Date: 2026-02-26
- Problem: Repo moved from `...\scripts\Utilities\` to `...\scripts\WhoIsUsingThis\`, leaving stale hardcoded paths.
- Root cause: Absolute paths in `WhoIsUsingThis.vbs` and `WhoIsUsingThis.reg`.
- Guardrail/rule: When repo path changes, update launcher (`.vbs`) and registry command paths together.
- Files affected: `WhoIsUsingThis.vbs`, `WhoIsUsingThis.reg`.
- Validation/tests run: Static path scan with `rg` after edits to verify old `...\scripts\Utilities\WhoIsUsingThis.*` references were removed.

### Entry - 2026-02-26 (Installer rollout)
- Date: 2026-02-26
- Problem: Tool lacked a reusable installer flow and required manual registry/path setup.
- Root cause: No dedicated installer script and no bundled dependency strategy.
- Guardrail/rule: Keep installer-generated deployment under `%LOCALAPPDATA%\WhoIsUsingThisContext`, bundle required runtime dependency (`assets\bin\handle.exe`), and clean legacy registry verbs (`WhoIsUsingThis` and `CheckLocks`) before writing active keys.
- Files affected: `Install.ps1`, `WhoIsUsingThis.ps1`, `WhoIsUsingThis.vbs`, `INSTALLER_PLAN.md`, `assets\bin\handle.exe`, `assets\icons\WhoIsUsingThis.ico`.
- Validation/tests run: `Parser::ParseFile` on edited/generated PowerShell scripts; static read-back checks for embedded registry cleanup/write specs in `Install.ps1`.

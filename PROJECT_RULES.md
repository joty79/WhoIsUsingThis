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

### Entry - 2026-02-26 (GitHub private fetch + MUIVerb normalization)
- Date: 2026-02-26
- Problem: GitHub install path could fall back to local source on private repo, and registry verify could mismatch on menu text.
- Root cause: Anonymous codeload zip URL for private repo returns 404; profile used escaped unicode sequence and leading-space-sensitive `MUIVerb`.
- Guardrail/rule: Installer download flow must try authenticated `gh`-based fallback when `Invoke-WebRequest` fails; `MUIVerb` strings should use normalized text (`Who is using this 🔎?`) without leading space and without `\ud...` escape literals.
- Files affected: `Install.ps1`, `WhoIsUsingThis.reg`, `PROJECT_RULES.md`.
- Validation/tests run: `Parser::ParseFile` on `Install.ps1`; static scan for `gh.exe repo archive` fallback and normalized `MUIVerb` values.

### Entry - 2026-02-26 (GitHub CLI compatibility fallback)
- Date: 2026-02-26
- Problem: Installer printed `unknown flag: --ref` during GitHub fallback and returned to local source.
- Root cause: Installed `gh` variant does not support `gh repo archive --ref ... --format ... --output ...`.
- Guardrail/rule: Use authenticated GitHub API zipball fallback via `gh auth token` + `Invoke-WebRequest https://api.github.com/repos/<repo>/zipball/<ref>` instead of `gh repo archive` flags.
- Files affected: `Install.ps1`, `PROJECT_RULES.md`.
- Validation/tests run: `Parser::ParseFile` on regenerated `Install.ps1`; static check for `gh auth token` + API zipball fallback path.

### Entry - 2026-02-26 (WSH command quoting fix)
- Date: 2026-02-26
- Problem: After install, context-menu click triggered `Windows Script Host failed` and icon presentation was inconsistent.
- Root cause: Installer profile wrote registry command with escaped literal sequences (`\"` and doubled `\\`) instead of plain quoted command text.
- Guardrail/rule: Registry command values must be stored as plain command strings (`wscript.exe "{InstallRoot}\WhoIsUsingThis.vbs" "%1"`), no literal backslash-escaped quote sequences; keep icon value aligned with stable shell icon source (`imageres.dll,-102`).
- Files affected: `Install.ps1`, `PROJECT_RULES.md`.
- Validation/tests run: `Parser::ParseFile` on regenerated `Install.ps1`; static read-back check of generated profile-embedded command/icon values.

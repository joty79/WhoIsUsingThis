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

### Entry - 2026-03-01 (Move verb under shared System Tools submenu)
- Date: 2026-03-01
- Problem: `WhoIsUsingThis` needed to appear under the shared `System Tools` submenu instead of as a top-level context-menu verb.
- Root cause: Registry definitions and installer profile still wrote standalone keys under `*\shell\WhoIsUsingThis` and `Directory\shell\WhoIsUsingThis`.
- Guardrail/rule: In this repo, install `WhoIsUsingThis` only as a child verb under `*\shell\SystemTools\shell\WhoIsUsingThis` and `Directory\shell\SystemTools\shell\WhoIsUsingThis`; cleanup must stay targeted to this repo's legacy standalone keys and owned child keys, never delete the shared `SystemTools` parent.
- Files affected: `WhoIsUsingThis.reg`, `Install.ps1`, `PROJECT_RULES.md`.
- Validation/tests run: Static review of registry paths in `.reg` and installer profile; `Parser::ParseFile` on `Install.ps1`.

### Entry - 2026-03-01 (GitHub ref autodetect)
- Date: 2026-03-01
- Problem: Installer required a hardcoded `github_ref`, which broke expected behavior when both `master` and `latest` existed.
- Root cause: GitHub install/update flow used a fixed default ref instead of resolving available remote branches.
- Guardrail/rule: For GitHub installs in this repo, if `-GitHubRef` is not explicitly provided, autodetect in this order: remote default branch, `master`, profile value, then `latest`; keep explicit `-GitHubRef` as an override.
- Files affected: `Install.ps1`, `PROJECT_RULES.md`.
- Validation/tests run: `Parser::ParseFile` on `Install.ps1`; static review of GitHub ref resolution flow.

### Entry - 2026-03-02 (Regenerated installer from InstallerCore)
- Date: 2026-03-02
- Problem: Local installer logic had drifted from the shared template and needed the restored branch picker, GitHub ref autodetect, clean Explorer restart flow, and current `System Tools` submenu profile values in one consistent generated file.
- Root cause: Repo installer had accumulated direct edits while `InstallerCore` was missing some of the intended template behavior.
- Guardrail/rule: Prefer regenerating `Install.ps1` from `InstallerCore` after template/profile fixes instead of hand-merging installer logic in this repo.
- Files affected: `Install.ps1`, `PROJECT_RULES.md`.
- Validation/tests run: `New-ToolInstaller.ps1` generation from `InstallerCore`; `Parser::ParseFile` on regenerated `Install.ps1`.

### Entry - 2026-03-02 (SystemTools file/folder host parents must exist)
- Date: 2026-03-02
- Problem: `WhoIsUsingThis` registration broke the shared `System Tools` submenu.
- Root cause: Live registry had only `Directory\Background\shell\SystemTools`; file (`*\shell`) and folder (`Directory\shell`) `SystemTools` parents were missing, so the child verb had no valid host submenu on those branches.
- Guardrail/rule: For this repo's `SystemTools` integration, ensure `*\shell\SystemTools` and `Directory\shell\SystemTools` exist as proper cascade parents (`MUIVerb`, `SubCommands`, `Icon`) before writing nested `shell\WhoIsUsingThis` child keys.
- Files affected: `WhoIsUsingThis.reg`, `Install.ps1`, `PROJECT_RULES.md`.
- Validation/tests run: Live `reg query` of installed `SystemTools` branches; regenerated installer/profile expected to emit file/folder parent keys plus child verb keys.

### Entry - 2026-03-02 (WhoIsUsingThis is child-only under System Tools)
- Date: 2026-03-02
- Problem: The quick local fix for missing `SystemTools` parents made `WhoIsUsingThis` act like a submenu host, which does not scale once multiple child tools participate.
- Root cause: Parent-key ownership was patched into this child repo instead of being fixed at the host repo (`SystemTools`).
- Guardrail/rule: `WhoIsUsingThis` must remain child-only under the shared submenu. It may register only `*\shell\SystemTools\shell\WhoIsUsingThis` and `Directory\shell\SystemTools\shell\WhoIsUsingThis`, plus targeted cleanup of its own legacy keys; parent `SystemTools` keys are owned by `SystemTools` and must not be emitted here.
- Files affected: `WhoIsUsingThis.reg`, `Install.ps1`, `PROJECT_RULES.md`.
- Validation/tests run: Regenerated installer from updated `InstallerCore` profile; static review of manual `.reg`.

### Entry - 2026-03-02 (Regenerated after empty-string registry helper hardening)
- Date: 2026-03-02
- Problem: Generated installer inherited a fragile template helper for empty-string registry writes.
- Root cause: `InstallerCore` template still used a literal `""` conversion pattern for empty `REG_SZ` values before the shared helper was hardened.
- Guardrail/rule: Keep `Install.ps1` generated from the current `InstallerCore` template after helper fixes; do not hand-maintain older generated copies once registry write semantics change in the template.
- Files affected: `Install.ps1`, `PROJECT_RULES.md`.
- Validation/tests run: Regenerated `Install.ps1` from `InstallerCore`; parser validation on generated installer; targeted scan confirmed the old literal `""` pattern is absent.

### Entry - 2026-03-02 (Add background support under shared System Tools submenu)
- Date: 2026-03-02
- Problem: `WhoIsUsingThis` worked on files/folders but not on folder background or desktop background under the shared `System Tools` submenu.
- Root cause: The repo only registered child verbs under `*\shell\SystemTools\shell\WhoIsUsingThis` and `Directory\shell\SystemTools\shell\WhoIsUsingThis`; no child keys existed for `Directory\Background\shell` or `DesktopBackground\Shell`.
- Guardrail/rule: Keep `WhoIsUsingThis` child-only, but mirror the verb on supported background branches too: `HKCU\Software\Classes\Directory\Background\shell\SystemTools\shell\WhoIsUsingThis` and `HKCU\Software\Classes\DesktopBackground\Shell\SystemTools\shell\WhoIsUsingThis`, using `%V` as the target path argument.
- Files affected: `WhoIsUsingThis.reg`, `Install.ps1`, `PROJECT_RULES.md`.
- Validation/tests run: PowerShell parser validation on `Install.ps1`; static review of background registry paths and command arguments.

### Entry - 2026-03-03 (InstallerCore profile drift removed background support on regenerate)
- Date: 2026-03-03
- Problem: Running local `Install.ps1` with `-PackageSource Local` still did not install the background verbs, while manual `WhoIsUsingThis.reg` did.
- Root cause: `Install.ps1` had been regenerated from `InstallerCore`, and the source profile there was missing the background branches. The local manual `.reg` fix existed only in this repo, not in the template/profile source of truth.
- Guardrail/rule: When local `.reg` behavior and generated installer behavior diverge, verify the matching `InstallerCore` profile before trusting a regenerate. For this repo, background support must exist in both the manual `.reg` and the `InstallerCore` profile before regeneration.
- Files affected: `Install.ps1`, `PROJECT_RULES.md`.
- Validation/tests run: Static comparison of generated installer content against `WhoIsUsingThis.reg`; profile audit in `InstallerCore`.

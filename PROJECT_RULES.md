# PROJECT_RULES - WhoIsUsingThis

## Scope
- Repo: `d:\Users\joty79\scripts\WhoIsUsingThis`
- Focus: context-menu integration and launch chain (`.reg` -> `.vbs` -> `.ps1`).

## Guardrails
- Keep `WhoIsUsingThis.ps1` as the main logic entrypoint.
- Keep `WhoIsUsingThis.vbs` as hidden launcher used by registry verbs.
- In `.reg`, prefer `HKCU\Software\Classes\*\shell` and `HKCU\Software\Classes\Directory\shell`.
- Use targeted legacy cleanup only for known old verb keys; avoid broad key deletion.
- For `UPDATEUI`, use the InstallerCore In-App Update UI Contract with the WT adapter. This app has no persistent main menu, so expose update status in the scan header and `U = Update App` on the no-target, path-error, and final scan screens. `U` must open an `Update App` actions submenu (`Run update now`, `Refresh update status`, `Back`), not run the updater directly.

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

### Entry - 2026-03-27 (Manual .reg must not depend on dev drive)
- Date: 2026-03-27
- Problem: Manual registry import could register a launcher path on `D:\...`, so the context-menu verb failed on machines/VMs that only had `C:`.
- Root cause: `WhoIsUsingThis.reg` still contained repo-specific absolute command paths instead of pointing at the installed runtime location.
- Guardrail/rule: Keep `WhoIsUsingThis.reg` drive-agnostic. Manual command values must target `%LOCALAPPDATA%\WhoIsUsingThisContext\WhoIsUsingThis.vbs` via `REG_EXPAND_SZ`, never a workstation-specific repo path.
- Files affected: `WhoIsUsingThis.reg`, `README.md`, `PROJECT_RULES.md`.
- Validation/tests run: Static read-back of `.reg` command values after edit; confirmed `D:\Users\joty79\...` launcher references are absent from the repo search.

### Entry - 2026-03-27 (VBS launcher must prefer pwsh for UTF-8 script)
- Date: 2026-03-27
- Problem: Context-menu click still failed even with correct install paths.
- Root cause: `WhoIsUsingThis.vbs` launched `powershell.exe` first, and `WhoIsUsingThis.ps1` is stored as UTF-8 without BOM and contains emoji/glyph literals, which Windows PowerShell 5.1 misparsed before the script could relaunch itself into `pwsh`/WT.
- Guardrail/rule: For this repo, the hidden VBS launcher must prefer `pwsh.exe` when available and only fall back to `powershell.exe`. Do not rely on a PS5 first-hop for UTF-8/no-BOM scripts that contain emoji or other non-ASCII literals.
- Files affected: `WhoIsUsingThis.vbs`, `PROJECT_RULES.md`.
- Validation/tests run: Local `Install.ps1 -Action Update -PackageSource Local`; direct `wscript.exe <installed vbs> <temp file>` smoke test confirmed the chain launched `WindowsTerminal` successfully; Explorer restart completed via installer.

### Entry - 2026-04-24 (InstallerCore UPDATEUI integration)
- Date: 2026-04-24
- Problem: The repo had an old generated installer and no app-side `Update App` UI, so updating `InstallerCore` alone would not give the scanner the canonical progress/recent-output/relaunch behavior.
- Root cause: `WhoIsUsingThis` predates the shared in-app update UI contract and is a context-menu scanner, not a persistent menu app.
- Guardrail/rule: Keep `Install.ps1` generated from `InstallerCore`, deploy/verify `app-metadata.json`, and implement `Update App` in `WhoIsUsingThis.ps1` with the WT adapter: header status, progress panel, recent installer output, relaunch, and old-host exit.
- Files affected: `WhoIsUsingThis.ps1`, `Install.ps1`, `app-metadata.json`, `CHANGELOG.md`, `.gitignore`, `README.md`, `PROJECT_RULES.md`, `InstallerCore\profiles\WhoIsUsingThis.json`.
- Validation/tests run: PowerShell parser validation for `WhoIsUsingThis.ps1` and `Install.ps1`; JSON validation for `app-metadata.json` and `InstallerCore\profiles\WhoIsUsingThis.json`; local-source installer update smoke completed with exit code `0`; installed file hash readback matched repo files; registry command readback passed for file, folder, folder background, and desktop background verbs.

### Entry - 2026-04-24 (UPDATEUI submenu and relaunch correction)
- Date: 2026-04-24
- Problem: The first `UPDATEUI` pass exposed a plain `U` prompt and update progress, but missed the canonical actions submenu and did not reliably relaunch with the original target context.
- Root cause: The scanner has no main menu, so the app-side adapter was simplified too far instead of keeping the WinAppManager `Update App` action model.
- Guardrail/rule: For this repo, `U = Update App` opens a real actions submenu. Successful relaunch must pass the same `targetPath` and use the target folder as `WorkingDirectory` so a folder/background scan resumes in the same context.
- Files affected: `WhoIsUsingThis.ps1`, `CHANGELOG.md`, `README.md`, `PROJECT_RULES.md`.
- Validation/tests run: PowerShell parser validation for `WhoIsUsingThis.ps1` and `Install.ps1`; local-source installer update smoke completed with exit code `0`; installed file hash readback matched repo files; registry command readback passed for file, folder, folder background, and desktop background verbs. Interactive `Run update now` relaunch still requires manual WT/UAC confirmation to observe end-to-end.

### Entry - 2026-04-25 (Runtime-only update menu expression and header colors)
- Date: 2026-04-25
- Problem: The `Update App` submenu printed a runtime error at `Relaunch target`, and the header still rendered too close to plain white instead of the canonical WinAppManager color treatment.
- Root cause: `Write-Host (if (...) { ... })` parses but PowerShell treats `if` as a command in that expression position at runtime; the header wrote padded strings without explicit foreground colors for the content spans.
- Guardrail/rule: Do not use inline `if` expressions inside `Write-Host` arguments in this codebase. Compute labels first, then write them. Header content must explicitly color title/version white, subtitle dark gray, and update status with status color.
- Files affected: `WhoIsUsingThis.ps1`, `PROJECT_RULES.md`.
- Validation/tests run: PowerShell parser validation for `WhoIsUsingThis.ps1` and `Install.ps1`; static scan confirmed the risky inline `Write-Host (if ...)` pattern is absent in repo and installed copy; local-source installer update smoke completed with exit code `0`; installed file hash readback matched repo files.

### Entry - 2026-04-25 (Header column alignment with colored spans)
- Date: 2026-04-25
- Problem: The colored header right border rendered one column left of the top/bottom border, making the UI look randomly broken.
- Root cause: The header was split across multiple `Write-Host -NoNewline` calls, and the content cell was padded to 76 characters while the 80-column border requires a 77-character cell after the left border+space.
- Guardrail/rule: For bordered PowerShell UI rows split into colored spans, compute the exact printable cell width and centralize padding/truncation in a helper. Do not hand-tune `PadRight()` values independently per span.
- Files affected: `WhoIsUsingThis.ps1`, `PROJECT_RULES.md`.
- Validation/tests run: PowerShell parser validation for `WhoIsUsingThis.ps1` and `Install.ps1`; header column length check confirmed top/content/update rows are all 80 printable characters; local-source installer update smoke completed with exit code `0`; installed file hash readback matched repo files.

### Entry - 2026-05-11 (Unicode-safe registry writes)
- Date: 2026-05-11
- Problem: GitHub install wrote `MUIVerb` as `Who is using this ???` and reported a registry mismatch for `HKCU\Software\Classes\*\shell\SystemTools\shell\WhoIsUsingThis`.
- Root cause: `reg.exe add /d` received the emoji-bearing string through the native command argument path and lost Unicode characters before the value reached the registry.
- Guardrail/rule: For installer-owned registry values, write and read back values through `Microsoft.Win32.RegistryKey` so Unicode labels and raw `REG_EXPAND_SZ` strings survive verification. Keep `reg.exe` for targeted cleanup only.
- Files affected: `Install.ps1`, `CHANGELOG.md`, `PROJECT_RULES.md`.
- Validation/tests run: PowerShell parser validation for `Install.ps1`; static review of registry write/readback helpers.

### Entry - 2026-05-11 (Commit-aware Update App status)
- Date: 2026-05-11
- Problem: The app-side `Update App` status compared only metadata versions, so same-version newer commits and stale cached `UpToDate` results could hide real updates.
- Root cause: The first WT adapter implementation predated the commit-aware `InstallerCore` update status contract and treated repo copies as archive overlays.
- Guardrail/rule: `WhoIsUsingThis` update status must track local/latest version, local/latest commit, source kind, and dirty state. Installed copies compare `state\install-meta.json` `github_commit` with the latest remote commit; git working copies update only with `git fetch` + fast-forward and refuse dirty workspaces; non-git portable copies may use `DownloadLatest -NoSelfRelaunch`; stale cached `UpToDate` must never be reused after a failed fresh remote check.
- Files affected: `WhoIsUsingThis.ps1`, `app-metadata.json`, `README.md`, `CHANGELOG.md`, `PROJECT_RULES.md`.
- Validation/tests run: Regenerated `Install.ps1` from corrected `InstallerCore` template; PowerShell parser validation for `WhoIsUsingThis.ps1` and `Install.ps1`; JSON validation for `app-metadata.json` and `InstallerCore\profiles\WhoIsUsingThis.json`; local-source installer update smoke completed with exit code `0`; installed file hash/readback matched repo files; registry readback preserved Unicode `MUIVerb`; targeted probes confirmed stale cached `UpToDate` is not reused after remote failure and dirty git workspaces are refused.

### Entry - 2026-05-11 (Update submenu must appear in scan action prompts)
- Date: 2026-05-11
- Problem: The scan header showed update status, but when a locking process was found the visible `[A]`, `[S]`, `[C]` prompt had no update entry.
- Root cause: `Update App` was wired only through final/no-target/path-error screens, not the mid-scan action prompts.
- Guardrail/rule: Every scan action prompt that can pause user input for lock handling must include `[U] Update App` and route to the same update submenu.
- Files affected: `WhoIsUsingThis.ps1`, `README.md`, `CHANGELOG.md`, `PROJECT_RULES.md`.
- Validation/tests run: PowerShell parser validation for `WhoIsUsingThis.ps1` and `Install.ps1`; static prompt probe confirmed the shared scan action prompt contains `[U] Update App`, routes to `Show-AppUpdateMenu`, and the old `[A]`/`[S]`/`[C]`-only prompt strings are absent.

### Entry - 2026-05-11 (WinAppManager-style scan action menus)
- Date: 2026-05-11
- Problem: The lock-action prompt still used letter shortcuts and mixed Greek/English labels, unlike the arrow-key main menu style used by WinAppManager.
- Root cause: The scanner flow predated the shared arrow menu pattern and only later gained a bolted-on `Update App` shortcut.
- Guardrail/rule: For lock-action pauses, use an arrow menu in this order: `Terminate all`, `Choose one-by-one`, `Skip`, `Update App`. Keep menu labels in English. `Choose one-by-one` must open an arrow process picker where `Enter` terminates the selected process and `Esc` skips the remaining processes in that scan stage.
- Files affected: `WhoIsUsingThis.ps1`, `app-metadata.json`, `README.md`, `CHANGELOG.md`, `PROJECT_RULES.md`.
- Validation/tests run: PowerShell parser validation for `WhoIsUsingThis.ps1` and `Install.ps1`; `git diff --check`; static probes confirmed scan action order, absence of old `[A]`/`[S]`/`[C]` action labels in current app/docs text, and no Greek visible UI strings remain in `WhoIsUsingThis.ps1`, `README.md`, or `CHANGELOG.md`; local-source installer update smoke and installed hash/readback verification completed.

### Entry - 2026-05-14 (Windows Utilities category move)

- Date: 2026-05-14
- Problem: `Who is using this?` was grouped under the shared `Explorer` category, which was too narrow for ownership/lock-inspection style utilities.
- Root cause: The shared SystemTools category name was chosen before the file-utility set expanded.
- Guardrail/rule: `WhoIsUsingThis` remains child-only, but its generated installer now targets `SystemTools\shell\WindowsUtilities\shell\WhoIsUsingThis`. Keep old `Explorer` child paths in cleanup during migration.
- Files affected: `Install.ps1`, `app-metadata.json`, `CHANGELOG.md`, `PROJECT_RULES.md`, `D:\Users\joty79\scripts\InstallerCore\profiles\WhoIsUsingThis.json`.
- Validation/tests run: Regenerated `Install.ps1` from `InstallerCore`; parser validation passed for generated installer.

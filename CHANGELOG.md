# Changelog

All notable user-facing changes for `WhoIsUsingThis` are recorded here.

## [2026-05-11] - 1.0.1

### Changed

- Bumped `app-metadata.json` to `1.0.1` for the commit-aware `Update App` behavior change.
- Made `Update App` status commit-aware with local/latest version, local/latest commit, source kind, and dirty workspace state.
- Changed git working-copy updates to use `git fetch` + fast-forward only and refuse dirty workspaces.
- Kept installed-copy updates on `UpdateGitHub` while comparing `state\install-meta.json` `github_commit` against the latest remote commit.
- Kept non-git portable-copy updates on `DownloadLatest -NoSelfRelaunch` so the scanner owns progress and relaunch.
- Prevented stale cached `UpToDate` status from being reused when a fresh remote check fails.

## [2026-05-11]

### Fixed

- Fixed installer registry writes for Unicode menu labels so `MUIVerb` keeps `Who is using this 🔎?` instead of degrading to `???` during verification.

## [2026-04-24]

### Added

- Added `app-metadata.json` as the canonical app name/version/repo metadata source.
- Added app-side `Update App` support using the `InstallerCore` In-App Update UI Contract with the WT adapter.
- Added a WinAppManager-style `Update App` actions submenu with `Run update now`, `Refresh update status`, and `Back`.

### Changed

- Regenerated `Install.ps1` from the current `InstallerCore` profile/template.
- Updated the scanner UI to show version/update status and offer `U = Update App` after a scan or when launched without a target.
- Updated post-update relaunch to preserve the original scan target and start the new host from the target's folder.

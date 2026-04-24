param([string]$targetPath)

# 🔸 Force UTF-8 Console Encoding
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 | Out-Null

$script:AppName = 'WhoIsUsingThis'
$script:AppVersion = '1.0.0'
$script:GitHubRepo = 'joty79/WhoIsUsingThis'
$script:MetadataPath = Join-Path $PSScriptRoot 'app-metadata.json'
$script:StatePath = Join-Path $PSScriptRoot 'state'
$script:InstallMetaPath = Join-Path $script:StatePath 'install-meta.json'
$script:UpdateStatusCachePath = Join-Path $script:StatePath 'app-update-status.json'
$script:UpdateStatusCacheTtlMinutes = 30
$script:UpdateStatus = [pscustomobject]@{
    Status = 'Unknown'
    Label = 'Status unavailable'
    LatestVersion = ''
    Branch = ''
    Message = 'Update status has not been checked yet.'
    CheckedAt = ''
}

function Initialize-AppMetadata {
    if (-not (Test-Path -LiteralPath $script:MetadataPath -PathType Leaf)) { return }
    try {
        $metadata = Get-Content -LiteralPath $script:MetadataPath -Raw -ErrorAction Stop | ConvertFrom-Json
        if ($metadata.PSObject.Properties['app_name'] -and -not [string]::IsNullOrWhiteSpace([string]$metadata.app_name)) {
            $script:AppName = [string]$metadata.app_name
        }
        if ($metadata.PSObject.Properties['version'] -and -not [string]::IsNullOrWhiteSpace([string]$metadata.version)) {
            $script:AppVersion = [string]$metadata.version
        }
        if ($metadata.PSObject.Properties['github_repo'] -and -not [string]::IsNullOrWhiteSpace([string]$metadata.github_repo)) {
            $script:GitHubRepo = [string]$metadata.github_repo
        }
    }
    catch {}
}

function ConvertTo-AppVersion {
    param([AllowEmptyString()][string]$VersionText)
    if ([string]::IsNullOrWhiteSpace($VersionText)) { return $null }
    try { return [version]$VersionText.Trim() }
    catch { return $null }
}

function New-UpdateStatus {
    param(
        [string]$Status = 'Unknown',
        [AllowEmptyString()][string]$LatestVersion = '',
        [AllowEmptyString()][string]$Branch = '',
        [string]$Message = 'Update status has not been checked yet.',
        [AllowEmptyString()][string]$CheckedAt = ''
    )

    $label = switch ($Status) {
        'UpToDate' { 'Up to date' }
        'UpdateAvailable' { if ([string]::IsNullOrWhiteSpace($LatestVersion)) { 'Update available' } else { "Update available ($LatestVersion)" } }
        'LocalAhead' { 'Local version ahead' }
        'Error' { 'Update check failed' }
        default { 'Status unavailable' }
    }

    [pscustomobject]@{
        Status = $Status
        Label = $label
        LatestVersion = $LatestVersion
        Branch = $Branch
        Message = $Message
        CheckedAt = $CheckedAt
    }
}

function Read-UpdateStatusCache {
    param([switch]$AllowStale)
    if (-not (Test-Path -LiteralPath $script:UpdateStatusCachePath -PathType Leaf)) { return $null }
    try {
        $cacheItem = Get-Item -LiteralPath $script:UpdateStatusCachePath -ErrorAction Stop
        if (-not $AllowStale -and ((Get-Date) - $cacheItem.LastWriteTime).TotalMinutes -gt $script:UpdateStatusCacheTtlMinutes) {
            return $null
        }
        $cache = Get-Content -LiteralPath $script:UpdateStatusCachePath -Raw -ErrorAction Stop | ConvertFrom-Json
        return (New-UpdateStatus -Status ([string]$cache.Status) -LatestVersion ([string]$cache.LatestVersion) -Branch ([string]$cache.Branch) -Message ([string]$cache.Message) -CheckedAt ([string]$cache.CheckedAt))
    }
    catch { return $null }
}

function Write-UpdateStatusCache {
    param([Parameter(Mandatory)]$Status)
    try {
        if (-not (Test-Path -LiteralPath $script:StatePath -PathType Container)) {
            New-Item -Path $script:StatePath -ItemType Directory -Force | Out-Null
        }
        $Status | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $script:UpdateStatusCachePath -Encoding UTF8
    }
    catch {}
}

function Get-RemoteAppMetadata {
    if ([string]::IsNullOrWhiteSpace($script:GitHubRepo)) { return $null }
    foreach ($branch in @('master', 'main')) {
        try {
            $metadata = Invoke-RestMethod -Uri "https://raw.githubusercontent.com/$($script:GitHubRepo)/$branch/app-metadata.json" -Method Get -Headers @{ 'User-Agent' = "$($script:AppName)/$($script:AppVersion)" } -TimeoutSec 8
            if ($null -ne $metadata) {
                return [pscustomobject]@{ Metadata = $metadata; Branch = $branch }
            }
        }
        catch {}
    }
    return $null
}

function Resolve-UpdateStatus {
    param([switch]$ForceRefresh)
    if (-not $ForceRefresh) {
        $cached = Read-UpdateStatusCache
        if ($null -ne $cached) {
            $script:UpdateStatus = $cached
            return $script:UpdateStatus
        }
    }

    $stale = Read-UpdateStatusCache -AllowStale
    $remote = Get-RemoteAppMetadata
    if ($null -eq $remote) {
        if ($null -ne $stale) {
            $stale.Message = 'Using cached update status because GitHub could not be reached.'
            $script:UpdateStatus = $stale
            return $script:UpdateStatus
        }
        $script:UpdateStatus = New-UpdateStatus -Status 'Error' -Message 'Could not reach GitHub to check updates.' -CheckedAt ((Get-Date).ToString('s'))
        return $script:UpdateStatus
    }

    $latestVersion = if ($remote.Metadata.PSObject.Properties['version']) { [string]$remote.Metadata.version } else { '' }
    $localVersion = ConvertTo-AppVersion -VersionText $script:AppVersion
    $remoteVersion = ConvertTo-AppVersion -VersionText $latestVersion
    $statusName = 'Unknown'
    $message = 'Update status is unavailable.'
    if ($null -ne $localVersion -and $null -ne $remoteVersion) {
        if ($localVersion -lt $remoteVersion) {
            $statusName = 'UpdateAvailable'
            $message = "Update available: v$latestVersion"
        }
        elseif ($localVersion -gt $remoteVersion) {
            $statusName = 'LocalAhead'
            $message = "Local version v$script:AppVersion is ahead of origin."
        }
        else {
            $statusName = 'UpToDate'
            $message = "App is up to date at v$latestVersion."
        }
    }

    $script:UpdateStatus = New-UpdateStatus -Status $statusName -LatestVersion $latestVersion -Branch ([string]$remote.Branch) -Message $message -CheckedAt ((Get-Date).ToString('s'))
    Write-UpdateStatusCache -Status $script:UpdateStatus
    return $script:UpdateStatus
}

function Get-RecentTextFileLines {
    param([AllowEmptyString()][string]$Path, [int]$TailCount = 10)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return @() }
    try { return @(Get-Content -LiteralPath $Path -Tail $TailCount -ErrorAction Stop | ForEach-Object { [string]$_ }) }
    catch { return @() }
}

function Get-TargetWorkingDirectory {
    if ([string]::IsNullOrWhiteSpace($targetPath)) { return $PSScriptRoot }
    try {
        if ([System.IO.Directory]::Exists($targetPath)) { return $targetPath }
        if ([System.IO.File]::Exists($targetPath)) {
            $parent = [System.IO.Path]::GetDirectoryName($targetPath)
            if (-not [string]::IsNullOrWhiteSpace($parent)) { return $parent }
        }
    }
    catch {}
    return $PSScriptRoot
}

function Write-AppSection {
    param([string]$Title)
    Write-Host ''
    Write-Host '◆ ' -ForegroundColor Cyan -NoNewline
    Write-Host $Title -ForegroundColor Cyan -NoNewline
    Write-Host (' ' + ('-' * 72)) -ForegroundColor DarkGray
}

function Read-AppConsoleKey {
    try {
        $key = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        return [pscustomobject]@{
            VirtualKeyCode = [int]$key.VirtualKeyCode
            Character = [char]$key.Character
        }
    }
    catch {
        $key = [Console]::ReadKey($true)
        return [pscustomobject]@{
            VirtualKeyCode = [int]$key.Key
            Character = [char]$key.KeyChar
        }
    }
}

function ConvertTo-NativeArgumentString {
    param([string[]]$ArgumentList)
    @($ArgumentList | ForEach-Object {
        $argument = [string]$_
        if ($argument -match '[\s"]') {
            '"' + ($argument -replace '"', '\"') + '"'
        }
        else {
            $argument
        }
    }) -join ' '
}

function Show-AppHeader {
    param([AllowEmptyString()][string]$Subtitle = 'Handle scan + DLL modules + ACL diagnostics')
    $updateColor = switch ($script:UpdateStatus.Status) {
        'UpToDate' { 'Green' }
        'UpdateAvailable' { 'Yellow' }
        'LocalAhead' { 'Yellow' }
        'Error' { 'Red' }
        default { 'DarkGray' }
    }

    Clear-Host
    Write-Host ''
    Write-Host "╔══════════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host '║ ' -ForegroundColor Cyan -NoNewline
    Write-Host ("{0} v{1}" -f $script:AppName, $script:AppVersion).PadRight(76) -ForegroundColor White -NoNewline
    Write-Host "║" -ForegroundColor Cyan
    Write-Host '║ ' -ForegroundColor Cyan -NoNewline
    Write-Host $Subtitle.PadRight(76) -ForegroundColor DarkGray -NoNewline
    Write-Host "║" -ForegroundColor Cyan
    Write-Host '║ ' -ForegroundColor Cyan -NoNewline
    Write-Host 'Update: ' -ForegroundColor Green -NoNewline
    Write-Host $script:UpdateStatus.Label.PadRight(68) -ForegroundColor $updateColor -NoNewline
    Write-Host "║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
}

function Show-AppUpdateResultPanel {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Good', 'Warn', 'Error')][string]$Level = 'Info',
        [string[]]$RecentLines = @(),
        [switch]$AutoRestart
    )
    $color = switch ($Level) {
        'Good' { 'Green' }
        'Warn' { 'Yellow' }
        'Error' { 'Red' }
        default { 'Cyan' }
    }
    Show-AppHeader -Subtitle 'Update App'
    Write-AppSection -Title 'Update App'
    Write-Host "  $Message" -ForegroundColor $color
    if (@($RecentLines).Count -gt 0) {
        Write-AppSection -Title 'Recent Output'
        foreach ($line in @($RecentLines | Select-Object -Last 10)) {
            $displayLine = [string]$line
            if ($displayLine.Length -gt 118) { $displayLine = $displayLine.Substring(0, 115) + '...' }
            Write-Host "  $displayLine" -ForegroundColor DarkGray
        }
    }
    Write-AppSection -Title 'Commands'
    if ($AutoRestart) {
        Write-Host "  Restarting $script:AppName with updated files..." -ForegroundColor Green
    }
    else {
        Write-Host '  ESC ' -ForegroundColor Red -NoNewline
        Write-Host 'back' -ForegroundColor DarkGray
    }
}

function Get-InstallerAction {
    $defaultInstallPath = [System.IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'WhoIsUsingThisContext')).TrimEnd('\')
    $currentRootPath = [System.IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\')
    if ($currentRootPath -ieq $defaultInstallPath) { return 'UpdateGitHub' }
    return 'DownloadLatest'
}

function Start-UpdatedAppHost {
    $appPath = Join-Path $PSScriptRoot 'WhoIsUsingThis.ps1'
    if (-not (Test-Path -LiteralPath $appPath -PathType Leaf)) { return $false }
    $pwshCommand = Get-Command pwsh.exe -ErrorAction SilentlyContinue
    if ($null -eq $pwshCommand) { return $false }
    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $appPath)
    if (-not [string]::IsNullOrWhiteSpace($targetPath)) { $arguments += $targetPath }
    $workingDirectory = Get-TargetWorkingDirectory
    $wtCommand = Get-Command wt.exe -ErrorAction SilentlyContinue
    try {
        if ($null -ne $wtCommand) {
            $wtArguments = @('-w', 'new', 'nt', '--title', 'Handle Scan', 'pwsh.exe') + $arguments
            Start-Process -FilePath $wtCommand.Source -ArgumentList (ConvertTo-NativeArgumentString -ArgumentList $wtArguments) -WorkingDirectory $workingDirectory | Out-Null
            return $true
        }
        Start-Process -FilePath $pwshCommand.Source -ArgumentList (ConvertTo-NativeArgumentString -ArgumentList $arguments) -WorkingDirectory $workingDirectory | Out-Null
        return $true
    }
    catch { return $false }
}

function Show-AppUpdateMenuOptions {
    param([string[]]$Options, [int]$SelectedIndex)
    Write-AppSection -Title 'Actions'
    for ($index = 0; $index -lt $Options.Count; $index++) {
        if ($index -eq $SelectedIndex) {
            Write-Host '  > ' -ForegroundColor White -NoNewline
            Write-Host $Options[$index] -ForegroundColor Black -BackgroundColor Gray
        }
        elseif ($Options[$index] -eq 'Back') {
            Write-Host "    $($Options[$index])" -ForegroundColor DarkGray
        }
        else {
            Write-Host "    $($Options[$index])" -ForegroundColor Cyan
        }
    }
}

function Show-AppUpdateMenu {
    $selectedIndex = 0
    $options = @('Run update now', 'Refresh update status', 'Back')

    while ($true) {
        $action = Get-InstallerAction
        $targetLabel = if ($action -eq 'DownloadLatest') { 'Workspace working copy' } else { 'Installed app copy' }
        $methodLabel = if ($action -eq 'DownloadLatest') { 'Repo-aware working-copy update' } else { 'Installer/GitHub in-place update' }
        $latestVersionLabel = if ([string]::IsNullOrWhiteSpace($script:UpdateStatus.LatestVersion)) { '--' } else { $script:UpdateStatus.LatestVersion }
        $branchLabel = if ([string]::IsNullOrWhiteSpace($script:UpdateStatus.Branch)) { '--' } else { $script:UpdateStatus.Branch }
        $checkedAtLabel = if ([string]::IsNullOrWhiteSpace($script:UpdateStatus.CheckedAt)) { '--' } else { $script:UpdateStatus.CheckedAt.Replace('T', ' ') }
        $statusColor = switch ($script:UpdateStatus.Status) {
            'UpToDate' { 'Green' }
            'UpdateAvailable' { 'Yellow' }
            'LocalAhead' { 'Yellow' }
            'Error' { 'Red' }
            default { 'DarkGray' }
        }

        Show-AppHeader -Subtitle 'Update App'
        Write-AppSection -Title 'Update App'
        Write-Host '  Current version: ' -ForegroundColor Gray -NoNewline
        Write-Host $script:AppVersion -ForegroundColor White
        Write-Host '  Latest version:  ' -ForegroundColor Gray -NoNewline
        Write-Host $latestVersionLabel -ForegroundColor White
        Write-Host '  Status:          ' -ForegroundColor Gray -NoNewline
        Write-Host $script:UpdateStatus.Label -ForegroundColor $statusColor
        Write-Host '  Repo:            ' -ForegroundColor Gray -NoNewline
        Write-Host $script:GitHubRepo -ForegroundColor White -NoNewline
        Write-Host '  |  Branch: ' -ForegroundColor DarkGray -NoNewline
        Write-Host $branchLabel -ForegroundColor White
        Write-Host '  Update target:   ' -ForegroundColor Gray -NoNewline
        Write-Host $targetLabel -ForegroundColor White
        Write-Host '  Method:          ' -ForegroundColor Gray -NoNewline
        Write-Host $methodLabel -ForegroundColor White
        Write-Host '  Relaunch target: ' -ForegroundColor Gray -NoNewline
        $relaunchTargetLabel = if ([string]::IsNullOrWhiteSpace($targetPath)) { '--' } else { $targetPath }
        Write-Host $relaunchTargetLabel -ForegroundColor White
        Write-Host '  Last check:      ' -ForegroundColor Gray -NoNewline
        Write-Host $checkedAtLabel -ForegroundColor White
        if (-not [string]::IsNullOrWhiteSpace($script:UpdateStatus.Message)) {
            Write-Host "  $($script:UpdateStatus.Message)" -ForegroundColor DarkGray
        }
        Show-AppUpdateMenuOptions -Options $options -SelectedIndex $selectedIndex
        Write-Host ''
        Write-Host '  ↑↓ navigate   Enter = select   Esc = back' -ForegroundColor DarkGray

        $key = Read-AppConsoleKey
        switch ($key.VirtualKeyCode) {
            38 { $selectedIndex = [Math]::Max(0, $selectedIndex - 1) }
            40 { $selectedIndex = [Math]::Min($options.Count - 1, $selectedIndex + 1) }
            27 { return $false }
            13 {
                switch ($selectedIndex) {
                    0 { return (Invoke-AppUpdate) }
                    1 { [void](Resolve-UpdateStatus -ForceRefresh) }
                    2 { return $false }
                }
            }
        }
    }
}

function Invoke-AppUpdate {
    $installerPath = Join-Path $PSScriptRoot 'Install.ps1'
    if (-not (Test-Path -LiteralPath $installerPath -PathType Leaf)) {
        Show-AppUpdateResultPanel -Message 'Install.ps1 was not found next to the app.' -Level 'Error'
        return $false
    }
    $pwshCommand = Get-Command pwsh.exe -ErrorAction SilentlyContinue
    if ($null -eq $pwshCommand) {
        Show-AppUpdateResultPanel -Message 'pwsh.exe was not found.' -Level 'Error'
        return $false
    }
    $action = Get-InstallerAction
    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $installerPath, '-Action', $action, '-Force')
    if ($action -eq 'UpdateGitHub') { $arguments += '-NoExplorerRestart' }
    if ($action -eq 'DownloadLatest') { $arguments += '-NoSelfRelaunch' }
    $progressMessage = if ($action -eq 'UpdateGitHub') { 'Updating from GitHub inside the current app session...' } else { 'Updating this working copy inside the current app session...' }
    $stdoutPath = Join-Path $env:TEMP ("WhoIsUsingThis_updater_out_{0}.log" -f [guid]::NewGuid().ToString('N'))
    $stderrPath = Join-Path $env:TEMP ("WhoIsUsingThis_updater_err_{0}.log" -f [guid]::NewGuid().ToString('N'))
    $installerLogPath = Join-Path $PSScriptRoot 'logs\installer.log'
    try {
        $process = Start-Process -FilePath $pwshCommand.Source -ArgumentList $arguments -WorkingDirectory $PSScriptRoot -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath -WindowStyle Hidden -PassThru -ErrorAction Stop
        while (-not $process.HasExited) {
            $recentLines = @((Get-RecentTextFileLines -Path $installerLogPath -TailCount 8) + (Get-RecentTextFileLines -Path $stderrPath -TailCount 3))
            Show-AppUpdateResultPanel -Message $progressMessage -Level 'Info' -RecentLines $recentLines
            Start-Sleep -Milliseconds 250
        }
        $process.Refresh()
        $exitCode = [int]$process.ExitCode
        $finalLines = @((Get-RecentTextFileLines -Path $installerLogPath -TailCount 8) + (Get-RecentTextFileLines -Path $stderrPath -TailCount 5))
        if ($exitCode -le 2) {
            Show-AppUpdateResultPanel -Message 'Update finished. Restarting the updated app host and closing this window...' -Level 'Good' -RecentLines $finalLines -AutoRestart
            Start-Sleep -Milliseconds 900
            if (Start-UpdatedAppHost) { exit 0 }
            Show-AppUpdateResultPanel -Message 'Update finished, but the app could not relaunch automatically.' -Level 'Warn' -RecentLines $finalLines
            return $false
        }
        Show-AppUpdateResultPanel -Message ("Update failed with exit code {0}." -f $exitCode) -Level 'Error' -RecentLines $finalLines
        return $false
    }
    catch {
        Show-AppUpdateResultPanel -Message "Could not start updater: $($_.Exception.Message)" -Level 'Error'
        return $false
    }
    finally {
        foreach ($tempPath in @($stdoutPath, $stderrPath)) {
            try {
                if (Test-Path -LiteralPath $tempPath -PathType Leaf) {
                    Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
                }
            }
            catch {}
        }
    }
}

function Read-CloseOrUpdate {
    Write-Host ''
    Write-Host 'U = Update App  |  any other key = close' -ForegroundColor Yellow
    $key = Read-AppConsoleKey
    if ($key.VirtualKeyCode -eq 85 -or [char]::ToUpperInvariant($key.Character) -eq 'U') {
        [void](Show-AppUpdateMenu)
        Write-Host ''
        Write-Host 'Press any key to close...' -ForegroundColor Yellow
        $null = Read-AppConsoleKey
    }
}

Initialize-AppMetadata
[void](Resolve-UpdateStatus)

if ($null -eq $targetPath) { $targetPath = '' }

# ---------------------------------------------------------
# 1. Trailing Slash / Quote Cleanup
# ---------------------------------------------------------
if (-not [string]::IsNullOrWhiteSpace($targetPath) -and $targetPath.EndsWith('\')) { $targetPath = $targetPath.TrimEnd('\') }
$targetPath = $targetPath -replace '"', ''

# ---------------------------------------------------------
# 2. Launch Mode: Normal → WT (emojis) | Safe Mode → PS5 (Nerd Font)
# ---------------------------------------------------------
$isWT = $null -ne $env:WT_SESSION
$isSafeMode = Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SafeBoot\Option'

# Normal mode: if WT exists but we're not in it yet → re-launch via WT
if (-not $isWT -and -not $isSafeMode -and (Get-Command wt.exe -ErrorAction SilentlyContinue)) {
    Start-Process wt.exe -ArgumentList "-w 0 --title `"Handle Scan`" pwsh.exe -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" `"$targetPath`"" -Verb RunAs
    exit
}

# Safe Mode / already in WT: elevate if needed
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $engine = if ($PSVersionTable.PSVersion.Major -ge 7) { 'pwsh.exe' } else { 'powershell.exe' }
    Start-Process $engine -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" `"$targetPath`"" -Verb RunAs
    exit
}

# If we're admin but launched hidden (from VBS) and NOT in WT → re-launch visible
if (-not $isWT -and $env:WHOISUSING_HIDDEN -eq '1') {
    $env:WHOISUSING_HIDDEN = $null  # clear flag to prevent loop
    $engine = if ($PSVersionTable.PSVersion.Major -ge 7) { 'pwsh.exe' } else { 'powershell.exe' }
    Start-Process $engine -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" `"$targetPath`""
    exit
}

# 🔸 Icon Set
if ($isWT) {
    $ico = @{
        N1='1️⃣'; N2='2️⃣'; N3='3️⃣'; N4='4️⃣'
        OK='✅'; ERR='❌'; WARN='⚠️'; LOCK='🔒'; DENY='🚫'
        HIT='🎯'; KILL='💀'; FIND='🔍'; TEST='🧪'; LINK='🔗'
        CBS='📋'; LOG='📜'; TIP='💡'; DOT='🔸'; WAIT='⏳'
        INFO='ℹ️'; TERM='⚠️'; DIAG='🔸'
    }
} else {
    # Nerd Font icons (CaskaydiaCove NFM)
    $ico = @{
        N1="$([char]0xF0C8) 1"; N2="$([char]0xF0C8) 2"; N3="$([char]0xF0C8) 3"; N4="$([char]0xF0C8) 4"
        OK="$([char]0xF00C)"; ERR="$([char]0xF00D)"; WARN="$([char]0xF071)"; LOCK="$([char]0xF023)"; DENY="$([char]0xF05E)"
        HIT="$([char]0xF05B)"; KILL="$([char]0xF00D)"; FIND="$([char]0xF002)"; TEST="$([char]0xF0C3)"; LINK="$([char]0xF0C1)"
        CBS="$([char]0xF1C0)"; LOG="$([char]0xF15C)"; TIP="$([char]0xF0EB)"; DOT="$([char]0xF054)"; WAIT="$([char]0xF017)"
        INFO="$([char]0xF05A)"; TERM="$([char]0xF120)"; DIAG="$([char]0xF054)"
    }
}

# ---------------------------------------------------------
# 3. Basic Checks
# ---------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($targetPath)) {
    Show-AppHeader
    Write-Host ''
    Write-Host 'No target path was provided.' -ForegroundColor Yellow
    Write-Host 'Launch from the context menu, or pass -targetPath from a terminal.' -ForegroundColor Gray
    Read-CloseOrUpdate
    exit
}

if (-not ([System.IO.File]::Exists($targetPath) -or [System.IO.Directory]::Exists($targetPath))) {
    Show-AppHeader
    Write-Host "`n$($ico.ERR) Error: Path not found." -ForegroundColor Red
    Write-Host "   $targetPath" -ForegroundColor Gray
    Read-CloseOrUpdate
    exit
}

$fileName = [System.IO.Path]::GetFileName($targetPath)
$isFolder = [System.IO.Directory]::Exists($targetPath)

# ---------------------------------------------------------
# Critical System Paths — Extra warning πριν takeown/delete
# ---------------------------------------------------------
$criticalPaths = @(
    'C:\Windows\WinSxS'
    'C:\Windows\System32'
    'C:\Windows\SysWOW64'
    'C:\Windows\Boot'
    'C:\Windows\servicing'
    'C:\Windows\security'
    'C:\Windows\assembly'
    'C:\Windows\Microsoft.NET'
    'C:\Windows\Installer'
    'C:\Windows\Fonts'
    'C:\Windows\SoftwareDistribution'
    'C:\Windows\inf'
    'C:\Windows\drivers'
    'C:\System Volume Information'
    'C:\Recovery'
    'C:\Program Files\WindowsApps'
)
$isCriticalPath = $false
foreach ($cp in $criticalPaths) {
    if ($targetPath -ieq $cp -or $targetPath.StartsWith($cp + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
        $isCriticalPath = $true; break
    }
}

Show-AppHeader
Write-Host "`n--- Scanning: $targetPath ---" -ForegroundColor Cyan

# Resolve handle.exe from bundled assets first, then fallback to PATH.
$handleExePath = $null
$handleCandidates = @(
    (Join-Path $PSScriptRoot 'assets\bin\handle.exe'),
    (Join-Path $PSScriptRoot 'handle.exe')
)
foreach ($candidate in $handleCandidates) {
    if (Test-Path -LiteralPath $candidate) {
        $handleExePath = $candidate
        break
    }
}
if (-not $handleExePath) {
    $handleCmd = Get-Command handle.exe -ErrorAction SilentlyContinue
    if ($handleCmd) {
        $handleExePath = $handleCmd.Source
    }
}
if (-not $handleExePath) {
    Write-Host "`n$($ico.ERR) Error: handle.exe was not found." -ForegroundColor Red
    Write-Host "   Expected: $PSScriptRoot\assets\bin\handle.exe" -ForegroundColor Gray
    Write-Host "`nPress any key to close..." -ForegroundColor Yellow
    $null = [Console]::ReadKey(); exit
}

# =========================================================
# ΜΕΘΟΔΟΣ 1: Handle.exe (Standard Locks)
# =========================================================
Write-Host "$($ico.N1)  Έλεγχος Handles (Standard Locks)..." -ForegroundColor Gray

# Μόνο ακριβής αναζήτηση (Χωρίς Spam)
$handleResults = & $handleExePath -a -u "$targetPath" 2>$null

if ($handleResults -match "pid:") {
    $processedPids = @()
    $lockList = @()

    # Συλλογή unique PIDs σε λίστα
    $lines = $handleResults | Where-Object { $_ -match "pid:\s+\d+" } | Select-Object -Unique
    
    foreach ($line in $lines) {
        if ($line -match "(?<procName>\S+)\s+pid:\s+(?<pid>\d+)") {
             $procName = $Matches['procName']
             $procPid  = $Matches['pid']
             if ($procPid -in $processedPids) { continue }
             $processedPids += $procPid

             $lockList += [PSCustomObject]@{
                 Name       = $procName
                 PID        = [int]$procPid
                 IsTerminal = $procName -match "pwsh|cmd|wt|WindowsTerminal|OpenConsole|conhost|powershell"
             }
        }
    }

    # --- Εμφάνιση πλήρους λίστας ---
    $termCount = ($lockList | Where-Object IsTerminal).Count
    $appCount  = $lockList.Count - $termCount
    Write-Host "   Βρέθηκαν $($lockList.Count) processes ($appCount apps, $termCount terminals):" -ForegroundColor White
    Write-Host ""

    $idx = 0
    foreach ($proc in $lockList) {
        $idx++
        if ($proc.IsTerminal) {
            Write-Host "   [$idx] $($ico.TERM) $($proc.Name) (PID: $($proc.PID))" -ForegroundColor Magenta
        } else {
            Write-Host "   [$idx] $($ico.HIT) $($proc.Name) (PID: $($proc.PID))" -ForegroundColor Yellow
        }
    }

    # --- Μενού Ενεργειών ---
    Write-Host ""
    Write-Host "   [A] Τερματισμός ΟΛΩΝ  |  [S] Παράλειψη  |  [C] Επιλογή ένα-ένα" -ForegroundColor Cyan
    $handleChoice = (Read-Host "   Επιλογή").ToUpper()

    switch ($handleChoice) {
        "A" {
            foreach ($proc in $lockList) {
                if ($proc.IsTerminal) {
                    Write-Host "   $($ico.WARN)  Skipped terminal: $($proc.Name) (PID: $($proc.PID))" -ForegroundColor Magenta
                    continue
                }
                try {
                    Stop-Process -Id $proc.PID -Force -ErrorAction Stop
                    Write-Host "   $($ico.KILL) $($proc.Name) (PID: $($proc.PID))" -ForegroundColor Red
                } catch {
                    Write-Host "   $($ico.ERR) Αποτυχία: $($proc.Name) (PID: $($proc.PID))" -ForegroundColor Red
                }
            }
            if ($termCount -gt 0) {
                Write-Host "   $($ico.TIP) $termCount terminal(s) παραλείφθηκαν — χρήσε [C] αν θες να τα κλείσεις" -ForegroundColor DarkYellow
            }
        }
        "C" {
            foreach ($proc in $lockList) {
                $label = if ($proc.IsTerminal) { "$($ico.TERM) Terminal:" } else { "$($ico.HIT)" }
                $ans = Read-Host "   $label $($proc.Name) (PID: $($proc.PID)) — Τερματισμός; (Y/N)"
                if ($ans -eq "Y") {
                    try {
                        Stop-Process -Id $proc.PID -Force -ErrorAction Stop
                        Write-Host "      $($ico.KILL) Τερματίστηκε." -ForegroundColor Red
                    } catch {
                        Write-Host "      $($ico.ERR) Αποτυχία." -ForegroundColor Red
                    }
                }
            }
        }
        default {
            Write-Host "   Παράλειψη." -ForegroundColor DarkGray
        }
    }
} else {
    Write-Host "   (Κανένα standard file handle)" -ForegroundColor DarkGray
}

# =========================================================
# ΜΕΘΟΔΟΣ 2: Module/Deep Scan (DLLs & Libraries)
# =========================================================
if (-not $isFolder) {
    # --- ΠΕΡΙΠΤΩΣΗ Α: ΑΡΧΕΙΟ (Tasklist Module Scan) ---
    Write-Host "`n$($ico.N2)  Έλεγχος Loaded Modules (DLLs)..." -ForegroundColor Gray
    
    # Χρήση Regex Escape για αρχεία με [ ]
    $moduleOutput = tasklist /m [Regex]::Escape($fileName) 2>$null
    $modList = @()
    
    if ($moduleOutput -match [Regex]::Escape($fileName)) {
        foreach ($line in ($moduleOutput | Select-Object -Skip 3)) {
            if ($line -match "^(\S+)\s+(\d+)") {
                $modList += [PSCustomObject]@{ Name = $matches[1]; PID = [int]$matches[2]; LockedFile = $fileName }
            }
        }
    }

    if ($modList.Count -gt 0) {
        Write-Host "   Βρέθηκαν $($modList.Count) processes με loaded module:" -ForegroundColor White
        Write-Host ""
        $idx = 0
        foreach ($m in $modList) {
            $idx++
            Write-Host "   [$idx] $($ico.HIT) $($m.Name) (PID: $($m.PID))" -ForegroundColor Yellow
        }
        Write-Host ""
        Write-Host "   [A] Τερματισμός ΟΛΩΝ  |  [S] Παράλειψη  |  [C] Επιλογή ένα-ένα" -ForegroundColor Cyan
        $modChoice = (Read-Host "   Επιλογή").ToUpper()

        switch ($modChoice) {
            "A" {
                foreach ($m in $modList) {
                    try {
                        Stop-Process -Id $m.PID -Force -ErrorAction Stop
                        Write-Host "   $($ico.KILL) $($m.Name) (PID: $($m.PID))" -ForegroundColor Red
                    } catch {
                        Write-Host "   $($ico.ERR) Αποτυχία: $($m.Name) (PID: $($m.PID))" -ForegroundColor Red
                    }
                }
            }
            "C" {
                foreach ($m in $modList) {
                    $ans = Read-Host "   $($ico.HIT) $($m.Name) (PID: $($m.PID)) — Τερματισμός; (Y/N)"
                    if ($ans -eq "Y") {
                        try {
                            Stop-Process -Id $m.PID -Force -ErrorAction Stop
                            Write-Host "      $($ico.KILL) Τερματίστηκε." -ForegroundColor Red
                        } catch {
                            Write-Host "      $($ico.ERR) Αποτυχία." -ForegroundColor Red
                        }
                    }
                }
            }
            default { Write-Host "   Παράλειψη." -ForegroundColor DarkGray }
        }
    } else { Write-Host "   (Κανένα loaded module)" -ForegroundColor DarkGray }

} else {
    # --- ΠΕΡΙΠΤΩΣΗ Β: ΦΑΚΕΛΟΣ (Deep Memory Scan) ---
    # Πιάνει Everything, AIMP, Upscayl που κλειδώνουν υποφάκελους
    Write-Host "`n$($ico.N2)  Deep Scan Φακέλου (Αναζήτηση σε όλα τα Processes)..." -ForegroundColor Gray
    Write-Host "   $($ico.WAIT) Σάρωση μνήμης..." -ForegroundColor DarkGray
    
    $deepList = @()
    $processes = Get-Process -ErrorAction SilentlyContinue
    
    foreach ($proc in $processes) {
        try {
            # Ψάχνουμε αν το process έχει φορτώσει κάτι από τον φάκελό μας
            $matchingModule = $proc.Modules | Where-Object { $_.FileName -and $_.FileName.StartsWith($targetPath, [System.StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
            if ($matchingModule) {
                $deepList += [PSCustomObject]@{
                    Name       = $proc.Name
                    PID        = $proc.Id
                    LockedFile = $matchingModule.FileName
                }
            }
        } catch { }
    }

    if ($deepList.Count -gt 0) {
        Write-Host "   Βρέθηκαν $($deepList.Count) processes με locked modules:" -ForegroundColor White
        Write-Host ""
        $idx = 0
        foreach ($d in $deepList) {
            $idx++
            Write-Host "   [$idx] $($ico.HIT) $($d.Name) (PID: $($d.PID))" -ForegroundColor Yellow
            Write-Host "       Locked: $($d.LockedFile)" -ForegroundColor DarkGray
        }
        Write-Host ""
        Write-Host "   [A] Τερματισμός ΟΛΩΝ  |  [S] Παράλειψη  |  [C] Επιλογή ένα-ένα" -ForegroundColor Cyan
        $deepChoice = (Read-Host "   Επιλογή").ToUpper()

        switch ($deepChoice) {
            "A" {
                foreach ($d in $deepList) {
                    try {
                        Stop-Process -Id $d.PID -Force -ErrorAction Stop
                        Write-Host "   $($ico.KILL) $($d.Name) (PID: $($d.PID))" -ForegroundColor Red
                    } catch {
                        Write-Host "   $($ico.ERR) Αποτυχία: $($d.Name) (PID: $($d.PID))" -ForegroundColor Red
                    }
                }
            }
            "C" {
                foreach ($d in $deepList) {
                    $ans = Read-Host "   $($ico.HIT) $($d.Name) (PID: $($d.PID)) — Τερματισμός; (Y/N)"
                    if ($ans -eq "Y") {
                        try {
                            Stop-Process -Id $d.PID -Force -ErrorAction Stop
                            Write-Host "      $($ico.KILL) Τερματίστηκε." -ForegroundColor Red
                        } catch {
                            Write-Host "      $($ico.ERR) Αποτυχία." -ForegroundColor Red
                        }
                    }
                }
            }
            default { Write-Host "   Παράλειψη." -ForegroundColor DarkGray }
        }
    } else { Write-Host "   (Κανένα κλειδωμένο module σε υποφάκελο)" -ForegroundColor DarkGray }
}

# =========================================================
# ΜΕΘΟΔΟΣ 3: ACL & Ownership Analysis
# Πιάνει: TrustedInstaller, SYSTEM, ACL-blocked αρχεία,
#          CBS InFlight (WinSxS pending updates)
# =========================================================
Write-Host "`n$($ico.N3)  Έλεγχος Ownership & ACL Permissions..." -ForegroundColor Gray

$aclIssuesFound = $false

try {
    $acl = Get-Acl -LiteralPath $targetPath -ErrorAction Stop
    $owner = $acl.Owner

    # --- Ownership Check ---
    $isSystemOwned = $owner -match "TrustedInstaller|NT SERVICE|SYSTEM|S-1-5-80"
    $isAdminOwned  = $owner -match "BUILTIN\\Administrators|Administrators"

    if ($isSystemOwned) {
        Write-Host "   $($ico.LOCK) Owner: $owner" -ForegroundColor Red
        Write-Host "      Ο φάκελος/αρχείο ανήκει στα Windows (TrustedInstaller/SYSTEM)" -ForegroundColor DarkYellow
        $aclIssuesFound = $true
    } elseif ($isAdminOwned) {
        Write-Host "   $($ico.OK) Owner: $owner (Administrators)" -ForegroundColor Green
    } else {
        Write-Host "   $($ico.INFO)  Owner: $owner" -ForegroundColor Cyan
    }

    # --- Access Rights Check (μπορούμε πραγματικά να σβήσουμε;) ---
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $hasDeleteAccess = $false
    foreach ($rule in $acl.Access) {
        $identity = $rule.IdentityReference.Value
        if ($identity -match "Administrators|$([regex]::Escape($currentUser))|Everyone|Users") {
            if ($rule.FileSystemRights -match "FullControl|Delete|Modify" -and $rule.AccessControlType -eq "Allow") {
                $hasDeleteAccess = $true; break
            }
        }
    }
    if (-not $hasDeleteAccess) {
        Write-Host "   $($ico.DENY) Ο τρέχων χρήστης ΔΕΝ έχει Delete permission!" -ForegroundColor Red
        $aclIssuesFound = $true
    }

    # --- Subfolder ACL Scan (μόνο για φακέλους) ---
    if ($isFolder) {
        Write-Host "   $($ico.FIND) Σάρωση sub-items για Access Denied..." -ForegroundColor DarkGray
        $deniedCount = 0
        $deniedSamples = @()
        $scanCap = 500          # Αρκετό για diagnostic — αποφυγή multi-minute scan
        $scanned = 0
        $hitCap = $false

        # ⚡ Streaming enumeration + Owner-only ACL (αντί dir /b /s + Get-Acl full)
        $enumOpts = [System.IO.EnumerationOptions]::new()
        $enumOpts.RecurseSubdirectories = $true
        $enumOpts.IgnoreInaccessible = $true
        $enumOpts.AttributesToSkip = [System.IO.FileAttributes]::None   # Σκανάρει και Hidden/System

        foreach ($item in [System.IO.Directory]::EnumerateFileSystemEntries($targetPath, '*', $enumOpts)) {
            $scanned++
            if ($scanned % 200 -eq 0) {
                Write-Progress -Activity "ACL Scan" -Status "$scanned items scanned ($deniedCount issues found)"
            }
            if ($scanned -ge $scanCap) { $hitCap = $true; break }

            try {
                # Ζητάμε ΜΟΝΟ Owner (παραλείπουμε DACL/SACL/Group — πολύ πιο γρήγορο)
                if ([System.IO.Directory]::Exists($item)) {
                    $sec = [System.IO.DirectoryInfo]::new($item).GetAccessControl(
                        [System.Security.AccessControl.AccessControlSections]::Owner
                    )
                } else {
                    $sec = [System.IO.FileInfo]::new($item).GetAccessControl(
                        [System.Security.AccessControl.AccessControlSections]::Owner
                    )
                }
                # ⚡ SID comparison αντί NTAccount — αποφυγή IdentityNotMappedException (false positives)
                # S-1-5-18 = SYSTEM | S-1-5-80-* = NT SERVICE (TrustedInstaller κλπ)
                $ownerSid = $sec.GetOwner([System.Security.Principal.SecurityIdentifier]).Value
                if ($ownerSid -match '^S-1-5-18$|^S-1-5-80-') {
                    $deniedCount++
                    if ($deniedSamples.Count -lt 3) {
                        # Resolve σε όνομα μόνο για display (best-effort)
                        try { $displayName = $sec.GetOwner([System.Security.Principal.NTAccount]).Value }
                        catch { $displayName = $ownerSid }
                        $deniedSamples += "      → $item (Owner: $displayName)"
                    }
                }
            } catch [System.UnauthorizedAccessException] {
                # Μόνο πραγματικό Access Denied — όχι unrelated exceptions
                $deniedCount++
                if ($deniedSamples.Count -lt 3) { $deniedSamples += "      → $item (Access Denied)" }
            } catch {
                # Αγνοούμε: long paths, corrupt ACL, orphaned SID κλπ — ΔΕΝ είναι access issues
            }
        }
        Write-Progress -Activity "ACL Scan" -Completed

        if ($deniedCount -gt 0) {
            $countLabel = if ($hitCap) { "τουλάχιστον $deniedCount (σάρωση: πρώτα $scanCap items)" } else { "$deniedCount" }
            Write-Host "   $($ico.DENY) Βρέθηκαν $countLabel αρχεία/φάκελοι με restricted access:" -ForegroundColor Red
            foreach ($s in $deniedSamples) { Write-Host $s -ForegroundColor DarkYellow }
            if ($deniedCount -gt 3) { Write-Host "      ... και $($deniedCount - 3) ακόμα" -ForegroundColor DarkGray }
            $aclIssuesFound = $true
        } else {
            Write-Host "   $($ico.OK) Όλα τα sub-items έχουν κανονικά permissions ($scanned scanned)" -ForegroundColor Green
        }
    }

    # --- CBS / WinSxS InFlight Detection ---
    if ($targetPath -match "WinSxS\\Temp\\InFlight|WinSxS\\Temp\\PendingDeletes|WinSxS\\Temp\\PendingRenames") {
        Write-Host "`n   $($ico.WARN)  CBS INFLIGHT DETECTED!" -ForegroundColor Magenta
        Write-Host "      Αυτός ο φάκελος περιέχει pending Windows Update components." -ForegroundColor DarkYellow
        Write-Host "      Τα αρχεία δεν κλειδώνονται από process — ανήκουν στον TrustedInstaller." -ForegroundColor DarkYellow

        # Check Windows Update & TrustedInstaller service status
        $wuStatus = (Get-Service wuauserv -ErrorAction SilentlyContinue).Status
        $tiStatus = (Get-Service TrustedInstaller -ErrorAction SilentlyContinue).Status
        Write-Host "      Windows Update: $wuStatus | TrustedInstaller: $tiStatus" -ForegroundColor Gray
        $aclIssuesFound = $true
    }

} catch {
    Write-Host "   $($ico.ERR) Αδυναμία ανάγνωσης ACL: $($_.Exception.Message)" -ForegroundColor Red
    $aclIssuesFound = $true
}

# =========================================================
# ΜΕΘΟΔΟΣ 4: Direct Access Test & Deep Diagnostics
# Πιάνει: Hard Links, CBS Pending, WRP, I/O Errors
#          Πραγματική δοκιμή πρόσβασης (όχι θεωρητική)
# =========================================================
Write-Host "`n$($ico.N4)  Direct Access Test & Deep Diagnostics..." -ForegroundColor Gray

$directIssuesFound = $false
$diagResults = @()

# --- 4A: Πραγματική Δοκιμή Διαγραφής (σε ένα αρχείο) ---
if ($isFolder) {
    # Βρες ένα αρχείο-δείγμα μέσα στο φάκελο για test
    $testFile = cmd /c "dir /b /s /a-d `"$targetPath`" 2>nul" 2>$null | Select-Object -First 1
} else {
    $testFile = $targetPath
}

if ($testFile) {
    Write-Host "   $($ico.TEST) Δοκιμή πρόσβασης: $([System.IO.Path]::GetFileName($testFile))" -ForegroundColor DarkGray

    # Test 1: Μπορούμε να ανοίξουμε το αρχείο;
    try {
        $stream = [System.IO.File]::Open($testFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $stream.Close()
        Write-Host "   $($ico.OK) File Access: OK (Read/Write)" -ForegroundColor Green
    } catch {
        $errMsg = $_.Exception.InnerException.Message
        if (-not $errMsg) { $errMsg = $_.Exception.Message }
        $hResult = $_.Exception.InnerException.HResult
        if (-not $hResult) { $hResult = $_.Exception.HResult }
        
        Write-Host "   $($ico.DENY) File Access BLOCKED!" -ForegroundColor Red
        Write-Host "      Error: $errMsg" -ForegroundColor DarkYellow
        
        # Αναγνώριση γνωστών HResult codes
        switch ($hResult) {
            -2147024891 { $diagResults += "ACCESS_DENIED (0x80070005) — Permissions ή WRP block" }
            -2147024864 { $diagResults += "SHARING_VIOLATION (0x80070020) — Άλλο process χρησιμοποιεί το αρχείο" }
            -2147024858 { $diagResults += "LOCK_VIOLATION (0x80070021) — Kernel-level file lock" }
            -2147024809 { $diagResults += "INVALID_PARAMETER — Πρόβλημα path ή reparse point" }
            default      { $diagResults += "HResult: 0x$([Convert]::ToString([uint32]$hResult, 16).ToUpper())" }
        }
        $directIssuesFound = $true
    }

    # Test 2: Hard Links Check
    Write-Host "   $($ico.LINK) Έλεγχος Hard Links..." -ForegroundColor DarkGray
    $hlResult = cmd /c "fsutil hardlink list `"$testFile`" 2>&1"
    if ($hlResult -and ($hlResult | Measure-Object).Count -gt 1) {
        $hlCount = ($hlResult | Measure-Object).Count
        Write-Host "   $($ico.WARN)  Hard Links Detected! ($hlCount links)" -ForegroundColor Magenta
        $hlResult | Select-Object -First 3 | ForEach-Object { Write-Host "      → $_" -ForegroundColor DarkYellow }
        if ($hlCount -gt 3) { Write-Host "      ... και $($hlCount - 3) ακόμα" -ForegroundColor DarkGray }
        $diagResults += "HARD_LINKS — Το αρχείο έχει $hlCount hard links (WinSxS component store)"
        $directIssuesFound = $true
    } elseif ($hlResult -match "Access is denied|Error") {
        Write-Host "   $($ico.WARN)  fsutil: $hlResult" -ForegroundColor Red
        $directIssuesFound = $true
    } else {
        Write-Host "   $($ico.OK) Μόνο 1 link (κανονικό αρχείο)" -ForegroundColor Green
    }
} else {
    Write-Host "   (Κανένα αρχείο για test)" -ForegroundColor DarkGray
}

# --- 4B: CBS Pending Transaction Check ---
if ($targetPath -match "WinSxS|InFlight|PendingDeletes") {
    Write-Host "   $($ico.CBS) Έλεγχος CBS Pending Transactions..." -ForegroundColor DarkGray
    
    $pendingXml = "C:\Windows\WinSxS\pending.xml"
    $rebootPending = "C:\Windows\WinSxS\reboot.xml"
    $cbsLogDir = "C:\Windows\Logs\CBS"
    
    if (Test-Path $pendingXml) {
        Write-Host "   $($ico.WARN)  pending.xml FOUND — Υπάρχει ημιτελής CBS transaction!" -ForegroundColor Red
        $diagResults += "CBS_PENDING — Ημιτελές Windows Update operation"
        $directIssuesFound = $true
    }
    if (Test-Path $rebootPending) {
        Write-Host "   $($ico.WARN)  reboot.xml FOUND — Pending reboot required!" -ForegroundColor Red
        $diagResults += "REBOOT_PENDING — Χρειάζεται restart για ολοκλήρωση"
        $directIssuesFound = $true
    }
    
    # Πρόσφατα CBS errors
    if (Test-Path $cbsLogDir) {
        $cbsLog = Get-ChildItem "$cbsLogDir\CBS.log" -ErrorAction SilentlyContinue
        if ($cbsLog) {
            $recentErrors = cmd /c "findstr /i /c:`"Error`" /c:`"Failed`" `"$($cbsLog.FullName)`" 2>nul" | Select-Object -Last 3
            if ($recentErrors) {
                Write-Host "   $($ico.LOG) Πρόσφατα CBS Errors:" -ForegroundColor DarkYellow
                $recentErrors | ForEach-Object { Write-Host "      $_" -ForegroundColor DarkGray }
            }
        }
    }
}

# --- 4C: Reparse Point / Junction Detection ---
try {
    $itemInfo = Get-Item -LiteralPath $targetPath -Force -ErrorAction Stop
    if ($itemInfo.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
        Write-Host "   $($ico.WARN)  REPARSE POINT detected! (Junction/Symlink)" -ForegroundColor Magenta
        $diagResults += "REPARSE_POINT — Symlink ή Junction point"
        $directIssuesFound = $true
    }
} catch { }

# --- Τελική Διάγνωση & Λύσεις ---
if ($directIssuesFound -or $aclIssuesFound) {
    Write-Host "`n   ─── Διάγνωση ───" -ForegroundColor Cyan
    foreach ($diag in $diagResults) {
        Write-Host "   $($ico.DOT) $diag" -ForegroundColor Yellow
    }

    Write-Host "`n   ─── Προτεινόμενη Λύση ───" -ForegroundColor Cyan

    if ($diagResults -match "SHARING_VIOLATION|LOCK_VIOLATION") {
        Write-Host "   $($ico.TIP) Κάποιο process κρατάει το αρχείο ανοιχτό (δεν εμφανίζεται στο handle.exe)" -ForegroundColor Yellow
        Write-Host "   $($ico.TIP) Δοκίμασε: Restart Explorer ή κλείσε εφαρμογές indexing (Everything, Search)" -ForegroundColor White
    }

    if ($diagResults -match "HARD_LINKS") {
        Write-Host "   $($ico.TIP) Τα hard links ανήκουν στο WinSxS Component Store" -ForegroundColor Yellow
        Write-Host "      DISM /Online /Cleanup-Image /StartComponentCleanup" -ForegroundColor White
        Write-Host "      Ή: DISM /Online /Cleanup-Image /StartComponentCleanup /ResetBase" -ForegroundColor DarkGray
    }

    if ($diagResults -match "CBS_PENDING|REBOOT_PENDING") {
        Write-Host "   $($ico.TIP) Κάνε RESTART πρώτα — μετά δοκίμασε ξανά" -ForegroundColor Yellow
        Write-Host "      Αν μετά το restart ο φάκελος υπάρχει ακόμα:" -ForegroundColor DarkGray
        Write-Host "      DISM /Online /Cleanup-Image /RestoreHealth" -ForegroundColor White
    }

    if ($targetPath -match "WinSxS") {
        Write-Host "   $($ico.TIP) Takeown + Unlock (Ξεκλείδωμα):" -ForegroundColor Yellow
        Write-Host "      takeown /F `"$targetPath`" /R /A /D Y" -ForegroundColor White
        Write-Host "      icacls `"$targetPath`" /grant Administrators:F /T /C" -ForegroundColor White
    } elseif (-not ($diagResults -match "SHARING_VIOLATION|LOCK_VIOLATION|HARD_LINKS|CBS_PENDING")) {
        Write-Host "   $($ico.TIP) Τρέξε σε elevated prompt:" -ForegroundColor Green
        Write-Host "      takeown /F `"$targetPath`" /R /A /D Y" -ForegroundColor White
        Write-Host "      icacls `"$targetPath`" /grant Administrators:F /T /C" -ForegroundColor White
    }

    # Προσφορά αυτόματης εκτέλεσης takeown
    $proceedWithFix = $false

    if ($isCriticalPath) {
        # ⛔ EXTRA WARNING για κρίσιμα system paths
        Write-Host ""
        Write-Host "   $($ico.DENY)$($ico.DENY)$($ico.DENY) CRITICAL SYSTEM PATH DETECTED $($ico.DENY)$($ico.DENY)$($ico.DENY)" -ForegroundColor Red
        Write-Host "   Αυτός ο φάκελος είναι κρίσιμος για τη λειτουργία των Windows!" -ForegroundColor Red
        Write-Host "   Η αλλαγή ownership/permissions μπορεί να προκαλέσει:" -ForegroundColor Red
        Write-Host "      → Αδυναμία εκκίνησης (unbootable system)" -ForegroundColor DarkRed
        Write-Host "      → Αποτυχία Windows Update" -ForegroundColor DarkRed
        Write-Host "      → Κατεστραμμένο Component Store" -ForegroundColor DarkRed
        Write-Host "      → Blue Screen of Death (BSOD)" -ForegroundColor DarkRed
        Write-Host ""
        Write-Host "   $($ico.TIP) Αν θες να καθαρίσεις WinSxS, χρησιμοποίησε αντ' αυτού:" -ForegroundColor Yellow
        Write-Host "      DISM /Online /Cleanup-Image /StartComponentCleanup" -ForegroundColor White
        Write-Host ""
        Write-Host "   Γράψε CONFIRM (κεφαλαία) για συνέχεια  |  Enter ή οτιδήποτε άλλο = Ακύρωση" -ForegroundColor Red
        $confirmInput = Read-Host "   Επιλογή"
        if ($confirmInput -ceq "CONFIRM") {
            $proceedWithFix = $true
        } else {
            Write-Host "   Ακυρώθηκε. Σωστή επιλογή!" -ForegroundColor Green
        }
    } else {
        $autoFix = Read-Host "`n      Θέλεις αυτόματο Takeown + Grant Access τώρα; (Y/N)"
        if ($autoFix -eq "Y") { $proceedWithFix = $true }
    }

    if ($proceedWithFix) {
        Write-Host "      $($ico.DOT) Εκτέλεση takeown..." -ForegroundColor Gray
        $tkResult = & takeown.exe /F $targetPath /R /A /D Y 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "      $($ico.OK) Ownership OK" -ForegroundColor Green
        } else {
            Write-Host "      $($ico.WARN) Takeown: $tkResult" -ForegroundColor Red
        }

        Write-Host "      $($ico.DOT) Εκτέλεση icacls..." -ForegroundColor Gray
        $icResult = & icacls.exe $targetPath /grant "Administrators:F" /T /C /Q 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "      $($ico.OK) Permissions Granted — Μπορείς πλέον να το διαχειριστείς χειροκίνητα" -ForegroundColor Green
        } else {
            Write-Host "      $($ico.WARN) Icacls: $icResult" -ForegroundColor Red
        }
    }
} else {
    Write-Host "   $($ico.OK) Direct Access OK — Κανένα deep πρόβλημα" -ForegroundColor Green
    if (-not $aclIssuesFound) {
        Write-Host "`n   $($ico.OK) Ownership & Permissions OK — Κανένα ACL πρόβλημα" -ForegroundColor Green
    }
}

Write-Host ("`n" + ("=" * 40)) -ForegroundColor Gray
Read-CloseOrUpdate

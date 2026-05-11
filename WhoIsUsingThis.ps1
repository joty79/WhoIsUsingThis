param([string]$targetPath)

# 🔸 Force UTF-8 Console Encoding
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 | Out-Null

$script:AppName = 'WhoIsUsingThis'
$script:AppVersion = '1.0.2'
$script:GitHubRepo = 'joty79/WhoIsUsingThis'
$script:MetadataPath = Join-Path $PSScriptRoot 'app-metadata.json'
$script:StatePath = Join-Path $PSScriptRoot 'state'
$script:InstallMetaPath = Join-Path $script:StatePath 'install-meta.json'
$script:UpdateStatusCachePath = Join-Path $script:StatePath 'app-update-status.json'
$script:UpdateStatusCacheTtlMinutes = 30
$script:UpdateStatus = [pscustomobject]@{
    Status = 'Unknown'
    Label = 'Status unavailable'
    LocalVersion = $script:AppVersion
    LatestVersion = ''
    LocalCommit = ''
    LatestCommit = ''
    SourceKind = 'Unknown'
    HasLocalChanges = $false
    Repo = $script:GitHubRepo
    Branch = ''
    Message = 'Update status has not been checked yet.'
    CheckedAt = ''
    Error = ''
}

function Get-OptionalObjectPropertyValue {
    param(
        [object]$InputObject,
        [string]$PropertyName,
        [object]$DefaultValue = $null
    )
    if ($null -eq $InputObject -or [string]::IsNullOrWhiteSpace($PropertyName)) { return $DefaultValue }
    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($null -eq $property) { return $DefaultValue }
    return $property.Value
}

function Read-JsonFile {
    param([Parameter(Mandatory)][string]$Path)
    return (Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop)
}

function Save-JsonFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][object]$InputObject
    )
    $parentPath = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($parentPath) -and -not (Test-Path -LiteralPath $parentPath -PathType Container)) {
        New-Item -Path $parentPath -ItemType Directory -Force | Out-Null
    }
    $InputObject | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Path -Encoding UTF8
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

function Get-ShortGitCommitText {
    param([AllowEmptyString()][string]$Commit)
    if ([string]::IsNullOrWhiteSpace($Commit)) { return '' }
    $normalizedCommit = $Commit.Trim()
    if ($normalizedCommit.Length -le 7) { return $normalizedCommit }
    return $normalizedCommit.Substring(0, 7)
}

function Get-CurrentAppSourceInfo {
    $result = [ordered]@{
        Commit = ''
        SourceKind = 'Portable'
        HasLocalChanges = $false
    }

    if (Test-Path -LiteralPath $script:InstallMetaPath -PathType Leaf) {
        try {
            $installMeta = Read-JsonFile -Path $script:InstallMetaPath
            $commit = [string](Get-OptionalObjectPropertyValue -InputObject $installMeta -PropertyName 'github_commit' -DefaultValue '')
            $result.Commit = $commit.Trim()
            $result.SourceKind = 'Installed'
            return [pscustomobject]$result
        }
        catch {
            $result.SourceKind = 'Installed'
            return [pscustomobject]$result
        }
    }

    if (Get-Command git.exe -ErrorAction SilentlyContinue) {
        try {
            $inside = (& git.exe -C $PSScriptRoot rev-parse --is-inside-work-tree 2>$null | Out-String).Trim()
            if ($inside -eq 'true') {
                $commit = (& git.exe -C $PSScriptRoot rev-parse HEAD 2>$null | Out-String).Trim()
                $dirty = (& git.exe -C $PSScriptRoot status --porcelain 2>$null | Out-String).Trim()
                $result.Commit = $commit
                $result.SourceKind = 'Workspace'
                $result.HasLocalChanges = (-not [string]::IsNullOrWhiteSpace($dirty))
                return [pscustomobject]$result
            }
        }
        catch {}
    }

    return [pscustomobject]$result
}

function Test-LocalGitCommitContainsRemoteCommit {
    param(
        [AllowEmptyString()][string]$RemoteCommit,
        [AllowEmptyString()][string]$LocalCommit
    )
    if (
        [string]::IsNullOrWhiteSpace($RemoteCommit) -or
        [string]::IsNullOrWhiteSpace($LocalCommit) -or
        -not (Get-Command git.exe -ErrorAction SilentlyContinue)
    ) {
        return $false
    }
    try {
        & git.exe -C $PSScriptRoot merge-base --is-ancestor $RemoteCommit $LocalCommit 2>$null
        return ($LASTEXITCODE -eq 0)
    }
    catch { return $false }
}

function New-UpdateStatus {
    param(
        [string]$LocalVersion = $script:AppVersion,
        [string]$Status = 'Unknown',
        [AllowEmptyString()][string]$LatestVersion = '',
        [AllowEmptyString()][string]$LocalCommit = '',
        [AllowEmptyString()][string]$LatestCommit = '',
        [AllowEmptyString()][string]$SourceKind = 'Unknown',
        [bool]$HasLocalChanges = $false,
        [AllowEmptyString()][string]$Repo = $script:GitHubRepo,
        [AllowEmptyString()][string]$Branch = '',
        [string]$Message = 'Update status has not been checked yet.',
        [AllowEmptyString()][string]$CheckedAt = '',
        [AllowEmptyString()][string]$Error = ''
    )

    $label = switch ($Status) {
        'UpToDate' { 'Up to date' }
        'UpdateAvailable' { if ([string]::IsNullOrWhiteSpace($LatestVersion)) { 'Update available' } else { "Update available ($LatestVersion)" } }
        'LocalAhead' { 'Local workspace ahead' }
        'WorkspaceModified' { 'Workspace modified' }
        'Error' { 'Update check failed' }
        default { 'Status unavailable' }
    }

    [pscustomobject]@{
        Status = $Status
        Label = $label
        LocalVersion = $LocalVersion
        LatestVersion = $LatestVersion
        LocalCommit = $LocalCommit
        LatestCommit = $LatestCommit
        SourceKind = $SourceKind
        HasLocalChanges = $HasLocalChanges
        Repo = $Repo
        Branch = $Branch
        Message = $Message
        CheckedAt = $CheckedAt
        Error = $Error
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
        $cache = Read-JsonFile -Path $script:UpdateStatusCachePath
        if (-not $AllowStale -and [string](Get-OptionalObjectPropertyValue -InputObject $cache -PropertyName 'Status' -DefaultValue '') -eq 'UpToDate') {
            return $null
        }

        $localCommit = [string](Get-OptionalObjectPropertyValue -InputObject $cache -PropertyName 'LocalCommit' -DefaultValue '')
        $latestCommit = [string](Get-OptionalObjectPropertyValue -InputObject $cache -PropertyName 'LatestCommit' -DefaultValue '')
        $sourceKind = [string](Get-OptionalObjectPropertyValue -InputObject $cache -PropertyName 'SourceKind' -DefaultValue '')
        if ([string]::IsNullOrWhiteSpace($sourceKind)) { return $null }

        $currentSource = Get-CurrentAppSourceInfo
        if ([string](Get-OptionalObjectPropertyValue -InputObject $cache -PropertyName 'LocalVersion' -DefaultValue '') -ne $script:AppVersion) {
            return $null
        }
        if ([string]$currentSource.SourceKind -ne $sourceKind) {
            return $null
        }
        if ([bool]$currentSource.HasLocalChanges -ne [bool](Get-OptionalObjectPropertyValue -InputObject $cache -PropertyName 'HasLocalChanges' -DefaultValue $false)) {
            return $null
        }
        if (
            -not [string]::IsNullOrWhiteSpace([string]$currentSource.Commit) -and
            -not [string]::IsNullOrWhiteSpace($localCommit) -and
            [string]$currentSource.Commit -ne $localCommit
        ) {
            return $null
        }

        return (New-UpdateStatus `
            -Status ([string](Get-OptionalObjectPropertyValue -InputObject $cache -PropertyName 'Status' -DefaultValue 'Unknown')) `
            -LocalVersion ([string](Get-OptionalObjectPropertyValue -InputObject $cache -PropertyName 'LocalVersion' -DefaultValue $script:AppVersion)) `
            -LatestVersion ([string](Get-OptionalObjectPropertyValue -InputObject $cache -PropertyName 'LatestVersion' -DefaultValue '')) `
            -LocalCommit $localCommit `
            -LatestCommit $latestCommit `
            -SourceKind $sourceKind `
            -HasLocalChanges ([bool](Get-OptionalObjectPropertyValue -InputObject $cache -PropertyName 'HasLocalChanges' -DefaultValue $false)) `
            -Repo ([string](Get-OptionalObjectPropertyValue -InputObject $cache -PropertyName 'Repo' -DefaultValue $script:GitHubRepo)) `
            -Branch ([string](Get-OptionalObjectPropertyValue -InputObject $cache -PropertyName 'Branch' -DefaultValue '')) `
            -Message ([string](Get-OptionalObjectPropertyValue -InputObject $cache -PropertyName 'Message' -DefaultValue '')) `
            -CheckedAt ([string](Get-OptionalObjectPropertyValue -InputObject $cache -PropertyName 'CheckedAt' -DefaultValue '')) `
            -Error ([string](Get-OptionalObjectPropertyValue -InputObject $cache -PropertyName 'Error' -DefaultValue '')))
    }
    catch { return $null }
}

function Write-UpdateStatusCache {
    param([Parameter(Mandatory)]$Status)
    try {
        if (-not (Test-Path -LiteralPath $script:StatePath -PathType Container)) {
            New-Item -Path $script:StatePath -ItemType Directory -Force | Out-Null
        }
        Save-JsonFile -Path $script:UpdateStatusCachePath -InputObject $Status
    }
    catch {}
}

function Get-GitHubApiHeaders {
    $headers = @{ 'User-Agent' = "$($script:AppName)/$($script:AppVersion)" }
    if (-not [string]::IsNullOrWhiteSpace($env:GITHUB_TOKEN)) {
        $headers['Authorization'] = "Bearer $($env:GITHUB_TOKEN)"
    }
    return $headers
}

function ConvertTo-GitHubRepoSlugFromRemoteUrl {
    param([AllowEmptyString()][string]$RemoteUrl)
    if ([string]::IsNullOrWhiteSpace($RemoteUrl)) { return '' }
    $match = [regex]::Match($RemoteUrl.Trim(), 'github\.com[:/](?<owner>[^/]+)/(?<repo>[^/#?]+?)(?:\.git)?(?:[/#?].*)?$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) { return '' }
    return ('{0}/{1}' -f $match.Groups['owner'].Value, $match.Groups['repo'].Value).ToLowerInvariant()
}

function Get-AppGitRemoteTarget {
    param([AllowEmptyString()][string]$Repo = $script:GitHubRepo)
    if ([string]::IsNullOrWhiteSpace($Repo) -or -not (Get-Command git.exe -ErrorAction SilentlyContinue)) {
        return ''
    }
    $expectedRepo = $Repo.Trim().ToLowerInvariant()
    try {
        $inside = (& git.exe -C $PSScriptRoot rev-parse --is-inside-work-tree 2>$null | Out-String).Trim()
        if ($inside -eq 'true') {
            foreach ($remoteName in @(& git.exe -C $PSScriptRoot remote 2>$null)) {
                $name = [string]$remoteName
                if ([string]::IsNullOrWhiteSpace($name)) { continue }
                $remoteUrl = (& git.exe -C $PSScriptRoot remote get-url $name 2>$null | Out-String).Trim()
                if ((ConvertTo-GitHubRepoSlugFromRemoteUrl -RemoteUrl $remoteUrl) -eq $expectedRepo) {
                    return $name.Trim()
                }
            }
        }
    }
    catch {}
    return ("https://github.com/{0}.git" -f $Repo.Trim())
}

function Resolve-RemoteAppCommit {
    param(
        [AllowEmptyString()][string]$Repo = $script:GitHubRepo,
        [AllowEmptyString()][string]$Ref = ''
    )
    if ([string]::IsNullOrWhiteSpace($Repo) -or [string]::IsNullOrWhiteSpace($Ref)) { return '' }
    if (Get-Command gh.exe -ErrorAction SilentlyContinue) {
        try {
            $commit = (& gh.exe api "repos/$Repo/commits/$Ref" --jq '.sha' 2>$null | Out-String).Trim()
            if (-not [string]::IsNullOrWhiteSpace($commit)) { return $commit }
        }
        catch {}
    }
    try {
        $commitInfo = Invoke-RestMethod -Uri ("https://api.github.com/repos/{0}/commits/{1}" -f $Repo, $Ref) -Headers (Get-GitHubApiHeaders) -TimeoutSec 5 -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace([string]$commitInfo.sha)) { return [string]$commitInfo.sha }
    }
    catch {}
    $gitRemoteTarget = Get-AppGitRemoteTarget -Repo $Repo
    if (-not [string]::IsNullOrWhiteSpace($gitRemoteTarget) -and (Get-Command git.exe -ErrorAction SilentlyContinue)) {
        foreach ($candidateRef in @("refs/heads/$Ref", $Ref)) {
            try {
                $remoteLine = (& git.exe -C $PSScriptRoot ls-remote $gitRemoteTarget $candidateRef 2>$null | Select-Object -First 1 | Out-String).Trim()
                if (-not [string]::IsNullOrWhiteSpace($remoteLine)) {
                    $commit = ($remoteLine -split '\s+')[0]
                    if (-not [string]::IsNullOrWhiteSpace($commit)) { return $commit }
                }
            }
            catch {}
        }
    }
    return ''
}

function Get-RemoteAppMetadataFromGit {
    param(
        [AllowEmptyString()][string]$Repo = $script:GitHubRepo,
        [Parameter(Mandatory)][System.Collections.Generic.List[string]]$BranchCandidates,
        [Parameter(Mandatory)][string]$MetadataRelativePath
    )
    if ([string]::IsNullOrWhiteSpace($Repo) -or -not (Get-Command git.exe -ErrorAction SilentlyContinue)) { return $null }
    $gitRemoteTarget = Get-AppGitRemoteTarget -Repo $Repo
    if ([string]::IsNullOrWhiteSpace($gitRemoteTarget)) { return $null }

    foreach ($branch in $BranchCandidates) {
        try {
            $remoteLine = (& git.exe -C $PSScriptRoot ls-remote $gitRemoteTarget "refs/heads/$branch" 2>$null | Select-Object -First 1 | Out-String).Trim()
            if ([string]::IsNullOrWhiteSpace($remoteLine)) { continue }
            $latestCommit = ($remoteLine -split '\s+')[0]
            $metadata = $null
            try {
                $metadataJson = (& git.exe -C $PSScriptRoot show "$($latestCommit):$MetadataRelativePath" 2>$null | Out-String).Trim()
                if (-not [string]::IsNullOrWhiteSpace($metadataJson)) {
                    $metadata = $metadataJson | ConvertFrom-Json
                }
            }
            catch {}
            if ($null -eq $metadata) {
                $tempRoot = Join-Path $env:TEMP ("WhoIsUsingThis_update_metadata_{0}" -f [guid]::NewGuid().ToString('N'))
                try {
                    & git.exe clone --quiet --depth 1 --branch $branch $gitRemoteTarget $tempRoot 2>$null
                    if ($LASTEXITCODE -eq 0) {
                        $metadataPath = Join-Path $tempRoot ($MetadataRelativePath -replace '/', [System.IO.Path]::DirectorySeparatorChar)
                        if (Test-Path -LiteralPath $metadataPath -PathType Leaf) {
                            $metadata = Read-JsonFile -Path $metadataPath
                        }
                    }
                }
                catch {}
                finally {
                    if (Test-Path -LiteralPath $tempRoot) {
                        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
                    }
                }
            }
            if ($null -eq $metadata) { $metadata = [pscustomobject]@{ version = '' } }
            return [pscustomobject]@{ Repo = $Repo; Branch = $branch; Commit = $latestCommit; Metadata = $metadata }
        }
        catch {}
    }
    return $null
}

function Get-RemoteAppMetadata {
    if ([string]::IsNullOrWhiteSpace($script:GitHubRepo)) { return $null }
    $headers = Get-GitHubApiHeaders
    $defaultBranch = ''
    if (Get-Command gh.exe -ErrorAction SilentlyContinue) {
        try {
            $repoJson = (& gh.exe api "repos/$script:GitHubRepo" 2>$null | Out-String).Trim()
            if (-not [string]::IsNullOrWhiteSpace($repoJson)) {
                $defaultBranch = [string](($repoJson | ConvertFrom-Json).default_branch)
            }
        }
        catch {}
    }
    try {
        if ([string]::IsNullOrWhiteSpace($defaultBranch)) {
            $repoInfo = Invoke-RestMethod -Uri ("https://api.github.com/repos/{0}" -f $script:GitHubRepo) -Headers $headers -TimeoutSec 5 -ErrorAction Stop
            $defaultBranch = [string]$repoInfo.default_branch
        }
    }
    catch {}

    $metadataRelativePath = ($script:MetadataPath.Substring($PSScriptRoot.Length).TrimStart('\')).Replace('\', '/')
    $branchCandidates = [System.Collections.Generic.List[string]]::new()
    foreach ($candidate in @($defaultBranch, 'master', 'main', 'latest')) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and -not $branchCandidates.Contains($candidate)) {
            $branchCandidates.Add($candidate)
        }
    }

    foreach ($branch in $branchCandidates) {
        if (Get-Command gh.exe -ErrorAction SilentlyContinue) {
            try {
                $contentJson = (& gh.exe api ("repos/{0}/contents/{1}?ref={2}" -f $script:GitHubRepo, $metadataRelativePath, $branch) 2>$null | Out-String).Trim()
                if (-not [string]::IsNullOrWhiteSpace($contentJson)) {
                    $contentInfo = $contentJson | ConvertFrom-Json
                    $encodedContent = [string]$contentInfo.content
                    if (-not [string]::IsNullOrWhiteSpace($encodedContent)) {
                        $metadata = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String(($encodedContent -replace '\s', ''))) | ConvertFrom-Json
                        return [pscustomobject]@{ Repo = $script:GitHubRepo; Branch = $branch; Commit = (Resolve-RemoteAppCommit -Ref $branch); Metadata = $metadata }
                    }
                }
            }
            catch {}
        }
        try {
            $metadata = Invoke-RestMethod -Uri "https://raw.githubusercontent.com/$($script:GitHubRepo)/$branch/$metadataRelativePath" -Method Get -Headers $headers -TimeoutSec 8 -ErrorAction Stop
            return [pscustomobject]@{ Repo = $script:GitHubRepo; Branch = $branch; Commit = (Resolve-RemoteAppCommit -Ref $branch); Metadata = $metadata }
        }
        catch {}
    }
    return (Get-RemoteAppMetadataFromGit -Repo $script:GitHubRepo -BranchCandidates $branchCandidates -MetadataRelativePath $metadataRelativePath)
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
        if ($null -ne $stale -and [string]$stale.Status -ne 'UpToDate') {
            $stale.Message = 'Using cached update status because GitHub could not be reached.'
            $script:UpdateStatus = $stale
            return $script:UpdateStatus
        }
        $currentSource = Get-CurrentAppSourceInfo
        $script:UpdateStatus = New-UpdateStatus -Status 'Error' -LocalVersion $script:AppVersion -LocalCommit ([string]$currentSource.Commit) -SourceKind ([string]$currentSource.SourceKind) -HasLocalChanges ([bool]$currentSource.HasLocalChanges) -Repo $script:GitHubRepo -Message 'Could not reach GitHub to check updates.' -CheckedAt ((Get-Date).ToString('s'))
        return $script:UpdateStatus
    }

    $latestVersion = if ($remote.Metadata.PSObject.Properties['version']) { [string]$remote.Metadata.version } else { '' }
    $localVersion = ConvertTo-AppVersion -VersionText $script:AppVersion
    $remoteVersion = ConvertTo-AppVersion -VersionText $latestVersion
    $sourceInfo = Get-CurrentAppSourceInfo
    $localCommit = [string]$sourceInfo.Commit
    $latestCommit = [string]$remote.Commit
    $sourceKind = [string]$sourceInfo.SourceKind
    $hasLocalChanges = [bool]$sourceInfo.HasLocalChanges
    $statusName = 'Unknown'
    $message = 'Update status is unavailable.'

    if ($sourceKind -eq 'Workspace' -and $hasLocalChanges) {
        $statusName = 'WorkspaceModified'
        $message = "This workspace has unpublished local changes. Local v$script:AppVersion at $(Get-ShortGitCommitText -Commit $localCommit); latest GitHub $($remote.Branch) is v$latestVersion at $(Get-ShortGitCommitText -Commit $latestCommit)."
    }
    elseif ($sourceKind -eq 'Workspace' -and $localCommit -ne $latestCommit -and (Test-LocalGitCommitContainsRemoteCommit -RemoteCommit $latestCommit -LocalCommit $localCommit)) {
        $statusName = 'LocalAhead'
        $message = "This workspace has local commits ahead of GitHub $($remote.Branch). Latest published commit is $(Get-ShortGitCommitText -Commit $latestCommit); local HEAD is $(Get-ShortGitCommitText -Commit $localCommit)."
    }
    elseif ($null -ne $localVersion -and $null -ne $remoteVersion) {
        if ($localVersion -lt $remoteVersion) {
            $statusName = 'UpdateAvailable'
            $message = "Update available from GitHub $($remote.Branch): v$latestVersion."
        }
        elseif ($localVersion -gt $remoteVersion) {
            $statusName = 'LocalAhead'
            $message = "Local version v$script:AppVersion is newer than GitHub $($remote.Branch) v$latestVersion."
        }
        elseif (
            -not [string]::IsNullOrWhiteSpace($localCommit) -and
            -not [string]::IsNullOrWhiteSpace($latestCommit) -and
            $localCommit -ne $latestCommit
        ) {
            $statusName = 'UpdateAvailable'
            $message = "Update available from GitHub $($remote.Branch): same version, newer commit $(Get-ShortGitCommitText -Commit $latestCommit)."
        }
        else {
            $statusName = 'UpToDate'
            $commitLabel = Get-ShortGitCommitText -Commit $latestCommit
            $message = if ([string]::IsNullOrWhiteSpace($commitLabel)) { "App is up to date with GitHub $($remote.Branch) at v$latestVersion." } else { "App is up to date with GitHub $($remote.Branch) at v$latestVersion ($commitLabel)." }
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($latestVersion) -and $latestVersion -eq $script:AppVersion) {
        if (
            -not [string]::IsNullOrWhiteSpace($localCommit) -and
            -not [string]::IsNullOrWhiteSpace($latestCommit) -and
            $localCommit -ne $latestCommit
        ) {
            $statusName = 'UpdateAvailable'
            $message = "Update available from GitHub $($remote.Branch): same version, newer commit $(Get-ShortGitCommitText -Commit $latestCommit)."
        }
        else {
            $statusName = 'UpToDate'
            $message = "App is up to date with GitHub $($remote.Branch) at v$latestVersion."
        }
    }
    elseif (
        [string]::IsNullOrWhiteSpace($latestVersion) -and
        -not [string]::IsNullOrWhiteSpace($localCommit) -and
        -not [string]::IsNullOrWhiteSpace($latestCommit)
    ) {
        if ($localCommit -ne $latestCommit) {
            $statusName = 'UpdateAvailable'
            $message = "Update available from GitHub $($remote.Branch): latest commit is $(Get-ShortGitCommitText -Commit $latestCommit); local is $(Get-ShortGitCommitText -Commit $localCommit)."
        }
        else {
            $statusName = 'UpToDate'
            $message = "App is up to date with GitHub $($remote.Branch) at commit $(Get-ShortGitCommitText -Commit $latestCommit)."
        }
    }

    $script:UpdateStatus = New-UpdateStatus -Status $statusName -LocalVersion $script:AppVersion -LatestVersion $latestVersion -LocalCommit $localCommit -LatestCommit $latestCommit -SourceKind $sourceKind -HasLocalChanges $hasLocalChanges -Repo ([string]$remote.Repo) -Branch ([string]$remote.Branch) -Message $message -CheckedAt ((Get-Date).ToString('s'))
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
        $keyName = switch ([int]$key.VirtualKeyCode) {
            13 { 'Enter' }
            27 { 'Escape' }
            38 { 'UpArrow' }
            40 { 'DownArrow' }
            default { [string]$key.VirtualKeyCode }
        }
        return [pscustomobject]@{
            Key = $keyName
            VirtualKeyCode = [int]$key.VirtualKeyCode
            Character = [char]$key.Character
        }
    }
    catch {
        $key = [Console]::ReadKey($true)
        $keyName = switch ($key.Key) {
            'Enter' { 'Enter' }
            'Escape' { 'Escape' }
            'UpArrow' { 'UpArrow' }
            'DownArrow' { 'DownArrow' }
            default { [string]$key.Key }
        }
        return [pscustomobject]@{
            Key = $keyName
            VirtualKeyCode = [int]$key.Key
            Character = [char]$key.KeyChar
        }
    }
}

function Clear-ConsoleInputBuffer {
    try {
        while ([Console]::KeyAvailable) {
            [void][Console]::ReadKey($true)
        }
    }
    catch {}
}

function Set-CursorVisibleSafe {
    param([bool]$Visible)
    try { [Console]::CursorVisible = $Visible }
    catch {}
}

function Write-MenuOption {
    param(
        [string]$Text,
        [string]$Color = 'White',
        [bool]$Selected = $false
    )

    if ($Selected) {
        Write-Host "  ❯ $Text" -ForegroundColor White -BackgroundColor DarkBlue
        return
    }

    Write-Host "    $Text" -ForegroundColor $Color
}

function Show-ArrowMenu {
    param(
        [string]$Title,
        [object[]]$Options,
        [int]$SelectedIndex = 0,
        [string]$HelpText = '↑↓ navigate   Enter = select   Esc = skip'
    )

    Write-AppSection -Title $Title
    for ($index = 0; $index -lt $Options.Count; $index++) {
        $option = $Options[$index]
        $color = if ($option.PSObject.Properties['Color']) { [string]$option.Color } else { 'White' }
        Write-MenuOption -Text ([string]$option.Label) -Color $color -Selected ($index -eq $SelectedIndex)
    }
    Write-Host ''
    Write-Host "  $HelpText" -ForegroundColor DarkGray
}

function Read-ArrowMenuChoice {
    param(
        [string]$Title,
        [object[]]$Options,
        [string]$HelpText = '↑↓ navigate   Enter = select   Esc = skip'
    )

    $selectedIndex = 0
    Set-CursorVisibleSafe -Visible $false
    try {
        while ($true) {
            $menuTop = [Console]::CursorTop
            Show-ArrowMenu -Title $Title -Options $Options -SelectedIndex $selectedIndex -HelpText $HelpText
            Write-Host "$([char]27)[J" -NoNewline

            $key = Read-AppConsoleKey
            switch ($key.Key) {
                'UpArrow' { $selectedIndex = [Math]::Max(0, $selectedIndex - 1) }
                'DownArrow' { $selectedIndex = [Math]::Min($Options.Count - 1, $selectedIndex + 1) }
                'Escape' { return $null }
                'Enter' { return $Options[$selectedIndex] }
            }

            try { [Console]::SetCursorPosition(0, $menuTop) }
            catch {}
        }
    }
    finally {
        Set-CursorVisibleSafe -Visible $true
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

function Format-AppHeaderCell {
    param([AllowEmptyString()][string]$Text, [int]$Width = 77)
    $value = if ($null -eq $Text) { '' } else { [string]$Text }
    if ($value.Length -gt $Width) {
        return $value.Substring(0, [Math]::Max(0, $Width - 3)) + '...'
    }
    return $value.PadRight($Width)
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
    Write-Host (Format-AppHeaderCell -Text ("{0} v{1}" -f $script:AppName, $script:AppVersion)) -ForegroundColor White -NoNewline
    Write-Host "║" -ForegroundColor Cyan
    Write-Host '║ ' -ForegroundColor Cyan -NoNewline
    Write-Host (Format-AppHeaderCell -Text $Subtitle) -ForegroundColor DarkGray -NoNewline
    Write-Host "║" -ForegroundColor Cyan
    Write-Host '║ ' -ForegroundColor Cyan -NoNewline
    Write-Host 'Update: ' -ForegroundColor Green -NoNewline
    Write-Host (Format-AppHeaderCell -Text $script:UpdateStatus.Label -Width 69) -ForegroundColor $updateColor -NoNewline
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

function Get-AppUpdateTargetInfo {
    $result = [ordered]@{
        Mode = 'Portable copy'
        Action = 'DownloadLatest'
        TargetLabel = 'Portable copy'
        MethodLabel = 'DownloadLatest in-place update'
        Branch = ''
        IsGitWorkingCopy = $false
        RefusesDirty = $false
    }

    $defaultInstallPath = [System.IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'WhoIsUsingThisContext')).TrimEnd('\')
    $currentRootPath = [System.IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\')
    if ($currentRootPath -ieq $defaultInstallPath -or (Test-Path -LiteralPath $script:InstallMetaPath -PathType Leaf)) {
        $result.Mode = 'Installed copy'
        $result.Action = 'UpdateGitHub'
        $result.TargetLabel = 'Installed app copy'
        $result.MethodLabel = 'Installer/GitHub in-place update'
        try {
            $installMeta = Read-JsonFile -Path $script:InstallMetaPath
            $githubRef = [string](Get-OptionalObjectPropertyValue -InputObject $installMeta -PropertyName 'github_ref' -DefaultValue '')
            if (-not [string]::IsNullOrWhiteSpace($githubRef)) {
                $result.Branch = $githubRef.Trim()
            }
        }
        catch {}
        return [pscustomobject]$result
    }

    if (Get-Command git.exe -ErrorAction SilentlyContinue) {
        try {
            $inside = (& git.exe -C $PSScriptRoot rev-parse --is-inside-work-tree 2>$null | Out-String).Trim()
            if ($inside -eq 'true') {
                $branch = (& git.exe -C $PSScriptRoot branch --show-current 2>$null | Out-String).Trim()
                if ([string]::IsNullOrWhiteSpace($branch)) {
                    $branch = (& git.exe -C $PSScriptRoot rev-parse --abbrev-ref HEAD 2>$null | Out-String).Trim()
                    if ($branch -eq 'HEAD') { $branch = '' }
                }
                $result.Mode = 'Git repo working copy'
                $result.Action = 'GitFastForward'
                $result.TargetLabel = 'Git repo working copy'
                $result.MethodLabel = 'git fetch + fast-forward only'
                $result.Branch = $branch
                $result.IsGitWorkingCopy = $true
                $result.RefusesDirty = $true
                return [pscustomobject]$result
            }
        }
        catch {}
    }

    return [pscustomobject]$result
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
        $updateTarget = Get-AppUpdateTargetInfo
        $targetLabel = [string]$updateTarget.TargetLabel
        $methodLabel = [string]$updateTarget.MethodLabel
        $latestVersionLabel = if ([string]::IsNullOrWhiteSpace($script:UpdateStatus.LatestVersion)) { '--' } else { $script:UpdateStatus.LatestVersion }
        $localCommitLabel = if ([string]::IsNullOrWhiteSpace($script:UpdateStatus.LocalCommit)) { '--' } else { Get-ShortGitCommitText -Commit $script:UpdateStatus.LocalCommit }
        $latestCommitLabel = if ([string]::IsNullOrWhiteSpace($script:UpdateStatus.LatestCommit)) { '--' } else { Get-ShortGitCommitText -Commit $script:UpdateStatus.LatestCommit }
        $sourceLabel = if ([bool]$script:UpdateStatus.HasLocalChanges) { "$($script:UpdateStatus.SourceKind) + local changes" } else { [string]$script:UpdateStatus.SourceKind }
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
        Write-Host '  Local commit:    ' -ForegroundColor Gray -NoNewline
        Write-Host $localCommitLabel -ForegroundColor White
        Write-Host '  Latest commit:   ' -ForegroundColor Gray -NoNewline
        Write-Host $latestCommitLabel -ForegroundColor White
        Write-Host '  Source:          ' -ForegroundColor Gray -NoNewline
        Write-Host $sourceLabel -ForegroundColor White
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
        if ([bool]$updateTarget.RefusesDirty) {
            Write-Host '  Safety:          ' -ForegroundColor Gray -NoNewline
            Write-Host 'dirty workspaces are refused' -ForegroundColor Yellow
        }
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

function Invoke-GitWorkingCopyUpdate {
    param([AllowEmptyString()][string]$Branch)

    $recentLines = [System.Collections.Generic.List[string]]::new()
    function Add-RecentLine {
        param([AllowEmptyString()][string]$Line)
        if ([string]::IsNullOrWhiteSpace($Line)) { return }
        [void]$recentLines.Add($Line)
        while ($recentLines.Count -gt 12) { $recentLines.RemoveAt(0) }
    }

    if (-not (Get-Command git.exe -ErrorAction SilentlyContinue)) {
        Add-RecentLine 'git.exe was not found in PATH.'
        return [pscustomobject]@{ ExitCode = 9001; RecentLines = @($recentLines) }
    }

    try {
        $inside = (& git.exe -C $PSScriptRoot rev-parse --is-inside-work-tree 2>&1 | Out-String).Trim()
        if ($inside -ne 'true') {
            Add-RecentLine 'This folder is not a git working copy.'
            return [pscustomobject]@{ ExitCode = 9002; RecentLines = @($recentLines) }
        }

        $dirty = (& git.exe -C $PSScriptRoot status --porcelain 2>&1 | Out-String).Trim()
        if (-not [string]::IsNullOrWhiteSpace($dirty)) {
            Add-RecentLine 'Working copy has local changes. Fast-forward update refused.'
            Add-RecentLine 'Commit, stash, or discard local changes before updating this repo copy.'
            return [pscustomobject]@{ ExitCode = 3; RecentLines = @($recentLines) }
        }

        if ([string]::IsNullOrWhiteSpace($Branch)) {
            $Branch = (& git.exe -C $PSScriptRoot branch --show-current 2>&1 | Out-String).Trim()
        }
        if ([string]::IsNullOrWhiteSpace($Branch)) {
            Add-RecentLine 'Could not determine the current git branch.'
            return [pscustomobject]@{ ExitCode = 9003; RecentLines = @($recentLines) }
        }

        Add-RecentLine ("Fetching origin/{0}..." -f $Branch)
        $fetchText = (& git.exe -C $PSScriptRoot fetch --prune origin $Branch 2>&1 | Out-String).Trim()
        foreach ($line in ($fetchText -split "`r?`n")) { Add-RecentLine $line }
        if ($LASTEXITCODE -ne 0) {
            return [pscustomobject]@{ ExitCode = $LASTEXITCODE; RecentLines = @($recentLines) }
        }

        $localHead = (& git.exe -C $PSScriptRoot rev-parse HEAD 2>&1 | Out-String).Trim()
        $remoteHead = (& git.exe -C $PSScriptRoot rev-parse "origin/$Branch" 2>&1 | Out-String).Trim()
        if (-not [string]::IsNullOrWhiteSpace($localHead) -and $localHead -eq $remoteHead) {
            Add-RecentLine ("Already up to date with origin/{0}." -f $Branch)
            return [pscustomobject]@{ ExitCode = 0; RecentLines = @($recentLines) }
        }

        Add-RecentLine ("Fast-forwarding to origin/{0}..." -f $Branch)
        $mergeText = (& git.exe -C $PSScriptRoot merge --ff-only "origin/$Branch" 2>&1 | Out-String).Trim()
        foreach ($line in ($mergeText -split "`r?`n")) { Add-RecentLine $line }
        return [pscustomobject]@{ ExitCode = $LASTEXITCODE; RecentLines = @($recentLines) }
    }
    catch {
        Add-RecentLine ("Git update failed: {0}" -f $_.Exception.Message)
        return [pscustomobject]@{ ExitCode = 9999; RecentLines = @($recentLines) }
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
    $updateTarget = Get-AppUpdateTargetInfo
    $action = [string]$updateTarget.Action
    if ($action -eq 'GitFastForward') {
        $progressMessage = 'Updating this git working copy with fetch + fast-forward...'
        Show-AppUpdateResultPanel -Message $progressMessage -Level 'Info' -RecentLines @("Branch: $($updateTarget.Branch)")
        $gitResult = Invoke-GitWorkingCopyUpdate -Branch ([string]$updateTarget.Branch)
        $exitCode = [int]$gitResult.ExitCode
        $finalLines = @($gitResult.RecentLines)
        if ($exitCode -eq 0) {
            Show-AppUpdateResultPanel -Message 'Update finished. Restarting the updated app host and closing this window...' -Level 'Good' -RecentLines $finalLines -AutoRestart
            Start-Sleep -Milliseconds 900
            if (Start-UpdatedAppHost) { exit 0 }
            Show-AppUpdateResultPanel -Message 'Update finished, but the app could not relaunch automatically.' -Level 'Warn' -RecentLines $finalLines
            return $false
        }
        Show-AppUpdateResultPanel -Message ("Update failed with exit code {0}." -f $exitCode) -Level 'Error' -RecentLines $finalLines
        return $false
    }

    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $installerPath, '-Action', $action, '-Force')
    if ($action -eq 'UpdateGitHub') {
        $arguments += '-NoExplorerRestart'
        if (-not [string]::IsNullOrWhiteSpace([string]$updateTarget.Branch)) {
            $arguments += @('-GitHubRef', [string]$updateTarget.Branch)
        }
    }
    if ($action -eq 'DownloadLatest') { $arguments += '-NoSelfRelaunch' }
    $progressMessage = if ($action -eq 'UpdateGitHub') { 'Updating from GitHub inside the current app session...' } else { 'Updating this portable copy inside the current app session...' }
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

function Read-ScanActionChoice {
    $options = @(
        [pscustomobject]@{ Label = 'Terminate all'; Action = 'TerminateAll'; Color = 'Red' },
        [pscustomobject]@{ Label = 'Choose one-by-one'; Action = 'ChooseOneByOne'; Color = 'Cyan' },
        [pscustomobject]@{ Label = 'Skip'; Action = 'Skip'; Color = 'DarkGray' },
        [pscustomobject]@{ Label = 'Update App'; Action = 'UpdateApp'; Color = 'Yellow' }
    )

    $choice = Read-ArrowMenuChoice -Title 'Actions' -Options $options -HelpText '↑↓ navigate   Enter = select   Esc = skip'
    if ($null -eq $choice) { return 'Skip' }
    if ([string]$choice.Action -eq 'UpdateApp') {
        [void](Show-AppUpdateMenu)
        Show-AppHeader
        Write-Host "`n--- Scanning: $targetPath ---" -ForegroundColor Cyan
        Write-Host "   Update App closed. Continuing scan with this step skipped." -ForegroundColor DarkGray
        return 'Skip'
    }
    return [string]$choice.Action
}

function Stop-LockingProcess {
    param(
        [Parameter(Mandatory)][object]$ProcessInfo,
        [switch]$AllowTerminal
    )

    if ($ProcessInfo.PSObject.Properties['IsTerminal'] -and [bool]$ProcessInfo.IsTerminal -and -not $AllowTerminal) {
        Write-Host "   $($ico.WARN) Skipped terminal: $($ProcessInfo.Name) (PID: $($ProcessInfo.PID))" -ForegroundColor Magenta
        return $false
    }

    try {
        Stop-Process -Id ([int]$ProcessInfo.PID) -Force -ErrorAction Stop
        Write-Host "   $($ico.KILL) Terminated: $($ProcessInfo.Name) (PID: $($ProcessInfo.PID))" -ForegroundColor Red
        return $true
    }
    catch {
        Write-Host "   $($ico.ERR) Failed: $($ProcessInfo.Name) (PID: $($ProcessInfo.PID))" -ForegroundColor Red
        return $false
    }
}

function Invoke-TerminateAllProcesses {
    param([object[]]$ProcessList)

    $terminalCount = 0
    foreach ($proc in @($ProcessList)) {
        if ($proc.PSObject.Properties['IsTerminal'] -and [bool]$proc.IsTerminal) {
            $terminalCount++
        }
        [void](Stop-LockingProcess -ProcessInfo $proc)
    }
    if ($terminalCount -gt 0) {
        Write-Host "   $($ico.TIP) $terminalCount terminal process(es) skipped. Use Choose one-by-one if you really want to terminate them." -ForegroundColor DarkYellow
    }
}

function Invoke-ChooseProcessMenu {
    param(
        [string]$Title,
        [object[]]$ProcessList
    )

    $remaining = [System.Collections.Generic.List[object]]::new()
    foreach ($proc in @($ProcessList)) { [void]$remaining.Add($proc) }

    while ($remaining.Count -gt 0) {
        $options = @(
            for ($index = 0; $index -lt $remaining.Count; $index++) {
                $proc = $remaining[$index]
                $terminalLabel = if ($proc.PSObject.Properties['IsTerminal'] -and [bool]$proc.IsTerminal) { ' [terminal]' } else { '' }
                $lockedLabel = if ($proc.PSObject.Properties['LockedFile'] -and -not [string]::IsNullOrWhiteSpace([string]$proc.LockedFile)) { " — $($proc.LockedFile)" } else { '' }
                [pscustomobject]@{
                    Label = "$($proc.Name) (PID: $($proc.PID))$terminalLabel$lockedLabel"
                    Index = $index
                    Color = if ($terminalLabel) { 'Magenta' } else { 'Yellow' }
                }
            }
        )

        $choice = Read-ArrowMenuChoice -Title $Title -Options $options -HelpText '↑↓ navigate   Enter = terminate selected   Esc = skip remaining'
        if ($null -eq $choice) { return }

        $selectedIndex = [int]$choice.Index
        $selectedProcess = $remaining[$selectedIndex]
        [void](Stop-LockingProcess -ProcessInfo $selectedProcess -AllowTerminal)
        $remaining.RemoveAt($selectedIndex)
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
# Critical System Paths — extra warning before takeown/delete
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
# Method 1: Handle.exe (standard locks)
# =========================================================
Write-Host "$($ico.N1)  Checking handles (standard locks)..." -ForegroundColor Gray

# Exact search only.
$handleResults = & $handleExePath -a -u "$targetPath" 2>$null

if ($handleResults -match "pid:") {
    $processedPids = @()
    $lockList = @()

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

    $termCount = ($lockList | Where-Object IsTerminal).Count
    $appCount  = $lockList.Count - $termCount
    Write-Host "   Found $($lockList.Count) process(es) ($appCount apps, $termCount terminals):" -ForegroundColor White
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

    Write-Host ""
    $handleChoice = Read-ScanActionChoice

    switch ($handleChoice) {
        "TerminateAll" { Invoke-TerminateAllProcesses -ProcessList $lockList }
        "ChooseOneByOne" { Invoke-ChooseProcessMenu -Title 'Choose Process' -ProcessList $lockList }
        default {
            Write-Host "   Skipped." -ForegroundColor DarkGray
        }
    }
} else {
    Write-Host "   (No standard file handle found)" -ForegroundColor DarkGray
}

# =========================================================
# Method 2: Module/deep scan (DLLs and libraries)
# =========================================================
if (-not $isFolder) {
    Write-Host "`n$($ico.N2)  Checking loaded modules (DLLs)..." -ForegroundColor Gray
    
    # Use Regex Escape for file names with [ ]
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
        Write-Host "   Found $($modList.Count) process(es) with loaded module:" -ForegroundColor White
        Write-Host ""
        $idx = 0
        foreach ($m in $modList) {
            $idx++
            Write-Host "   [$idx] $($ico.HIT) $($m.Name) (PID: $($m.PID))" -ForegroundColor Yellow
        }
        Write-Host ""
        $modChoice = Read-ScanActionChoice

        switch ($modChoice) {
            "TerminateAll" { Invoke-TerminateAllProcesses -ProcessList $modList }
            "ChooseOneByOne" { Invoke-ChooseProcessMenu -Title 'Choose Process' -ProcessList $modList }
            default { Write-Host "   Skipped." -ForegroundColor DarkGray }
        }
    } else { Write-Host "   (No loaded module found)" -ForegroundColor DarkGray }

} else {
    Write-Host "`n$($ico.N2)  Deep folder scan (searching all processes)..." -ForegroundColor Gray
    Write-Host "   $($ico.WAIT) Scanning process memory..." -ForegroundColor DarkGray
    
    $deepList = @()
    $processes = Get-Process -ErrorAction SilentlyContinue
    
    foreach ($proc in $processes) {
        try {
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
        Write-Host "   Found $($deepList.Count) process(es) with locked modules:" -ForegroundColor White
        Write-Host ""
        $idx = 0
        foreach ($d in $deepList) {
            $idx++
            Write-Host "   [$idx] $($ico.HIT) $($d.Name) (PID: $($d.PID))" -ForegroundColor Yellow
            Write-Host "       Locked: $($d.LockedFile)" -ForegroundColor DarkGray
        }
        Write-Host ""
        $deepChoice = Read-ScanActionChoice

        switch ($deepChoice) {
            "TerminateAll" { Invoke-TerminateAllProcesses -ProcessList $deepList }
            "ChooseOneByOne" { Invoke-ChooseProcessMenu -Title 'Choose Process' -ProcessList $deepList }
            default { Write-Host "   Skipped." -ForegroundColor DarkGray }
        }
    } else { Write-Host "   (No locked module found under this folder)" -ForegroundColor DarkGray }
}

# =========================================================
# Method 3: ACL & Ownership Analysis
# Catches: TrustedInstaller, SYSTEM, ACL-blocked files,
#          CBS InFlight (WinSxS pending updates)
# =========================================================
Write-Host "`n$($ico.N3)  Checking ownership and ACL permissions..." -ForegroundColor Gray

$aclIssuesFound = $false

try {
    $acl = Get-Acl -LiteralPath $targetPath -ErrorAction Stop
    $owner = $acl.Owner

    # --- Ownership Check ---
    $isSystemOwned = $owner -match "TrustedInstaller|NT SERVICE|SYSTEM|S-1-5-80"
    $isAdminOwned  = $owner -match "BUILTIN\\Administrators|Administrators"

    if ($isSystemOwned) {
        Write-Host "   $($ico.LOCK) Owner: $owner" -ForegroundColor Red
        Write-Host "      This path belongs to Windows (TrustedInstaller/SYSTEM)" -ForegroundColor DarkYellow
        $aclIssuesFound = $true
    } elseif ($isAdminOwned) {
        Write-Host "   $($ico.OK) Owner: $owner (Administrators)" -ForegroundColor Green
    } else {
        Write-Host "   $($ico.INFO)  Owner: $owner" -ForegroundColor Cyan
    }

    # --- Access Rights Check (can we really delete it?) ---
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
        Write-Host "   $($ico.DENY) The current user does NOT have Delete permission!" -ForegroundColor Red
        $aclIssuesFound = $true
    }

    # --- Subfolder ACL Scan (folders only) ---
    if ($isFolder) {
        Write-Host "   $($ico.FIND) Scanning sub-items for Access Denied..." -ForegroundColor DarkGray
        $deniedCount = 0
        $deniedSamples = @()
        $scanCap = 500          # Enough for diagnostics while avoiding multi-minute scans
        $scanned = 0
        $hitCap = $false

        # Streaming enumeration + Owner-only ACL instead of dir /b /s + full Get-Acl
        $enumOpts = [System.IO.EnumerationOptions]::new()
        $enumOpts.RecurseSubdirectories = $true
        $enumOpts.IgnoreInaccessible = $true
        $enumOpts.AttributesToSkip = [System.IO.FileAttributes]::None   # Scan Hidden/System too

        foreach ($item in [System.IO.Directory]::EnumerateFileSystemEntries($targetPath, '*', $enumOpts)) {
            $scanned++
            if ($scanned % 200 -eq 0) {
                Write-Progress -Activity "ACL Scan" -Status "$scanned items scanned ($deniedCount issues found)"
            }
            if ($scanned -ge $scanCap) { $hitCap = $true; break }

            try {
                # Request only Owner and skip DACL/SACL/Group for speed
                if ([System.IO.Directory]::Exists($item)) {
                    $sec = [System.IO.DirectoryInfo]::new($item).GetAccessControl(
                        [System.Security.AccessControl.AccessControlSections]::Owner
                    )
                } else {
                    $sec = [System.IO.FileInfo]::new($item).GetAccessControl(
                        [System.Security.AccessControl.AccessControlSections]::Owner
                    )
                }
                # SID comparison instead of NTAccount to avoid IdentityNotMappedException false positives
                # S-1-5-18 = SYSTEM | S-1-5-80-* = NT SERVICE (TrustedInstaller, etc.)
                $ownerSid = $sec.GetOwner([System.Security.Principal.SecurityIdentifier]).Value
                if ($ownerSid -match '^S-1-5-18$|^S-1-5-80-') {
                    $deniedCount++
                    if ($deniedSamples.Count -lt 3) {
                        # Resolve to a display name only (best effort)
                        try { $displayName = $sec.GetOwner([System.Security.Principal.NTAccount]).Value }
                        catch { $displayName = $ownerSid }
                        $deniedSamples += "      → $item (Owner: $displayName)"
                    }
                }
            } catch [System.UnauthorizedAccessException] {
                # Only real Access Denied, not unrelated exceptions
                $deniedCount++
                if ($deniedSamples.Count -lt 3) { $deniedSamples += "      → $item (Access Denied)" }
            } catch {
                # Ignore long paths, corrupt ACL, orphaned SID, etc.; they are not access issues
            }
        }
        Write-Progress -Activity "ACL Scan" -Completed

        if ($deniedCount -gt 0) {
            $countLabel = if ($hitCap) { "at least $deniedCount (scanned first $scanCap items)" } else { "$deniedCount" }
            Write-Host "   $($ico.DENY) Found $countLabel files/folders with restricted access:" -ForegroundColor Red
            foreach ($s in $deniedSamples) { Write-Host $s -ForegroundColor DarkYellow }
            if ($deniedCount -gt 3) { Write-Host "      ... and $($deniedCount - 3) more" -ForegroundColor DarkGray }
            $aclIssuesFound = $true
        } else {
            Write-Host "   $($ico.OK) All sub-items have normal permissions ($scanned scanned)" -ForegroundColor Green
        }
    }

    # --- CBS / WinSxS InFlight Detection ---
    if ($targetPath -match "WinSxS\\Temp\\InFlight|WinSxS\\Temp\\PendingDeletes|WinSxS\\Temp\\PendingRenames") {
        Write-Host "`n   $($ico.WARN)  CBS INFLIGHT DETECTED!" -ForegroundColor Magenta
        Write-Host "      This folder contains pending Windows Update components." -ForegroundColor DarkYellow
        Write-Host "      The files are not locked by a process; they belong to TrustedInstaller." -ForegroundColor DarkYellow

        # Check Windows Update & TrustedInstaller service status
        $wuStatus = (Get-Service wuauserv -ErrorAction SilentlyContinue).Status
        $tiStatus = (Get-Service TrustedInstaller -ErrorAction SilentlyContinue).Status
        Write-Host "      Windows Update: $wuStatus | TrustedInstaller: $tiStatus" -ForegroundColor Gray
        $aclIssuesFound = $true
    }

} catch {
    Write-Host "   $($ico.ERR) Could not read ACL: $($_.Exception.Message)" -ForegroundColor Red
    $aclIssuesFound = $true
}

# =========================================================
# Method 4: Direct Access Test & Deep Diagnostics
# Catches: Hard Links, CBS Pending, WRP, I/O Errors
#          Real access test, not only theoretical checks
# =========================================================
Write-Host "`n$($ico.N4)  Direct Access Test & Deep Diagnostics..." -ForegroundColor Gray

$directIssuesFound = $false
$diagResults = @()

# --- 4A: Real delete/access test on one file ---
if ($isFolder) {
    # Find one sample file inside the folder for testing
    $testFile = cmd /c "dir /b /s /a-d `"$targetPath`" 2>nul" 2>$null | Select-Object -First 1
} else {
    $testFile = $targetPath
}

if ($testFile) {
    Write-Host "   $($ico.TEST) Access test: $([System.IO.Path]::GetFileName($testFile))" -ForegroundColor DarkGray

    # Test 1: Can we open the file?
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
        
        # Recognize known HResult codes
        switch ($hResult) {
            -2147024891 { $diagResults += "ACCESS_DENIED (0x80070005) — permissions or WRP block" }
            -2147024864 { $diagResults += "SHARING_VIOLATION (0x80070020) — another process is using the file" }
            -2147024858 { $diagResults += "LOCK_VIOLATION (0x80070021) — Kernel-level file lock" }
            -2147024809 { $diagResults += "INVALID_PARAMETER — path problem or reparse point" }
            default      { $diagResults += "HResult: 0x$([Convert]::ToString([uint32]$hResult, 16).ToUpper())" }
        }
        $directIssuesFound = $true
    }

    # Test 2: Hard Links Check
    Write-Host "   $($ico.LINK) Checking hard links..." -ForegroundColor DarkGray
    $hlResult = cmd /c "fsutil hardlink list `"$testFile`" 2>&1"
    if ($hlResult -and ($hlResult | Measure-Object).Count -gt 1) {
        $hlCount = ($hlResult | Measure-Object).Count
        Write-Host "   $($ico.WARN)  Hard Links Detected! ($hlCount links)" -ForegroundColor Magenta
        $hlResult | Select-Object -First 3 | ForEach-Object { Write-Host "      → $_" -ForegroundColor DarkYellow }
        if ($hlCount -gt 3) { Write-Host "      ... and $($hlCount - 3) more" -ForegroundColor DarkGray }
        $diagResults += "HARD_LINKS — the file has $hlCount hard links (WinSxS component store)"
        $directIssuesFound = $true
    } elseif ($hlResult -match "Access is denied|Error") {
        Write-Host "   $($ico.WARN)  fsutil: $hlResult" -ForegroundColor Red
        $directIssuesFound = $true
    } else {
        Write-Host "   $($ico.OK) Only 1 link (normal file)" -ForegroundColor Green
    }
} else {
    Write-Host "   (No file available for test)" -ForegroundColor DarkGray
}

# --- 4B: CBS Pending Transaction Check ---
if ($targetPath -match "WinSxS|InFlight|PendingDeletes") {
    Write-Host "   $($ico.CBS) Checking CBS pending transactions..." -ForegroundColor DarkGray
    
    $pendingXml = "C:\Windows\WinSxS\pending.xml"
    $rebootPending = "C:\Windows\WinSxS\reboot.xml"
    $cbsLogDir = "C:\Windows\Logs\CBS"
    
    if (Test-Path $pendingXml) {
        Write-Host "   $($ico.WARN)  pending.xml FOUND — unfinished CBS transaction detected!" -ForegroundColor Red
        $diagResults += "CBS_PENDING — unfinished Windows Update operation"
        $directIssuesFound = $true
    }
    if (Test-Path $rebootPending) {
        Write-Host "   $($ico.WARN)  reboot.xml FOUND — Pending reboot required!" -ForegroundColor Red
        $diagResults += "REBOOT_PENDING — restart is required to finish"
        $directIssuesFound = $true
    }
    
    # Recent CBS errors
    if (Test-Path $cbsLogDir) {
        $cbsLog = Get-ChildItem "$cbsLogDir\CBS.log" -ErrorAction SilentlyContinue
        if ($cbsLog) {
            $recentErrors = cmd /c "findstr /i /c:`"Error`" /c:`"Failed`" `"$($cbsLog.FullName)`" 2>nul" | Select-Object -Last 3
            if ($recentErrors) {
                Write-Host "   $($ico.LOG) Recent CBS errors:" -ForegroundColor DarkYellow
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
        $diagResults += "REPARSE_POINT — symlink or junction point"
        $directIssuesFound = $true
    }
} catch { }

# --- Final Diagnosis & Solutions ---
if ($directIssuesFound -or $aclIssuesFound) {
    Write-Host "`n   --- Diagnosis ---" -ForegroundColor Cyan
    foreach ($diag in $diagResults) {
        Write-Host "   $($ico.DOT) $diag" -ForegroundColor Yellow
    }

    Write-Host "`n   --- Suggested Solution ---" -ForegroundColor Cyan

    if ($diagResults -match "SHARING_VIOLATION|LOCK_VIOLATION") {
        Write-Host "   $($ico.TIP) A process is holding the file open but did not appear in handle.exe" -ForegroundColor Yellow
        Write-Host "   $($ico.TIP) Try restarting Explorer or closing indexing apps (Everything, Search)" -ForegroundColor White
    }

    if ($diagResults -match "HARD_LINKS") {
        Write-Host "   $($ico.TIP) These hard links belong to the WinSxS Component Store" -ForegroundColor Yellow
        Write-Host "      DISM /Online /Cleanup-Image /StartComponentCleanup" -ForegroundColor White
        Write-Host "      Or: DISM /Online /Cleanup-Image /StartComponentCleanup /ResetBase" -ForegroundColor DarkGray
    }

    if ($diagResults -match "CBS_PENDING|REBOOT_PENDING") {
        Write-Host "   $($ico.TIP) Restart first, then try again" -ForegroundColor Yellow
        Write-Host "      If the folder still exists after restart:" -ForegroundColor DarkGray
        Write-Host "      DISM /Online /Cleanup-Image /RestoreHealth" -ForegroundColor White
    }

    if ($targetPath -match "WinSxS") {
        Write-Host "   $($ico.TIP) Takeown + Unlock:" -ForegroundColor Yellow
        Write-Host "      takeown /F `"$targetPath`" /R /A /D Y" -ForegroundColor White
        Write-Host "      icacls `"$targetPath`" /grant Administrators:F /T /C" -ForegroundColor White
    } elseif (-not ($diagResults -match "SHARING_VIOLATION|LOCK_VIOLATION|HARD_LINKS|CBS_PENDING")) {
        Write-Host "   $($ico.TIP) Run in an elevated prompt:" -ForegroundColor Green
        Write-Host "      takeown /F `"$targetPath`" /R /A /D Y" -ForegroundColor White
        Write-Host "      icacls `"$targetPath`" /grant Administrators:F /T /C" -ForegroundColor White
    }

    # Offer automatic takeown
    $proceedWithFix = $false

    if ($isCriticalPath) {
        # Extra warning for critical system paths
        Write-Host ""
        Write-Host "   $($ico.DENY)$($ico.DENY)$($ico.DENY) CRITICAL SYSTEM PATH DETECTED $($ico.DENY)$($ico.DENY)$($ico.DENY)" -ForegroundColor Red
        Write-Host "   This folder is critical for Windows." -ForegroundColor Red
        Write-Host "   Changing ownership/permissions can cause:" -ForegroundColor Red
        Write-Host "      -> Boot failure (unbootable system)" -ForegroundColor DarkRed
        Write-Host "      -> Windows Update failure" -ForegroundColor DarkRed
        Write-Host "      -> Corrupted Component Store" -ForegroundColor DarkRed
        Write-Host "      → Blue Screen of Death (BSOD)" -ForegroundColor DarkRed
        Write-Host ""
        Write-Host "   $($ico.TIP) If you want to clean WinSxS, use this instead:" -ForegroundColor Yellow
        Write-Host "      DISM /Online /Cleanup-Image /StartComponentCleanup" -ForegroundColor White
        Write-Host ""
        Write-Host "   Type CONFIRM (uppercase) to continue  |  Enter or anything else = cancel" -ForegroundColor Red
        $confirmInput = Read-Host "   Choice"
        if ($confirmInput -ceq "CONFIRM") {
            $proceedWithFix = $true
        } else {
            Write-Host "   Cancelled. Good choice." -ForegroundColor Green
        }
    } else {
        $autoFix = Read-Host "`n      Run automatic Takeown + Grant Access now? (Y/N)"
        if ($autoFix -eq "Y") { $proceedWithFix = $true }
    }

    if ($proceedWithFix) {
        Write-Host "      $($ico.DOT) Running takeown..." -ForegroundColor Gray
        $tkResult = & takeown.exe /F $targetPath /R /A /D Y 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "      $($ico.OK) Ownership OK" -ForegroundColor Green
        } else {
            Write-Host "      $($ico.WARN) Takeown: $tkResult" -ForegroundColor Red
        }

        Write-Host "      $($ico.DOT) Running icacls..." -ForegroundColor Gray
        $icResult = & icacls.exe $targetPath /grant "Administrators:F" /T /C /Q 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "      $($ico.OK) Permissions granted. You can now manage it manually." -ForegroundColor Green
        } else {
            Write-Host "      $($ico.WARN) Icacls: $icResult" -ForegroundColor Red
        }
    }
} else {
    Write-Host "   $($ico.OK) Direct Access OK — no deep issue found" -ForegroundColor Green
    if (-not $aclIssuesFound) {
        Write-Host "`n   $($ico.OK) Ownership & Permissions OK — no ACL issue found" -ForegroundColor Green
    }
}

Write-Host ("`n" + ("=" * 40)) -ForegroundColor Gray
Read-CloseOrUpdate

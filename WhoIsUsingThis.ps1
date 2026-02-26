param([string]$targetPath)

# 🔸 Force UTF-8 Console Encoding
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 | Out-Null

# ---------------------------------------------------------
# 1. Trailing Slash / Quote Cleanup
# ---------------------------------------------------------
if ($targetPath.EndsWith('\')) { $targetPath = $targetPath.TrimEnd('\') }
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
if (-not ([System.IO.File]::Exists($targetPath) -or [System.IO.Directory]::Exists($targetPath))) {
    Write-Host "`n$($ico.ERR) Error: Path not found." -ForegroundColor Red
    Write-Host "   $targetPath" -ForegroundColor Gray
    Write-Host "`nPress any key to close..." -ForegroundColor Yellow
    $null = [Console]::ReadKey(); exit
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
Write-Host "Πάτα οποιοδήποτε πλήκτρο για να κλείσει..." -ForegroundColor Yellow
$null = [Console]::ReadKey()

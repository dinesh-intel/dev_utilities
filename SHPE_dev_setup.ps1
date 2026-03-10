#Requires -Version 5
<#
.SYNOPSIS
        DevMachine-Setup.ps1 v6.0 - Automated development environment setup
.DESCRIPTION
        Installs: VS Code, Miniforge (conda), Git, GitHub CLI, Intel dt
    Configures: VS Code proxy settings, VS Code extensions
        Creates: Python 3.11 environment called 'py_learn' with common packages for data + Excel/Word/PPT/PDF/image workflows (fully automated)
        Installs: common Python packages including pymupdf for PDF processing using conda run
    Logs: detailed transcript to %TEMP%

        Changelog:
        - v6.0
            * Added Intel dt bootstrap flow:
                - Downloads dt.exe to Downloads folder
                - Runs 'dt.exe install'
                - Runs mandatory 'dt setup' after GitHub CLI setup
            * Converted py_learn environment creation and package installation
                from manual Miniforge Prompt instructions to automated conda commands
                executed directly from PowerShell
            * Updated startup banner and summary text for new tooling and automation
        - v2
            * Bug fixes (B-01..B-08), pre-flight checks, interactive mode,
                per-package pip install, version detection, visual progress, dry-run support,
                per-component skip flags, PS 5.1 compatibility.
.PARAMETER Proxy
    Proxy server URL (default: http://proxy-chain.intel.com:912) - Used only for VS Code settings and manual instructions
.PARAMETER NoProxy
    NO_PROXY list for bypassing proxy (informational only)
.PARAMETER InstallIntelCerts
    Import Intel certificate chain into Windows trust store
.PARAMETER InstallBuildTools
    Install Visual Studio Build Tools for compiling Python packages
.PARAMETER UseExternalPyPI
    Prefer external PyPI during package installs (informational only)
.PARAMETER Interactive
    Stop and ask for user confirmation (Yes/No/Skip) at each step
.PARAMETER DryRun
    Print all planned actions without executing them
.PARAMETER SkipVSCode
    Skip Visual Studio Code installation
.PARAMETER SkipNodeJS
    Skip Node.js LTS installation (deprecated - Node.js no longer installed)
.PARAMETER SkipPython
    Skip Miniforge installation
.PARAMETER SkipGit
    Skip Git installation
.PARAMETER SkipGitHubCLI
    Skip GitHub CLI installation
.PARAMETER SkipPythonPackages
    Skip Python package installation
.PARAMETER SkipExtensions
    Skip VS Code extension installation
.PARAMETER SkipGitHubAuth
    Skip GitHub CLI authentication
.PARAMETER SkipDt
    Skip Intel dt installation and setup
.EXAMPLE
    .\DevMachine-Setup.ps1 -InstallIntelCerts
.EXAMPLE
    .\DevMachine-Setup.ps1 -Interactive
.EXAMPLE
    .\DevMachine-Setup.ps1 -DryRun
.EXAMPLE
    .\DevMachine-Setup.ps1 -Proxy "http://proxy.iind.intel.com:911"
#>

[CmdletBinding()]
param(
    [Parameter(Position=0)]
    [string]$Proxy = "http://proxy-chain.intel.com:912",
    [string]$NoProxy = "localhost,127.0.0.0/8,172.16.0.0/20,192.168.0.0/16,10.0.0.0/8,intel.com",
    [switch]$InstallIntelCerts,
    [switch]$InstallBuildTools,
    [switch]$UseExternalPyPI,
    [Alias("i")]
    [switch]$Interactive,
    [switch]$DryRun,
    [switch]$SkipVSCode,
    [switch]$SkipNodeJS,  # Deprecated - kept for backward compatibility
    [switch]$SkipPython,
    [switch]$SkipGit,
    [switch]$SkipGitHubCLI,
    [switch]$SkipPythonPackages,
    [switch]$SkipExtensions,
    [switch]$SkipGitHubAuth,
    [switch]$SkipDt
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ========================================
# PARAMETER VALIDATION AND CORRECTION
# ========================================
# Handle common parameter parsing issues
if ($Proxy -eq "DryRun" -or $Proxy -eq "Interactive") {
    Write-Warning "Detected parameter parsing issue. Correcting..."
    if ($Proxy -eq "DryRun") {
        $DryRun = $true
        $Proxy = "http://proxy-chain.intel.com:912"
        Write-Host "Set -DryRun = $true and reset Proxy to default" -ForegroundColor Yellow
    }
    if ($Proxy -eq "Interactive") {
        $Interactive = $true
        $Proxy = "http://proxy-chain.intel.com:912"
        Write-Host "Set -Interactive = $true and reset Proxy to default" -ForegroundColor Yellow
    }
}

# ========================================
# GLOBAL CONFIGURATION
# ========================================
$script:MaxRetries      = 3
$script:RetryDelaySeconds = 5
$script:TotalSteps      = 0
$script:CurrentStep     = 0
$script:SuccessfulSteps = @()
$script:FailedSteps     = @()
$script:SkippedSteps    = @()
$script:StepTimings     = @()          # Per-step elapsed time tracking
$script:StartTime       = Get-Date
$script:VersionMatrix   = @()          # Post-run version table

# ========================================
# LOGGING SETUP
# ========================================
$stamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile = Join-Path $env:TEMP "DevSetup_$stamp.log"
# Stop any dangling transcript from a prior failed run, then start fresh
try { Stop-Transcript -ErrorAction SilentlyContinue | Out-Null } catch { }
try { Start-Transcript -Path $LogFile -Append -ErrorAction Stop | Out-Null }
catch { Write-Warning "Could not start transcript: $_" }

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $line  = "[{0}] [{1}] {2}" -f (Get-Date -Format "s"), $Level, $Message
    $color = switch ($Level) {
        "ERROR"    { "Red"     }
        "WARN"     { "Yellow"  }
        "SUCCESS"  { "Green"   }
        "PROGRESS" { "Cyan"    }
        "HEADER"   { "Magenta" }
        default    { "White"   }
    }
    Write-Host $line -ForegroundColor $color
}

# --------------- visual helpers ---------------
function Write-Banner {
    param([string]$Title, [string]$Level = "HEADER")
    $border = "=" * 60
    Write-Host ""
    Write-Host "  $border" -ForegroundColor Magenta
    Write-Host ("  " + $Title.PadLeft([math]::Floor(($border.Length + $Title.Length) / 2)).PadRight($border.Length)) -ForegroundColor Magenta
    Write-Host "  $border" -ForegroundColor Magenta
    Write-Host ""
}

function Write-StepBanner {
    param(
        [int]$Number,
        [int]$Total,
        [string]$Name,
        [string]$Upcoming = ""
    )
    $elapsed  = (Get-Date) - $script:StartTime
    $pct      = [math]::Round(($Number / $Total) * 100, 0)
    $bar      = ("#" * [math]::Floor($pct / 5)).PadRight(20, "-")
    $ts       = "{0:hh\:mm\:ss}" -f $elapsed

    Write-Host ""
    Write-Host "  +--------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "  |  STEP $Number / $Total   [$bar]  $pct%   Elapsed: $ts" -ForegroundColor Cyan
    Write-Host "  |  >> $Name" -ForegroundColor White
    if ($Upcoming) {
        Write-Host "  |  (next: $Upcoming)" -ForegroundColor DarkGray
    }
    Write-Host "  +--------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host ""

    # PowerShell progress bar (visible in taskbar)
    Write-Progress -Activity "DevMachine Setup" -Status "Step $Number / $Total : $Name" `
        -PercentComplete $pct -Id 1
}

function Write-StepProgress {
    param(
        [string]$StepName,
        [string]$UpcomingStep = ""
    )
    $script:CurrentStep++
    Write-StepBanner -Number $script:CurrentStep -Total $script:TotalSteps `
        -Name $StepName -Upcoming $UpcomingStep
}

function Write-StepResult {
    param([string]$Result, [string]$StepName, [timespan]$Elapsed)
    $ts = "{0:mm\:ss\.ff}" -f $Elapsed
    switch ($Result) {
        "SUCCESS" { Write-Host "  [OK]   $StepName  ($ts)" -ForegroundColor Green }
        "SKIPPED" { Write-Host "  [SKIP] $StepName  ($ts)" -ForegroundColor Yellow }
        "FAILED"  { Write-Host "  [FAIL] $StepName  ($ts)" -ForegroundColor Red }
    }
}

# ========================================
# INTERACTIVE PROMPT
# ========================================
function Request-StepApproval {
    param([string]$StepName)
    if (-not $Interactive) { return "Yes" }

    Write-Host ""
    Write-Host "  +=========================================================+" -ForegroundColor Yellow
    Write-Host "  |  INTERACTIVE: Proceed with '$StepName'?                 " -ForegroundColor Yellow
    Write-Host "  |  [Y] Yes   [N] No (abort)   [S] Skip this step         " -ForegroundColor Yellow
    Write-Host "  +=========================================================+" -ForegroundColor Yellow

    while ($true) {
        $key = Read-Host "  Enter choice (Y/N/S)"
        switch ($key.ToUpper()) {
            "Y" { return "Yes"  }
            "N" { return "No"   }
            "S" { return "Skip" }
            default { Write-Host "  Invalid input. Please enter Y, N, or S." -ForegroundColor Red }
        }
    }
}

function Request-VersionConflictResolution {
    param(
        [string]$ToolName,
        [string]$CurrentVersion,
        [string]$ExpectedVersion
    )
    
    Write-Host ""
    Write-Host "  +=========================================================+" -ForegroundColor Red
    Write-Host "  |  VERSION CONFLICT DETECTED                              |" -ForegroundColor Red
    Write-Host "  |                                                         |" -ForegroundColor Red
    Write-Host "  |  Tool: $($ToolName.PadRight(47)) |" -ForegroundColor Red
    Write-Host "  |  Current:  $($CurrentVersion.PadRight(43)) |" -ForegroundColor Red
    Write-Host "  |  Expected: $($ExpectedVersion.PadRight(43)) |" -ForegroundColor Red
    Write-Host "  |                                                         |" -ForegroundColor Red
    Write-Host "  |  Please manually remove the existing version and       |" -ForegroundColor Red
    Write-Host "  |  restart this script, or choose to keep current.       |" -ForegroundColor Red
    Write-Host "  |                                                         |" -ForegroundColor Red
    Write-Host "  |  [K] Keep current   [A] Abort (recommended)            |" -ForegroundColor Red
    Write-Host "  +=========================================================+" -ForegroundColor Red

    while ($true) {
        $key = Read-Host "  Enter choice (K/A)"
        switch ($key.ToUpper()) {
            "K" { return "Keep" }
            "A" { return "Abort" }
            default { Write-Host "  Invalid input. Please enter K or A." -ForegroundColor Red }
        }
    }
}

# ========================================
# ADMIN CHECK
# ========================================
function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p  = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# ========================================
# PRE-FLIGHT CHECKS
# ========================================
function Invoke-PreFlightChecks {
    Write-Banner "PRE-FLIGHT CHECKS"

    # --- Parameter validation display ---
    Write-Log "Parameter Values:"
    Write-Log "  Proxy         : $Proxy (for VS Code settings and manual instructions only)"
    Write-Log "  DryRun        : $DryRun"
    Write-Log "  Interactive   : $Interactive"
    Write-Log "  SkipPython    : $SkipPython"
    Write-Log "  SkipVSCode    : $SkipVSCode"
    Write-Log ""

    # --- Machine context header (INF-06) ---
    Write-Log "Hostname      : $env:COMPUTERNAME"
    Write-Log "Username      : $env:USERNAME"
    Write-Log "Windows       : $([Environment]::OSVersion.VersionString)"
    Write-Log "PS Version    : $($PSVersionTable.PSVersion)"
    Write-Log "Admin         : $(Test-IsAdmin)"
    Write-Log "Architecture  : $([Environment]::Is64BitOperatingSystem | ForEach-Object { if ($_) { 'x64' } else { 'x86' } })"
    Write-Log ""

    # --- OS version check (RS-04 / P0-07) ---
    $osVer = [Environment]::OSVersion.Version
    if ($osVer.Major -lt 10) {
        Write-Log "Windows 10 or later is required for WinGet. Detected: $osVer" "ERROR"
        throw "Unsupported OS version: $osVer"
    }
    Write-Log "OS version check passed ($osVer)" "SUCCESS"

    # --- Admin warning (UF-05 / P1-11) ---
    if (-not (Test-IsAdmin)) {
        Write-Log "Not running as Administrator. Some steps may be limited:" "WARN"
        Write-Log "  - Certificate import (-InstallIntelCerts) requires Admin" "WARN"
        if ($InstallIntelCerts) {
            Write-Log "You specified -InstallIntelCerts but are NOT Admin. This step will fail." "ERROR"
            throw "Admin required when -InstallIntelCerts is specified"
        }
    } else {
        Write-Log "Running as Administrator" "SUCCESS"
    }

    # --- Disk space check (RS-05 / P1-10) ---
    $sysDrive  = $env:SystemDrive
    $freeBytes = (Get-PSDrive ($sysDrive.TrimEnd(':'))).Free
    $freeGB    = [math]::Round($freeBytes / 1GB, 2)
    if ($freeGB -lt 5) {
        Write-Log "Low disk space on $sysDrive : ${freeGB} GB free (minimum 5 GB recommended)" "ERROR"
        throw "Insufficient disk space: ${freeGB} GB"
    }
    Write-Log "Disk space check passed (${freeGB} GB free on $sysDrive)" "SUCCESS"

    # --- Proxy information ---
    if (-not $DryRun) {
        Write-Log "Proxy will be configured manually in command sessions: $Proxy" "SUCCESS"
    } else {
        Write-Log "Skipping proxy validation in DryRun mode" "WARN"
    }
    Write-Log ""
}

# ========================================
# RETRY LOGIC  (B-07 fix: no retry for interactive gh auth)
# ========================================
function Invoke-WithRetry {
    param(
        [string]$StepName,
        [scriptblock]$ScriptBlock,
        [int]$MaxAttempts = $script:MaxRetries,
        [int]$DelaySeconds = $script:RetryDelaySeconds,
        [switch]$Optional,
        [string]$UpcomingStep = ""
    )

    # --- Interactive gate ---
    $approval = Request-StepApproval -StepName $StepName
    if ($approval -eq "No") {
        Write-Log "User chose to ABORT at step: $StepName" "ERROR"
        throw "Aborted by user at step: $StepName"
    }
    if ($approval -eq "Skip") {
        $script:CurrentStep++
        Write-StepBanner -Number $script:CurrentStep -Total $script:TotalSteps `
            -Name $StepName -Upcoming $UpcomingStep
        Write-Log "User chose to SKIP: $StepName" "WARN"
        $script:SkippedSteps += $StepName
        $script:StepTimings += @{ Step = $StepName; Elapsed = [timespan]::Zero; Result = "SKIPPED" }
        return $true
    }

    # --- DryRun gate ---
    if ($DryRun) {
        $script:CurrentStep++
        Write-StepBanner -Number $script:CurrentStep -Total $script:TotalSteps `
            -Name $StepName -Upcoming $UpcomingStep
        Write-Log "[DRY RUN] Would execute: $StepName" "WARN"
        $script:SkippedSteps += "$StepName (dry-run)"
        $script:StepTimings += @{ Step = $StepName; Elapsed = [timespan]::Zero; Result = "DRYRUN" }
        return $true
    }

    Write-StepProgress -StepName $StepName -UpcomingStep $UpcomingStep

    $stepStart = Get-Date
    $attempt   = 0
    $success   = $false
    $lastError = $null

    while ($attempt -lt $MaxAttempts -and -not $success) {
        $attempt++
        try {
            if ($attempt -gt 1) {
                Write-Log "  Retry $attempt / $MaxAttempts for: $StepName" "WARN"
            }
            & $ScriptBlock
            $success = $true
            $script:SuccessfulSteps += $StepName
        }
        catch {
            $lastError = $_
            Write-Log "  Attempt $attempt failed: $($_.Exception.Message)" "WARN"
            if ($attempt -lt $MaxAttempts) {
                Write-Log "  Retrying in $DelaySeconds seconds..." "WARN"
                Start-Sleep -Seconds $DelaySeconds
            }
        }
    }

    $stepElapsed = (Get-Date) - $stepStart

    if ($success) {
        Write-StepResult -Result "SUCCESS" -StepName $StepName -Elapsed $stepElapsed
        $script:StepTimings += @{ Step = $StepName; Elapsed = $stepElapsed; Result = "SUCCESS" }
    }
    elseif ($Optional) {
        Write-StepResult -Result "SKIPPED" -StepName $StepName -Elapsed $stepElapsed
        $script:SkippedSteps += $StepName
        $script:StepTimings += @{ Step = $StepName; Elapsed = $stepElapsed; Result = "SKIPPED" }
    }
    else {
        Write-StepResult -Result "FAILED" -StepName $StepName -Elapsed $stepElapsed
        $script:FailedSteps += @{ Step = $StepName; Error = $lastError.Exception.Message; Attempt = $attempt }
        $script:StepTimings += @{ Step = $StepName; Elapsed = $stepElapsed; Result = "FAILED" }
        throw $lastError
    }

    return $success
}

# ========================================
# UTILITY FUNCTIONS
# ========================================
function Refresh-Path {
    # B-06 fix: include Process scope entries to capture in-session installer mutations
    $machine = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $user    = [Environment]::GetEnvironmentVariable("Path", "User")
    $process = [Environment]::GetEnvironmentVariable("Path", "Process")

    $combined = @()
    foreach ($p in ($machine, $user, $process)) {
        if ($p) { $combined += $p -split ";" }
    }
    $env:Path = ($combined | Select-Object -Unique) -join ";"
    Write-Log "PATH refreshed for current session"
}

# ========================================
# CONDA EXECUTABLE HELPER
# ========================================
function Get-CondaExecutable {
    $miniforgeDirs = @(
        (Join-Path $env:USERPROFILE "miniforge3"),
        (Join-Path $env:LOCALAPPDATA "miniforge3")
    )
    $condaPaths = @(
        (Join-Path $miniforgeDirs[0] "condabin\conda.bat"),
        (Join-Path $miniforgeDirs[0] "Scripts\conda.exe"),
        (Join-Path $miniforgeDirs[0] "Scripts\conda.bat"),
        (Join-Path $miniforgeDirs[0] "bin\conda"),
        (Join-Path $miniforgeDirs[1] "condabin\conda.bat"),
        (Join-Path $miniforgeDirs[1] "Scripts\conda.exe"),
        (Join-Path $miniforgeDirs[1] "Scripts\conda.bat"),
        (Join-Path $miniforgeDirs[1] "bin\conda")
    )
    
    foreach ($path in $condaPaths) {
        if (Test-Path $path) {
            Write-Log "Found conda executable at: $path" "SUCCESS"
            return $path
        }
    }
    
    Write-Log "Conda executable not found in any expected location:" "WARN"
    foreach ($path in $condaPaths) {
        Write-Log "  Checked: $path" "WARN"
    }
    return $null
}

function Resolve-CondaExecutable {
    $condaPath = Get-CondaExecutable
    if ($condaPath) {
        return $condaPath
    }

    $cmd = Get-Command conda -ErrorAction SilentlyContinue
    if ($cmd) {
        Write-Log "Using conda from PATH: $($cmd.Source)" "SUCCESS"
        return $cmd.Source
    }

    return $null
}

# ========================================
# VERSION DETECTION
# ========================================
function Get-InstalledVersion {
    param([string]$Tool)
    try {
        switch ($Tool) {
            "code"   { 
                $v = & code --version 2>$null | Select-Object -First 1
                return $v
            }
            "conda"  { 
                $condaPath = Get-CondaExecutable
                if ($condaPath) {
                    $v = & $condaPath --version 2>$null
                    if ($v -match "conda (\d+\.\d+\.\d+)") {
                        return $matches[1]
                    }
                    return $v
                }
                return "not installed"
            }
            "python" { 
                $v = $null
                if (Get-Command python -ErrorAction SilentlyContinue) {
                    $v = & python --version 2>$null | Select-Object -First 1
                    if ($v -match "Python (\d+\.\d+\.\d+)") {
                        return $matches[1]
                    }
                }

                $condaPath = Resolve-CondaExecutable
                if ($condaPath) {
                    $v = & $condaPath run -n py_learn python --version 2>$null | Select-Object -First 1
                    if ($v -match "Python (\d+\.\d+\.\d+)") {
                        return $matches[1]
                    }

                    $v = & $condaPath run -n base python --version 2>$null | Select-Object -First 1
                    if ($v -match "Python (\d+\.\d+\.\d+)") {
                        return $matches[1]
                    }
                }

                return "not installed"
            }
            "python-py_learn" {
                $condaPath = Resolve-CondaExecutable
                if (-not $condaPath) { return "not installed" }

                $out = & $condaPath run -n py_learn python --version 2>&1
                foreach ($line in $out) {
                    if ($line -match "Python (\d+\.\d+\.\d+)") {
                        return $matches[1]
                    }
                }
                return "not installed"
            }
            "git"    { 
                $v = & git --version 2>$null
                if ($v -match "git version (\d+\.\d+\.\d+)") {
                    return $matches[1]
                }
                return $v
            }
            "gh"     { 
                $v = & gh --version 2>$null | Select-Object -First 1
                if ($v -match "gh version (\d+\.\d+\.\d+)") {
                    return $matches[1]
                }
                return $v
            }
            "pip"    { 
                $v = $null
                if (Get-Command python -ErrorAction SilentlyContinue) {
                    $v = & python -m pip --version 2>$null | Select-Object -First 1
                    if ($v -match "pip (\d+\.\d+\.\d+)") {
                        return $matches[1]
                    }
                }

                $condaPath = Resolve-CondaExecutable
                if ($condaPath) {
                    $v = & $condaPath run -n py_learn python -m pip --version 2>$null | Select-Object -First 1
                    if ($v -match "pip (\d+\.\d+\.\d+)") {
                        return $matches[1]
                    }

                    $v = & $condaPath run -n base python -m pip --version 2>$null | Select-Object -First 1
                    if ($v -match "pip (\d+\.\d+\.\d+)") {
                        return $matches[1]
                    }
                }

                return "not installed"
            }
            "pip-py_learn" {
                $condaPath = Resolve-CondaExecutable
                if (-not $condaPath) { return "not installed" }

                $out = & $condaPath run -n py_learn python -m pip --version 2>&1
                foreach ($line in $out) {
                    if ($line -match "pip (\d+\.\d+\.\d+)") {
                        return $matches[1]
                    }
                }
                return "not installed"
            }
            "dt"     {
                $dtExe = Get-DtExecutablePath
                if (Test-Path $dtExe) {
                    $v = Get-DtVersion -DtExecutable $dtExe
                    if ($v) { return $v }
                    return "installed"
                }
                return "not installed"
            }
            "winget" { 
                $v = & winget --version 2>$null
                return $v
            }
            default  { return "N/A" }
        }
    }
    catch { return "not installed" }
}

function Get-WingetInstalledVersion {
    param([string]$Id)
    try {
        $out = winget list --id $Id --accept-source-agreements 2>$null
        # parse version from tabular output
        foreach ($line in $out) {
            if ($line -match $Id) {
                $parts = $line -split '\s{2,}'
                if ($parts.Count -ge 3) { return $parts[2].Trim() }
            }
        }
        return $null
    }
    catch { return $null }
}

function Test-MiniforgePath {
    $miniforgeDir = Join-Path $env:USERPROFILE "miniforge3"
    $condaExe = Get-CondaExecutable
    return ($condaExe -ne $null)
}

function Get-MiniforgePythonVersion {
    try {
        $pythonCandidates = @(
            (Join-Path $env:USERPROFILE "miniforge3\python.exe"),
            (Join-Path $env:LOCALAPPDATA "miniforge3\python.exe")
        )
        foreach ($pythonExe in $pythonCandidates) {
            if (Test-Path $pythonExe) {
                $v = & $pythonExe --version 2>$null
                if ($v -match "Python (\d+\.\d+\.\d+)") {
                    return $matches[1]
                }
            }
        }
        return "not found"
    }
    catch { return "not found" }
}

function Test-CondaEnvironmentExists {
    param([string]$EnvName)
    try {
        $condaPath = Resolve-CondaExecutable
        if ($condaPath) {
            $envList = & $condaPath env list 2>$null
            foreach ($line in $envList) {
                if ($line -match "^\s*$EnvName\s" -or $line -match "\\envs\\$EnvName\s*$") {
                    Write-Log "Conda environment '$EnvName' found in conda env list" "SUCCESS"
                    return $true
                }
            }
        }

        # Check multiple possible locations for conda environments
        $possiblePaths = @(
            (Join-Path $env:USERPROFILE "miniforge3\envs\$EnvName"),           # User profile location
            (Join-Path $env:LOCALAPPDATA "miniforge3\envs\$EnvName"),          # AppData\Local location
            (Join-Path $env:APPDATA "miniforge3\envs\$EnvName")                # AppData\Roaming location (fallback)
        )
        
        foreach ($envPath in $possiblePaths) {
            if (Test-Path $envPath) {
                Write-Log "Found conda environment directory: $envPath" "SUCCESS"
                return $true
            } else {
                Write-Log "Checked conda environment path: $envPath (not found)" "WARN"
            }
        }
        
        Write-Log "Conda environment '$EnvName' not found in any expected location:" "WARN"
        foreach ($path in $possiblePaths) {
            Write-Log "  - $path" "WARN"
        }
        
        return $false
    }
    catch { 
        Write-Log "Error checking conda environment: $($_.Exception.Message)" "WARN"
        return $false 
    }
}

# ========================================
# WINGET MANAGEMENT
# ========================================
function Ensure-WinGet {
    # Phase 1: check if winget is already present
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Write-Log "WinGet already installed: $(winget --version)" "SUCCESS"
        return
    }

    Write-Log "WinGet not found. Installing via Microsoft.WinGet.Client (PSGallery)..." "WARN"

    # Phase 2: bootstrap via PSGallery / Repair-WinGetPackageManager
    try {
        $global:ProgressPreference = 'SilentlyContinue'

        Write-Log "  Installing NuGet package provider..."
        Install-PackageProvider -Name NuGet -Force -ErrorAction Stop | Out-Null

        Write-Log "  Installing Microsoft.WinGet.Client module from PSGallery..."
        Install-Module -Name Microsoft.WinGet.Client -Force -Repository PSGallery -ErrorAction Stop | Out-Null

        Write-Log "  Running Repair-WinGetPackageManager -AllUsers ..."
        Repair-WinGetPackageManager -AllUsers -ErrorAction Stop

        $global:ProgressPreference = 'Continue'
    }
    catch {
        $global:ProgressPreference = 'Continue'
        Write-Log "  PSGallery install path failed: $($_.Exception.Message)" "WARN"
        Write-Log "  Falling back to App Installer AppX registration..." "WARN"

        # Phase 3: fallback - re-register the Desktop App Installer AppX package
        try {
            Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe -ErrorAction Stop | Out-Null
            Start-Sleep -Seconds 3
        }
        catch {
            Write-Log "  AppX registration also failed: $($_.Exception.Message)" "ERROR"
            Write-Log "  Please install/update 'App Installer' from the Microsoft Store and rerun." "ERROR"
            throw
        }
    }

    # Final verification
    Refresh-Path
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Log "WinGet still not available after installation attempts." "ERROR"
        throw "WinGet unavailable after install"
    }

    Write-Log "WinGet is now available: $(winget --version)" "SUCCESS"
}

# ========================================
# CERTIFICATE INSTALLATION  (RS-07 fix: admin check first)
# ========================================
function Install-IntelCertificateBundles {
    if (-not (Test-IsAdmin)) {
        Write-Log "Certificate import needs Admin. Re-run in elevated PowerShell or skip -InstallIntelCerts." "ERROR"
        throw "Admin required for certificate import"
    }

    Write-Log "Installing Intel certificate bundles into Windows certificate stores..."

    $TempDir = Join-Path $env:TEMP "intel_certs_$stamp"
    New-Item -ItemType Directory -Path $TempDir -Force | Out-Null

    $zip1 = Join-Path $TempDir "IntelSHA2RootChain-Base64.zip"
    $zip2 = Join-Path $TempDir "IntelSHA384TrustChain-Base64.zip"

    Invoke-WebRequest -Uri "http://certificates.intel.com/repository/certificates/IntelSHA2RootChain-Base64.zip" -OutFile $zip1 -UseBasicParsing
    Invoke-WebRequest -Uri "https://certificates.intel.com/repository/certificates/TrustBundles/IntelSHA384TrustChain-Base64.zip" -OutFile $zip2 -UseBasicParsing

    Expand-Archive $zip1 -DestinationPath $TempDir -Force
    Expand-Archive $zip2 -DestinationPath $TempDir -Force

    $certCount = 0
    Get-ChildItem -Path $TempDir -Recurse -Filter *.crt | ForEach-Object {
        Write-Log "  Importing cert: $($_.Name)"
        Import-Certificate -FilePath $_.FullName -CertStoreLocation "Cert:\LocalMachine\Root" | Out-Null
        Import-Certificate -FilePath $_.FullName -CertStoreLocation "Cert:\LocalMachine\CA"   | Out-Null
        $certCount++
    }

    Remove-Item -Recurse -Force $TempDir
    Write-Log "Intel certificates imported: $certCount certificates"
}

# ========================================
# MINIFORGE INSTALLATION
# ========================================
function Install-Miniforge {
    # Check if Miniforge is already installed
    if (Test-MiniforgePath) {
        $currentVersion = Get-MiniforgePythonVersion
        Write-Log "Miniforge is already installed (Python $currentVersion)" "SUCCESS"
        Write-Log "Skipping Miniforge installation - using existing installation"
        return
    }

    Write-Log "Installing Miniforge (conda-forge distribution)..."
    
    $TempDir = Join-Path $env:TEMP "miniforge_install_$stamp"
    New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
    
    $installerPath = Join-Path $TempDir "Miniforge3-Windows-x86_64.exe"
    $downloadUrl = "https://github.com/conda-forge/miniforge/releases/latest/download/Miniforge3-Windows-x86_64.exe"
    
    Write-Log "  Downloading Miniforge installer from: $downloadUrl"
    Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath -UseBasicParsing
    
    Write-Log "  Running Miniforge installer (silent mode)..."
    $installArgs = @(
        "/InstallationType=JustMe"
        "/RegisterPython=1"
        "/S"
        "/D=$env:USERPROFILE\miniforge3"
    )
    
    $process = Start-Process -FilePath $installerPath -ArgumentList $installArgs -Wait -PassThru
    
    if ($process.ExitCode -ne 0) {
        throw "Miniforge installation failed with exit code: $($process.ExitCode)"
    }
    
    # Add conda to PATH for current session
    $condaPath = Join-Path $env:USERPROFILE "miniforge3\Scripts"
    $condaBinPath = Join-Path $env:USERPROFILE "miniforge3"
    $condaBinPath2 = Join-Path $env:USERPROFILE "miniforge3\condabin"
    $env:PATH = "$condaBinPath2;$condaPath;$condaBinPath;$env:PATH"
    
    # Clean up installer
    Remove-Item -Recurse -Force $TempDir
    
    Write-Log "Miniforge installed successfully"
}

function Create-PythonEnvironment {
    # Check if py_learn environment already exists
    if (Test-CondaEnvironmentExists -EnvName "py_learn") {
        Write-Log "Python environment 'py_learn' already exists" "SUCCESS"
        Write-Log "Skipping environment creation - using existing environment"
        return
    }

    $condaPath = Resolve-CondaExecutable
    if (-not $condaPath) {
        throw "Conda executable not found. Install Miniforge first."
    }

    Write-Log "Creating Python environment 'py_learn' automatically via conda..."
    $env:http_proxy = $Proxy
    $env:https_proxy = $Proxy
    & $condaPath create -n py_learn python=3.11 -y 2>&1 | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to create conda environment 'py_learn'"
    }

    if (Test-CondaEnvironmentExists -EnvName "py_learn") {
        Write-Log "Python environment 'py_learn' created successfully" "SUCCESS"
    } else {
        throw "Could not verify py_learn environment creation"
    }
}

function Install-PythonPackages {
    # Verify py_learn environment exists
    if (-not (Test-CondaEnvironmentExists -EnvName "py_learn")) {
        Write-Log "py_learn environment not found; skipping Python packages." "WARN"
        throw "py_learn environment not available"
    }

    $condaPath = Resolve-CondaExecutable
    if (-not $condaPath) {
        throw "Conda executable not found. Install Miniforge first."
    }

    Write-Log "Installing Python packages automatically in py_learn using conda run..."
    $env:http_proxy = $Proxy
    $env:https_proxy = $Proxy

    $installCommands = @(
        @("python", "-m", "pip", "install", "--upgrade", "pip", "setuptools", "wheel"),
        @("python", "-m", "pip", "install", "numpy", "pandas", "scipy", "matplotlib", "seaborn", "scikit-learn"),
        @("python", "-m", "pip", "install", "jupyter", "ipykernel"),
        @("python", "-m", "pip", "install", "openpyxl", "xlrd", "xlsxwriter"),
        @("python", "-m", "pip", "install", "python-docx", "python-pptx"),
        @("python", "-m", "pip", "install", "pypdf", "reportlab", "pymupdf"),
        @("python", "-m", "pip", "install", "pillow"),
        @("python", "-m", "pip", "install", "requests", "tqdm", "rich")
    )

    foreach ($cmdArgs in $installCommands) {
        Write-Log ("  conda run -n py_learn " + ($cmdArgs -join " "))
        & $condaPath run -n py_learn @cmdArgs 2>&1 | Out-Host
        if ($LASTEXITCODE -ne 0) {
            throw "Package install command failed: conda run -n py_learn $($cmdArgs -join ' ')"
        }
    }

    # Validate that core imports work.
    & $condaPath run -n py_learn python -c "import pandas,numpy,pymupdf,jupyter,openpyxl,requests; print('Package validation passed')" 2>&1 | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "Python package validation failed in py_learn"
    }

    Write-Log "Python packages installed and validated successfully" "SUCCESS"
}

# ========================================
# DT INSTALLATION
# ========================================
function Get-DtExecutablePath {
    return (Join-Path $env:USERPROFILE "Downloads\dt.exe")
}

function Get-DtVersion {
    param([string]$DtExecutable)

    $versionArgs = @(
        @("version"),
        @("-v")
    )

    foreach ($args in $versionArgs) {
        try {
            $out = & $DtExecutable @args 2>&1
            if ($LASTEXITCODE -eq 0 -and $out) {
                return ($out | Select-Object -First 1)
            }
        }
        catch {
            # Try the next supported variant.
        }
    }

    try {
        $helpOut = & $DtExecutable --help 2>&1
        if ($LASTEXITCODE -eq 0 -or $helpOut) {
            return "installed"
        }
    }
    catch {
        return $null
    }

    return $null
}

function Install-Dt {
    $dtUrl = "https://gfx-assets.intel.com/artifactory/gfx-build-assets/build-tools/devtool-go/latest/artifacts/win64/dt.exe"
    $dtExe = Get-DtExecutablePath
    $downloadsDir = Split-Path -Parent $dtExe

    New-Item -ItemType Directory -Path $downloadsDir -Force | Out-Null

    Write-Log "Downloading Intel dt from: $dtUrl"
    Invoke-WebRequest -Uri $dtUrl -OutFile $dtExe -UseBasicParsing

    if (-not (Test-Path $dtExe)) {
        throw "dt.exe download failed"
    }

    Write-Log "Running dt install from: $dtExe"
    & $dtExe install 2>&1 | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "dt install failed with exit code $LASTEXITCODE"
    }

    Write-Log "Intel dt installed successfully" "SUCCESS"
}

function Setup-Dt {
    $dtExe = Get-DtExecutablePath
    if (-not (Test-Path $dtExe)) {
        throw "dt.exe not found at $dtExe"
    }

    Write-Log "Preparing environment for mandatory dt setup"
    $env:http_proxy = $Proxy
    $env:https_proxy = $Proxy

    if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
        throw "GitHub CLI not found. dt setup requires GitHub authentication context."
    }

    try {
        gh auth status 2>&1 | Out-Null
    }
    catch {
        throw "GitHub CLI is not authenticated. Run 'gh auth login' and re-run the script for mandatory dt setup."
    }

    Write-Log "Running mandatory command: dt setup"
    & $dtExe setup 2>&1 | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "dt setup failed with exit code $LASTEXITCODE"
    }

    Write-Log "dt setup completed successfully" "SUCCESS"
}

function Verify-Setup {
    if (-not $SkipPython) {
        $condaPath = Resolve-CondaExecutable
        if (-not $condaPath) {
            throw "Conda executable not found for verification"
        }

        Write-Log "Verification 1/3: Checking Python version in py_learn"
        $pyVersion = (& $condaPath run -n py_learn python --version 2>&1 | Select-Object -First 1)
        $pyVersion | Out-Host
        if (-not ($pyVersion -match "Python 3\.11\.")) {
            throw "Expected Python 3.11.x in py_learn, got: $pyVersion"
        }
    }

    if (-not $SkipPythonPackages) {
        $condaPath = Resolve-CondaExecutable
        if (-not $condaPath) {
            throw "Conda executable not found for package verification"
        }

        Write-Log "Verification 2/3: Testing core Python package imports"
        & $condaPath run -n py_learn python -c "import pandas,numpy,pymupdf; print('Success!')" 2>&1 | Out-Host
        if ($LASTEXITCODE -ne 0) {
            throw "Python package import verification failed"
        }
    }

    if (-not $SkipDt) {
        $dtExe = Get-DtExecutablePath
        if (-not (Test-Path $dtExe)) {
            throw "dt.exe not found for verification at $dtExe"
        }

        Write-Log "Verification 3/3: Checking dt availability"
        $dtVersion = Get-DtVersion -DtExecutable $dtExe
        if (-not $dtVersion) {
            throw "Unable to verify dt installation from $dtExe"
        }
        Write-Log "  dt detected: $dtVersion"
    }

    Write-Log "All applicable verification checks completed successfully" "SUCCESS"
}

# ========================================
# WINGET INSTALLATION WITH VERSION CHECKING
# ========================================
function Winget-Install {
    param([string]$Id)

    # Report currently installed version
    $currentVer = Get-WingetInstalledVersion -Id $Id
    if ($currentVer) {
        Write-Log "  Currently installed: $Id v$currentVer"
        
        # For most tools, we'll let winget handle version management
        # Only show warning for major version differences
        Write-Log "  Proceeding with winget install/update (winget will handle version management)"
    } else {
        Write-Log "  $Id is not currently installed"
    }

    Write-Log "  Installing / updating $Id via winget..."
    $result = winget install -e --id $Id --silent --accept-source-agreements --accept-package-agreements 2>&1
    $exitCode = $LASTEXITCODE

    # Display output
    $result | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }

    if ($exitCode -eq 0) {
        $newVer = Get-WingetInstalledVersion -Id $Id
        Write-Log "  $Id installed successfully$(if ($newVer) { " (v$newVer)" })" "SUCCESS"
    }
    elseif ($exitCode -eq -1978335189 -or                          # APPINSTALLER_CLI_ERROR_UPDATE_NOT_APPLICABLE
            ($result -match "already installed") -or
            ($result -match "No applicable update")) {
        Write-Log "  $Id is already at the latest version$(if ($currentVer) { " (v$currentVer)" })"
    }
    else {
        Write-Log "  winget returned exit code $exitCode for $Id" "ERROR"
        throw "winget install failed for $Id (exit code $exitCode)"
    }
}

# ========================================
# VS CODE CONFIGURATION  (B-02 fix + B-08 PS 5.1 compat)
# ========================================
function Get-CodeCmd {
    $cmd = (Get-Command code -ErrorAction SilentlyContinue)
    if ($cmd) { return $cmd.Source }

    # B-02 fix: avoid ${env:ProgramFiles(x86)} syntax
    $progX86 = [Environment]::GetFolderPath("ProgramFilesX86")
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA "Programs\Microsoft VS Code\bin\code.cmd"),
        (Join-Path $env:ProgramFiles "Microsoft VS Code\bin\code.cmd")
    )
    if ($progX86) {
        $candidates += (Join-Path $progX86 "Microsoft VS Code\bin\code.cmd")
    }

    $found = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    return $found
}

function ConvertFrom-JsonToHashtable {
    # B-08 fix: PS-5.1-compatible JSON-to-hashtable conversion
    param([string]$Json)
    $obj = $Json | ConvertFrom-Json
    $ht  = @{}
    foreach ($prop in $obj.PSObject.Properties) {
        $ht[$prop.Name] = $prop.Value
    }
    return $ht
}

function Set-VSCodeProxySetting {
    param([string]$ProxyValue)

    $settingsDir  = Join-Path $env:APPDATA "Code\User"
    $settingsFile = Join-Path $settingsDir "settings.json"
    New-Item -ItemType Directory -Path $settingsDir -Force | Out-Null

    $settings = @{}
    if (Test-Path $settingsFile) {
        try {
            $raw = Get-Content $settingsFile -Raw -ErrorAction Stop
            if ($raw.Trim()) {
                $settings = ConvertFrom-JsonToHashtable $raw    # B-08 fix
            }
        }
        catch {
            Write-Log "Existing VS Code settings.json isn't valid JSON; backing it up and recreating." "WARN"
            Copy-Item $settingsFile "$settingsFile.bak_$stamp" -Force
            $settings = @{}
        }
    }

    $settings["http.proxy"]        = $ProxyValue
    $settings["http.proxySupport"] = "on"

    ($settings | ConvertTo-Json -Depth 10) | Set-Content -Path $settingsFile -Encoding UTF8
    Write-Log "VS Code proxy configured in $settingsFile"
}

function Install-VSCodeExtensions {
    $codeCmd = Get-CodeCmd
    if (-not $codeCmd) {
        Write-Log "VS Code command-line 'code' not found. Extensions install will be skipped." "WARN"
        throw "VS Code CLI not found"
    }

    Write-Log "Installing VS Code extensions..."

    $extensions = @(
        "GitHub.copilot",
        "GitHub.copilot-chat",
        "ms-python.python",
        "ms-toolsai.jupyter"
    )

    foreach ($ext in $extensions) {
        Write-Log "  Installing extension: $ext"
        & $codeCmd --install-extension $ext --force 2>&1 | Out-Null
    }

    Write-Log "VS Code extensions installed"
}

# ========================================
# GITHUB AUTHENTICATION  (B-07 fix: no retry for interactive flow)
# ========================================
function GitHub-Auth {
    if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
        Write-Log "GitHub CLI not found; skipping gh auth." "WARN"
        throw "GitHub CLI not found"
    }

    Write-Log "Checking GitHub CLI auth status..."
    $authOk = $false
    try {
        $status = gh auth status 2>&1
        $status | Out-Host
        Write-Log "GitHub CLI authentication verified"
        $authOk = $true
    }
    catch {
        Write-Log "" "WARN"
        Write-Host "  +=========================================================+" -ForegroundColor Yellow
        Write-Host "  |  GitHub CLI is not authenticated.                       |" -ForegroundColor Yellow
        Write-Host "  |  A browser / device-code flow will open.                |" -ForegroundColor Yellow
        Write-Host "  |  Complete it with your GitHub account to continue.      |" -ForegroundColor Yellow
        Write-Host "  +=========================================================+" -ForegroundColor Yellow

        # Do not pipe interactive gh login output; piping can break TTY behavior.
        & gh auth login
        if ($LASTEXITCODE -ne 0) {
            throw "gh auth login failed with exit code $LASTEXITCODE"
        }
    }

    try {
        gh config set -h github.com git_protocol https 2>&1 | Out-Null
        Write-Log "GitHub CLI configured to use HTTPS for git operations."
    }
    catch {
        Write-Log "Could not set GitHub CLI git protocol (non-critical)" "WARN"
    }
}

# ========================================
# VERSION MATRIX  (INF-05 / P1-07)
# ========================================
function Show-VersionMatrix {
    Write-Banner "INSTALLED TOOL VERSIONS"

    $tools = @(
        @{ Name = "VS Code";    Cmd = "code"   },
        @{ Name = "Conda";      Cmd = "conda"  },
        @{ Name = "Python (py_learn)"; Cmd = "python-py_learn" },
        @{ Name = "Git";        Cmd = "git"    },
        @{ Name = "GitHub CLI"; Cmd = "gh"     },
        @{ Name = "pip (py_learn)"; Cmd = "pip-py_learn" },
        @{ Name = "Intel dt";   Cmd = "dt"     },
        @{ Name = "WinGet";     Cmd = "winget" }
    )

    Write-Host ("  {0,-18} {1}" -f "Tool", "Detected Version") -ForegroundColor Cyan
    Write-Host ("  {0,-18} {1}" -f "----", "----------------") -ForegroundColor Cyan

    foreach ($t in $tools) {
        $ver = Get-InstalledVersion -Tool $t.Cmd
        if (-not $ver -or $ver -eq "not installed") {
            Write-Host ("  {0,-18} {1}" -f $t.Name, "not installed") -ForegroundColor DarkGray
        } else {
            Write-Host ("  {0,-18} {1}" -f $t.Name, $ver) -ForegroundColor White
        }
    }
    
    # Show py_learn environment status
    if (Test-CondaEnvironmentExists -EnvName "py_learn") {
        Write-Host ("  {0,-18} {1}" -f "py_learn env", "exists") -ForegroundColor Green
    } else {
        Write-Host ("  {0,-18} {1}" -f "py_learn env", "not found") -ForegroundColor DarkGray
    }
    
    Write-Host ""
}

# ========================================
# SUMMARY GENERATION  (enhanced with per-step timing + log path repeat)
# ========================================
function Show-Summary {
    param([bool]$Success)

    $duration    = (Get-Date) - $script:StartTime
    $durationStr = "{0:hh\:mm\:ss}" -f $duration

    Write-Banner "SETUP SUMMARY"

    $statusColor = if ($Success) { "Green" } else { "Red" }
    $statusText  = if ($Success) { "SUCCESS" } else { "FAILED" }
    Write-Host "  Status   : $statusText" -ForegroundColor $statusColor
    Write-Host "  Duration : $durationStr" -ForegroundColor White
    Write-Host "  Started  : $($script:StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor White
    Write-Host "  Ended    : $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor White
    Write-Host ""

    # --- Per-step timing table (INF-01 / P1-06) ---
    if ($script:StepTimings.Count -gt 0) {
        Write-Host ("  {0,-50} {1,-10} {2}" -f "Step", "Result", "Time") -ForegroundColor Cyan
        Write-Host ("  {0,-50} {1,-10} {2}" -f ("-" * 50), ("-" * 10), ("-" * 8)) -ForegroundColor Cyan
        foreach ($st in $script:StepTimings) {
            $resColor = switch ($st.Result) {
                "SUCCESS" { "Green"  }
                "SKIPPED" { "Yellow" }
                "DRYRUN"  { "DarkGray" }
                default   { "Red"    }
            }
            $ts = "{0:mm\:ss}" -f $st.Elapsed
            Write-Host ("  {0,-50} {1,-10} {2}" -f $st.Step, $st.Result, $ts) -ForegroundColor $resColor
        }
        Write-Host ""
    }

    # --- Failed steps detail (INF-08: distinguish optional vs mandatory) ---
    if ($script:FailedSteps.Count -gt 0) {
        Write-Host "  FAILED STEPS ($($script:FailedSteps.Count)):" -ForegroundColor Red
        foreach ($f in $script:FailedSteps) {
            Write-Host "    [X] $($f.Step)" -ForegroundColor Red
            Write-Host "       Error    : $($f.Error)" -ForegroundColor DarkRed
            Write-Host "       Attempts : $($f.Attempt)" -ForegroundColor DarkRed
        }
        Write-Host ""
    }

    if ($script:SkippedSteps.Count -gt 0) {
        Write-Host "  SKIPPED STEPS ($($script:SkippedSteps.Count)):" -ForegroundColor Yellow
        foreach ($s in $script:SkippedSteps) {
            Write-Host "    [>] $s" -ForegroundColor Yellow
        }
        Write-Host ""
    }

    # --- Config details ---
    Write-Host "  Configuration:" -ForegroundColor Cyan
    Write-Host "    Proxy            : $Proxy (manual setup only)"
    Write-Host "    NO_PROXY         : $NoProxy"
    Write-Host "    Intel Certs      : $InstallIntelCerts"
    Write-Host "    Build Tools      : $InstallBuildTools"
    Write-Host "    External PyPI    : $UseExternalPyPI"
    Write-Host "    Interactive      : $Interactive"
    Write-Host "    DryRun           : $DryRun"
    Write-Host "    SkipDt           : $SkipDt"
    Write-Host ""

    # --- Log file path (repeated here per UF-06 / P1-08) ---
    Write-Host "  Log File : $LogFile" -ForegroundColor Cyan
    Write-Host ""

    # --- Version matrix ---
    if (-not $DryRun) {
        Show-VersionMatrix
    }

    if ($Success) {
        Write-Banner "NEXT STEPS"
        Write-Host ""
        Write-Host "  Core verification (Python env, packages, dt) was completed automatically." -ForegroundColor Green
        Write-Host ""
        Write-Host "  1. GIT PROXY SETUP (if needed):" -ForegroundColor Yellow
        Write-Host "     • In regular Command Prompt or PowerShell:" -ForegroundColor White
        Write-Host "       git config --global http.proxy $Proxy" -ForegroundColor DarkGray
        Write-Host "       git config --global https.proxy $Proxy" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  2. GITHUB COPILOT SETUP:" -ForegroundColor Yellow
        Write-Host "     • Ensure you have GitHub Copilot entitlement" -ForegroundColor White
        Write-Host "     • Complete required training and onboarding" -ForegroundColor White
        Write-Host ""
        Write-Host "  3. START CODING:" -ForegroundColor Yellow
        Write-Host "     • Open VS Code: code ." -ForegroundColor White
        Write-Host "     • Sign in to GitHub when prompted by Copilot extensions" -ForegroundColor White
        Write-Host "     • Select py_learn as Python interpreter:" -ForegroundColor White
        Write-Host "       Ctrl+Shift+P > Python: Select Interpreter" -ForegroundColor DarkGray
        Write-Host "     • Verify Copilot status in VS Code status bar" -ForegroundColor White
        Write-Host ""
        Write-Host "  4. QUICK TEST:" -ForegroundColor Yellow
        Write-Host "     • Create a new .py file in VS Code" -ForegroundColor White
        Write-Host "     • Type: # Create a pandas dataframe" -ForegroundColor White
        Write-Host "     • Copilot should suggest code completion" -ForegroundColor White
        Write-Host ""
        Write-Host "  5. PROXY COMMANDS REFERENCE:" -ForegroundColor Yellow
        Write-Host "     • For command sessions:" -ForegroundColor White
        Write-Host "       set http_proxy=$Proxy && set https_proxy=$Proxy" -ForegroundColor DarkGray
        Write-Host "     • For Git operations:" -ForegroundColor White
        Write-Host "       git config --global http.proxy $Proxy" -ForegroundColor DarkGray
        Write-Host ""
    } else {
        Write-Banner "TROUBLESHOOTING"
        Write-Host "  1) Review the error details above" -ForegroundColor Yellow
        Write-Host "  2) Check the full log: $LogFile" -ForegroundColor Yellow
        Write-Host "  3) Verify proxy settings: $Proxy" -ForegroundColor Yellow
        Write-Host "  4) Ensure you're running PowerShell as Administrator" -ForegroundColor Yellow
        Write-Host "  5) Check internet connectivity" -ForegroundColor Yellow
        Write-Host "  6) Re-run this script -- it will resume/retry failed steps" -ForegroundColor Yellow
        Write-Host ""
    }

    Write-Progress -Activity "DevMachine Setup" -Completed -Id 1
}

# ========================================
# MAIN EXECUTION
# ========================================
function Invoke-DevSetup {
    try {
        Write-Banner "DevMachine Setup v6.0 -- VS Code + Copilot + Miniforge + dt"

        Write-Log "Log File : $LogFile"
        Write-Log "Proxy    : $Proxy (for manual configuration only)"
        Write-Log "NO_PROXY : $NoProxy"
        Write-Log "Mode     : $(if ($DryRun) { 'DRY RUN' } elseif ($Interactive) { 'INTERACTIVE' } else { 'AUTOMATIC' })"
        Write-Log "Retries  : $script:MaxRetries per step"
        Write-Log ""

        # Show deprecation warning for SkipNodeJS
        if ($SkipNodeJS) {
            Write-Log "Note: -SkipNodeJS parameter is deprecated (Node.js is no longer installed by this script)" "WARN"
        }

        # --- Pre-flight checks ---
        Invoke-PreFlightChecks

        # --- Build dynamic step list (B-03 fix: no more hard-coded TotalSteps) ---
        $steps = @()
        $steps += "Ensure WinGet is available"
        if ($InstallIntelCerts) { $steps += "Install Intel certificate bundles" }
        if (-not $SkipVSCode)       { $steps += "Install Visual Studio Code" }
        if (-not $SkipPython)       { $steps += "Install Miniforge" }
        if (-not $SkipPython)       { $steps += "Create Python 3.11 environment (py_learn)" }
        if (-not $SkipGit)          { $steps += "Install Git" }
        if (-not $SkipGitHubCLI)    { $steps += "Install GitHub CLI" }
        if ($InstallBuildTools)     { $steps += "Install Visual Studio Build Tools" }
        if (-not $SkipPythonPackages) { $steps += "Install Python packages in py_learn environment" }
        $steps += "Configure VS Code proxy settings"
        if (-not $SkipExtensions)   { $steps += "Install VS Code extensions" }
        if (-not $SkipGitHubAuth)   { $steps += "Configure GitHub CLI authentication" }
        if (-not $SkipDt)           { $steps += "Install Intel dt" }
        if (-not $SkipDt)           { $steps += "Run mandatory dt setup" }
        $steps += "Run automated verification checks"

        $script:TotalSteps = $steps.Count

        # --- Pre-run plan banner (UF-02 / P1-02) ---
        Write-Banner "EXECUTION PLAN ($($steps.Count) steps)"
        for ($idx = 0; $idx -lt $steps.Count; $idx++) {
            Write-Host ("  {0,2}. {1}" -f ($idx + 1), $steps[$idx]) -ForegroundColor White
        }
        Write-Host ""
        Write-Log "Total Steps: $script:TotalSteps" "PROGRESS"
        Write-Log ""
        Write-Log "NOTE: Proxy settings will be configured manually in command sessions" "WARN"
        Write-Log "      No system-level environment variables will be modified" "WARN"
        Write-Log ""

        if (-not $DryRun -and -not $Interactive) {
            Write-Host "  Starting in 5 seconds... (press Ctrl+C to cancel)" -ForegroundColor Yellow
            Start-Sleep -Seconds 5
        }

        # --- Execution policy check ---
        $pol = Get-ExecutionPolicy -Scope CurrentUser
        if ($pol -eq "Restricted") {
            Write-Log "CurrentUser execution policy is Restricted. Setting to RemoteSigned..." "WARN"
            Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
        }

        # Helper: get the next step name for "upcoming" display
        $stepIdx = 0
        function Get-NextStepName {
            if (($stepIdx + 1) -lt $steps.Count) { return $steps[$stepIdx + 1] }
            return ""
        }

        # ------ STEP EXECUTION ------

        $null = Invoke-WithRetry -StepName $steps[$stepIdx] -UpcomingStep (Get-NextStepName) -ScriptBlock {
            Ensure-WinGet
        }
        $stepIdx++

        if ($InstallIntelCerts) {
            $null = Invoke-WithRetry -StepName $steps[$stepIdx] -UpcomingStep (Get-NextStepName) -ScriptBlock {
                Install-IntelCertificateBundles
            }
            $stepIdx++
        }

        if (-not $SkipVSCode) {
            $null = Invoke-WithRetry -StepName $steps[$stepIdx] -UpcomingStep (Get-NextStepName) -ScriptBlock {
                Winget-Install -Id "Microsoft.VisualStudioCode"
            }
            $stepIdx++
        }

        if (-not $SkipPython) {
            $null = Invoke-WithRetry -StepName $steps[$stepIdx] -UpcomingStep (Get-NextStepName) -ScriptBlock {
                Install-Miniforge
            }
            $stepIdx++

            $null = Invoke-WithRetry -StepName $steps[$stepIdx] -UpcomingStep (Get-NextStepName) -ScriptBlock {
                Create-PythonEnvironment
            }
            $stepIdx++
        }

        if (-not $SkipGit) {
            $null = Invoke-WithRetry -StepName $steps[$stepIdx] -UpcomingStep (Get-NextStepName) -ScriptBlock {
                Winget-Install -Id "Git.Git"
            }
            $stepIdx++
        }

        if (-not $SkipGitHubCLI) {
            $null = Invoke-WithRetry -StepName $steps[$stepIdx] -UpcomingStep (Get-NextStepName) -ScriptBlock {
                Winget-Install -Id "GitHub.cli"
            }
            $stepIdx++
        }

        if ($InstallBuildTools) {
            $null = Invoke-WithRetry -StepName $steps[$stepIdx] -UpcomingStep (Get-NextStepName) -Optional -ScriptBlock {
                Winget-Install -Id "Microsoft.VisualStudio.2022.BuildTools"
            }
            $stepIdx++
        }

        # Refresh PATH after installations
        Refresh-Path
        Start-Sleep -Seconds 2

        if (-not $SkipPythonPackages) {
            $null = Invoke-WithRetry -StepName $steps[$stepIdx] -UpcomingStep (Get-NextStepName) -Optional -ScriptBlock {
                Install-PythonPackages
            }
            $stepIdx++
        }

        $null = Invoke-WithRetry -StepName $steps[$stepIdx] -UpcomingStep (Get-NextStepName) -ScriptBlock {
            Set-VSCodeProxySetting -ProxyValue $Proxy
        }
        $stepIdx++

        if (-not $SkipExtensions) {
            $null = Invoke-WithRetry -StepName $steps[$stepIdx] -UpcomingStep (Get-NextStepName) -Optional -ScriptBlock {
                Install-VSCodeExtensions
            }
            $stepIdx++
        }

        # B-07 fix: GitHub auth uses MaxAttempts=1 (no retry for interactive flow)
        if (-not $SkipGitHubAuth) {
            $null = Invoke-WithRetry -StepName $steps[$stepIdx] -UpcomingStep (Get-NextStepName) -MaxAttempts 1 -ScriptBlock {
                GitHub-Auth
            }
            $stepIdx++
        }

        if (-not $SkipDt) {
            $null = Invoke-WithRetry -StepName $steps[$stepIdx] -UpcomingStep (Get-NextStepName) -ScriptBlock {
                Install-Dt
            }
            $stepIdx++

            $null = Invoke-WithRetry -StepName $steps[$stepIdx] -UpcomingStep "" -MaxAttempts 1 -ScriptBlock {
                Setup-Dt
            }
            $stepIdx++
        }

        $null = Invoke-WithRetry -StepName $steps[$stepIdx] -UpcomingStep "" -ScriptBlock {
            Verify-Setup
        }
        $stepIdx++

        Write-Log ""
        Show-Summary -Success $true

    }
    catch {
        Write-Log ""
        Write-Log "FATAL ERROR: $($_.Exception.Message)" "ERROR"
        Write-Log "Stack Trace: $($_.ScriptStackTrace)" "ERROR"
        Write-Log ""
        Show-Summary -Success $false
        throw
    }
    finally {
        try { Stop-Transcript -ErrorAction SilentlyContinue | Out-Null } catch { }
    }
}

# Execute main function
Invoke-DevSetup
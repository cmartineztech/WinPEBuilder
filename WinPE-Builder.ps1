<#
.SYNOPSIS
    Build customized WinPE media (ISO / USB). Version 2.0

.DESCRIPTION
    WinPE Builder 2.0
    -Init / -Mount / -Build / -Save / -Discard / -Clean
    -ISO / -USB (with -Build)
    -WorkDirectory alone resumes a project (no rebuild).

    Project path resolve order (this PowerShell window):
      1) -WorkDirectory
      2) Session from a previous call
      3) Current directory if it looks like a project
      4) Folder containing this script

.EXAMPLE
    .\WinPE-Builder.ps1 -Init -WorkDirectory C:\MyWinPE-Project
    .\WinPE-Builder.ps1 -Build -ISO -Force

.EXAMPLE
    .\WinPE-Builder.ps1 -WorkDirectory C:\MyWinPE-Project
    .\WinPE-Builder.ps1 -Build -ISO -PCA2023 -Force

.EXAMPLE
    # Edit packages.pe (uncomment OCs), then install that list:
    .\WinPE-Builder.ps1 -Build -ISO -Force -AddPackage packages.pe

.NOTES
    Version: 2.0
    Project: https://github.com/cmartinezone/WinPEBuilder
    -AddPackage accepts OC names and/or a packages.pe list file (active lines = install order).

.LINK
    https://github.com/cmartinezone/WinPEBuilder
#>

#Requires -Version 5.1
#Requires -RunAsAdministrator

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string]$WorkDirectory,
    [ValidateSet('amd64', 'x86', 'arm64')][string]$Architecture = 'amd64',
    [switch]$Init,
    [switch]$Mount,
    [switch]$Build,
    # OC names and/or a packages.pe list file (uncommented lines = install set/order)
    [string[]]$AddPackage,
    [switch]$ISO,
    [ValidateScript({
            if ([string]::IsNullOrWhiteSpace($_)) { return $true }
            if ($_ -notmatch '^[A-Za-z]:$') { throw 'USB must look like F:' }
            $true
        })][string]$USB,
    [switch]$PCA2023,
    [switch]$Clean,
    [ValidateSet(32, 64, 128, 256, 512)][int]$ScratchSpace = 256,
    [Alias('CustomBootWimPath')][string]$BootWimPath,
    [switch]$Save,
    [switch]$Discard,
    [switch]$Force,
    [switch]$NoDefaultPackages,
    [ValidatePattern('^[a-z]{2}-[A-Za-z]{2}$')][string]$Language = 'en-us',
    [ValidateNotNullOrEmpty()][string]$TimeZone = 'Eastern Standard Time',
    # Host / WPF: suppress console UI; result object (+ log file) still returned
    [Alias('NonInteractive')][switch]$Quiet
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module Dism -ErrorAction Stop

$script:Version = '2.0'
$script:QuietMode = [bool]$Quiet
$script:LastPreflightError = $null
$script:Started = Get-Date
$script:ScriptRoot = if ($PSScriptRoot) {
    $PSScriptRoot
}
else {
    Split-Path -Parent $MyInvocation.MyCommand.Path
}
$script:ScriptRoot = [IO.Path]::GetFullPath($script:ScriptRoot).TrimEnd('\')
$script:WorkDirectorySource = 'ScriptRoot'

# Dependency-safe install order (Microsoft WinPE OC reference).
# PowerShell stack: WMI > NetFx > Scripting > PowerShell > cmdlet packages.
$DefaultPackages = @(
    'WinPE-WMI',
    'WinPE-NetFx',
    'WinPE-Scripting',
    'WinPE-PowerShell',
    'WinPE-DismCmdlets',
    'WinPE-StorageWMI',
    'WinPE-SecureBootCmdlets',
    'WinPE-HTA',
    'WinPE-FMAPI',
    'WinPE-SecureStartup',
    'WinPE-EnhancedStorage'
)

# Foundation packages required for PowerShell and WMI execution in WinPE.
$script:MandatoryFoundationPackages = @(
    'WinPE-WMI',
    'WinPE-NetFx',
    'WinPE-Scripting',
    'WinPE-PowerShell'
)

# --- project path / session ---

function Assert-NoCmdMetacharacters {
    # Reject cmd.exe metacharacters before passing paths or arguments to ADK batch files.
    param(
        [string]$Value,
        [string]$ParameterName = 'Path'
    )
    if ([string]::IsNullOrEmpty($Value)) { return }
    if ($Value -match '[&|<>^%!\r\n]') {
        throw "Refused $ParameterName containing unsafe cmd.exe metacharacters ([&|<>^%!]): $Value"
    }
}

function Test-WinPeProjectRoot {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    $full = Resolve-FullPath $Path
    if (-not $full) { return $false }
    if (-not (Test-Path -LiteralPath $full -PathType Container)) { return $false }
    if (Test-Path -LiteralPath (Join-Path $full 'WinPE-Root\media\sources\boot.wim')) { return $true }
    return (
        (Test-Path -LiteralPath (Join-Path $full 'WinPE-Root')) -and
        (Test-Path -LiteralPath (Join-Path $full 'Add-Drivers')) -and
        (Test-Path -LiteralPath (Join-Path $full 'Add-Scripts'))
    )
}

function Resolve-FullPath {
    # Relative paths must use PowerShell location, not .NET Environment.CurrentDirectory.
    # Elevated PowerShell often leaves process CWD as C:\Windows\System32 even after
    # Set-Location / prompt shows Desktop — so [IO.Path]::GetFullPath('.\x') alone is wrong.
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $p = $Path.Trim().TrimEnd('\')
    try {
        if (-not [IO.Path]::IsPathRooted($p)) {
            # Join to PowerShell's location (matches the prompt path)
            $here = (Get-Location -PSProvider FileSystem).ProviderPath
            $p = Join-Path $here $p
        }
        return [IO.Path]::GetFullPath($p).TrimEnd('\')
    }
    catch {
        return $null
    }
}

function Assert-SafeWorkDirectory {
    # Refuse project roots at drive roots (C:\), user profile root, or under system/program directories
    # (%WINDIR%, %SystemRoot%, %ProgramFiles%, %ProgramFiles(x86)%, %ProgramData%).
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    Assert-NoCmdMetacharacters -Value $Path -ParameterName 'WorkDirectory'

    $trimmed = $Path.Trim()
    if ($trimmed -match '^[A-Za-z]:[\\/]?$' -or $trimmed -eq '/' -or $trimmed -eq '\') {
        throw "Refused WorkDirectory at drive root: $Path"
    }

    $normInput = $trimmed.Replace('/', '\').TrimEnd('\')
    $full = $normInput
    try {
        $resolvedFull = [IO.Path]::GetFullPath($trimmed)
        $pathRoot = [IO.Path]::GetPathRoot($resolvedFull)
        if ($pathRoot -and $resolvedFull.TrimEnd('\', '/').Equals($pathRoot.TrimEnd('\', '/'), [StringComparison]::OrdinalIgnoreCase)) {
            throw "Refused WorkDirectory at drive root: $resolvedFull"
        }
        $full = $resolvedFull.Replace('/', '\').TrimEnd('\')
    }
    catch {
        if ($_.Exception.Message -match '^Refused WorkDirectory') { throw }
    }

    if ($env:USERPROFILE) {
        $userProfNorm = $env:USERPROFILE.Trim().Replace('/', '\').TrimEnd('\')
        if ($userProfNorm -and (
                $full.Equals($userProfNorm, [StringComparison]::OrdinalIgnoreCase) -or
                $normInput.Equals($userProfNorm, [StringComparison]::OrdinalIgnoreCase)
            )) {
            throw "Refused WorkDirectory at USERPROFILE root: $Path"
        }
    }

    $win = if ($env:SystemRoot) { $env:SystemRoot } elseif ($env:WINDIR) { $env:WINDIR } else { 'C:\Windows' }
    $rawBlockedPrefixes = @(
        $win,
        $env:WINDIR,
        $env:SystemRoot,
        'C:\Windows',
        (Join-Path $win 'System32'),
        'C:\Windows\System32',
        $env:ProgramFiles,
        'C:\Program Files',
        ${env:ProgramFiles(x86)},
        'C:\Program Files (x86)',
        $env:ProgramData,
        'C:\ProgramData'
    )

    $blocked = New-Object System.Collections.Generic.List[string]
    foreach ($candidate in $rawBlockedPrefixes) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        $cNorm = $candidate.Trim().Replace('/', '\').TrimEnd('\')
        if ($cNorm -and -not $blocked.Contains($cNorm)) {
            [void]$blocked.Add($cNorm)
        }
        try {
            $cFull = [IO.Path]::GetFullPath($candidate).Replace('/', '\').TrimEnd('\')
            if ($cFull -and -not $blocked.Contains($cFull)) {
                [void]$blocked.Add($cFull)
            }
        }
        catch { }
    }

    foreach ($b in $blocked) {
        foreach ($target in @($full, $normInput)) {
            if ([string]::IsNullOrWhiteSpace($target)) { continue }
            if ($target.Equals($b, [StringComparison]::OrdinalIgnoreCase) -or
                $target.StartsWith($b + '\', [StringComparison]::OrdinalIgnoreCase)) {
                throw @"
Refused WorkDirectory under protected system directory ($b):
  $target

That is not a safe project folder. Elevated PowerShell often starts in System32;
if the path was relative, it may have resolved there by mistake.

Use a dedicated project path outside system folders, for example:

  .\WinPE-Builder.ps1 -Init -WorkDirectory C:\MyWinPE-Project
"@
            }
        }
    }
}

function Get-SessionWorkDirectory {
    if (-not (Get-Variable -Name WinPEBuilderWorkDirectory -Scope Global -ErrorAction SilentlyContinue)) {
        return $null
    }
    $raw = $global:WinPEBuilderWorkDirectory
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    return (Resolve-FullPath $raw)
}

function Set-SessionWorkDirectory {
    param([string]$Path)
    $global:WinPEBuilderWorkDirectory = Resolve-FullPath $Path
}

function Resolve-WorkDirectoryPath {
    param([string]$RequestedPath, [bool]$PathWasBound)

    if ($PathWasBound -and -not [string]::IsNullOrWhiteSpace($RequestedPath)) {
        $full = Resolve-FullPath $RequestedPath
        if (-not $full) { throw "Invalid WorkDirectory path: $RequestedPath" }
        return [pscustomobject]@{ Path = $full; Source = 'Parameter' }
    }

    $session = Get-SessionWorkDirectory
    if ($session -and (Test-Path -LiteralPath $session -PathType Container)) {
        return [pscustomobject]@{ Path = $session; Source = 'Session' }
    }

    $cwd = Resolve-FullPath (Get-Location).Path
    if ($cwd -and (Test-WinPeProjectRoot -Path $cwd)) {
        return [pscustomobject]@{ Path = $cwd; Source = 'CurrentDirectory' }
    }

    return [pscustomobject]@{ Path = $script:ScriptRoot; Source = 'ScriptRoot' }
}

$resolved = Resolve-WorkDirectoryPath `
    -RequestedPath $WorkDirectory `
    -PathWasBound:$PSBoundParameters.ContainsKey('WorkDirectory')
$WorkDirectory = $resolved.Path
$script:WorkDirectorySource = $resolved.Source
Assert-SafeWorkDirectory -Path $WorkDirectory
if ($BootWimPath) {
    Assert-NoCmdMetacharacters -Value $BootWimPath -ParameterName 'BootWimPath'
}
Set-SessionWorkDirectory -Path $WorkDirectory

# Project layout
$DirDrivers = Join-Path $WorkDirectory 'Add-Drivers'
$DirScripts = Join-Path $WorkDirectory 'Add-Scripts'
$DirUpdates = Join-Path $WorkDirectory 'Add-Updates'
$DirIso = Join-Path $WorkDirectory 'WinPE-ISO'
$DirPeRoot = Join-Path $WorkDirectory 'WinPE-Root'
$DirMount = Join-Path $DirPeRoot 'mount'
$DirMedia = Join-Path $DirPeRoot 'media'
$DirLogs = Join-Path $WorkDirectory 'WinPE-Logs'
$FileBootWim = Join-Path $DirPeRoot 'media\sources\boot.wim'

if (-not $WhatIfPreference) {
    if (-not (Test-Path -LiteralPath $DirLogs)) {
        New-Item -ItemType Directory -Path $DirLogs -Force | Out-Null
    }
    $script:LogFile = Join-Path $DirLogs ("WinPEBuilder_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))
}
else {
    $script:LogFile = $null
}

# --- console (respect -Quiet for WPF / host apps; always log to file) ---

function Write-Ui {
    param([string]$Message, $Color = 'Gray')
    if ($script:QuietMode) { return }
    Write-Host $Message -ForegroundColor $Color
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Ok', 'Warn', 'Error', 'Step')][string]$Level = 'Info'
    )
    $tag = @{ Info = 'INFO'; Ok = 'OK'; Warn = 'WARN'; Error = 'ERROR'; Step = 'STEP' }[$Level]
    $color = @{ Info = 'Cyan'; Ok = 'Green'; Warn = 'Yellow'; Error = 'Red'; Step = 'White' }[$Level]
    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'HH:mm:ss'), $tag, $Message
    Write-Ui $line $color
    if ($script:LogFile) { Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 }
}

function Write-Banner {
    param([string]$Action)
    if ($script:QuietMode) { return }
    $srcLabel = switch ($script:WorkDirectorySource) {
        'Parameter' { 'from -WorkDirectory' }
        'Session' { 'this PowerShell session' }
        'CurrentDirectory' { 'current folder (project detected)' }
        'ScriptRoot' { 'script folder' }
        default { $script:WorkDirectorySource }
    }
    Write-Host ''
    Write-Host '========================================' -ForegroundColor DarkGray
    Write-Host " WinPE Builder $script:Version" -ForegroundColor White
    Write-Host " Action : $Action" -ForegroundColor Cyan
    Write-Host " Folder : $WorkDirectory" -ForegroundColor Cyan
    Write-Host "         ($srcLabel)" -ForegroundColor DarkGray
    Write-Host " Arch   : $Architecture" -ForegroundColor Cyan
    Write-Host " Log    : $script:LogFile" -ForegroundColor DarkGray
    Write-Host '========================================' -ForegroundColor DarkGray
    Write-Host ''
}

function Write-Phase {
    param([string]$Title, [string]$Hint = '')
    if ($script:QuietMode) { return }
    Write-Host ''
    Write-Host ">> $Title" -ForegroundColor Yellow
    if ($Hint) { Write-Host "   $Hint" -ForegroundColor DarkGray }
}

function Write-Done {
    param([string]$Summary)
    if ($script:QuietMode) { return }
    Write-Host ''
    Write-Host "DONE: $Summary" -ForegroundColor Green
    Write-Host ''
}

function New-BuildResult {
    param(
        [bool]$Ok,
        [string]$Action,
        [string]$Err = $null,
        [string]$Iso = $null,
        [string]$Usb = $null,
        [bool]$Mounted = $false,
        [int]$Pkgs = 0,
        [int]$Drvs = 0,
        [int]$Upds = 0
    )
    [pscustomobject]@{
        Success       = $Ok
        Action        = $Action
        Version       = $script:Version
        WorkDirectory = $WorkDirectory
        Architecture  = $Architecture
        BootWim       = $FileBootWim
        MountPath     = $DirMount
        IsMounted     = $Mounted
        ISOPath       = $Iso
        USBDrive      = $Usb
        PackagesAdded = $Pkgs
        DriversAdded  = $Drvs
        UpdatesAdded  = $Upds
        LogPath       = $script:LogFile
        Started       = $script:Started
        Finished      = Get-Date
        Error         = $Err
    }
}

# --- ADK ---

function Get-AdkInfo {
    $kits = $null
    foreach ($k in @(
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows Kits\Installed Roots',
            'HKLM:\SOFTWARE\Microsoft\Windows Kits\Installed Roots'
        )) {
        if (Test-Path $k) {
            $kits = (Get-ItemProperty $k -ErrorAction SilentlyContinue).KitsRoot10
            if ($kits) { break }
        }
    }
    if (-not $kits) { $kits = 'C:\Program Files (x86)\Windows Kits\10' }
    $adk = Join-Path $kits 'Assessment and Deployment Kit'
    $pe = Join-Path $adk 'Windows Preinstallation Environment'
    [pscustomobject]@{
        EnvBat    = Join-Path $adk 'Deployment Tools\DandISetEnv.bat'
        Copype    = Join-Path $pe 'copype.cmd'
        MakeMedia = Join-Path $pe 'MakeWinPEMedia.cmd'
        OcRoot    = Join-Path $pe "$Architecture\WinPE_OCs"
        PeArch    = Join-Path $pe $Architecture
    }
}

function Test-AdkReady {
    param($Adk, [switch]$RequireWim)

    $adkUrl = 'https://learn.microsoft.com/en-us/windows-hardware/get-started/adk-install'

    # Deployment Tools (base ADK)
    $deployMissing = @()
    if (-not (Test-Path -LiteralPath $Adk.EnvBat)) { $deployMissing += $Adk.EnvBat }

    # WinPE add-on (copype, MakeWinPEMedia, architecture OCs)
    $peMissing = @()
    foreach ($p in @($Adk.Copype, $Adk.MakeMedia, $Adk.OcRoot, $Adk.PeArch)) {
        if (-not (Test-Path -LiteralPath $p)) { $peMissing += $p }
    }

    if ($deployMissing -or $peMissing) {
        $msg = @()
        $msg += 'Required software is not installed (or the install is incomplete).'
        $msg += ''

        if ($deployMissing -and $peMissing) {
            $msg += 'Not installed: Windows ADK  AND  Windows PE add-on'
        }
        elseif ($deployMissing) {
            $msg += 'Not installed: Windows ADK (Deployment Tools)'
        }
        else {
            $msg += 'Not installed: Windows PE add-on for the ADK'
            $msg += '(The main ADK is present, but the WinPE add-on is missing.)'
        }
        $msg += ''

        if ($deployMissing) {
            $msg += 'ADK files not found:'
            foreach ($p in $deployMissing) { $msg += "  - $p" }
            $msg += ''
        }
        if ($peMissing) {
            $msg += 'WinPE add-on files not found:'
            foreach ($p in $peMissing) { $msg += "  - $p" }
            $msg += ''
        }

        $msg += 'What to do:'
        $msg += '  1. Open this page in your browser:'
        $msg += "     $adkUrl"
        $msg += '  2. Download and install Windows ADK'
        $msg += '  3. Download and install Windows PE add-on for the ADK'
        $msg += '     (same release / same Windows version as the ADK)'
        $msg += '  4. Close PowerShell, open a new Administrator PowerShell'
        $msg += '  5. Run WinPE-Builder.ps1 again'
        $msg += ''
        $msg += 'You need BOTH installs. The WinPE add-on is a separate download on that page.'

        throw ($msg -join "`n")
    }

    if ($RequireWim -and -not (Test-Path -LiteralPath $FileBootWim)) {
        throw "boot.wim not found. Run -Init first: $FileBootWim"
    }
}

function Test-BootExSupport {
    param($Adk)
    $adkUrl = 'https://learn.microsoft.com/en-us/windows-hardware/get-started/adk-install'
    if ((Get-Content $Adk.MakeMedia -Raw -ErrorAction Stop) -notmatch 'bootex') {
        throw @"
Your ADK is too old for -PCA2023 (MakeWinPEMedia has no /bootex).

What to do:
  1. Open: $adkUrl
  2. Install a current Windows ADK + Windows PE add-on (Dec 2024 or later)
  3. Open a new Administrator PowerShell and run the command again

Or omit -PCA2023 to build with the default (PCA 2011) signing.
"@
    }
}

function Invoke-AdkPreflight {
    <#
      First-gate check: ADK Deployment Tools + WinPE add-on.
      Silent when OK. Only prints and stops when something is missing.
      Does not require boot.wim (that is checked later for Mount/Build).
    #>
    try {
        $adk = Get-AdkInfo
        Test-AdkReady -Adk $adk
        return $adk
    }
    catch {
        $errText = $_.Exception.Message
        if ($script:LogFile) {
            Add-Content -Path $script:LogFile -Value ("[PREFLIGHT] " + $errText) -Encoding UTF8
        }
        if (-not $script:QuietMode) {
            Write-Host ''
            Write-Host '========================================' -ForegroundColor DarkGray
            Write-Host " WinPE Builder $script:Version" -ForegroundColor White
            Write-Host '========================================' -ForegroundColor DarkGray
            Write-Host ' Windows ADK / WinPE add-on is not installed.' -ForegroundColor Red
            Write-Host ''
            Write-Host $errText -ForegroundColor Yellow
            Write-Host ''
            Write-Host "Log: $script:LogFile" -ForegroundColor DarkGray
            Write-Host ''
        }
        # Stash for host consumers when Quiet
        $script:LastPreflightError = $errText
        return $null
    }
}

function Invoke-AdkCommand {
    # Progress lines often go to stderr — do not treat as terminating errors
    param(
        [string]$Tool,
        [string]$Arguments,
        [string]$EnvBat,
        [string]$Name,
        [string]$WaitHint = ''
    )
    Assert-NoCmdMetacharacters -Value $WorkDirectory -ParameterName 'WorkDirectory'
    Assert-NoCmdMetacharacters -Value $BootWimPath -ParameterName 'BootWimPath'
    Assert-NoCmdMetacharacters -Value $EnvBat -ParameterName 'EnvBat'
    Assert-NoCmdMetacharacters -Value $Tool -ParameterName 'Tool'
    Assert-NoCmdMetacharacters -Value $Arguments -ParameterName 'Arguments'

    $cmd = 'call "{0}" && call "{1}" {2}' -f $EnvBat, $Tool, $Arguments
    Write-Log "Running $Name ..." -Level Step
    if ($WaitHint) { Write-Ui "   $WaitHint" 'DarkYellow' }
    Write-Ui '   (tool output below - please wait)' 'DarkGray'

    $prevEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $exitCode = 0
    try {
        $wrapped = '{0} & call echo __WPB_EXIT__%ERRORLEVEL%' -f $cmd
        cmd.exe /c $wrapped 2>&1 | ForEach-Object {
            $text = if ($_ -is [System.Management.Automation.ErrorRecord]) {
                $_.ToString()
            }
            else {
                "$_"
            }
            if (-not $text) { return }
            if ($text -match '^__WPB_EXIT__(\d+)\s*$') {
                $exitCode = [int]$Matches[1]
                return
            }
            Write-Ui "   | $text" 'DarkGray'
            if ($script:LogFile) { Add-Content -Path $script:LogFile -Value "  $text" -Encoding UTF8 }
        }
    }
    finally {
        $ErrorActionPreference = $prevEap
    }

    if ($exitCode -ne 0) {
        throw "$Name failed (exit $exitCode). See log: $script:LogFile"
    }
    Write-Log "$Name finished." -Level Ok
}

function Invoke-DismOnImage {
    param([string]$MountPath, [string[]]$DismArgs, [string]$Name)
    $cleanMount = $MountPath.Trim('"')
    $quotedArgs = New-Object System.Collections.Generic.List[string]
    [void]$quotedArgs.Add(('/Image:"{0}"' -f $cleanMount))
    foreach ($a in @($DismArgs)) {
        if ([string]::IsNullOrWhiteSpace($a)) { continue }
        if ($a -match '^(/[^:]+):(.+)$') {
            $opt = $Matches[1]
            $val = $Matches[2].Trim('"')
            if ($val -match '\s') {
                [void]$quotedArgs.Add(('{0}:"{1}"' -f $opt, $val))
            }
            else {
                [void]$quotedArgs.Add(('{0}:{1}' -f $opt, $val))
            }
        }
        elseif ($a -match '\s' -and $a -notmatch '^".*"$') {
            [void]$quotedArgs.Add(('"{0}"' -f $a.Trim('"')))
        }
        else {
            [void]$quotedArgs.Add($a)
        }
    }
    $argLine = $quotedArgs -join ' '
    Write-Log "dism $Name ($argLine) ..." -Level Info
    $tmpOut = [IO.Path]::GetTempFileName()
    $tmpErr = [IO.Path]::GetTempFileName()
    try {
        $p = Start-Process -FilePath 'dism.exe' -ArgumentList $argLine -Wait -PassThru -NoNewWindow `
            -RedirectStandardOutput $tmpOut -RedirectStandardError $tmpErr
        foreach ($tmpFile in @($tmpOut, $tmpErr)) {
            if ($tmpFile -and (Test-Path -LiteralPath $tmpFile)) {
                foreach ($line in @(Get-Content -LiteralPath $tmpFile -ErrorAction SilentlyContinue)) {
                    if ([string]::IsNullOrWhiteSpace($line)) { continue }
                    Write-Ui "   | $line" 'DarkGray'
                    if ($script:LogFile) {
                        Add-Content -Path $script:LogFile -Value "  [dism] $line" -Encoding UTF8
                    }
                }
            }
        }
        if ($null -eq $p) { throw "dism $Name failed to start" }
        if ($p.ExitCode -ne 0) { throw "dism $Name failed (exit $($p.ExitCode))" }
    }
    finally {
        foreach ($tmpFile in @($tmpOut, $tmpErr)) {
            if ($tmpFile -and (Test-Path -LiteralPath $tmpFile)) {
                Remove-Item -LiteralPath $tmpFile -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

# --- mounts ---

function Get-ActiveMounts {
    $root = [IO.Path]::GetFullPath($DirPeRoot).TrimEnd('\')
    @(Get-WindowsImage -Mounted -ErrorAction SilentlyContinue) | Where-Object {
        $mPath = [IO.Path]::GetFullPath($_.Path).TrimEnd('\')
        $mPath.Equals($root, [StringComparison]::OrdinalIgnoreCase) -or
        $mPath.StartsWith($root + '\', [StringComparison]::OrdinalIgnoreCase)
    }
}

function Dismount-ProjectImage {
    param([switch]$Save)
    $mounts = @(Get-ActiveMounts)
    if ($mounts.Count -eq 0) { return }
    foreach ($m in $mounts) {
        if ($Save) {
            Dismount-WindowsImage -Path $m.Path -Save -ErrorAction Stop | Out-Null
        }
        else {
            try {
                Dismount-WindowsImage -Path $m.Path -Discard -ErrorAction Stop | Out-Null
            }
            catch {
                Write-Log "Dismount-WindowsImage (-Discard) failed for '$($m.Path)': $($_.Exception.Message). Running Clear-WindowsCorruptMountPoint..." -Level Error
                Clear-WindowsCorruptMountPoint -ErrorAction SilentlyContinue | Out-Null
            }
        }
    }
    $remaining = @(Get-ActiveMounts)
    if ($remaining.Count -gt 0) {
        $remPaths = ($remaining | ForEach-Object { $_.Path }) -join ', '
        throw "Active WIM mount(s) remain under '$DirPeRoot' after dismount attempt: $remPaths"
    }
}

function Assert-NoReparsePointsOrActiveMounts {
    # Prevent recursive deletion of active WIM mount points or NTFS junctions/symlinks.
    param([Parameter(Mandatory = $true)][string]$LiteralPath)
    $fullTarget = [IO.Path]::GetFullPath($LiteralPath).TrimEnd('\')
    $activeMounts = @(Get-ActiveMounts)
    if ($activeMounts.Count -gt 0) {
        $mList = ($activeMounts | ForEach-Object { $_.Path }) -join ', '
        throw "Refused recursive delete of '$fullTarget': active project WIM mount(s) detected ($mList)."
    }
    foreach ($m in @(Get-WindowsImage -Mounted -ErrorAction SilentlyContinue)) {
        if (-not $m.Path) { continue }
        $mPath = [IO.Path]::GetFullPath($m.Path).TrimEnd('\')
        if ($mPath.Equals($fullTarget, [StringComparison]::OrdinalIgnoreCase) -or
            $mPath.StartsWith($fullTarget + '\', [StringComparison]::OrdinalIgnoreCase) -or
            $fullTarget.StartsWith($mPath + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw "Refused recursive delete of '$fullTarget': overlaps active WIM mount '$mPath'."
        }
    }
    if (-not (Test-Path -LiteralPath $fullTarget)) { return }
    $rootItem = Get-Item -LiteralPath $fullTarget -Force -ErrorAction Stop
    if ([bool]($rootItem.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
        throw "Refused recursive delete: reparse point detected at '$($rootItem.FullName)'."
    }
    if ($rootItem.PSIsContainer) {
        $queue = New-Object System.Collections.Generic.Queue[string]
        $queue.Enqueue($rootItem.FullName)
        while ($queue.Count -gt 0) {
            $curr = $queue.Dequeue()
            foreach ($child in @(Get-ChildItem -LiteralPath $curr -Force -ErrorAction Stop)) {
                if ([bool]($child.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
                    throw "Refused recursive delete of '$fullTarget': reparse point detected at '$($child.FullName)'."
                }
                if ($child.PSIsContainer) {
                    $queue.Enqueue($child.FullName)
                }
            }
        }
    }
}

function Clear-ProjectMount {
    param([switch]$Global)
    Write-Log 'Clearing mounts...' -Level Step
    Dismount-ProjectImage
    if ($Global) { Clear-WindowsCorruptMountPoint -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $DirMount) {
        Assert-NoReparsePointsOrActiveMounts -LiteralPath $DirMount
        Get-ChildItem -LiteralPath $DirMount -Force -ErrorAction SilentlyContinue |
        Remove-Item -Recurse -Force -ErrorAction Stop
    }
    Write-Log 'Mounts cleared.' -Level Ok
}

function Mount-BootImage {
    Test-AdkReady -Adk (Get-AdkInfo) -RequireWim
    $open = @(Get-ActiveMounts)
    if ($open.Count -gt 0) {
        Write-Log "Already mounted: $($open[0].Path)" -Level Info
        return [IO.Path]::GetFullPath($open[0].Path).TrimEnd('\')
    }

    if (-not (Test-Path -LiteralPath $DirMount)) {
        New-Item -ItemType Directory -Path $DirMount -Force | Out-Null
    }
    else {
        $leftover = @(Get-ChildItem -LiteralPath $DirMount -Force -ErrorAction SilentlyContinue)
        if ($leftover.Count -gt 0) {
            Write-Log 'Mount folder has leftover files; cleaning before remount...' -Level Warn
            Assert-NoReparsePointsOrActiveMounts -LiteralPath $DirMount
            $leftover | Remove-Item -Recurse -Force -ErrorAction Stop
        }
    }

    Write-Phase 'Mount boot.wim' "ReadWrite -> $DirMount"
    Mount-WindowsImage -ImagePath $FileBootWim -Index 1 -Path $DirMount -ErrorAction Stop | Out-Null
    $path = [IO.Path]::GetFullPath($DirMount).TrimEnd('\')
    Write-Done "Mounted at $path"
    return $path
}

# --- init / customize ---

function New-ProjectFolders {
    # Always create project layout on -Init (including Logs + Root placeholder).
    # copype later rebuilds WinPE-Root with the full ADK tree.
    foreach ($d in @($DirDrivers, $DirScripts, $DirUpdates, $DirIso, $DirLogs, $DirPeRoot)) {
        if (-not (Test-Path -LiteralPath $d)) {
            New-Item -ItemType Directory -Path $d -Force | Out-Null
            Write-Log "Created $d" -Level Info
        }
    }
}

# Optional OCs known to ship as separate cabs under WinPE_OCs (not winre-only).
# Used when ADK is not available to scan at Init time. Real Init prefers scanning OcRoot.
$script:KnownOptionalPackages = @(
    'WinPE-Dot3Svc',
    'WinPE-Fonts-Legacy',
    'WinPE-FontSupport-JA-JP',
    'WinPE-FontSupport-KO-KR',
    'WinPE-FontSupport-WinRE',
    'WinPE-FontSupport-ZH-CN',
    'WinPE-FontSupport-ZH-HK',
    'WinPE-FontSupport-ZH-TW',
    'WinPE-GamingPeripherals',
    'WinPE-HSP-Driver',
    'WinPE-LegacySetup',
    'WinPE-MDAC',
    'WinPE-PlatformId',
    'WinPE-PmemCmdlets',
    'WinPE-PPPoE',
    'WinPE-RNDIS',
    'WinPE-Setup',
    'WinPE-Setup-ASZ',
    'WinPE-Setup-Client',
    'WinPE-Setup-Server',
    'WinPE-WDS-Tools',
    'WinPE-WinReCfg',
    'WinPE-x64-Support'
)

function Get-WinPeOcPackageNames {
    param([string]$OcRoot)
    if (-not $OcRoot -or -not (Test-Path -LiteralPath $OcRoot -PathType Container)) {
        return @()
    }
    return @(
        Get-ChildItem -LiteralPath $OcRoot -Filter '*.cab' -File -ErrorAction SilentlyContinue |
            ForEach-Object { $_.BaseName } |
            Sort-Object -Unique
    )
}

function Get-PackagesPeTemplate {
    param([string]$OcRoot)

    $available = @(Get-WinPeOcPackageNames -OcRoot $OcRoot)
    $defaultLookup = @{}
    foreach ($d in $DefaultPackages) {
        $defaultLookup[$d.ToLowerInvariant()] = $true
    }

    $defaultLines = New-Object System.Collections.Generic.List[string]
    foreach ($name in $DefaultPackages) {
        if ($available.Count -gt 0) {
            $match = $available | Where-Object { $_.Equals($name, [StringComparison]::OrdinalIgnoreCase) } |
                Select-Object -First 1
            if ($match) { [void]$defaultLines.Add([string]$match) }
            else { [void]$defaultLines.Add($name) } # keep default; build skips if missing
        }
        else {
            [void]$defaultLines.Add($name)
        }
    }

    $optionalLines = New-Object System.Collections.Generic.List[string]
    if ($available.Count -gt 0) {
        foreach ($cab in $available) {
            if ($defaultLookup.ContainsKey($cab.ToLowerInvariant())) { continue }
            [void]$optionalLines.Add("# $cab")
        }
    }
    else {
        foreach ($name in $script:KnownOptionalPackages) {
            if ($defaultLookup.ContainsKey($name.ToLowerInvariant())) { continue }
            [void]$optionalLines.Add("# $name")
        }
    }

    $nl = [Environment]::NewLine
    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add('# packages.pe - WinPE optional components for this project')
    [void]$lines.Add('# Uncomment a line to install it. Order = install order.')
    [void]$lines.Add('# Docs: https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/winpe-add-packages--optional-components-reference')
    [void]$lines.Add('#')
    [void]$lines.Add('# Usage:')
    [void]$lines.Add('#   .\WinPE-Builder.ps1 -Build -ISO -Force -AddPackage packages.pe')
    [void]$lines.Add('#')
    [void]$lines.Add('# Language packs under WinPE_OCs/<lang>/ are added automatically when present.')
    [void]$lines.Add('')
    [void]$lines.Add('# Default Components')
    foreach ($line in $defaultLines) { [void]$lines.Add($line) }
    [void]$lines.Add('')
    [void]$lines.Add('# Optional Components')
    foreach ($line in $optionalLines) { [void]$lines.Add($line) }
    return (($lines -join $nl) + $nl)
}

function New-PackagesPeFile {
    param(
        [string]$Directory = $WorkDirectory,
        [string]$OcRoot
    )
    $path = Join-Path $Directory 'packages.pe'
    if (Test-Path -LiteralPath $path) {
        Write-Log "packages.pe already exists (kept): $path" -Level Info
        return $path
    }
    if (-not $OcRoot) {
        try { $OcRoot = (Get-AdkInfo).OcRoot } catch { $OcRoot = $null }
    }
    $text = Get-PackagesPeTemplate -OcRoot $OcRoot
    # UTF-8 without BOM for broad editor compatibility
    [IO.File]::WriteAllText($path, $text)
    $src = if ($OcRoot -and (Test-Path -LiteralPath $OcRoot)) { "from ADK ($OcRoot)" } else { 'built-in list' }
    Write-Log "Created packages.pe ($src): $path" -Level Ok
    return $path
}

function Test-LooksLikePackageListFile {
    param([string]$Entry)
    if ([string]::IsNullOrWhiteSpace($Entry)) { return $false }
    $e = $Entry.Trim().Trim('"')
    if ($e -match '\.pe$') { return $true }
    if ($e -match '[\\/]') { return $true }
    return $false
}

function Resolve-PackageListFilePath {
    param([string]$Entry)
    $e = $Entry.Trim().Trim('"')
    $candidates = [System.Collections.Generic.List[string]]::new()
    if ([IO.Path]::IsPathRooted($e)) {
        $r = Resolve-FullPath $e
        if ($r) { [void]$candidates.Add($r) }
    }
    else {
        if ($WorkDirectory) {
            [void]$candidates.Add((Join-Path $WorkDirectory $e))
        }
        $r = Resolve-FullPath $e
        if ($r) { [void]$candidates.Add($r) }
    }
    foreach ($c in $candidates) {
        if ($c -and (Test-Path -LiteralPath $c -PathType Leaf)) {
            return [IO.Path]::GetFullPath($c)
        }
    }
    return $null
}

function Assert-ValidPackageName {
    # Enforce strict package basename allowlist to prevent path traversal outside OcRoot.
    param([Parameter(Mandatory = $true)][string]$Name)
    if ($Name -notmatch '^[a-zA-Z0-9._-]+$' -or $Name -match '\.\.') {
        throw "Invalid WinPE package name '$Name'. Package names must match '^[a-zA-Z0-9._-]+$' and cannot contain path separators or '..'."
    }
}

function Assert-PathUnderRoot {
    # Verify canonical path containment inside the expected root directory.
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Root
    )
    $fullPath = [IO.Path]::GetFullPath($Path).TrimEnd('\')
    $fullRoot = [IO.Path]::GetFullPath($Root).TrimEnd('\')
    if (-not $fullPath.StartsWith($fullRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path traversal blocked: '$fullPath' is outside allowed root '$fullRoot'."
    }
    return $fullPath
}

function Write-ArtifactDigestLog {
    # Log SHA-256 digest of external artifacts (boot.wim, drivers, updates) before injection.
    param(
        [Parameter(Mandatory = $true)][string]$LiteralPath,
        [Parameter(Mandatory = $true)][string]$Category
    )
    if (-not (Test-Path -LiteralPath $LiteralPath -PathType Leaf)) { return $null }
    $hashObj = Get-FileHash -LiteralPath $LiteralPath -Algorithm SHA256 -ErrorAction Stop
    Write-Log ("SHA-256 [{0}] {1} = {2}" -f $Category, $LiteralPath, $hashObj.Hash) -Level Info
    return $hashObj.Hash
}

function Read-PackageListFile {
    param([string]$Path)
    $names = [System.Collections.Generic.List[string]]::new()
    foreach ($raw in Get-Content -LiteralPath $Path -ErrorAction Stop) {
        if ($null -eq $raw) { continue }
        $line = $raw.Trim()
        if (-not $line) { continue }
        if ($line.StartsWith('#')) { continue }
        # First token only (package basenames have no spaces in ADK cabs)
        $name = ($line -split '\s+')[0]
        if ($name) {
            Assert-ValidPackageName -Name $name
            [void]$names.Add($name)
        }
    }
    return @($names)
}

function Expand-AddPackageArguments {
    param([string[]]$Entries)
    $names = [System.Collections.Generic.List[string]]::new()
    $seen = @{}
    $usedListFile = $false

    foreach ($entry in @($Entries)) {
        if ([string]::IsNullOrWhiteSpace($entry)) { continue }
        $e = $entry.Trim()
        $looksLikeFile = Test-LooksLikePackageListFile $e
        $resolved = Resolve-PackageListFilePath $e

        if ($resolved) {
            $usedListFile = $true
            $active = @(Read-PackageListFile -Path $resolved)
            Write-Log "Package list file: $resolved ($($active.Count) active)" -Level Info
            foreach ($n in $active) {
                if (-not $n) { continue }
                Assert-ValidPackageName -Name $n
                if ($seen.ContainsKey($n)) { continue }
                $seen[$n] = $true
                [void]$names.Add($n)
            }
        }
        elseif ($looksLikeFile) {
            throw "Package list file not found: $e (checked WorkDirectory and current location)"
        }
        else {
            Assert-ValidPackageName -Name $e
            if ($seen.ContainsKey($e)) { continue }
            $seen[$e] = $true
            [void]$names.Add($e)
        }
    }

    return [pscustomobject]@{
        Names        = @($names)
        UsedListFile = $usedListFile
    }
}

function New-WinPeTree {
    param($Adk)
    Write-Phase '1/2 Create folders' 'Add-Drivers, Add-Scripts, Add-Updates, WinPE-ISO, WinPE-Logs, WinPE-Root'
    New-ProjectFolders
    Test-AdkReady -Adk $Adk
    # Seed packages.pe from real WinPE_OCs cabs when possible (never overwrite existing).
    New-PackagesPeFile -OcRoot $Adk.OcRoot | Out-Null

    Write-Phase '2/2 Stage WinPE tree (copype)' 'Copies ADK boot files into WinPE-Root.'
    if (Test-Path -LiteralPath $DirPeRoot) {
        Write-Log 'Removing old WinPE-Root before copype...' -Level Warn
        Dismount-ProjectImage
        Assert-NoReparsePointsOrActiveMounts -LiteralPath $DirPeRoot
        Remove-Item -LiteralPath $DirPeRoot -Recurse -Force
    }

    # copype may briefly mount/unmount — ADK staging, not our -Mount
    Invoke-AdkCommand -Tool $Adk.Copype `
        -Arguments ("{0} `"{1}`"" -f $Architecture, $DirPeRoot) `
        -EnvBat $Adk.EnvBat -Name 'copype' `
        -WaitHint '1-3 min. Mounting/Unmounting from copype is normal.'

    if (Get-ActiveMounts) {
        Write-Log 'Cleaning leftover mount from failed/partial copype...' -Level Warn
        Dismount-ProjectImage
    }

    if ($BootWimPath) {
        Assert-NoCmdMetacharacters -Value $BootWimPath -ParameterName 'BootWimPath'
        if (-not (Test-Path -LiteralPath $BootWimPath -PathType Leaf)) { throw "Custom boot.wim not found: $BootWimPath" }
        $null = Write-ArtifactDigestLog -LiteralPath $BootWimPath -Category 'BootWim'
        Copy-Item -LiteralPath $BootWimPath -Destination $FileBootWim -Force
        Write-Log 'Custom boot.wim applied.' -Level Ok
    }
    if (-not (Test-Path -LiteralPath $FileBootWim)) { throw "copype done but boot.wim missing: $FileBootWim" }
    if (Get-ActiveMounts) { throw 'Unexpected mount still open. Run -Discard.' }

    Write-Done "WinPE tree ready at $DirPeRoot"
    Write-Host ''
    Write-Host 'Add your custom content here (before -Build):' -ForegroundColor Cyan
    Write-Host "  Drivers : $DirDrivers" -ForegroundColor White
    Write-Host "  Scripts : $DirScripts" -ForegroundColor White
    Write-Host '            (startnet.cmd, winpe.jpg, unattend.xml, tools...)' -ForegroundColor DarkGray
    Write-Host "  Updates : $DirUpdates" -ForegroundColor White
    Write-Host '            (.msu / .cab packages)' -ForegroundColor DarkGray
    Write-Host "  Packages: $(Join-Path $WorkDirectory 'packages.pe')" -ForegroundColor White
    Write-Host '            (uncomment OCs, then -AddPackage packages.pe)' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host "This PowerShell window will keep using: $WorkDirectory" -ForegroundColor DarkGray
    Write-Host 'Next:  .\WinPE-Builder.ps1 -Build -ISO -Force' -ForegroundColor Cyan
    Write-Host '   or:  .\WinPE-Builder.ps1 -Build -ISO -Force -AddPackage packages.pe' -ForegroundColor Cyan
    Write-Host ''
}

function Add-OptionalPackages {
    param([string]$MountPath, [string]$OcRoot)
    $expanded = Expand-AddPackageArguments -Entries $AddPackage
    $list = New-Object System.Collections.Generic.List[string]
    $seen = @{}

    if ($expanded.UsedListFile) {
        # List file is the source of truth (active lines only; defaults not re-merged).
        foreach ($n in @($expanded.Names)) {
            if (-not $n -or $seen.ContainsKey($n)) { continue }
            Assert-ValidPackageName -Name $n
            $seen[$n] = $true
            [void]$list.Add($n)
        }
    }
    else {
        if (-not $NoDefaultPackages) {
            foreach ($n in $DefaultPackages) {
                if ($seen.ContainsKey($n)) { continue }
                Assert-ValidPackageName -Name $n
                $seen[$n] = $true
                [void]$list.Add($n)
            }
        }
        foreach ($n in @($expanded.Names)) {
            if (-not $n -or $seen.ContainsKey($n)) { continue }
            Assert-ValidPackageName -Name $n
            $seen[$n] = $true
            [void]$list.Add($n)
        }
    }

    # The ADK boot.wim is en-us by default. Other languages need their base
    # WinPE language pack before localized optional-component packages and
    # before Set-AllIntl runs later in the build.
    if ($Language -ne 'en-us') {
        $langRoot = Join-Path $OcRoot $Language
        $mainLanguageCab = Assert-PathUnderRoot -Path (Join-Path $langRoot 'lp.cab') -Root $OcRoot
        if (-not (Test-Path -LiteralPath $mainLanguageCab -PathType Leaf)) {
            throw "Main WinPE language pack not found: $mainLanguageCab"
        }

        Write-Log "Adding main language pack: $Language" -Level Info
        Add-WindowsPackage -Path $MountPath -PackagePath $mainLanguageCab -ErrorAction Stop | Out-Null
    }

    if ($list.Count -eq 0) { return 0 }

    Write-Log "Adding $($list.Count) package(s)..." -Level Step
    $n = 0
    $skipped = 0
    foreach ($name in $list) {
        Assert-ValidPackageName -Name $name
        $isMandatory = [bool]($script:MandatoryFoundationPackages | Where-Object { $_.Equals($name, [StringComparison]::OrdinalIgnoreCase) })
        $cab = Assert-PathUnderRoot -Path (Join-Path $OcRoot ($name + '.cab')) -Root $OcRoot
        if (-not (Test-Path -LiteralPath $cab)) {
            # Case-insensitive match (e.g. PlatformId vs PlatformID)
            $hit = $null
            if (Test-Path -LiteralPath $OcRoot -PathType Container) {
                $hit = Get-ChildItem -LiteralPath $OcRoot -Filter '*.cab' -File -ErrorAction SilentlyContinue |
                    Where-Object { $_.BaseName -eq $name } |
                    Select-Object -First 1
            }
            if (-not $hit) {
                if ($isMandatory) {
                    throw "Mandatory foundation package not found in '$OcRoot': $name"
                }
                Write-Log "Package not found (skipped): $name" -Level Warn
                $skipped++
                continue
            }
            $cab = Assert-PathUnderRoot -Path $hit.FullName -Root $OcRoot
            $name = $hit.BaseName
        }

        Write-Log "  $name" -Level Info
        try {
            Add-WindowsPackage -Path $MountPath -PackagePath $cab -ErrorAction Stop | Out-Null
        }
        catch {
            $errMsg = $_.Exception.Message
            if ($isMandatory) {
                throw "Mandatory foundation package install failed for '$name': $errMsg"
            }
            Write-Log "Package install failed (skipped): $name - $errMsg" -Level Warn
            $skipped++
            continue
        }

        $langCab = Assert-PathUnderRoot -Path (Join-Path $OcRoot (Join-Path $Language ($name + '_' + $Language + '.cab'))) -Root $OcRoot
        $langDir = Join-Path $OcRoot $Language
        if (-not (Test-Path -LiteralPath $langCab) -and (Test-Path -LiteralPath $langDir -PathType Container)) {
            $want = $name + '_' + $Language
            $langHit = Get-ChildItem -LiteralPath $langDir -Filter '*.cab' -File -ErrorAction SilentlyContinue |
                Where-Object { $_.BaseName -eq $want } |
                Select-Object -First 1
            if ($langHit) {
                $langCab = Assert-PathUnderRoot -Path $langHit.FullName -Root $OcRoot
            }
        }
        if (Test-Path -LiteralPath $langCab) {
            try {
                Add-WindowsPackage -Path $MountPath -PackagePath $langCab -ErrorAction Stop | Out-Null
            }
            catch {
                $errMsg = $_.Exception.Message
                Write-Log "Language cab failed (continued): $name ($Language) - $errMsg" -Level Warn
            }
        }
        $n++
    }

    if ($skipped -gt 0) {
        Write-Log "Packages: $n added, $skipped skipped" -Level Ok
    }
    else {
        Write-Log "Packages: $n" -Level Ok
    }
    return $n
}

function Add-ProjectDrivers {
    param([string]$MountPath)
    if (-not (Test-Path -LiteralPath $DirDrivers)) { return 0 }
    $infs = @(Get-ChildItem -LiteralPath $DirDrivers -Recurse -Filter '*.inf' -File -ErrorAction SilentlyContinue)
    if ($infs.Count -eq 0) { return 0 }
    $cats = @(Get-ChildItem -LiteralPath $DirDrivers -Recurse -Filter '*.cat' -File -ErrorAction SilentlyContinue)
    Write-Log "Adding drivers ($($infs.Count))..." -Level Step
    foreach ($drvFile in @($infs + $cats)) {
        $null = Write-ArtifactDigestLog -LiteralPath $drvFile.FullName -Category 'Driver'
    }
    Add-WindowsDriver -Path $MountPath -Driver $DirDrivers -Recurse -ErrorAction Stop | Out-Null
    Write-Log "Drivers: $($infs.Count)" -Level Ok
    return $infs.Count
}

function Copy-ProjectScripts {
    param([string]$MountPath)
    if (-not (Test-Path -LiteralPath $DirScripts)) { return }
    Write-Log 'Copying scripts...' -Level Step
    $sys32 = Join-Path $MountPath 'Windows\System32'
    $copied = 0

    foreach ($map in @(
            @{ Src = 'System32'; Dst = $sys32; Label = 'System32 -> image' }
            @{ Src = 'Root'; Dst = $MountPath; Label = 'Root -> X:\' }
            @{ Src = 'Media'; Dst = $DirMedia; Label = 'Media -> boot media' }
        )) {
        $srcPath = Join-Path $DirScripts $map.Src
        if (-not (Test-Path -LiteralPath $srcPath)) { continue }
        $items = @(Get-ChildItem -LiteralPath $srcPath -Force -ErrorAction SilentlyContinue)
        if ($items.Count -eq 0) { continue }
        Copy-Item (Join-Path $srcPath '*') $map.Dst -Recurse -Force
        $copied += $items.Count
        Write-Log "  $($map.Label) ($($items.Count) item(s))" -Level Info
    }

    # Copy root-level files in Add-Scripts\ (excluding special files and subdirectories) into Windows\System32.
    $specialRootFiles = @('startnet.cmd', 'winpe.jpg', 'unattend.xml')
    $rootFiles = @(
        Get-ChildItem -LiteralPath $DirScripts -File -Force -ErrorAction SilentlyContinue |
            Where-Object { $specialRootFiles -notcontains $_.Name.ToLowerInvariant() }
    )
    foreach ($rf in $rootFiles) {
        Copy-Item -LiteralPath $rf.FullName -Destination (Join-Path $sys32 $rf.Name) -Force
        $copied++
    }
    if ($rootFiles.Count -gt 0) {
        Write-Log "  Add-Scripts root files -> System32 ($($rootFiles.Count) item(s))" -Level Info
    }

    foreach ($n in @('Startnet.cmd', 'startnet.cmd')) {
        $src = Join-Path $DirScripts $n
        if (Test-Path -LiteralPath $src) {
            Copy-Item -LiteralPath $src -Destination (Join-Path $sys32 'Startnet.cmd') -Force
            Write-Log '  startnet.cmd applied' -Level Info
            $copied++
            break
        }
    }

    $bg = Join-Path $DirScripts 'winpe.jpg'
    if (Test-Path -LiteralPath $bg) {
        $dest = Join-Path $sys32 'winpe.jpg'
        if (Test-Path -LiteralPath $dest) {
            $ownArgs = '/f "{0}" /a' -f $dest
            $own = Start-Process -FilePath 'takeown.exe' -ArgumentList $ownArgs -Wait -PassThru -NoNewWindow
            if ($null -eq $own -or $own.ExitCode -ne 0) {
                $code = if ($null -ne $own) { $own.ExitCode } else { -1 }
                throw "takeown failed for winpe.jpg (exit $code)"
            }
            $aclArgs = '"{0}" /grant "*S-1-5-32-544:F"' -f $dest
            $acl = Start-Process -FilePath 'icacls.exe' -ArgumentList $aclArgs -Wait -PassThru -NoNewWindow
            if ($null -eq $acl -or $acl.ExitCode -ne 0) {
                $code = if ($null -ne $acl) { $acl.ExitCode } else { -1 }
                throw "icacls failed for winpe.jpg (exit $code)"
            }
        }
        Copy-Item -LiteralPath $bg -Destination $dest -Force
        Write-Log '  winpe.jpg background applied' -Level Ok
        $copied++
    }

    $ua = Join-Path $DirScripts 'unattend.xml'
    if (Test-Path -LiteralPath $ua) {
        Copy-Item -LiteralPath $ua -Destination (Join-Path $DirMedia 'unattend.xml') -Force
        Write-Log '  unattend.xml -> media' -Level Info
        $copied++
    }

    if ($copied -eq 0) {
        Write-Log 'No script content found under Add-Scripts (optional).' -Level Info
    }
    else {
        Write-Log "Scripts/content items applied: $copied" -Level Ok
    }
}

function Add-ProjectUpdates {
    param([string]$MountPath)
    if (-not (Test-Path -LiteralPath $DirUpdates)) { return 0 }
    $files = @(Get-ChildItem -LiteralPath $DirUpdates -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -match '^\.(msu|cab)$' })
    if ($files.Count -eq 0) { return 0 }

    Write-Log "Adding updates ($($files.Count))..." -Level Step
    foreach ($f in $files) {
        $null = Write-ArtifactDigestLog -LiteralPath $f.FullName -Category 'Update'
        Write-Log "  $($f.Name)" -Level Info
        Add-WindowsPackage -Path $MountPath -PackagePath $f.FullName -ErrorAction Stop | Out-Null
    }
    Write-Log 'Component cleanup...' -Level Info
    Repair-WindowsImage -Path $MountPath -StartComponentCleanup -ResetBase -ErrorAction Stop | Out-Null
    Write-Log "Updates: $($files.Count)" -Level Ok
    return $files.Count
}

function Test-TimeZoneId {
    param([string]$Name)
    $ids = foreach ($line in @(tzutil.exe /l 2>$null)) {
        $t = $line.Trim()
        if ($t -and $t -notmatch '^\(') { $t }
    }
    return [bool]($ids | Where-Object { $_ -eq $Name })
}

function Set-ImageScratchSpace {
    param([string]$MountPath, [int]$SizeMB)
    Write-Log "Scratch space: $SizeMB MB" -Level Info
    Invoke-DismOnImage -MountPath $MountPath -DismArgs @("/Set-ScratchSpace:$SizeMB") -Name 'Set-ScratchSpace'
}

function Set-ImageRegionalSettings {
    param([string]$MountPath)
    Write-Log "Locale (Set-AllIntl): $Language" -Level Info
    Invoke-DismOnImage -MountPath $MountPath -DismArgs @("/Set-AllIntl:$Language") -Name 'Set-AllIntl'
    if (-not (Test-TimeZoneId -Name $TimeZone)) {
        throw "Unknown time zone '$TimeZone'. List valid IDs with: tzutil /l"
    }
    Write-Log "Time zone: $TimeZone" -Level Info
    Invoke-DismOnImage -MountPath $MountPath -DismArgs @("/Set-TimeZone:`"$TimeZone`"") -Name 'Set-TimeZone'
}

# --- media ---

function New-WinPeIso {
    param($Adk)
    if (-not (Test-Path -LiteralPath $DirIso)) {
        New-Item -ItemType Directory -Path $DirIso -Force | Out-Null
    }
    $iso = Join-Path $DirIso ("WinPE_{0}_{1:yyyyMMdd_HHmm}.iso" -f $Architecture, (Get-Date))
    if ((Test-Path -LiteralPath $iso) -and -not $Force) { throw "ISO exists. Use -Force: $iso" }

    $argstring = '/ISO'
    if ($Force) { $argstring += ' /f' }
    $argstring += ' "{0}" "{1}"' -f $DirPeRoot, $iso
    if ($PCA2023) {
        Test-BootExSupport -Adk $Adk
        $argstring += ' /bootex'
    }

    Write-Phase 'Create ISO' $iso
    Invoke-AdkCommand -Tool $Adk.MakeMedia -Arguments $argstring -EnvBat $Adk.EnvBat `
        -Name 'MakeWinPEMedia ISO' -WaitHint 'Building ISO - can take a few minutes.'
    Write-Done "ISO ready: $iso"
    return $iso
}

function New-WinPeUsb {
    param($Adk)
    $letter = $USB.Substring(0, 1).ToUpperInvariant()
    $sys = if ($env:SystemDrive) { $env:SystemDrive.Substring(0, 1).ToUpperInvariant() } else { 'C' }
    if ($letter -eq $sys) { throw "Refused: system drive ${letter}:" }

    $part = Get-Partition -DriveLetter $letter -ErrorAction SilentlyContinue
    if (-not $part) { throw "Drive not found: ${letter}:" }
    $disk = Get-Disk -Number $part.DiskNumber -ErrorAction Stop
    if ($disk.IsBoot -or $disk.IsSystem) { throw "Refused: boot/system disk $($disk.Number)" }
    if ($disk.IsOffline) { throw "Refused: disk $($disk.Number) is offline" }
    if ($disk.IsReadOnly) { throw "Refused: disk $($disk.Number) is read-only" }
    if ($disk.BusType -ne 'USB') { throw "Refused: bus is $($disk.BusType), not USB" }
    $maxUsbBytes = 32GB
    if ($disk.Size -gt $maxUsbBytes) {
        throw ("Refused: USB disk {0} is {1:N1} GB (> 32 GB FAT32 limit for MakeWinPEMedia /UFD)." -f $disk.Number, ($disk.Size / 1GB))
    }
    if ($PCA2023) { Test-BootExSupport -Adk $Adk }

    $diskIdentity = "Number=$($disk.Number) FriendlyName='$($disk.FriendlyName)' SerialNumber='$($disk.SerialNumber)' UniqueId='$($disk.UniqueId)' Size=$($disk.Size) PartitionStyle='$($disk.PartitionStyle)'"
    Write-Log "Target USB disk identity: $diskIdentity" -Level Info

    Write-Phase 'Create USB' ("Drive {0}:  Disk {1}  {2:N1} GB  (WILL FORMAT)" -f $letter, $disk.Number, ($disk.Size / 1GB))
    $argstring = '/UFD /f "{0}" {1}:' -f $DirPeRoot, $letter
    if ($PCA2023) { $argstring += ' /bootex' }
    Invoke-AdkCommand -Tool $Adk.MakeMedia -Arguments $argstring -EnvBat $Adk.EnvBat `
        -Name 'MakeWinPEMedia USB' -WaitHint 'Formatting USB and copying boot files...'
    Write-Done "USB ready: ${letter}:"
    return "${letter}:"
}

# --- main ---

if ($Discard -and ($Build -or $Save -or $Init)) { throw '-Discard cannot mix with -Build/-Save/-Init' }
if ($Save -and ($ISO -or $USB)) { throw '-Save cannot mix with -ISO/-USB' }
if ($ISO -and -not $Build) { throw '-ISO requires -Build' }
if ($USB -and -not $Build) { throw '-USB requires -Build' }
if ($USB -and -not $Force) { throw '-USB requires -Force' }

$workDirWasPassed = $PSBoundParameters.ContainsKey('WorkDirectory') -and
    -not [string]::IsNullOrWhiteSpace($PSBoundParameters['WorkDirectory'])
$hasAction = $Init -or $Mount -or $Build -or $Save -or $Discard -or $Clean

# Early ADK / WinPE gate — before banner and before any Init/Mount/Build work
$needsAdk = $Init -or $Mount -or $Build
$adk = $null
if ($needsAdk) {
    $adk = Invoke-AdkPreflight
    if ($null -eq $adk) {
        $preErr = if ($script:LastPreflightError) {
            $script:LastPreflightError
        } else {
            'Windows ADK and/or WinPE add-on not installed. Install both from https://learn.microsoft.com/en-us/windows-hardware/get-started/adk-install then retry.'
        }
        return New-BuildResult -Ok $false -Action 'Failed' -Err $preErr
    }
}

if (-not $hasAction) {
    if ($workDirWasPassed) {
        if (-not (Test-WinPeProjectRoot -Path $WorkDirectory)) {
            throw @"
Not an existing WinPE project: $WorkDirectory
  Expected WinPE-Root (with boot.wim) or Add-Drivers + Add-Scripts + WinPE-Root.
  Create it once with:  .\WinPE-Builder.ps1 -Init -WorkDirectory '$WorkDirectory'
  Do not use -Init to resume — it rebuilds the tree and can overwrite work.
"@
        }
        Write-Banner 'Resume'
        Write-Log "Resumed project for this session: $WorkDirectory" -Level Ok
        Write-Host 'Project selected. Nothing was rebuilt or overwritten.' -ForegroundColor Green
        Write-Host ''
        Write-Host 'Next examples (same window):' -ForegroundColor Cyan
        Write-Host '  .\WinPE-Builder.ps1 -Build -ISO -Force' -ForegroundColor White
        Write-Host '  .\WinPE-Builder.ps1 -Mount' -ForegroundColor White
        Write-Host ''
        return New-BuildResult -Ok $true -Action 'Resume' -Mounted ([bool](Get-ActiveMounts))
    }

    Write-Banner 'Help'
    Write-Host 'Pick one action:' -ForegroundColor White
    Write-Host '  -Init              Create / rebuild project tree (overwrites WinPE-Root)'
    Write-Host '  -WorkDirectory X   Resume existing project for this PowerShell window'
    Write-Host '  -Mount             Mount boot.wim for manual edits'
    Write-Host '  -Build             Add packages/drivers/scripts/updates'
    Write-Host '  -Save              Unmount and keep changes'
    Write-Host '  -Discard           Unmount and drop changes'
    Write-Host '  -Build -ISO -Force Create ISO'
    Write-Host '  -Build -USB F: -Force  Create USB (formats drive)'
    Write-Host ''
    Write-Host 'Packages:' -ForegroundColor White
    Write-Host '  -AddPackage Name   Extra OC name(s), and/or packages.pe list file' -ForegroundColor DarkGray
    Write-Host '  -NoDefaultPackages Skip built-in defaults (list file still applies if passed)' -ForegroundColor DarkGray
    Write-Host '  packages.pe        Created on -Init; uncomment lines, then -AddPackage packages.pe' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host 'Secure Boot media (with -Build -ISO / -USB):' -ForegroundColor White
    Write-Host '  (default)          Boot files signed with Windows Production PCA 2011' -ForegroundColor DarkGray
    Write-Host '  -PCA2023           Boot files signed with Windows UEFI CA 2023' -ForegroundColor DarkGray
    Write-Host '                     (only for devices that already trust the 2023 CA)' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host 'Project folder:' -ForegroundColor White
    Write-Host '  Default = script folder (or current folder if already a project)' -ForegroundColor DarkGray
    Write-Host '  -WorkDirectory alone = resume (no -Init)' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host 'Example: .\WinPE-Builder.ps1 -Init' -ForegroundColor Cyan
    Write-Host '         .\WinPE-Builder.ps1 -Build -ISO -Force' -ForegroundColor Cyan
    Write-Host '         .\WinPE-Builder.ps1 -Build -ISO -Force -AddPackage packages.pe' -ForegroundColor Cyan
    Write-Host '         .\WinPE-Builder.ps1 -Build -ISO -PCA2023 -Force' -ForegroundColor Cyan
    Write-Host '         .\WinPE-Builder.ps1 -WorkDirectory C:\MyWinPE-Project   # resume later' -ForegroundColor Cyan
    Write-Host '         https://github.com/cmartinezone/WinPEBuilder' -ForegroundColor DarkGray
    Write-Host ''
    return New-BuildResult -Ok $false -Action 'None' -Err 'No action'
}

$actionName = if ($Init -and -not $Mount -and -not $Build) { 'Init' }
elseif ($Mount -and -not $Build) { 'Mount' }
elseif ($Build -and $ISO -and $USB) { 'Build + ISO + USB' }
elseif ($Build -and $ISO) { 'Build + ISO' }
elseif ($Build -and $USB) { 'Build + USB' }
elseif ($Build) { 'Build' }
elseif ($Save) { 'Save' }
elseif ($Discard) { 'Discard' }
elseif ($Clean) { 'Clean' }
else { 'Run' }

Write-Banner $actionName
if ($null -eq $adk) { $adk = Get-AdkInfo }
$isoPath = $null
$usbPath = $null
$pkgN = 0; $drvN = 0; $updN = 0
$buildMountedImage = $false

try {
    if ($Clean -and -not $Mount -and -not $Build) {
        if (-not $PSCmdlet.ShouldProcess($DirMount, 'Clean project mounts')) {
            return New-BuildResult -Ok $true -Action 'Clean'
        }
        Write-Phase 'Clean mounts' $(if ($Force) { 'Including global DISM cleanup' } else { 'Project mounts only' })
        Clear-ProjectMount -Global:$Force
        Write-Done 'Mounts cleaned'
        return New-BuildResult -Ok $true -Action 'Clean'
    }

    if ($Discard) {
        if (-not $PSCmdlet.ShouldProcess($DirMount, 'Discard active mount')) {
            return New-BuildResult -Ok $true -Action 'Discard'
        }
        Write-Phase 'Discard mount' 'Drops uncommitted changes'
        Clear-ProjectMount
        Write-Done 'Changes discarded; not mounted'
        return New-BuildResult -Ok $true -Action 'Discard'
    }

    if ($Init) {
        if (-not $PSCmdlet.ShouldProcess($WorkDirectory, 'Initialize WinPE project tree')) {
            if (-not $Mount -and -not $Build) {
                return New-BuildResult -Ok $true -Action 'Init' -Mounted $false
            }
        }
        else {
            New-WinPeTree -Adk $adk
            if (-not $Mount -and -not $Build) {
                return New-BuildResult -Ok $true -Action 'Init' -Mounted $false
            }
        }
    }

    if ($Save -and -not $Build -and -not $Mount) {
        if (-not (Get-ActiveMounts)) { throw 'Nothing mounted. Run -Mount or -Build first.' }
        if (-not $PSCmdlet.ShouldProcess($DirMount, 'Save and dismount boot.wim')) {
            return New-BuildResult -Ok $true -Action 'Save' -Mounted $true
        }
        Write-Phase 'Save image' 'Unmount and keep changes'
        Dismount-ProjectImage -Save
        Write-Done 'Image saved and unmounted'
        return New-BuildResult -Ok $true -Action 'Save' -Mounted $false
    }

    if ($Mount -and -not $Build) {
        if (-not $PSCmdlet.ShouldProcess($FileBootWim, "Mount boot.wim to $DirMount")) {
            return New-BuildResult -Ok $true -Action 'Mount' -Mounted ([bool](Get-ActiveMounts))
        }
        if ($Clean) { Clear-ProjectMount }
        $null = Mount-BootImage
        Write-Host "Edit files in: $DirMount" -ForegroundColor Cyan
        Write-Host 'Then: -Save (keep) or -Discard (drop)' -ForegroundColor DarkGray
        Write-Host ''
        return New-BuildResult -Ok $true -Action 'Mount' -Mounted $true
    }

    if ($Build) {
        if (-not $PSCmdlet.ShouldProcess($WorkDirectory, $actionName)) {
            return New-BuildResult -Ok $true -Action 'Build' -Mounted ([bool](Get-ActiveMounts))
        }
        if ($Clean) { Clear-ProjectMount }
        if (-not (Test-TimeZoneId -Name $TimeZone)) {
            throw "Unknown time zone '$TimeZone'. List valid IDs with: tzutil /l"
        }
        if ($PCA2023 -and ($ISO -or $USB)) { Test-BootExSupport -Adk $adk }
        # boot.wim must exist (ADK install already verified in preflight)
        if (-not (Test-Path -LiteralPath $FileBootWim)) {
            throw "boot.wim not found. Run -Init first: $FileBootWim"
        }

        Write-Phase 'Build customization' 'Packages, drivers, scripts, updates'
        Write-Log "Language=$Language TimeZone='$TimeZone' ScratchSpace=${ScratchSpace}MB PCA2023=$PCA2023" -Level Info
        $wasAlreadyMounted = (@(Get-ActiveMounts).Count -gt 0)
        $mountPath = Mount-BootImage
        if (-not $wasAlreadyMounted) {
            $buildMountedImage = $true
        }
        Set-ImageScratchSpace -MountPath $mountPath -SizeMB $ScratchSpace
        $pkgN = Add-OptionalPackages -MountPath $mountPath -OcRoot $adk.OcRoot
        Set-ImageRegionalSettings -MountPath $mountPath
        $drvN = Add-ProjectDrivers -MountPath $mountPath
        Copy-ProjectScripts -MountPath $mountPath
        $updN = Add-ProjectUpdates -MountPath $mountPath
        Write-Log "Build summary: packages=$pkgN drivers=$drvN updates=$updN" -Level Info

        if ($ISO -or $USB -or $Save) {
            Write-Phase 'Commit image' 'Saving changes into boot.wim'
            Dismount-WindowsImage -Path $mountPath -Save -ErrorAction Stop | Out-Null
            $buildMountedImage = $false
            Write-Log 'Image committed.' -Level Ok
        }

        if ($ISO -and $PSCmdlet.ShouldProcess($DirIso, 'Create WinPE ISO')) {
            $isoPath = New-WinPeIso -Adk $adk
        }
        if ($USB -and $PSCmdlet.ShouldProcess("USB Drive $USB", 'Format and create WinPE USB')) {
            $usbPath = New-WinPeUsb -Adk $adk
        }

        if (-not ($ISO -or $USB -or $Save)) {
            $buildMountedImage = $false
            Write-Done 'Build applied; image still mounted (use -Save when ready)'
        }
        else {
            if ($isoPath) { Write-Host "ISO: $isoPath" -ForegroundColor Green }
            if ($usbPath) { Write-Host "USB: $usbPath" -ForegroundColor Green }
            Write-Done 'Build finished'
        }
    }

    return New-BuildResult -Ok $true -Action 'Build' -Iso $isoPath -Usb $usbPath `
        -Mounted ([bool](Get-ActiveMounts)) -Pkgs $pkgN -Drvs $drvN -Upds $updN
}
catch {
    $primaryErr = $_.Exception.Message
    Write-Host ''
    Write-Host 'FAILED' -ForegroundColor Red
    Write-Log $primaryErr -Level Error
    if ($_.ScriptStackTrace) { Write-Log $_.ScriptStackTrace -Level Error }
    if ($Build -and $buildMountedImage) {
        Write-Log 'Build failed after mounting boot.wim; discarding uncommitted mount...' -Level Warn
        try {
            Dismount-ProjectImage
        }
        catch {
            Write-Log "Automatic discard after build failure encountered an error: $($_.Exception.Message)" -Level Error
        }
    }
    $isMounted = $false
    try { $isMounted = [bool](Get-ActiveMounts) } catch { }
    Write-Host ''
    Write-Host "Log: $script:LogFile" -ForegroundColor Yellow
    if ($isMounted) {
        Write-Host 'Image may still be mounted. Run -Discard if needed.' -ForegroundColor Yellow
    }
    Write-Host ''
    return New-BuildResult -Ok $false -Action 'Failed' -Err $primaryErr -Mounted $isMounted
}

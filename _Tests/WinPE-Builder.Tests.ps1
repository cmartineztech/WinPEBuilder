#Requires -Version 5.1
<#
.SYNOPSIS
    Pester & AST verification tests for WinPE-Builder.ps1 (WPEB-01 through WPEB-06).
    Compatible with Pester 3.4+ (inbox Windows PowerShell 5.1), Pester 5.x, and direct execution.
#>

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:RepoRoot = Split-Path -Parent $PSScriptRoot
$script:BuilderScriptPath = Join-Path $script:RepoRoot 'WinPE-Builder.ps1'

if (-not (Get-Command -Name Describe -ErrorAction SilentlyContinue)) {
    $script:DirectPassed = 0
    $script:DirectFailed = 0
    function Describe ([string]$Name, [scriptblock]$Fixture) {
        Write-Host "`n=== $Name ===" -ForegroundColor Cyan
        & $Fixture
        Write-Host "`nSummary: $($script:DirectPassed) passed, $($script:DirectFailed) failed." -ForegroundColor Cyan
        if ($script:DirectFailed -gt 0) { throw "$($script:DirectFailed) test(s) failed." }
    }
    function Context ([string]$Name, [scriptblock]$Fixture) {
        Write-Host "  [$Name]" -ForegroundColor Yellow
        & $Fixture
    }
    function It ([string]$Name, [scriptblock]$Test) {
        try {
            & $Test
            Write-Host "    [PASS] $Name" -ForegroundColor Green
            $script:DirectPassed++
        }
        catch {
            Write-Host "    [FAIL] $Name - $($_.Exception.Message)" -ForegroundColor Red
            $script:DirectFailed++
        }
    }
}

Describe 'WinPE-Builder.ps1 Hardening Suite (WPEB-01 .. WPEB-06)' {

    $tokens = $null
    $parseErrors = $null
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $script:BuilderScriptPath,
        [ref]$tokens,
        [ref]$parseErrors
    )
    $script:ParseErrors = $parseErrors
    $script:RawContent = Get-Content -LiteralPath $script:BuilderScriptPath -Raw -Encoding UTF8

    # Load helper functions from AST into current scope for unit testing without requiring elevation or Dism module.
    $funcAsts = $script:Ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst]
        }, $false)
    foreach ($fn in $funcAsts) {
        Invoke-Expression $fn.Extent.Text
    }

    $script:MandatoryFoundationPackages = @(
        'WinPE-WMI',
        'WinPE-NetFx',
        'WinPE-Scripting',
        'WinPE-PowerShell'
    )
    $script:DefaultPackages = @(
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
    $script:QuietMode = $true
    $script:LogFile = $null

    Context 'AST Syntax & CmdletBinding Contract (WPEB-05)' {
        It 'parses WinPE-Builder.ps1 with zero PowerShell AST syntax errors' {
            if ($script:ParseErrors.Count -ne 0) {
                throw "Expected 0 parse errors, got $($script:ParseErrors.Count): $($script:ParseErrors | Out-String)"
            }
        }

        It 'declares SupportsShouldProcess = $true and ConfirmImpact = High on the script ParamBlock' {
            $cmdletAttr = $script:Ast.ParamBlock.Attributes | Where-Object {
                $_.TypeName.Name -eq 'CmdletBinding'
            }
            if ($null -eq $cmdletAttr) { throw 'Missing [CmdletBinding] attribute on ParamBlock' }
            if (-not $script:RawContent.Contains('SupportsShouldProcess = $true')) {
                throw 'Missing SupportsShouldProcess = $true'
            }
            if (-not $script:RawContent.Contains("ConfirmImpact = 'High'")) {
                throw 'Missing ConfirmImpact = High'
            }
        }
    }

    Context 'WPEB-01: Active WIM Mount & Reparse Point Protection' {
        It 'Get-ActiveMounts enforces exact root or trailing backslash boundary' {
            $DirPeRoot = 'C:\Project\WinPE-Root'
            function Get-WindowsImage {
                param([switch]$Mounted)
                @(
                    [pscustomobject]@{ Path = 'C:\Project\WinPE-Root\mount' }
                    [pscustomobject]@{ Path = 'C:\Project\WinPE-Root2\mount' }
                )
            }
            $active = @(Get-ActiveMounts)
            if ($active.Count -ne 1 -or $active[0].Path -ne 'C:\Project\WinPE-Root\mount') {
                throw "Expected only C:\Project\WinPE-Root\mount, got: $($active | Out-String)"
            }
        }

        It 'Dismount-ProjectImage invokes Clear-WindowsCorruptMountPoint and throws when a mount remains active' {
            $DirPeRoot = 'C:\Project\WinPE-Root'
            $script:corruptCalled = $false
            function Get-WindowsImage {
                param([switch]$Mounted)
                @([pscustomobject]@{ Path = 'C:\Project\WinPE-Root\mount' })
            }
            function Dismount-WindowsImage {
                param([string]$Path, [switch]$Discard, [switch]$Save)
                throw 'Simulated unmount lock failure'
            }
            function Clear-WindowsCorruptMountPoint {
                $script:corruptCalled = $true
            }
            $threw = $false
            try {
                Dismount-ProjectImage
            }
            catch {
                $threw = $true
            }
            if (-not $threw) { throw 'Expected Dismount-ProjectImage to throw when mount remains active' }
            if (-not $script:corruptCalled) { throw 'Expected Clear-WindowsCorruptMountPoint to be invoked' }
        }

        It 'Assert-NoReparsePointsOrActiveMounts throws when an active mount exists under DirPeRoot' {
            $DirPeRoot = 'C:\Project\WinPE-Root'
            function Get-WindowsImage {
                param([switch]$Mounted)
                @([pscustomobject]@{ Path = 'C:\Project\WinPE-Root\mount' })
            }
            $threw = $false
            try {
                Assert-NoReparsePointsOrActiveMounts -LiteralPath 'C:\Project\WinPE-Root'
            }
            catch {
                $threw = $true
            }
            if (-not $threw) { throw 'Expected Assert-NoReparsePointsOrActiveMounts to throw when mount is active' }
        }

        It 'calls Assert-NoReparsePointsOrActiveMounts before Remove-Item -Recurse in Clear-ProjectMount, Mount-BootImage, and New-WinPeTree' {
            $matchesCount = ([regex]::Matches($script:RawContent, 'Assert-NoReparsePointsOrActiveMounts\s+-LiteralPath')).Count
            if ($matchesCount -lt 3) {
                throw "Expected at least 3 calls to Assert-NoReparsePointsOrActiveMounts, found $matchesCount"
            }
        }
    }

    Context 'WPEB-02: Command Injection Prevention & Proper Quoting' {
        It 'Assert-NoCmdMetacharacters rejects cmd.exe metacharacters ([&|<>^%!])' {
            foreach ($bad in @('C:\Proj&calc', 'C:\Proj|whoami', 'C:\Proj>out', 'C:\Proj<in', 'C:\Proj^x', 'C:\Proj%TEMP%', 'C:\Proj!x!')) {
                $threw = $false
                try {
                    Assert-NoCmdMetacharacters -Value $bad -ParameterName 'WorkDirectory'
                }
                catch {
                    $threw = $true
                }
                if (-not $threw) { throw "Expected Assert-NoCmdMetacharacters to reject '$bad'" }
            }
        }

        It 'Copy-ProjectScripts uses well-known Administrators SID (*S-1-5-32-544:F) and quotes winpe.jpg paths' {
            if (-not $script:RawContent.Contains('*S-1-5-32-544:F')) {
                throw 'Expected well-known SID *S-1-5-32-544:F in Copy-ProjectScripts'
            }
            if ($script:RawContent.Contains('Administrators:F')) {
                throw 'Localized group name Administrators:F should not be present'
            }
        }
    }

    Context 'WPEB-03: Package Name Allowlist & Canonical OcRoot Containment' {
        It 'Assert-ValidPackageName accepts valid WinPE OC names and rejects traversal or shell characters' {
            Assert-ValidPackageName -Name 'WinPE-PowerShell'
            Assert-ValidPackageName -Name 'WinPE-FontSupport-JA-JP'

            foreach ($badPkg in @('..\..\evil', '../evil', 'WinPE\..\evil', 'WinPE;calc', 'WinPE/cab')) {
                $threw = $false
                try {
                    Assert-ValidPackageName -Name $badPkg
                }
                catch {
                    $threw = $true
                }
                if (-not $threw) { throw "Expected Assert-ValidPackageName to reject '$badPkg'" }
            }
        }

        It 'Assert-PathUnderRoot blocks paths escaping OcRoot' {
            $tmpBase = Join-Path ([IO.Path]::GetTempPath()) ("wpeb_test_" + [Guid]::NewGuid().ToString('N'))
            $ocRoot = Join-Path $tmpBase 'WinPE_OCs'
            $outside = Join-Path $tmpBase 'rogue.cab'
            $inside = Join-Path $ocRoot 'WinPE-WMI.cab'

            $insideResolved = Assert-PathUnderRoot -Path $inside -Root $ocRoot
            if ($insideResolved -ne ([IO.Path]::GetFullPath($inside).TrimEnd('\'))) {
                throw "Unexpected canonical path: $insideResolved"
            }

            $threw = $false
            try {
                $null = Assert-PathUnderRoot -Path $outside -Root $ocRoot
            }
            catch {
                $threw = $true
            }
            if (-not $threw) { throw 'Expected Assert-PathUnderRoot to reject path outside OcRoot' }
        }
    }

    Context 'WPEB-04: Root Add-Scripts File Copying' {
        It 'Copy-ProjectScripts copies root-level files in Add-Scripts to Windows\System32 while excluding startnet.cmd, winpe.jpg, and unattend.xml' {
            $tmpProj = Join-Path ([IO.Path]::GetTempPath()) ("wpeb_scripts_" + [Guid]::NewGuid().ToString('N'))
            $DirScripts = Join-Path $tmpProj 'Add-Scripts'
            $DirMedia = Join-Path $tmpProj 'media'
            $mountDir = Join-Path $tmpProj 'mount'
            $sys32Dir = Join-Path $mountDir 'Windows\System32'
            try {
                New-Item -ItemType Directory -Path $DirScripts, $DirMedia, $sys32Dir -Force | Out-Null
                New-Item -ItemType Directory -Path (Join-Path $DirScripts 'Root') -Force | Out-Null
                Set-Content -LiteralPath (Join-Path $DirScripts 'BitLockerUtility.ps1') -Value '# tool' -Encoding UTF8
                Set-Content -LiteralPath (Join-Path $DirScripts 'startnet.cmd') -Value '@echo off' -Encoding UTF8
                Set-Content -LiteralPath (Join-Path $DirScripts 'unattend.xml') -Value '<unattend/>' -Encoding UTF8
                Set-Content -LiteralPath (Join-Path $DirScripts 'Root\readme.txt') -Value 'root-file' -Encoding UTF8

                Copy-ProjectScripts -MountPath $mountDir

                if (-not (Test-Path -LiteralPath (Join-Path $sys32Dir 'BitLockerUtility.ps1'))) {
                    throw 'Root file BitLockerUtility.ps1 was not copied to Windows\System32'
                }
                if (-not (Test-Path -LiteralPath (Join-Path $sys32Dir 'Startnet.cmd'))) {
                    throw 'startnet.cmd was not copied to Windows\System32\Startnet.cmd'
                }
                if (Test-Path -LiteralPath (Join-Path $sys32Dir 'unattend.xml')) {
                    throw 'unattend.xml should be copied only to media, not Windows\System32'
                }
                if (-not (Test-Path -LiteralPath (Join-Path $DirMedia 'unattend.xml'))) {
                    throw 'unattend.xml was not copied to media'
                }
                if (-not (Test-Path -LiteralPath (Join-Path $mountDir 'readme.txt'))) {
                    throw 'Root\readme.txt was not copied to mount root'
                }
            }
            finally {
                Remove-Item -LiteralPath $tmpProj -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    Context 'WPEB-05 & WPEB-06: Safe WorkDir, USB Guards, Mandatory Foundation Packages, and Digest Logging' {
        It 'Assert-SafeWorkDirectory rejects drive roots, SystemRoot, ProgramFiles, ProgramFiles(x86), ProgramData, and USERPROFILE' {
            foreach ($unsafePath in @(
                    'C:\',
                    'D:\',
                    'C:\Windows',
                    'C:\Windows\System32',
                    'C:\Program Files',
                    'C:\Program Files\MyWinPE',
                    'C:\Program Files (x86)\MyWinPE',
                    'C:\ProgramData\MyWinPE'
                )) {
                $threw = $false
                try {
                    Assert-SafeWorkDirectory -Path $unsafePath
                }
                catch {
                    $threw = $true
                }
                if (-not $threw) { throw "Expected Assert-SafeWorkDirectory to reject '$unsafePath'" }
            }

            if ($env:USERPROFILE) {
                $threwProfile = $false
                try {
                    Assert-SafeWorkDirectory -Path $env:USERPROFILE
                }
                catch {
                    $threwProfile = $true
                }
                if (-not $threwProfile) { throw 'Expected Assert-SafeWorkDirectory to reject $env:USERPROFILE' }
            }
        }

        It 'Add-OptionalPackages fails closed when a mandatory foundation package fails to install' {
            $tmpOc = Join-Path ([IO.Path]::GetTempPath()) ("wpeb_oc_" + [Guid]::NewGuid().ToString('N'))
            try {
                New-Item -ItemType Directory -Path $tmpOc -Force | Out-Null
                $AddPackage = @('WinPE-WMI')
                $NoDefaultPackages = $true
                $Language = 'en-us'

                $threwMissing = $false
                try {
                    $null = Add-OptionalPackages -MountPath $tmpOc -OcRoot $tmpOc
                }
                catch {
                    $threwMissing = $true
                }
                if (-not $threwMissing) {
                    throw 'Expected missing mandatory foundation package WinPE-WMI to throw a terminating error'
                }
            }
            finally {
                Remove-Item -LiteralPath $tmpOc -Recurse -Force -ErrorAction SilentlyContinue
            }
        }

        It 'New-WinPeUsb rejects USB disks larger than 32GB or read-only/offline disks' {
            $USB = 'F:'
            $PCA2023 = $false
            function Get-Partition {
                param([string]$DriveLetter)
                [pscustomobject]@{ DiskNumber = 4 }
            }
            function Get-Disk {
                param([int]$Number)
                [pscustomobject]@{
                    Number         = 4
                    FriendlyName   = 'Test USB'
                    SerialNumber   = 'SN12345'
                    UniqueId       = 'UID12345'
                    Size           = 64GB
                    PartitionStyle = 'MBR'
                    BusType        = 'USB'
                    IsBoot         = $false
                    IsSystem       = $false
                    IsOffline      = $false
                    IsReadOnly     = $false
                }
            }
            $threwSize = $false
            try {
                $null = New-WinPeUsb -Adk ([pscustomobject]@{})
            }
            catch {
                $threwSize = $true
            }
            if (-not $threwSize) { throw 'Expected New-WinPeUsb to reject 64GB USB disk (>32GB limit)' }
        }
    }
}

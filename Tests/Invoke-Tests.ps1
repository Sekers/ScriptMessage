<#
    .SYNOPSIS
    Runs the Pester tests under Windows PowerShell 5.1 and PowerShell 7, and exits with a non-zero code if any
    test fails.

    .DESCRIPTION
    Each edition runs the tests in its own child process, so both import the module from scratch. Both editions
    need Pester 5 or later. When either edition does not have it, or when -UsePinnedPester is given, both use a
    pinned Pester version, which is saved once to a temporary folder with Save-PSResource (this needs PowerShell 7).

    .PARAMETER Edition
    Which PowerShell editions to run the tests under: Both (the default), Core, or Desktop.

    .PARAMETER UsePinnedPester
    Use the pinned Pester version even when Pester 5 or later is already installed, so every run uses the same
    version. The CI workflow uses this.

    .EXAMPLE
    .\Tests\Invoke-Tests.ps1

    .EXAMPLE
    .\Tests\Invoke-Tests.ps1 -Edition Core
#>

[CmdletBinding()]
param(
    [ValidateSet('Both', 'Core', 'Desktop')]
    [string]$Edition = 'Both',

    [switch]$UsePinnedPester,

    # The script starts itself in each edition's process with -Stage to do one part of the work there.
    [Parameter(DontShow = $true)]
    [ValidateSet('CheckPester', 'SavePester', 'Run')]
    [string]$Stage,

    # For the SavePester stage, the folder to save Pester into. For the Run stage, the pinned Pester manifest.
    [Parameter(DontShow = $true)]
    [string]$PesterPath
)

$ErrorActionPreference = 'Stop'

$PinnedPesterVersion = '6.1.0'

switch ($Stage)
{
    'CheckPester'
    {
        if (Get-Module -ListAvailable -Name Pester | Where-Object { $_.Version -ge [version]'5.0' })
        {
            exit 0
        }
        exit 1
    }
    'SavePester'
    {
        New-Item -ItemType Directory -Force -Path $PesterPath | Out-Null
        Save-PSResource -Name Pester -Version $PinnedPesterVersion -Path $PesterPath -TrustRepository
        exit 0
    }
    'Run'
    {
        $ProgressPreference = 'SilentlyContinue'
        if ($PesterPath)
        {
            Import-Module -Name $PesterPath
        }
        else
        {
            Import-Module -Name Pester -MinimumVersion 5.0
        }

        $Configuration = New-PesterConfiguration
        $Configuration.Run.Path = $PSScriptRoot
        $Configuration.Run.PassThru = $true
        $Configuration.Output.Verbosity = 'Detailed'
        $Result = Invoke-Pester -Configuration $Configuration

        # Every stage must exit here; otherwise the child process would continue into the code below and start
        # test runs of its own.
        if ($Result.Result -eq 'Passed')
        {
            exit 0
        }
        exit 1
    }
}

# Starts this script in another PowerShell process to run one stage there.
function Invoke-Stage
{
    param(
        [string]$Executable,
        [string]$StageName,
        [string]$StagePesterPath
    )

    $Arguments = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath, '-Stage', $StageName)
    if ($StagePesterPath)
    {
        $Arguments += @('-PesterPath', $StagePesterPath)
    }
    & $Executable @Arguments
}

$OnWindows = ($PSVersionTable.PSEdition -eq 'Desktop') -or $IsWindows

$Editions = [ordered]@{}
if ($Edition -in 'Both', 'Core')
{
    $Editions['PowerShell 7'] = 'pwsh'
}
if ($Edition -in 'Both', 'Desktop')
{
    if ($OnWindows)
    {
        $Editions['Windows PowerShell 5.1'] = 'powershell'
    }
    elseif ($Edition -eq 'Desktop')
    {
        throw 'Windows PowerShell 5.1 is only available on Windows.'
    }
    else
    {
        Write-Warning 'Skipping Windows PowerShell 5.1, which is only available on Windows.'
    }
}

foreach ($Executable in $Editions.Values)
{
    if (-not (Get-Command -Name $Executable -CommandType Application -ErrorAction SilentlyContinue))
    {
        throw "Cannot find '$Executable'. Install it, or use -Edition to run only the other edition."
    }
}

# Use the installed Pester when every edition has version 5 or later. Otherwise use the pinned version for all of
# them, so both editions always run the same Pester.
$NeedsPinnedPester = [bool]$UsePinnedPester
if (-not $NeedsPinnedPester)
{
    foreach ($Executable in $Editions.Values)
    {
        Invoke-Stage -Executable $Executable -StageName 'CheckPester'
        if ($LASTEXITCODE -ne 0)
        {
            $NeedsPinnedPester = $true
        }
    }
}

$PesterManifest = ''
if ($NeedsPinnedPester)
{
    $SaveRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath 'ScriptMessage-TestModules'
    $PesterManifest = Join-Path -Path (Join-Path -Path (Join-Path -Path $SaveRoot -ChildPath 'Pester') -ChildPath $PinnedPesterVersion) -ChildPath 'Pester.psd1'
    if (-not (Test-Path -LiteralPath $PesterManifest))
    {
        if (-not (Get-Command -Name 'pwsh' -CommandType Application -ErrorAction SilentlyContinue))
        {
            throw "Pester 5 or later is not installed, and PowerShell 7 is needed to save Pester $PinnedPesterVersion. Install either one."
        }

        Write-Host "Saving Pester $PinnedPesterVersion to $SaveRoot"
        Invoke-Stage -Executable 'pwsh' -StageName 'SavePester' -StagePesterPath $SaveRoot
        if (($LASTEXITCODE -ne 0) -or (-not (Test-Path -LiteralPath $PesterManifest)))
        {
            throw "Could not save Pester $PinnedPesterVersion to $SaveRoot."
        }
    }
}

$ExitCodes = [ordered]@{}
foreach ($Name in $Editions.Keys)
{
    Write-Host ''
    Write-Host "===== $Name ====="
    Invoke-Stage -Executable $Editions[$Name] -StageName 'Run' -StagePesterPath $PesterManifest
    $ExitCodes[$Name] = $LASTEXITCODE
}

Write-Host ''
Write-Host '===== Summary ====='
$AnyFailed = $false
foreach ($Name in $ExitCodes.Keys)
{
    if ($ExitCodes[$Name] -eq 0)
    {
        Write-Host "PASS  $Name"
    }
    else
    {
        Write-Host "FAIL  $Name (exit code $($ExitCodes[$Name]))"
        $AnyFailed = $true
    }
}

if ($AnyFailed)
{
    exit 1
}
exit 0

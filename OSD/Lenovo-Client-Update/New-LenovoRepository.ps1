#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Builds a local Lenovo Update Retriever style repository during a Configuration Manager
    OSD task sequence.

.DESCRIPTION
    Runs as a "Run PowerShell Script" step (Windows PowerShell 5.1) after Setup Windows and
    Configuration Manager. Downloads all package types (applications, drivers, BIOS and
    firmware) for the machine type of the device being imaged into the task sequence data
    path, ready for Install-LenovoBIOSFirmware.ps1 and Install-LenovoDrivers.ps1 to install.

    Assumes the Lenovo.Client.Update module is already present on the device and resolvable
    via PSModulePath.

.PARAMETER RepositoryPath
    Target repository directory. Defaults to <_SMSTSMDataPath>\LenovoUpdates, which resolves
    to C:\_SMSTaskSequence\LenovoUpdates and is cleaned up when the task sequence completes.

.PARAMETER RebootTypes
    Reboot types to download. 0=none, 3=requires reboot, 5=delayed forced reboot.

.NOTES
    FileName:   New-LenovoRepository.ps1
    Author:     Philip Jorgensen
    Created:    2026-08-27

    Task sequence step: Run PowerShell Script
        Script source:               Enter a PowerShell script (this file pasted inline)
        PowerShell execution policy: Bypass
        Parameters:                  empty, or -RebootTypes "0,3"
        Success codes:               0

    No package is attached to the step, so an earlier step must place the
    Lenovo.Client.Update module on the device where PSModulePath can resolve it.
#>

[CmdletBinding()]
param (
    [Parameter(HelpMessage = 'Target repository directory')]
    [string] $RepositoryPath,

    [Parameter(HelpMessage = 'Reboot types: 0=none, 3=requires reboot, 5=delayed forced reboot')]
    [string] $RebootTypes = '0,3,5'
)

$ErrorActionPreference = 'Continue'

try
{
    $tsEnv = New-Object -ComObject Microsoft.SMS.TSEnvironment
}
catch
{
    Write-Error "Unable to connect to the task sequence environment: $($_.Exception.Message)"
    exit 1
}

if (-not $RepositoryPath)
{
    $RepositoryPath = Join-Path -Path $tsEnv.Value('_SMSTSMDataPath') -ChildPath 'LenovoUpdates'
}
Write-Output "Repository path: $RepositoryPath"

try
{
    Get-LnvUpdatesRepo -RepositoryPath $RepositoryPath -RebootTypes $RebootTypes -ErrorAction Stop
}
catch
{
    Write-Error "Failed to build the repository: $($_.Exception.Message)"
    exit 1
}

# Get-LnvUpdatesRepo creates the repository folder before it downloads anything, so the
# folder existing proves nothing. database.xml is only written once the catalog is populated.
$database = Join-Path -Path $RepositoryPath -ChildPath 'database.xml'
if (-not (Test-Path -LiteralPath $database))
{
    Write-Error "Repository build did not produce '$database'. No packages were downloaded."
    exit 1
}

$packageCount = @(Get-ChildItem -LiteralPath $RepositoryPath -Directory -ErrorAction SilentlyContinue).Count
Write-Output "Repository built: $packageCount package folder(s), catalog present at $database."

exit 0

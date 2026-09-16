#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Installs Lenovo BIOS and firmware updates during a Configuration Manager OSD task sequence.

.DESCRIPTION
    Runs as a "Run PowerShell Script" step (Windows PowerShell 5.1) after Setup Windows and
    Configuration Manager, and after the repository step has populated the local repository.

    Installs BIOS and firmware packages only, and sets the RebootMandatory task sequence
    variable so a subsequent Restart Computer step can be conditioned on it. Drivers are
    handled by a separate script.

    Assumes the Lenovo.Client.Update module is already present on the device and resolvable
    via PSModulePath.

.PARAMETER RepositoryPath
    Existing repository/scratch directory holding the Lenovo update packages. Defaults to
    <_SMSTSMDataPath>\LenovoUpdates, which resolves to C:\_SMSTaskSequence\LenovoUpdates
    and is cleaned up when the task sequence completes. Leave the step's Parameters field
    empty unless the repository step wrote somewhere else.

.NOTES
    FileName:   Install-LenovoBIOSFirmware.ps1
    Author:     Philip Jorgensen
    Created:    2026-08-27

    Task sequence step: Run PowerShell Script
        Script source:               Enter a PowerShell script (this file pasted inline)
        PowerShell execution policy: Bypass
        Parameters:                  empty, or -RepositoryPath "<custom path>"
        Success codes:               0 3010

    No package is attached to the step, so an earlier step must place the
    Lenovo.Client.Update module on the device where PSModulePath can resolve it.
#>

[CmdletBinding()]
param (
    [Parameter(HelpMessage = 'Repository path containing Lenovo update packages')]
    [string] $RepositoryPath
)

# Left at Continue on purpose: Install-LnvUpdate emits Write-Error and continues when it
# skips a package (e.g. signature check), and Stop would abort the whole run.
$ErrorActionPreference = 'Continue'
$exitCode = 0

try
{
    $tsEnv = New-Object -ComObject Microsoft.SMS.TSEnvironment
}
catch
{
    Write-Error "Unable to connect to the task sequence environment: $($_.Exception.Message)"
    exit 1
}

# Default the variable up front so a later step's condition always has something to evaluate.
$tsEnv.Value('RebootMandatory') = 'False'

if (-not $RepositoryPath)
{
    $RepositoryPath = Join-Path -Path $tsEnv.Value('_SMSTSMDataPath') -ChildPath 'LenovoUpdates'
}

# The repository is built by an earlier task sequence step. Test for the catalog rather
# than the folder: Get-LnvUpdatesRepo creates the folder before downloading, so a failed
# build leaves an empty directory that would otherwise pass as "no updates found".
$database = Join-Path -Path $RepositoryPath -ChildPath 'database.xml'
if (-not (Test-Path -LiteralPath $database))
{
    Write-Error "No repository catalog at '$database'. Confirm the repository step ran and succeeded."
    exit 1
}
Write-Output "Repository path: $RepositoryPath"

try
{
    $updates = @(Get-LnvUpdate -Repository $RepositoryPath -ScratchDirectory $RepositoryPath -ErrorAction Stop |
            Where-Object { $_.Type -in 'BIOS', 'Firmware' })
}
catch
{
    Write-Error "Failed to query the update repository: $($_.Exception.Message)"
    exit 1
}

if ($updates.Count -eq 0)
{
    Write-Output 'No applicable BIOS or firmware updates found.'
    exit 0
}

Write-Output "$($updates.Count) BIOS/firmware update(s) to install."

$results = @($updates | Install-LnvUpdate -Path $RepositoryPath -SaveBIOSUpdateInfoToRegistry -ExportToWMI)

foreach ($result in $results)
{
    $status = if ($result.Success) { 'SUCCESS' } else { "FAILED ($($result.FailureReason))" }
    Write-Output "$($result.ID) | $($result.Title) | $status | ExitCode=$($result.ExitCode) | PendingAction=$($result.PendingAction)"
}

$installed = @($results | Where-Object { $_.Success })
$failed = @($results | Where-Object { -not $_.Success })
# REBOOT_MANDATORY only. A SHUTDOWN pending action would power the machine off and break
# the task sequence, so it is deliberately not treated as a reboot here.
$rebootPending = @($installed | Where-Object { $_.PendingAction -eq 'REBOOT_MANDATORY' })

Write-Output "Summary: $($installed.Count) installed, $($failed.Count) failed, $($results.Count) attempted."

if ($rebootPending.Count -gt 0)
{
    $tsEnv.Value('RebootMandatory') = 'True'
    Write-Output "$($rebootPending.Count) update(s) require a reboot. RebootMandatory set to True."
    $exitCode = 3010
}

if ($failed.Count -gt 0)
{
    Write-Warning "$($failed.Count) update(s) failed. See %ProgramData%\Lenovo\Lenovo.Client.Update\History for details."
}

exit $exitCode

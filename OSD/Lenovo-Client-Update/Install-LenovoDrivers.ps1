#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Installs Lenovo driver updates from a prebuilt local repository during a Configuration
    Manager OSD task sequence.

.DESCRIPTION
    Runs as a "Run PowerShell Script" step (Windows PowerShell 5.1) after Setup Windows and
    Configuration Manager, and after the repository step has populated the local repository.

    Installs driver packages only. BIOS and firmware are handled by a separate script, and
    reboots are left to a Restart Computer step later in the sequence.

    Assumes the Lenovo.Client.Update module is already present on the device and resolvable
    via PSModulePath.

.PARAMETER RepositoryPath
    Existing repository directory holding the Lenovo update packages. Defaults to
    <_SMSTSMDataPath>\LenovoUpdates, which resolves to C:\_SMSTaskSequence\LenovoUpdates
    and is cleaned up when the task sequence completes. Leave the step's Parameters field
    empty unless the repository step wrote somewhere else.

.NOTES
    FileName:   Install-LenovoDrivers.ps1
    Author:     Philip Jorgensen
    Created:    2026-08-27

    Task sequence step: Run PowerShell Script
        Script source:               Enter a PowerShell script (this file pasted inline)
        PowerShell execution policy: Bypass
        Parameters:                  empty, or -RepositoryPath "<custom path>"
        Success codes:               0

    No package is attached to the step, so an earlier step must place the
    Lenovo.Client.Update module on the device where PSModulePath can resolve it.
#>

[CmdletBinding()]
param (
    [Parameter(HelpMessage = 'Existing repository directory')]
    [string] $RepositoryPath
)

# Left at Continue on purpose: Install-LnvUpdate emits Write-Error and continues when it
# skips a package (e.g. signature check), and Stop would abort the whole run.
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
            Where-Object { $_.Type -eq 'Driver' })
}
catch
{
    Write-Error "Failed to query the update repository: $($_.Exception.Message)"
    exit 1
}

if ($updates.Count -eq 0)
{
    Write-Output 'No applicable driver updates found.'
    exit 0
}

Write-Output "$($updates.Count) driver update(s) to install."

$results = @($updates | Install-LnvUpdate -Path $RepositoryPath -ExportToWMI)

foreach ($result in $results)
{
    $status = if ($result.Success) { 'SUCCESS' } else { "FAILED ($($result.FailureReason))" }
    Write-Output "$($result.ID) | $($result.Title) | $status | ExitCode=$($result.ExitCode)"
}

$installed = @($results | Where-Object { $_.Success })
$failed = @($results | Where-Object { -not $_.Success })

Write-Output "Summary: $($installed.Count) installed, $($failed.Count) failed, $($results.Count) attempted."

if ($failed.Count -gt 0)
{
    Write-Warning "$($failed.Count) driver(s) failed. See %ProgramData%\Lenovo\Lenovo.Client.Update\History for details."
}

exit 0

#Requires -Version 5.1

<#
.SYNOPSIS
    Sets the ApplicableUpdates task sequence variable based on the Lenovo updates still
    pending in the local repository.

.DESCRIPTION
    Runs as a "Run PowerShell Script" step (Windows PowerShell 5.1) in the Final Pass
    Catch-All group of the Lenovo Client Update child task sequence, after the second
    driver pass.

    Defaults ApplicableUpdates to False, re-queries the repository built by the repository
    step, and sets the variable to True when any update is still applicable. The 3rd Pass -
    Drivers step is conditioned on ApplicableUpdates equals True, so it runs only when
    something remains to install.

    Assumes the Lenovo.Client.Update module is already present on the device and resolvable
    via PSModulePath.

.NOTES
    FileName:   Set-ApplicableUpdatesVariable.ps1
    Author:     Philip Jorgensen
    Created:    2026-09-16

    Task sequence step: Run PowerShell Script
        Script source:               Enter a PowerShell script (this file pasted inline)
        PowerShell execution policy: Bypass
        Parameters:                  empty
        Success codes:               0
#>

# Initialize TS environment
$tsEnv = New-Object -ComObject Microsoft.SMS.TSEnvironment

# Default the variable up front so a later step's condition always has something to evaluate.
$tsEnv.Value('ApplicableUpdates') = 'False'

# Set repository path
$RepositoryPath = Join-Path -Path $tsEnv.Value('_SMSTSMDataPath') -ChildPath 'LenovoUpdates'

# Query the repository in this session - Get-LnvUpdate runs under Windows PowerShell 5.1
$updates = @(Get-LnvUpdate -Repository $RepositoryPath -ScratchDirectory $RepositoryPath |
        Where-Object { $_.IsApplicable })

if ($updates.Count -gt 0)
{
    $tsEnv.Value('ApplicableUpdates') = 'True'
}

Write-Output "$($updates.Count) applicable update(s). ApplicableUpdates set to $($tsEnv.Value('ApplicableUpdates'))."

#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Installs the Lenovo Client Update and Lenovo Client Scripting modules from the PowerShell
    Gallery during a Configuration Manager OSD task sequence.

.DESCRIPTION
    Runs as a "Run PowerShell Script" step (Windows PowerShell 5.1) at the start of the
    Lenovo Client Update child task sequence, ahead of every step that calls LCU cmdlets.

    Queries the PowerShell Gallery API for the latest version of each module, downloads the
    package directly, and expands it into a versioned folder under
    %ProgramFiles%\WindowsPowerShell\Modules. Downloading the package directly avoids
    Install-Module, which prompts to install the NuGet provider on a freshly imaged device.

    A module is skipped when the installed version is already current.

.NOTES
    FileName:   Install-LenovoModule.ps1
    Author:     Philip Jorgensen
    Created:    2026-08-27

    Task sequence step: Run PowerShell Script
        Script source:               Enter a PowerShell script (this file pasted inline)
        PowerShell execution policy: Bypass
        Parameters:                  empty
        Success codes:               0

    The device must reach www.powershellgallery.com during the task sequence. Later steps
    rely on module auto-loading, since %ProgramFiles%\WindowsPowerShell\Modules is on the
    default PSModulePath.
#>

$modules = @(
    'Lenovo.Client.Update'
    'Lenovo.Client.Scripting'
)

# The gallery requires TLS 1.2, which Windows PowerShell 5.1 does not negotiate by default.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$moduleRoot = "$env:ProgramFiles\WindowsPowerShell\Modules"
Write-Output "Modules will be installed to $moduleRoot"

# Ensure the target root is on PSModulePath for this session.
if (-not ($env:PSModulePath -split ';' -contains $moduleRoot))
{
    $env:PSModulePath = $env:PSModulePath + ";$moduleRoot"
    Write-Output "Added $moduleRoot to PSModulePath"
}

foreach ($module in $modules)
{
    $params = @{
        Uri  = 'https://www.powershellgallery.com/api/v2/FindPackagesById()'
        Body = @{
            '$filter' = 'IsLatestVersion eq true'
            id        = "'$module'" # Extra quotes are required.
        }
    }
    $searchResult = Invoke-RestMethod @params
    if (-not $searchResult)
    {
        Write-Warning "$module was not found in the gallery. Skipping."
        continue
    }

    $installedVersion = Get-Module -Name $module -ListAvailable |
        Sort-Object -Property Version -Descending |
        Select-Object -First 1 -ExpandProperty Version

    if ($installedVersion -and [version]$installedVersion -ge [version]$searchResult.properties.version)
    {
        Write-Output "$module $installedVersion is already installed. Skipping."
        continue
    }

    # Download the module archive
    Invoke-RestMethod -Uri $searchResult.content.src -OutFile module.zip

    # Expand the archive to a versioned folder under the module root.
    $params = @{
        Path        = 'module.zip'
        Destination = $moduleRoot |
            Join-Path -ChildPath $module |
            Join-Path -ChildPath $searchResult.properties.version
        Force       = $true
    }
    Expand-Archive @params

    # Clean up redundant package files.
    Get-ChildItem -Path $params.Destination |
        Where-Object Name -In @(
            '_rels'
            'package'
            '[Content_Types].xml'
            "$module.nuspec"
        ) |
        Remove-Item -Recurse
    Write-Output "Installed $module $($searchResult.properties.version) to $($params.Destination)"

    # Remove downloaded archive
    Remove-Item -Path module.zip
}

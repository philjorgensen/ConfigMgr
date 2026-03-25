<#
.SYNOPSIS
    Creates ConfigMgr Device Collections for Lenovo systems

.DESCRIPTION
    Script that checks the ConfigMgr database for Lenovo branded systems and creates a Device Collection based off the friendly name.

    Example: ThinkPad X1 Carbon Gen 13

.PARAMETER SiteServer
    Fully qualified domain name of Site Server

.EXAMPLE
    .\New-LenovoDeviceCollections-Updated.ps1 -SiteServer cm01.domain.com

.NOTES
    Author:     Philip Jorgensen
    Created:    2021-12-16
    Updated:    2026-02-26
    Filename:   New-LenovoDeviceCollections-Updated.ps1

    Version history:
    1.0 - Initial script development and testing
    2.0 - Refactored code for improved error handling, added progress indicators, and enhanced collection management logic

    Run script from Site Server or system with the ConfigMgr console installed.
#>

[CmdletBinding()]
[OutputType([String])]
param (
    [Parameter(Mandatory = $true,
        HelpMessage = "FQDN of Site Server",
        ValueFromPipeline = $false)]
    [ValidateNotNullOrEmpty()]
    $SiteServer
)

$ErrorActionPreference = "Stop"

function Connect-SCCMSite
{
    param (
        [string]$SiteServer
    )

    try
    {
        Write-Host "Importing ConfigMgr module" -ForegroundColor Yellow
        Import-Module $env:SMS_ADMIN_UI_PATH.Replace("bin\i386", "bin\ConfigurationManager.psd1") -Force
    }
    catch
    {
        throw "Failed to import ConfigMgr module"
    }

    try
    {
        # Retrieve Site Code from SCCM server
        $SiteCode = Get-CimInstance -ComputerName $SiteServer -Namespace root/SMS -ClassName SMS_ProviderLocation -ErrorAction Stop |
            Select-Object -ExpandProperty SiteCode -First 1

        if (-not $SiteCode)
        {
            throw "Failed to retrieve Site Code from $SiteServer."
        }

        # Display progress while connecting
        Write-Progress -Activity "Connecting to Site Server" -Status "Creating PSDrive for SCCM"

        # Create PSDrive for SCCM
        if (-not (Get-PSDrive -PSProvider CMSite -Name $SiteCode -ErrorAction SilentlyContinue))
        {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer -Description "Primary Site Server" -Scope Global -ErrorAction Stop | Out-Null
        }

        # Set location to SCCM PSDrive
        Set-Location -Path "$SiteCode`:"

        Write-Host "Successfully connected to SCCM Site Server: $SiteServer ($SiteCode`:)" -ForegroundColor Green
        return $SiteCode
    }
    catch
    {
        throw "Error connecting to SCCM Site Server: $_"
    }
}

function New-LenovoDeviceCollections
{
    param (
        [string]$SiteServer,
        [string]$SiteCode
    )
    $Vendor = "Lenovo"
    $Subfolder = Get-CMFolder -ObjectTypeName SMS_Collection_Device -Name $Vendor -ErrorAction SilentlyContinue
    $Models = Get-CimInstance -ComputerName $SiteServer -Namespace "root\SMS\site_$($SiteCode)" -Query "Select * From SMS_G_System_COMPUTER_SYSTEM_PRODUCT Where Vendor = 'LENOVO'" | Select-Object -Property Version | Sort-Object -Unique -Property Version

    if ($Models.Count -ne 0)
    {
        if ($null -eq $Subfolder)
        {
            $Subfolder = New-CMFolder -Name $Vendor -ParentFolderPath 'DeviceCollection' -ErrorAction Stop | Out-Null
        }

        $i = 1
        $iModel = $Models.Count
        $Days = @('Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday')
        foreach ($Model in $Models)
        {
            $ModelVersion = $Model.Version
            Write-Progress "Adding devices to named device collections." -Status "Updating the $ModelVersion collection. $i of $iModel" -PercentComplete ($i++ / $iModel * 100)

            # Generate a unique staggered schedule per collection
            $RandomDay = Get-Random -InputObject $Days
            $RandomHour = Get-Random -Minimum 0 -Maximum 24
            $RandomMinute = Get-Random -InputObject @(0, 15, 30, 45)
            $Schedule = New-CMSchedule -DayOfWeek $RandomDay -Start (Get-Date -Hour $RandomHour -Minute $RandomMinute -Second 0)

            try
            {
                if (-not (Get-CMDeviceCollection -Name $ModelVersion))
                {
                    Write-Host "Creating collection for $ModelVersion" -ForegroundColor Cyan
                    $NewCollection = New-CMDeviceCollection -Name "$ModelVersion" -LimitingCollectionName 'All Systems' -RefreshType Periodic -RefreshSchedule $Schedule

                    Move-CMObject -InputObject $NewCollection -FolderPath $SiteCode":\DeviceCollection\$Vendor"

                }
                $CollectionQuery = "Select * From SMS_R_System inner join SMS_G_System_COMPUTER_SYSTEM_PRODUCT on SMS_G_System_COMPUTER_SYSTEM_PRODUCT.ResourceId = SMS_R_System.ResourceId where SMS_G_System_COMPUTER_SYSTEM_PRODUCT.Version = '$ModelVersion'"
                if (-not (Get-CMDeviceCollectionQueryMembershipRule -CollectionName $ModelVersion -RuleName $ModelVersion))
                {
                    Add-CMDeviceCollectionQueryMembershipRule -CollectionName "$ModelVersion" -QueryExpression $CollectionQuery -RuleName $ModelVersion
                }
            }
            catch
            {
                Write-Warning "Failed to process collection for '$ModelVersion': $_"
            }
        }
    }
    else
    {
        Write-Host "No Lenovo Models detected in the CM database..." -ForegroundColor Red
    }
}

function Disconnect-ConfigMgrSite
{
    param (
        [string]$SiteCode
    )
    Write-Host "Device Collections updated." -ForegroundColor Green
    Write-Host "Disconnecting from Site..." -ForegroundColor Yellow
    Set-Location -Path $env:USERPROFILE
    Remove-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue
}

# Main
$SiteCode = Connect-SCCMSite -SiteServer $SiteServer
if ($SiteCode)
{
    New-LenovoDeviceCollections -SiteServer $SiteServer -SiteCode $SiteCode
    Disconnect-ConfigMgrSite -SiteCode $SiteCode
}
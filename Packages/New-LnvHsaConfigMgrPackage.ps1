#requires -modules ConfigurationManager
#requires -runasadministrator

<#
  .SYNOPSIS
  Downloads Lenovo HSA pack from the Think Deploy catalog

  .DESCRIPTION
  This cmdlet will download the HSA Pack based on the specified machine
  type to a temporary directory, extract the contents, and move to a ConfigMgr
  share. A ConfigMgr Package will then be created with necessary fields set.
  
  Choose an available pack from the Out-GridView window. If a pack isn't available,
  script will end.

  .PARAMETER SiteServer
  Mandaory: True


  .PARAMETER MachineType
  Mandatory: True
  First 4 characters of Machine Type Model

  .PARAMETER HsaPackSourceLocation
  Mandatory: True
  UNC path to file share where contents will be stored

  .EXAMPLE
  New-LnvHsaConfigMgrPackage -SiteServer \\siteserver.corp.com -HsaPackSourceLocation \\fileshare.corp.com\Content\HSAs\Lenovo -MachineType 21DD

  .NOTES
    FileName:    New-LnvHsaConfigMgrPackage.ps1
    Author:      Philip Jorgensen
    Created:     2024-3-12

#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory = $true, HelpMessage = "FQDN of Site Server")]
    [ValidateNotNullOrEmpty()]
    [string]$SiteServer,

    [Parameter(Mandatory = $true, HelpMessage = "Path to share where contents are stored")]
    [ValidateNotNullOrEmpty()]
    [string]$HsaPackSourceLocation,

    [Parameter(Mandatory = $true, HelpMessage = "First 4 characters of Machine Type Model")]
    [ValidateLength(4, 4)]
    [String] $MachineType
)

begin

{
    $ErrorActionPreference = 'SilentlyContinue'

    $StartingLocation = (Get-Location).Path

    # Format MTM
    $MachineType = $MachineType.ToUpper()

    # Initialize flag for CM connection
    $ConnectionEstablished = $false
}

process
{
    # Variable for Catalog
    $CatalogUrl = "https://download.lenovo.com/cdrt/td/catalogv2.xml"

    try
    {
        [xml]$Catalog = (New-Object -TypeName System.Net.WebClient).DownloadString($CatalogUrl)
    }
    catch
    {
        Write-Output "Could not obtain the driver automation catalog."
        break
    }

    $Node = $Catalog.ModelList.Model | Where-Object { $_.Types.Type -eq "$MachineType" }
        
    $HsaPackExists = $Node | Get-Member -MemberType Property | Where-Object { $_.Name -eq "HSA" }
    if ($null -eq $HsaPackExists)
    {
        Write-Output "No HSA pack exists for $MachineType"
        break
    }
    else
    {
        # Variables to hold pack data
        $HsaSelections = $Node.HSA | Out-GridView -PassThru -Title "Select the HSA pack to download"
        if ($null -eq $HsaSelections)
        {
            Write-Output "Selection cancelled"
            break
        }
    }

    ForEach ($HsaSelection in $HsaSelections)
    {

        $WindowsVersion = ($HsaSelection).os
        $WindowsBuild = ($HsaSelection).version
        $HsaPackUrl = ($HsaSelection).'#text'
        $HsaPackVersion = $HsaPackUrl.Split("_")[-1].Split(".")[0]
        $HsaPackName = $(Split-Path -Path $HsaPackUrl -Leaf).Replace(".exe", "")

        # Temp directory to extract HSA contents
        $HsaTempDirectory = Join-Path -Path $env:TEMP -ChildPath $HsaPackName -Verbose

        # Variable to store pack executable
        $HsaExe = ($HsaTempDirectory, ".exe" -join "")
    
        # Temp directory to download pack
        Write-Output "Downloading $HsaPackName"
        Invoke-WebRequest -Uri $HsaPackUrl -OutFile $HsaExe -UseBasicParsing

        # Extract Pack
        Write-Output "Extract HSA pack contents"
        Start-Process -FilePath (Get-ChildItem -Path $HsaExe -Filter "*hsa*.exe") -ArgumentList "/VERYSILENT /DIR=$HsaTempDirectory" -Wait -Verbose

        #region CREATEHSAINSTALLSCRIPT
        Write-Output "Creating HSA installation script"

        $InstallHsaScriptPresence = Get-ChildItem -Path $HsaTempDirectory -Filter "Install-HSAs.ps1"
        if (-not($InstallHsaScriptPresence.IsPresent))
        {
            # Create HSA installation script
            $InstallHsaScript = {
                #######################
                #  SCRIPT PARAMETERS  #
                #######################
                [CmdletBinding(DefaultParameterSetName = 'GetList')]
                Param(
                    [Parameter(ParameterSetName = 'GetList')]
                    [switch]$List,
                    [Parameter(ParameterSetName = 'GetList')]
                    [switch]$Export,
                    [Parameter(ParameterSetName = 'InstallOffline')]
                    [switch]$Offline,
                    [Parameter(ParameterSetName = 'InstallOffline')]
                    [ValidateNotNullOrEmpty()]
                    [string]$Name,
                    [Parameter(ParameterSetName = 'InstallOffline')]
                    [ValidateNotNullOrEmpty()]
                    [string]$File,
                    [Parameter(ParameterSetName = 'InstallOffline')]
                    [switch]$All,
                    [Parameter(ParameterSetName = 'InstallOffline')]
                    [ValidateNotNullOrEmpty()]
                    [string]$NoSMSTS,
                    [Parameter(ParameterSetName = 'GetList')]
                    [Parameter(ParameterSetName = 'InstallOffline')]
                    [switch]$DebugInformation
                )
                ###############
                #  FUNCTIONS  #
                ###############
                #Install-HSA
                Function Install-HSA
                {
                    [CmdletBinding()]
                    param (
                        [PSCustomObject]$HSAPackage,
                        [String]$HSAName
                    )
                    $OutDep = $Null
                    If ((($HSAName -contains $HSAPackage.hsa) -and ($Null -ne $HSAName)) -or $All)
                    {
                        ForEach ($Dep in $HSAPackage.Dependencies)
                        {
                            $OutDep += " /DependencyPackagePath:`"$($HSAPackage.JSONPath)\$($Dep)`""
                        }
                        $DISMLog = ""
                        If ($DebugInformation)
                        {
                            If (!(Test-Path -Path "$($LogPath)\DISM"))
                            {
                                New-Item -Path "$($LogPath)\DISM" -ItemType Directory
                            }
                            $DISMLog = " /LogLevel:4 /LogPath:`"$LogPath\DISM\$($HSAPackage.hsa).log`""
                        }
                        $DISMArgs = "/Add-ProvisionedAppxPackage /PackagePath:`"$($HSAPackage.JSONPath)\$($HSAPackage.appx)`" /LicensePath:`"$($HSAPackage.JSONPath)\$($HSAPackage.license)`"$($OutDep) /Region:`"All`"$DISMLog"
                        Write-Host "Offline DISM - $($HSAPackage.hsa)"
                        If ($NoSMSTSPresent)
                        {
                            Write-Host "Using string data from NoSMSTS parameter to define the root drive letter for the DISM /Image parameter."
                        }
                        $DISMArgs = "/Image:$($Drive)\ $($DISMArgs)"
                        Write-Host "$env:windir\system32\Dism.exe $DISMArgs"
                        Start-Process -FilePath "$env:windir\system32\Dism.exe" -ArgumentList $DISMArgs -Wait
                    }
                }
                ##################
                #  SCRIPT SETUP  #
                ##################
                $FilePresent = $false
                $NamePresent = $false
                $NoSMSTSPresent = $false
                If (($PSBoundParameters.ContainsKey('List') -and $PSBoundParameters.ContainsKey('Offline')))
                {
                    Write-Host "Use just one from the following list of parameters: -List or -Offline.  Review the script usage information for using these parameters."
                    Return 1
                }
                If ($Offline)
                {
                    If (($PSBoundParameters.ContainsKey('All') -and $PSBoundParameters.ContainsKey('File') -and $PSBoundParameters.ContainsKey('Name')) -or ($PSBoundParameters.ContainsKey('All') -and $PSBoundParameters.ContainsKey('File')) -or ($PSBoundParameters.ContainsKey('All') -and $PSBoundParameters.ContainsKey('Name')) -or ($PSBoundParameters.ContainsKey('File') -and $PSBoundParameters.ContainsKey('Name')) -or ((!($PSBoundParameters.ContainsKey('All')) -and (!($PSBoundParameters.ContainsKey('File'))) -and (!($PSBoundParameters.ContainsKey('Name'))))))
                    {
                        Write-Host "Use just one from the following list of parameters: -All, -Name, or -File.  Review the script usage information for using these parameters."
                        Return 2
                    }
                    ElseIf ($PSBoundParameters.ContainsKey('File'))
                    {
                        $FilePresent = $true
                    }
                    ElseIf ($PSBoundParameters.ContainsKey('Name'))
                    {
                        $NamePresent = $true
                    }
                }
                If ($PSBoundParameters.ContainsKey('NoSMSTS'))
                {
                    $NoSMSTSPresent = $true
                }
                #Setup Vars
                $ScriptDir = Split-Path $Script:MyInvocation.MyCommand.Path
                If (($Offline) -and (!($NoSMSTSPresent)))
                {
                    $TSenv = New-Object -ComObject Microsoft.SMS.TSEnvironment
                    If ($TSEnv.value("OSDTargetSystemDrive") -ne "")
                    {
                        $Drive = "$($TSEnv.value("OSDTargetSystemDrive"))"
                    }
                    Else
                    {
                        $Drive = "$($TSEnv.value("OSDisk"))"
                    }
                }
                ElseIf (($Offline) -and ($NoSMSTSPresent))
                {
                    $Drive = $NoSMSTS
                }
                ElseIf ($List)
                {
                    $Drive = $env:SystemDrive
                }
                $LogDate = Get-Date -Format yyyyMMddHHmmss
                If ($DebugInformation)
                {
                    $LogPath = "$($Drive)\Windows\Logs"
                    $LogFile = "$LogPath\$($myInvocation.MyCommand)_$($LogDate).log"
    
                    ######################
                    #  START TRANSCRIPT  #
                    ######################
                    Start-Transcript $LogFile -Append -NoClobber
                    Write-Host "Debug enabled"
                }
                ##########
                #  MAIN  #
                ##########
                $MFJs = Get-ChildItem -Path $ScriptDir -Recurse -File -Include "*_manifest.json"
                If ($Null -eq $MFJs)
                {
                    Write-Host "No HSA_Manifest.JSON files found in the subfolder structure."
                    Return 3
                }
                Else
                {
                    $HSAPackages = @()
                    ForEach ($MFJ in $MFJs)
                    {
                        $MFJData = Get-Content -Path "$($MFJ.FullName)" | ConvertFrom-Json
                        $HSAPackages += New-Object PSObject -Property @{'JSONPath' = $MFJ.DirectoryName; 'HSA' = $MFJData.HSA; 'Appx' = $MFJData.Appx; 'License' = $MFJData.License; 'Dependencies' = $MFJData.Dependencies }
                    }
                }
                If ($List -or $PSBoundParameters.Count -eq 0 -or ($DebugInformation -and $PSBoundParameters.Count -eq 1))
                {
                    If ($PSBoundParameters.ContainsKey('Export'))
                    {
                        ForEach ($Package in $HSAPackages)
                        {
                            $Package.hsa | Out-File "$ScriptDir\Export_$LogDate.txt" -Append -NoClobber
                        }
                    }
                    Else
                    {
                        ForEach ($Package in $HSAPackages)
                        {
                            Write-Host $Package.hsa
                        }
                    }
                }
                If ($Offline)
                {
                    If ($All)
                    {
                        Write-Host "Installing all HSAs found in the folder structure."
                    }
                    ElseIf ($NamePresent)
                    {
                        Write-Host "Installing the $Name HSA."
                    }
                    ElseIf ($FilePresent)
                    {
                        Write-Host "Reading the list of HSAs from $File."
                        If (!(Test-Path -Path "$ScriptDir\$File"))
                        {
                            Write-Host "File: $File not found in $ScriptDir"
                            Return 4
                        }
                    }
                    ForEach ($Package in $HSAPackages)
                    {
                        If ($All)
                        {
                            Install-HSA -HSAPackage $Package
                        }
                        ElseIf ($FilePresent -or $NamePresent)
                        { 
                            If ($FilePresent)
                            {
                                $InstallFileArray = @()
                                $InstallFileArray = Get-Content -Path "$ScriptDir\$File"
                                ForEach ($InstallFile in $InstallFileArray)
                                {
                                    Install-HSA -HSAPackage $Package -HSAName $InstallFile
                                }
                            }
                            If ($NamePresent)
                            {
                                Install-HSA -HSAPackage $Package -HSAName $Name
                            }
                        }
                        $InstallFileArray = $Null
                    }
                }
                If ($DebugInformation)
                {
                    #####################
                    #  STOP TRANSCRIPT  #
                    #####################
                    Stop-Transcript
                }
            }

            $Script = New-Item -Path $HsaTempDirectory -Name "Install-HSAs.ps1" -ItemType File -Value $InstallHsaScript -Force

        }
        #endregion

        #region MOVECONTENTS
        # Move pack contents to share
        $PackageSourcePath = (Join-Path -Path $HsaPackSourceLocation -ChildPath $HsaPackName)

        if (-not(Test-Path -Path $PackageSourcePath))
        {
            try
            {
                Write-Output "Moving HSA contents to ConfigMgr share"
                Move-Item -Path $HsaTempDirectory -Destination "Microsoft.Powershell.Core\FileSystem::$HsaPackSourceLocation" -Force -Verbose
                Write-Output "Contents arrived at destination"
            }
            catch
            {
                throw "Failed to move contents to share"
            }
        }
        else
        {
            Write-Output "HSA pack contents are already in the specified source location"
        }
        #endregion

        # Check if connection to ConfigMgr site server has been established
        if (-not $ConnectionEstablished)
        {
            try
            {
                Write-Output "Importing ConfigMgr module"
                Import-Module $env:SMS_ADMIN_UI_PATH.Replace("bin\i386", "bin\ConfigurationManager.psd1") -Force
            }
            catch
            {
                throw "Failed to import ConfigMgr module"
            }
            #Connect to ConfigMgr Site 
            $SiteCode = $(Get-CimInstance -ComputerName $SiteServer -Namespace root/SMS -ClassName SMS_ProviderLocation).SiteCode

            if (-not(Get-PSDrive $SiteCode))
            {
                Write-Progress -Activity "Connecting to Site Server"
                New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer -Description "Primary Site Server"
            }

            Set-Location -Path (((Get-PSDrive -PSProvider CMSite -Verbose:$false).Name) + ":")
        
            # Set flag to indicate connection has been established
            $ConnectionEstablished = $true
        }

        #region BUILDCONFIGMGRPKG
        # Check if a ConfigMgr Package already exists
        $PackageExists = Get-CMPackage -Fast | Where-Object { $_.Name -eq ($($Node.Name)) -and $_.Version -eq $HsaPackVersion -and $_.MifFileName -eq "HSA" -and $_.MifName -eq $WindowsVersion }
    
        if (-not($PackageExists))
        {
            # Build ConfigMgr HSA Package
            $newPackageSplatParams = @{
                Description = $($Node.Types.Type -join ",")
                Name        = $($Node.name)
                Path        = (Join-Path -Path $HsaPackSourceLocation -ChildPath $HsaPackName)
            }

            Write-Output "Creating ConfigMgr HSA Package"
            $NewCmPackage = New-CMPackage @newPackageSplatParams

            $packageSplatParams = @{
                MifFileName = "HSA"
                MifName     = $WindowsVersion
                Version     = $HsaPackVersion
            }
            
            # Add conditional parameters based on $WindowsVersion
            if ($WindowsVersion -ne "win10")
            {
                $packageSplatParams.Add("MifVersion", "$WindowsBuild" + "_23H2")
            }
            else
            {
                $packageSplatParams.Add("MifVersion", $WindowsBuild)
            }
            Write-Output "Setting helpful package data"
            Get-CMPackage -Id $NewCmPackage.PackageID -Fast | Set-CMPackage @packageSplatParams -Verbose
        }
        else
        {
            Write-Output "ConfigMgr Package for this HSA Pack already exists"
        }
        #endregion

        # Clean up
        Write-Output "Removing temp files"
        Remove-Item -Path $HsaExe -Force
        Remove-Item -Path $HsaTempDirectory -Recurse -Force
    }
}

end

{
    Write-Output "ConfigMgr Packages ready to distribute"

    # Disconnect from CM Site
    Write-Output "Disconnecting from $SiteServer"
    Remove-PSDrive -Name $SiteCode -PSProvider CMSite -Force -Verbose
    Set-Location -Path $StartingLocation  
}
<#
DISCLAIMER:

These sample scripts are not supported under any Lenovo standard support

program or service. The sample scripts are provided AS IS without warranty

of any kind. Lenovo further disclaims all implied warranties including,

without limitation, any implied warranties of merchantability or of fitness for

a particular purpose. The entire risk arising out of the use or performance of

the sample scripts and documentation remains with you. In no event shall

Lenovo, its authors, or anyone else involved in the creation, production, or

delivery of the scripts be liable for any damages whatsoever (including,

without limitation, damages for loss of business profits, business interruption,

loss of business information, or other pecuniary loss) arising out of the use

of or inability to use the sample scripts or documentation, even if Lenovo

has been advised of the possibility of such damages.
#>

<#
.SYNOPSIS
    Script to be executed from SCCM which will read the database.xml from an Update Retriever repository to find any available BIOS update for the target system.

.DESCRIPTION
    Script reads the XML for the update to determine if it is a newer version than what is on the target device.
    If applicable, the update is executed silently with reboots suppressed (e.g. Winuptp.exe -s)

.PARAMETER Path
    UNC path to Update Retriever repository
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [parameter(Mandatory = $true, HelpMessage = "Specify the UNC path to the Update Retriever repository")]
    [ValidateNotNullOrEmpty()]
    [string]$Path
)

<# Function
    If absent, creates the Lenovo WMI Namespace and Lenovo_Updates Class with the following properties:
        AdditionalInfo
        PackageID
        Status
        Title
        Version
#>
function CreateClass {
    $ns = [wmiclass]'root:__NAMESPACE'
    $sc = $ns.CreateInstance()
    $sc.Name = 'Lenovo'
    $sc.Put()

    $class = New-Object System.Management.ManagementClass ("root\Lenovo", [string]::Empty, $null)
    $class["__CLASS"] = "Lenovo_Updates"
    $class.Qualifiers.Add("SMS_Report", $true)
    $class.Qualifiers.Add("SMS_Group_Name", "Lenovo_Updates")
    $class.Qualifiers.Add("SMS_Class_Id", "Lenovo_Updates")

    $class.Properties.Add("PackageID", [System.Management.CimType]::String, $false)
    $class.Properties.Add("Title", [System.Management.CimType]::String, $false)
    $class.Properties.Add("Status", [System.Management.CimType]::String, $false)
    $class.Properties.Add("AdditionalInfo", [System.Management.CimType]::String, $false)
    $class.Properties.Add("Version", [System.Management.CimType]::String, $false)

    $class.Properties["PackageID"].Qualifiers.Add("Key", $true)
    $class.Properties["PackageID"].Qualifiers.Add("SMS_Report", $true)
    $class.Properties["Title"].Qualifiers.Add("SMS_Report", $true)
    $class.Properties["Status"].Qualifiers.Add("SMS_Report", $true)
    $class.Properties["AdditionalInfo"].Qualifiers.Add("SMS_Report", $true)
    $class.Properties["Version"].Qualifiers.Add("SMS_Report", $true)

    $class.Put()
}
<# Function
    Updates Lenovo_Updates Class with BIOS package ID and install status
#>
function AddStatus {

    $winuptpLog = ("$extractDir" + "\Winuptp.log")

    ForEach-Object {
        $packageid = $pkg
        $title = $nodes.description
        $status = ((Get-Content -Tail 3 -Path $winuptpLog) | Out-String).Trim()
        $version = $nodes.Version
        try {
            $update = Get-WmiObject -Namespace root\Lenovo -Class Lenovo_Updates -Filter "PackageID = '$packageid'"
            if ($update.PackageID -eq $packageid) {
                if ($update.Status -ne $status -or $update.Title -ne $title -or $update.Version -ne $version) {
                    $update.Status = $status
                    $update.Title = $title
                    $update.Version = $version
                    $update.Put()
                }
            }
            else {
                Set-WmiInstance -Namespace root\Lenovo -Class Lenovo_updates -Arguments @{PackageID = $packageid; Title = $title; Status = $status; Version = $version } -PutType CreateOnly
            }
        }
        catch {
            "Did not add"
            $packageid + " " + $title + " " + $status
        }
    }
}
# // Create the Lenovo WMI Namespace // #
[void](Get-WmiObject -Namespace root\Lenovo -Class Lenovo_Updates -ErrorAction SilentlyContinue -ErrorVariable wmiclasserror)
if ($wmiclasserror) {
    try {
        Write-Output "================================="
        Write-Output "Creating the Lenovo WMI Namespace"
        Write-Output "================================="
        CreateClass
    }
    catch {
        Write-Warning -Message "Could not create WMI class" ; Exit 1
    }
}

# Set variable for first 4 characters of MTM
$bios = ((Get-WmiObject -Namespace root\cimv2 -ClassName Win32_Bios).SMBIOSBIOSVersion.Split('(')[1].Split(')') | Out-String).Trim()

# Set variable for first 4 characters of BIOS
$mtm = ((Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty Model).SubString(0, 4)).Trim()

# Set variable for database.xml
$dbXML = [xml](Get-Content -Path "$Path\database.xml")

# Locate the package and package XML
$nodes = $dbXML.SelectNodes("/Database/Package/SystemCompatibility/System[@mtm='$mtm']/../..") | Where-Object { $_.description -match "BIOS" } | Select-Object -Last 1

# Set variable for package path
$pkgPath = Join-Path $Path $nodes.id

# Set variable for package ID
$pkg = $nodes.id

# Set variable for package title
$pkgTitle = $nodes.description

# Set variable for package XML path
$pkgXMLPath = Join-Path $Path $nodes.LocalPath

# Compare BIOS versions from package XML to client
if ($nodes.Version -match $bios) {
    Write-Output "==============="
    Write-Output "BIOS is current"
    Write-Output "==============="
}
else {

    # Set variable for package XML
    $pkgXML = [xml](Get-Content -Path "$pkgXMLPath")

    # Pull BIOS version from package XML
    $currentVer = $pkgXML.SelectSingleNode("//Package") | Select-Object -ExpandProperty version
    Write-Output "==============================================="
    Write-Output "BIOS will be updated to $currentVer"
    Write-Output "==============================================="

    # Suspend BitLocker if enabled
    $bde = Get-BitLockerVolume -MountPoint $env:SystemDrive | Select-Object -Property ProtectionStatus

    if ($bde.ProtectionStatus -eq "On") {
        Write-Output "========================================="
        Write-Output "Suspending BitLocker prior to BIOS update"
        Write-Output "========================================="
        Suspend-BitLocker -MountPoint $env:SystemDrive
    }

    # Set extraction point
    $extractDir = ("$env:HOMEDRIVE" + "\BIOS" + "\$pkg")

    # Set variable for flash utility
    $install = "winuptp64.exe"

    # Set variable for silent switch
    $silentSwitch = "-s"

    try {
        Write-Output "======================="
        Write-Output "Extracting BIOS package"
        Write-Output "======================="

        Start-Process -FilePath ("$pkgPath\$pkg" + ".exe") -ArgumentList "/VERYSILENT /DIR=$extractDir /EXTRACT=YES" -PassThru -Wait

        Write-Output "================"
        Write-Output "Updating BIOS..."
        Write-Output "================"

        $flash = Start-Process -FilePath $install -WorkingDirectory $extractDir -ArgumentList $silentswitch -PassThru -Wait

        $winuptplog = Get-Content -Path ($extractDir + "\winuptp.log")
        Write-Output $winuptplog

        if ($flash.ExitCode -eq 1) {
            Write-Output "====================================="
            Write-Output "BIOS update complete..."
            Write-Output "====================================="
            AddStatus
        }
    }
    catch [System.Exception] {
        Write-Warning -Message "An error occured during BIOS update..."
        Resume-BitLocker -MountPoint $env:SystemDrive; Exit 1
    }
}

# Clean up
Remove-Item -Path ("$env:HOMEDRIVE" + "\BIOS") -Recurse -Force -ErrorAction SilentlyContinue

# Display balloon tip to reboot system

if ($flash.ExitCode -eq 1) {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

    $objNotifyIcon = New-Object System.Windows.Forms.NotifyIcon

    $objNotifyIcon.Icon = [Drawing.Icon]::ExtractAssociatedIcon((Get-Command powershell).Path)
    $objNotifyIcon.BalloonTipText = "A reboot is required to complete BIOS update"
    $objNotifyIcon.BalloonTipTitle = "BIOS Update Complete"
    $objNotifyIcon.Visible = $true
    $objNotifyIcon.ShowBalloonTip(10000)
}
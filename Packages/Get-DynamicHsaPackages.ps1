[CmdletBinding()]
param (

    [Parameter(ValueFromPipelineByPropertyName,
        Position = 0)]
    [string]
    $MatchProperty = 'Description',

    [Parameter(ValueFromPipelineByPropertyName,
        Position = 1)]
    [string]
    $MachineType = (Get-CimInstance -Namespace root/CIMV2 -ClassName Win32_ComputerSystemProduct).Name.Substring(0, 4).Trim(),

    [Parameter(ValueFromPipelineByPropertyName,
        Position = 2)]
    [string]
    $PackageXMLLibrary = ".\_Packages.xml",

    
    [Parameter(ValueFromPipelineByPropertyName,
        Position = 3)]
    [ValidateSet("win10", "win11")]
    [string]
    $WindowsVersion = "",

    [Parameter(ValueFromPipelineByPropertyName,
        Position = 4)]
    [ValidateSet("1709", "1803", "1809", "1903", "1909", "2004", "20H2", "21H1", "21H2", "22H2", "23H2", "24H2")]
    [string]
    $WindowsBuild = ""
)

#interesting properties pkgsourcepath, Description, ISVData, ISVString, Manufacturer, MifFileName, MifName, MifPublisher, MIFVersion, Name, PackageID, ShareName, Version
[xml]$Packages = Get-Content -Path $PackageXMLLibrary

#environment variable call for task sequence only

try
{
    $tsenv = New-Object -ComObject Microsoft.SMS.TSEnvironment
    $tsenvInitialized = $true
}
catch
{
    Write-Host -Object 'Not executing in a tasksequence'
    $tsenvInitialized = $false
}

$PackageID = (Import-Clixml $PackageXMLLibrary | Where-Object { $_.$MatchProperty.Split(',').Contains($MachineType) -and $_.MifFileName -eq "HSA" -and $_.MifName -eq $WindowsVersion -and $_.MifVersion -match $WindowsBuild }).PackageID
$PackageID
if ($tsenvInitialized)
{
    $tsenv.Value('OSDDownloadDownloadPackages') = $PackageID
}
$ErrorActionPreference = "SilentlyContinue"

If (Get-AppxPackage -Name E046963F.LenovoSettingsforEnterprise -all) {

    If (Get-Service -Name ImControllerService) {

        If (Get-Service -Name LenovoVantageService) {

            # Check for older of version of Vantage Service that causes UAC prompt. This is due to an expired certificate.  
            $minVersion = "3.8.23.0"
            $path = ${env:ProgramFiles(x86)} + "\Lenovo\VantageService\*\LenovoVantageService.exe"
            $version = (Get-ChildItem -Path $path).VersionInfo.FileVersion
            
            if ([version]$version -ge [version]$minVersion) {
                
                Write-Host "All Commercial Vantage Services and App are installed..."
            }

            else {
            
            }
        
        }
    
    }

}
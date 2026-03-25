# Device Collections

Automatically creates and maintains ConfigMgr device collections organized by Lenovo system model and version.

## Requirements

- System Center Configuration Manager (SCCM) site server or system with ConfigMgr console installed
- PowerShell 5.0 or later
- Administrator privileges
- Hardware inventory enabled with **Computer System Product** class populated
- SMS Provider access to the SCCM site

## Files

**New-LenovoDeviceCollections.ps1**
- Queries the SCCM database for all Lenovo-branded systems
- Creates a device collection for each unique Lenovo model version
- Organizes collections in a "Lenovo" subfolder within Device Collections
- Applies dynamic WQL membership rules based on system product version
- Schedules periodic collection refreshes (default: Monday at 3:00 AM)

## Parameters

| Parameter | Required | Description |
|---|---|---|
| `-SiteServer` | Yes | Fully qualified domain name (FQDN) of the ConfigMgr site server |

## Usage

```powershell
# Basic execution
.\New-LenovoDeviceCollections.ps1 -SiteServer cm01.domain.com

# With verbose output
.\New-LenovoDeviceCollections.ps1 -SiteServer cm01.domain.com -Verbose
```

## Notes

- Run script from the site server or a system with ConfigMgr console installed
- First execution creates the "Lenovo" folder and all collections; may take several minutes
- Script safely handles re-runs; skips creating collections that already exist
- Collections use dynamic WQL queries on `SMS_G_System_COMPUTER_SYSTEM_PRODUCT` for membership
- All collections are limited to "All Systems" collection by default
- Collection names are derived from the Computer System Product Version (friendly name) property (e.g., "ThinkPad X1 Carbon Gen 13")

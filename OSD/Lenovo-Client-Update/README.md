# Lenovo Client Update

Installs current BIOS, firmware, and drivers on Lenovo commercial devices during a Configuration Manager OSD task sequence, using the [Lenovo Client Update](https://docs.lenovocdrt.com/guides/lcu/) (LCU) PowerShell module.

Each script is pasted inline into a **Run PowerShell Script** step in a child task sequence. No CM package is required.

For the full walkthrough and screenshots, see the blog post: [Updating Lenovo BIOS, Drivers, and Firmware in a Configuration Manager Task Sequence Using the Lenovo Client Update Module](https://blog.lenovocdrt.com/updating-lenovo-bios-drivers-and-firmware-in-a-configuration-manager-task-sequence-using-the-lenovo-client-update-module/).

## Requirements

- Configuration Manager OSD task sequence targeting Lenovo commercial devices
- Windows PowerShell 5.1
- Internet access from the device during the task sequence to:
    - `www.powershellgallery.com` (module download)
    - Lenovo's update servers (package download)

## Task Sequence Layout

Place the child task sequence after **Setup Windows and Configuration Manager** in the parent sequence.

```
Lenovo Client Update
├─ Download Lenovo Modules               Install-LenovoModule.ps1
├─ Populate Repository                   New-LenovoRepository.ps1
├─ 1st Pass - BIOS/Firmware
│   ├─ Install Applicable BIOS/Firmware  Install-LenovoBIOSFirmware.ps1
│   └─ Restart Computer                  condition: RebootMandatory equals True
├─ 2nd Pass - Drivers
│   └─ 2nd Pass - Drivers                Install-LenovoDrivers.ps1
└─ Final Pass Catch-All
    ├─ Set ApplicableUpdates Variable    Set-ApplicableUpdatesVariable.ps1
    └─ 3rd Pass - Drivers                Install-LenovoDrivers.ps1
                                         condition: ApplicableUpdates equals True
```

## Files

| File | Step | Purpose |
|---|---|---|
| `Install-LenovoModule.ps1` | Download Lenovo Modules | Installs `Lenovo.Client.Update` and `Lenovo.Client.Scripting` from the PowerShell Gallery |
| `New-LenovoRepository.ps1` | Populate Repository | Downloads all update packages for the machine type to `<_SMSTSMDataPath>\LenovoUpdates` |
| `Install-LenovoBIOSFirmware.ps1` | Install Applicable BIOS/Firmware | Installs BIOS and firmware packages and sets `RebootMandatory` |
| `Install-LenovoDrivers.ps1` | 2nd Pass - Drivers, 3rd Pass - Drivers | Installs driver packages |
| `Set-ApplicableUpdatesVariable.ps1` | Set ApplicableUpdates Variable | Sets `ApplicableUpdates` when any update is still pending |

## Step Settings

All steps use the same settings except for the success codes.

| Setting | Value |
|---|---|
| Script source | Enter a PowerShell script |
| PowerShell execution policy | Bypass |
| Parameters | Empty |
| Success codes | `0 3010` on Install Applicable BIOS/Firmware and 3rd Pass - Drivers; `0` on all other steps |

## Task Sequence Variables

| Variable | Set by | Used by |
|---|---|---|
| `RebootMandatory` | `Install-LenovoBIOSFirmware.ps1` | Restart Computer condition |
| `ApplicableUpdates` | `Set-ApplicableUpdatesVariable.ps1` | 3rd Pass - Drivers condition |

Both scripts set their variable to `False` before doing any work, so the conditions always evaluate against a defined value.

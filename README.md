![LibreOffice Header](https://it.libreoffice.org/themes/libreofficenew/img/discover_.jpg)

# LibreOffice Upgrader

## Overview

The `Invoke-UpgradeLibreOffice.ps1` script has been developed to automatically install the latest available version of LibreOffice without requiring administrative privileges. This solution is ideal for enterprise environments that need to keep LibreOffice up-to-date on numerous workstations with a centralized approach.

**Author:** Lorenzo Lione
**GitHub:** https://github.com/lorenz1974
**LinkedIn:** https://www.linkedin.com/in/lorenzolione/
**Version:** 1.0.0
**Creation Date:** April 23, 2025

## Key Features

- **Installation without administrative privileges**: Allows users to update LibreOffice without being system administrators
- **Automatic version detection**: Compares the local version with the latest available online
- **Silent installation**: No interaction required from the user
- **Comprehensive logging**: All operations are recorded in a CSV file for audit and troubleshooting
- **Object-oriented design**: Modular approach with class structure to ensure flexibility and reusability
- **Active Directory compatibility**: Designed to be deployed via GPO

## How the Script Works

The script follows these steps:

1. **Version check**: Verifies if LibreOffice is installed and compares the local version with the one available online
2. **MSI file download**: If an update is needed, downloads the MSI file of the latest version from the GARR mirror
3. **Silent installation**: Performs the installation in silent mode using msiexec
4. **Logging**: Records all operations in a CSV file for traceability and audit

## Usage Methods

### Direct Execution

```powershell
.\Invoke-UpgradeLibreOffice.ps1
```

Runs the script in automatic mode, checking and updating if necessary.

### Execution with Custom Logging

```powershell
.\Invoke-UpgradeLibreOffice.ps1 -logFile "C:\Logs\LibreOffice_upgrade.csv"
```

Specifies a custom path for the log file.

### Dot-Sourcing the Script

```powershell
. .\Invoke-UpgradeLibreOffice.ps1
Start-LibreOfficeUpgrader -LogFile "C:\Logs\LibreOffice_upgrade.csv"
```

Loads the script functionality without executing it, allowing manual function calls.

### Advanced Usage with the Class

```powershell
. .\Invoke-UpgradeLibreOffice.ps1
$upgrader = [LibreOfficeUpgrader]::new()
$upgrader.Upgrade($true)  # Force upgrade even if not needed
```

Direct use of the class for custom scenarios.

## Deployment via GPO in Active Directory

The script is designed to be deployed through Active Directory Group Policy Objects (GPO), allowing centralized updating of LibreOffice across the organization without requiring administrative privileges from users.

### GPO Configuration

1. **Repository Preparation**:

   - Copy the `Invoke-UpgradeLibreOffice.ps1` script to a network share accessible to all users (e.g., in the Netlogon folder)
   - Ensure users have read and execute permissions for the script

2. **GPO Creation**:

   - Open Group Policy Management Console (GPMC)
   - Create a new GPO or select an existing one
   - Right-click on the desired GPO and select "Edit"

3. **Login Script Configuration**:

   - Navigate to the `User Configuration\Policies\Windows Settings\Scripts (Logon/Logoff)` section
   - In the results panel, double-click on `Logon`
   - In the dialog box, click on `Add`
   - In the `Script Name` field, enter the full path of the script (UNC path, e.g., `\\domain.local\netlogon\Invoke-UpgradeLibreOffice.ps1`)
   - Optionally, add parameters in the `Script Parameters` field (e.g., `-logFile "\\server\logs\LibreOffice_upgrade.csv"`)
   - Click `OK` to save

4. **GPO Application**:
   - Associate the GPO to the relevant Organizational Units (OUs) to run the script at login for designated users
   - Execution can also occur via computer startup scripts (`Computer Configuration\Policies\Windows Settings\Scripts (Startup/Shutdown)`)

### Important Considerations

- **Execution as User**: Logon scripts run in the user context, not as administrator, so it's essential that the script is designed to work without elevated privileges
- **Centralized Log Files**: For better management, use a network path for log files when configuring the script via GPO
- **Asynchronous Execution**: By default, scripts run asynchronously, allowing for faster login
- **Script Visibility**: Scripts run asynchronously are not visible to the user

### Microsoft Documentation References

The script is deployed following the official Microsoft procedures for Logon scripts in Group Policy, as described in the documentation: [Using Startup, Shutdown, Logon, and Logoff Scripts in Group Policy](<https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2012-r2-and-2012/dn789196(v=ws.11)>)

## Solution Benefits

- **Reduced Administrative Burden**: No manual intervention needed for updates
- **Software Compliance**: Ensures all users use the same version of LibreOffice
- **Centralized Control**: All updates are managed by IT via GPO
- **Complete Logging**: Ability to monitor all updates through log files
- **Zero Disruption**: Updates occur without requiring user intervention

## Troubleshooting

In case of problems with script execution via GPO:

1. **Check Permissions**: Ensure users have read and execute access to the script
2. **Check Logs**: Examine the log file to identify any errors
3. **PowerShell Execution Policy**: Ensure the PowerShell execution policy allows script execution
4. **GPO Delays**: Group policy updates may take time; force an update with `gpupdate /force`
5. **Check Network**: Verify workstations have Internet access to download the MSI

## Support

For assistance, contact:
**Email**: lorenzo.lione gmail.com
**Extension**: 2025

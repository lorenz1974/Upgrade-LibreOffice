<#
.SYNOPSIS
    Script to upgrade LibreOffice to the latest available version.
.DESCRIPTION
    This script retrieves the latest available version of LibreOffice from the GARR mirror,
    downloads the MSI file and prepares it for installation.

    The script uses a modular class-based approach that allows for flexible usage:
    - Automatic upgrading when executed directly
    - Component reuse through dot-sourcing
    - Selective method execution for advanced scenarios
.PARAMETER logFile
    Path to a custom log file. If not provided, logs will be stored in the system temp directory.
.NOTES
    Version: 1.0.0
    Author: Lorenzo Lione
    GitHub: https://github.com/lorenz1974
    LinkedIn: https://www.linkedin.com/in/lorenzolione/
    Creation Date: April 23, 2025

    - Object-oriented design with class structure
    - Smart version detection and comparison
    - Silent installation capability using msiexec
    - Comprehensive error handling and logging
    - CSV-based activity logging for auditing
    - Auto-detection of dot-sourcing execution mode

    GPO DEPLOYMENT INFORMATION:
    This script has been specifically designed to install the latest version of LibreOffice without
    requiring administrative privileges from the end-user. This allows for centralized management
    and automatic updates across the organization.

    The deployment is handled through Active Directory Group Policy Objects (GPO) with the following approach:

    1. The script is stored in a network share accessible by all domain users (read-only)
    2. A GPO is configured to run the script at user logon or as a scheduled task
    3. The installation runs in the user context, leveraging the MSI's ability to install in user mode
    4. No administrative intervention is required on individual workstations
    5. Updates are managed centrally by simply updating the script on the network share

    This approach significantly reduces administrative overhead and ensures all users
    have access to the latest version of LibreOffice without IT staff involvement for each update.
    The CSV logging makes it easy to monitor deployment progress and troubleshoot any issues.

    https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2012-r2-and-2012/dn789196(v=ws.11)


.EXAMPLE
    .\Invoke-UpgradeLibreOffice.ps1
    # Runs the script in automatic mode, checking and upgrading if needed

.EXAMPLE
    .\Invoke-UpgradeLibreOffice.ps1 -logFile "C:\Logs\LibreOffice_upgrade.csv"
    # Runs the script with logging to a custom file location

.EXAMPLE
    . .\Invoke-UpgradeLibreOffice.ps1
    Start-LibreOfficeUpgrader -LogFile "C:\Logs\LibreOffice_upgrade.csv"
    # Dot-sources the script and manually runs the upgrader with custom logging

.EXAMPLE
    . .\Invoke-UpgradeLibreOffice.ps1
    $upgrader = [LibreOfficeUpgrader]::new()
    $upgrader.Upgrade($true)
    # Forces an upgrade even if not needed
#>

param(
    [String]$logFile
)

class LibreOfficeUpgrader {
    # Class instance properties for storing configuration and state
    # All variables are defined at the beginning for easy reference
    [string] $baseUrl          # Base URL of the LibreOffice download mirror
    [string] $latestVersion    # Latest version available online
    [string] $msiFileName      # Name of the MSI file to download
    [string] $downloadUrl      # Complete URL for the MSI file
    [string] $targetDir        # Directory where MSI will be saved
    [string] $outputPath       # Full path to the downloaded MSI
    [string] $localVersion     # Currently installed LibreOffice version
    [string] $localVersionPath # Path to executable used to check version
    [bool]   $needsUpgrade     # Flag indicating if upgrade is required
    [bool]   $versionChecked   # Flag to prevent redundant version checks

    # Constructor - Sets up initial property values when a new instance is created
    LibreOfficeUpgrader() {
        # Initialize the mirror URL for LibreOffice stable releases
        $this.baseUrl = "https://tdf.mirror.garr.it/libreoffice/stable/"

        # Set path to the executable used for version detection
        # We use quickstart.exe as it's typically present in all LibreOffice installations
        $this.localVersionPath = "C:\Program Files\LibreOffice\program\quickstart.exe"

        # Initialize upgrade flags - we don't assume upgrade is needed by default
        $this.needsUpgrade = $false
        $this.versionChecked = $false

        # Version checking is deferred until explicitly requested to avoid
        # unnecessary web requests when the object is instantiated
    }

    # Compare the local LibreOffice version with the latest available version online
    # This method forms the foundation for the upgrade decision logic
    [void] CheckLocalVersion() {
        # Skip redundant checking if we've already performed this operation
        # This prevents unnecessary web requests and improves performance
        if ($this.versionChecked) {
            return
        }

        try {
            # First, check if LibreOffice is installed by looking for the quickstart executable
            if (Test-Path -Path $this.localVersionPath) {
                # Get the file version information from the executable's properties
                $fileVersion = (Get-Item -Path $this.localVersionPath).VersionInfo.FileVersion

                # Extract just the version number without build information
                # This ensures we have a clean number for comparison
                $this.localVersion = ($fileVersion -split ' ')[0]
                Write-Host "Current LibreOffice version: $($this.localVersion)"

                # Query the online mirror for the latest available version
                $onlineVersion = $this.GetLatestVersion()

                # Compare local version with latest online version to determine if upgrade is needed
                # Uses PowerShell's version comparison to properly handle semantic versioning
                if ([version]$this.localVersion -lt [version]$onlineVersion) {
                    Write-Host "Local version ($($this.localVersion)) is older than online version ($onlineVersion). Upgrade needed." -ForegroundColor Yellow
                    $this.needsUpgrade = $true
                }
                else {
                    Write-Host "LibreOffice is already at the latest version ($($this.localVersion))." -ForegroundColor Green
                    $this.needsUpgrade = $false
                }
            }
            else {
                # LibreOffice is not installed, so we don't perform any action
                # This prevents automatic installation on systems where it's not present
                Write-Host "LibreOffice is not installed. No action will be taken." -ForegroundColor Yellow
                $this.needsUpgrade = $false
                $this.localVersion = "0.0.0"  # Set placeholder version
            }

            # Mark the version check as completed to prevent redundant checks
            $this.versionChecked = $true
        }
        catch {
            # Handle any errors during version checking
            Write-Warning "Error checking local version: $_"

            # If we encounter errors, don't assume upgrade is needed
            # This is a conservative approach to prevent unintended installations
            $this.needsUpgrade = $false
            $this.localVersion = "0.0.0"
            $this.versionChecked = $true
        }
    }

    # Query the LibreOffice mirror for a list of available versions and determine the latest
    # This method performs web scraping of the mirror directory structure
    [string] GetLatestVersion() {
        try {
            Write-Host "Retrieving the list of available versions..."

            # Make a web request to the mirror base URL containing version directories
            $response = Invoke-WebRequest -Uri $this.baseUrl -ErrorAction Stop

            # Parse the HTML links to extract version numbers
            # We use regex matching to identify version format directories (e.g., "7.5.2/")
            $versions = ($response.Links | Where-Object { $_.href -match '^\d+\.\d+(\.\d+)?\/$' }).href -replace '/', ''

            # Ensure we found at least one valid version
            if (-not $versions) {
                throw "No versions found on the page."
            }

            # Sort versions semantically and select the most recent one
            # This ensures proper version comparison (e.g., 7.10 is newer than 7.2)
            $this.latestVersion = $versions | Sort-Object { [version]$_ } -Descending | Select-Object -First 1
            Write-Host "Latest available version: $($this.latestVersion)"
            return $this.latestVersion
        }
        catch {
            # Propagate any errors to the caller for proper handling
            throw "Error retrieving the version list: $_"
        }
    }

    # Set up environment for downloading the MSI file
    # This prepares all necessary paths and URLs based on version information
    [void] InitializeEnvironment() {
        # Ensure we have the latest version information before proceeding
        if (-not $this.latestVersion) {
            $this.GetLatestVersion()
        }

        try {
            # Format the MSI filename based on the version number
            # We use the 64-bit MSI by default
            $this.msiFileName = "LibreOffice_$($this.latestVersion)_Win_x86-64.msi"

            # Construct the full download URL by combining the base URL with version-specific path
            $this.downloadUrl = "$($this.baseUrl)$($this.latestVersion)/win/x86_64/$($this.msiFileName)"

            # Use the system temporary directory to store the downloaded MSI file
            # This leverages the built-in temp location rather than creating custom folders
            $this.targetDir = [System.IO.Path]::GetTempPath()
            $this.outputPath = Join-Path -Path $this.targetDir -ChildPath $this.msiFileName

            Write-Host "Will save downloaded file to: $($this.outputPath)"
        }
        catch {
            throw "Error initializing environment: $_"
        }
    }

    # Download the MSI file from the LibreOffice mirror
    [void] DownloadMSI() {
        # Ensure environment is initialized before attempting download
        if (-not $this.downloadUrl -or -not $this.outputPath) {
            $this.InitializeEnvironment()
        }

        try {
            # Perform the actual download using Invoke-WebRequest
            # This downloads directly to the specified output path
            Write-Host "Downloading MSI file from: $($this.downloadUrl)"
            Invoke-WebRequest -Uri $this.downloadUrl -OutFile $this.outputPath -ErrorAction Stop
            Write-Host "Download completed: $($this.outputPath)"
        }
        catch {
            throw "Error downloading the MSI file: $_"
        }
    }

    # Install LibreOffice silently using the Windows Installer service (msiexec)
    # This allows for unattended installation without requiring user interaction
    [void] InstallSilently([string]$AdditionalOptions = "") {
        # Verify the MSI file exists before attempting installation
        if (-not (Test-Path -Path $this.outputPath)) {
            throw "MSI file not found. Please download it first using the DownloadMSI method."
        }

        try {
            # Construct the msiexec command with silent install parameters
            # /i = install mode
            # /qn = completely silent mode with no user interface
            # Additional options can be passed to customize the installation
            $msiExecCommand = "msiexec.exe /i `"$($this.outputPath)`" /qn $AdditionalOptions"
            Write-Host "Installing LibreOffice silently using command: $msiExecCommand"

            # Execute the command and wait for it to complete
            # -Wait ensures we don't continue until installation is done
            # -PassThru returns process object so we can check exit code
            # -NoNewWindow keeps the process in the same console
            $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $msiExecCommand" -Wait -PassThru -NoNewWindow

            # Check the exit code to determine if installation was successful
            # 0 = success, anything else indicates some issue
            if ($process.ExitCode -eq 0) {
                Write-Host "Silent installation completed successfully." -ForegroundColor Green
            }
            else {
                Write-Warning "Installation completed with exit code: $($process.ExitCode)"
            }
        }
        catch {
            throw "Error during silent installation: $_"
        }
    }

    # Main upgrade method that orchestrates the entire upgrade process
    # Returns the path to the MSI file if upgrade was performed, empty string otherwise
    [string] Upgrade([bool]$ForceUpgrade = $false) {
        try {
            # Check if upgrade is needed (unless force flag is set)
            if (-not $this.needsUpgrade -and -not $ForceUpgrade) {
                $this.CheckLocalVersion()
            }

            # Proceed with upgrade if needed or forced
            if ($this.needsUpgrade -or $ForceUpgrade) {
                # Ensure we have the latest version information
                if (-not $this.latestVersion) {
                    $this.GetLatestVersion()
                }

                # Execute the upgrade process step by step
                $this.InitializeEnvironment()  # Set up paths and URLs
                $this.DownloadMSI()            # Download the MSI file
                $this.InstallSilently()        # Perform silent installation

                Write-Host "Upgrade process completed successfully." -ForegroundColor Green
                return $this.outputPath  # Return the path to the MSI file
            }
            else {
                # No upgrade needed
                Write-Host "No upgrade needed. LibreOffice is already at the latest version ($($this.localVersion))." -ForegroundColor Green
                return ""  # Empty string indicates no upgrade was performed
            }
        }
        catch {
            # Propagate any errors that occurred during the upgrade process
            Write-Error $_
            throw
        }
    }

    # Convenience method to check if an upgrade is needed
    # This is useful for scripts that want to check before taking action
    [bool] IsUpgradeNeeded() {
        # Perform version check if not already done
        if (-not $this.versionChecked) {
            $this.CheckLocalVersion()
        }
        # Return the flag indicating if upgrade is needed
        return $this.needsUpgrade
    }
}

# Main script execution function
# This encapsulates the logic for the automatic upgrading process
function Start-LibreOfficeUpgrader {
    param(
        [String]$LogFile
    )

    try {
        # Enable verbose output for detailed logging
        # This helps with troubleshooting and monitoring the upgrade process
        $VerbosePreference = 'Continue'

        # Determine the log file path - use parameter if provided, otherwise use default
        if (-not $LogFile) {
            $LogFile = Join-Path -Path $env:TEMP -ChildPath "LibreOfficeUpgrade_log.csv"
        }

        # Log script start with timestamp for tracking
        Write-Host "LibreOffice Updater - Started $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
        Write-Host "Running from: $($MyInvocation.MyCommand.Path)" -ForegroundColor Cyan
        Write-Host "Log file: $LogFile" -ForegroundColor Cyan

        # Create a new instance of the upgrader class
        $upgrader = [LibreOfficeUpgrader]::new()

        # Check if an upgrade is needed
        if ($upgrader.IsUpgradeNeeded()) {
            Write-Host "Upgrading LibreOffice from version $($upgrader.localVersion) to $($upgrader.latestVersion)" -ForegroundColor Yellow

            # Start the upgrade process
            $result = $upgrader.Upgrade()

            # Log the result of the upgrade
            if ($result) {
                Write-Host "LibreOffice was successfully upgraded to version $($upgrader.latestVersion)" -ForegroundColor Green

                # Create a log entry in CSV format for tracking and auditing
                # This includes timestamp, action type, result, and version information
                $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),Upgrade,Success,From $($upgrader.localVersion) to $($upgrader.latestVersion),Computer=$env:COMPUTERNAME"
                Add-Content -Path $LogFile -Value $logEntry
            }
        }
        else {
            Write-Host "No upgrade required. LibreOffice is already at the latest version ($($upgrader.localVersion))." -ForegroundColor Green

            # Log check operations as well for audit trail
            # This helps track when the system was verified but no upgrade was needed
            $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),Check,NoActionNeeded,Version=$($upgrader.localVersion),Computer=$env:COMPUTERNAME"
            Add-Content -Path $LogFile -Value $logEntry
        }
    }
    catch {
        # Handle and log any errors that occurred during the upgrade process
        Write-Error "LibreOffice upgrade process failed: $_"

        # Ensure LogFile is set to prevent errors
        if (-not $LogFile) {
            $LogFile = Join-Path -Path $env:TEMP -ChildPath "LibreOfficeUpgrade_log.csv"
        }

        # Log the error with details for troubleshooting
        $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),Upgrade,Failed,$_,Computer=$env:COMPUTERNAME"
        Add-Content -Path $LogFile -Value $logEntry

        exit 1  # Exit with non-zero code to indicate failure
    }
    finally {
        # Always log the end time, even if exceptions occurred
        # This ensures the log contains complete timing information
        Write-Host "LibreOffice Updater - Completed $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
    }
}

# Script execution control logic
# This determines how the script behaves when run in different contexts
# ----------------------------------------------------------------
# Detect if the script is being dot-sourced or run directly
# $MyInvocation.InvocationName will be '.' if the script is being dot-sourced
if ($MyInvocation.InvocationName -ne '.') {
    # When script is executed directly, run the upgrader automatically
    # This is the default behavior when running the script as a standalone tool
    Start-LibreOfficeUpgrader -LogFile $logFile
}
else {
    # When script is dot-sourced, only load the functions and class without running
    # This allows the functionality to be imported into other scripts
    Write-Host "LibreOffice Updater has been loaded. Use Start-LibreOfficeUpgrader to run the update process." -ForegroundColor Cyan
}

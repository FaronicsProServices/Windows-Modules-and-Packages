# PowerShell Script to Download, Install 7-Zip (if needed), Extract, and Install Read&Write 12 US Version (MSI-based)

# --- GLOBAL SCRIPT SETTINGS ---
# Define a dedicated log directory for all script outputs
$LogPath = "C:\ProgramData\AppInstallLogs\ReadWrite12"
$LogFile = Join-Path $LogPath "ReadWrite12_FullInstall_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Function to write messages to console and log file
Function Write-Log {
    Param(
        [string]$Message,
        [string]$Level = "INFO" # INFO, WARN, ERROR
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "$Timestamp [$Level] $Message"
    Write-Host $LogEntry # Output to console
    Add-Content -Path $LogFile -Value $LogEntry # Write to log file
}

# --- SCRIPT START ---
Write-Log "--- Starting Comprehensive Read&Write 12 Installation Script ---" -Level "INFO"
Write-Log "Current Time: $(Get-Date)" -Level "INFO"
Write-Log "Log file location: $LogFile" -Level "INFO"
Write-Log "Note: This script will download and install 7-Zip first (if not present), then download, extract (1 GB file), and install Read&Write 12." -Level "INFO"
Write-Log "This entire process may take significant time (up to 3-4 hours depending on network speed and system performance). Ensure deployment tool timeouts are set accordingly." -Level "WARN"

# Create log directory if it doesn't exist
if (-not (Test-Path $LogPath)) {
    try {
        New-Item -ItemType Directory -Path $LogPath -ErrorAction Stop | Out-Null
        Write-Log "Created log directory: $LogPath" -Level "INFO"
    }
    catch {
        Write-Log "Failed to create log directory $LogPath. Error: $_" -Level "ERROR"
        # Continue without logging to file if directory creation fails, but output to console
    }
}

# --- 7-Zip Installation Configuration ---
$7ZipDownloadURL = "https://www.7-zip.org/a/7z2301-x64.msi" # Standard 64-bit 7-Zip MSI (version 23.01)
$7ZipDownloadPath = "$env:TEMP\7zinstaller.msi"
$7ZipInstallArgs = "/qn /norestart" # /qn for quiet, /norestart to prevent immediate reboot

# --- Main Application Configuration ---
$ReadWriteDownloadURL = "https://fastdownloads2.texthelp.com/readwrite12/installers/us/setup.zip"
$ReadWriteDownloadPath = "$env:TEMP\setup.zip"
$ReadWriteExtractPath = "$env:TEMP\ReadWrite12_Installer"
$ReadWriteMsiFileName = "setup.msi" # Confirmed MSI file name inside the zip

# MSI Silent Install Arguments for Read&Write
$ReadWriteMsiInstallArgs = "/i `"$ReadWriteMsiFileName`" /qn /norestart" # /i for install, /qn for quiet, /norestart to prevent immediate reboot

# --- SCRIPT STEPS ---

# STEP 1: Check and Install 7-Zip (if not present)
Write-Log "Checking for 7-Zip installation..." -Level "INFO"
$7ZipExePath_x64 = "$env:ProgramFiles\7-Zip\7z.exe"
$7ZipExePath_x86 = "$env:ProgramFiles(x86)\7-Zip\7z.exe"
$7ZipInstalled = $false

if (Test-Path $7ZipExePath_x64) {
    $7ZipInstalled = $true
    Write-Log "7-Zip (64-bit) found at $7ZipExePath_x64." -Level "INFO"
    $Current7ZipPath = $7ZipExePath_x64
}
elseif (Test-Path $7ZipExePath_x86) {
    $7ZipInstalled = $true
    Write-Log "7-Zip (32-bit) found at $7ZipExePath_x86." -Level "INFO"
    $Current7ZipPath = $7ZipExePath_x86
}

if (-not $7ZipInstalled) {
    Write-Log "7-Zip not found. Proceeding to download and install 7-Zip..." -Level "INFO"

    # Download 7-Zip MSI
    Write-Log "Downloading 7-Zip from $7ZipDownloadURL to $7ZipDownloadPath..." -Level "INFO"
    try {
        Invoke-WebRequest -Uri $7ZipDownloadURL -OutFile $7ZipDownloadPath -ErrorAction Stop
        Write-Log "7-Zip download complete." -Level "INFO"
    }
    catch {
        Write-Log "Failed to download 7-Zip. Error: $_" -Level "ERROR"
        exit 1
    }

    # Install 7-Zip MSI
    Write-Log "Installing 7-Zip..." -Level "INFO"
    try {
        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$7ZipDownloadPath`" $7ZipInstallArgs" -Wait -NoNewWindow -PassThru -ErrorAction Stop
        $exitCode = $process.ExitCode
        Write-Log "7-Zip MSI installer finished with exit code: $exitCode. Common success codes: 0, 3010 (requires reboot)." -Level "INFO"

        if ($exitCode -ne 0 -and $exitCode -ne 3010) {
            Write-Log "7-Zip MSI installation returned a non-success exit code. Exiting." -Level "ERROR"
            exit 1
        }
        # Update current 7-Zip path after installation
        if (Test-Path $7ZipExePath_x64) { $Current7ZipPath = $7ZipExePath_x64 }
        elseif (Test-Path $7ZipExePath_x86) { $Current7ZipPath = $7ZipExePath_x86 }
        else {
             Write-Log "7-Zip was installed, but executable not found at expected paths. This might cause issues. Continuing anyway." -Level "WARN"
             # Attempt to find it, might be needed for later step
             $Current7ZipPath = Get-Command 7z.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source
             if (-not $Current7ZipPath) {
                 Write-Log "Failed to locate 7z.exe after installation. Extraction will likely fail." -Level "ERROR"
                 exit 1
             }
        }
    }
    catch {
        Write-Log "Failed to install 7-Zip. Error: $_" -Level "ERROR"
        exit 1
    }
    # Clean up 7-Zip installer
    try {
        Remove-Item -Path $7ZipDownloadPath -Force -ErrorAction SilentlyContinue
        Write-Log "Cleaned up 7-Zip installer." -Level "INFO"
    }
    catch {
        Write-Log "Failed to clean up 7-Zip installer. Error: $_" -Level "WARN"
    }
}
else {
    Write-Log "7-Zip is already installed. Skipping 7-Zip installation." -Level "INFO"
}


# STEP 2: Ensure the application download and extract directories exist
Write-Log "Ensuring application download and extract directories exist..." -Level "INFO"
if (-not (Test-Path $ReadWriteExtractPath)) {
    try {
        New-Item -ItemType Directory -Path $ReadWriteExtractPath -ErrorAction Stop | Out-Null
        Write-Log "Created extraction directory: $ReadWriteExtractPath" -Level "INFO"
    }
    catch {
        Write-Log "Failed to create extraction directory $ReadWriteExtractPath. Error: $_" -Level "ERROR"
        exit 1
    }
}

# STEP 3: Download the Read&Write ZIP file
Write-Log "Downloading Read&Write ZIP from $ReadWriteDownloadURL to $ReadWriteDownloadPath..." -Level "INFO"
Write-Log "This is a large file (approx. 1 GB) and may take significant time depending on network speed." -Level "INFO"
try {
    Invoke-WebRequest -Uri $ReadWriteDownloadURL -OutFile $ReadWriteDownloadPath -ErrorAction Stop
    Write-Log "Read&Write ZIP download complete." -Level "INFO"
}
catch {
    Write-Log "Failed to download the Read&Write ZIP file. Error: $_" -Level "ERROR"
    exit 1
}

# STEP 4: Extract the Read&Write ZIP file using 7-Zip
Write-Log "Extracting $ReadWriteDownloadPath to $ReadWriteExtractPath using 7-Zip..." -Level "INFO"
Write-Log "This might take a while due to the large file size and numerous files." -Level "INFO"
try {
    # 'x' for extract with full paths, '-o' for output directory, '-y' for assume yes (overwrite)
    $7ZipArgs = "x `"$ReadWriteDownloadPath`" -o`"$ReadWriteExtractPath`" -y"
    $process = Start-Process -FilePath $Current7ZipPath -ArgumentList $7ZipArgs -Wait -NoNewWindow -PassThru -ErrorAction Stop
    $exitCode = $process.ExitCode

    if ($exitCode -eq 0) {
        Write-Log "7-Zip extraction complete." -Level "INFO"
    }
    else {
        Write-Log "7-Zip extraction failed with exit code: $exitCode. Please review 7-Zip documentation for exit codes." -Level "ERROR"
        exit 1
    }
}
catch {
    Write-Log "Failed to extract the Read&Write ZIP file using 7-Zip. Error: $_" -Level "ERROR"
    exit 1
}

# STEP 5: Locate and run the Read&Write MSI installer
$ReadWriteMsiPath = Join-Path $ReadWriteExtractPath $ReadWriteMsiFileName

if (Test-Path $ReadWriteMsiPath) {
    Write-Log "Found Read&Write MSI at: $ReadWriteMsiPath" -Level "INFO"
    Write-Log "Running MSI installer with arguments: msiexec.exe $ReadWriteMsiInstallArgs (relative to MSI path)" -Level "INFO"
    Write-Log "Read&Write installation may take a while..." -Level "INFO"

    # Change directory to where the MSI is located for proper execution if the MSI has dependencies
    Push-Location $ReadWriteExtractPath

    try {
        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $ReadWriteMsiInstallArgs -Wait -NoNewWindow -PassThru -ErrorAction Stop
        $exitCode = $process.ExitCode
        Write-Log "Read&Write MSI installer finished with exit code: $exitCode. Common success codes: 0, 3010 (requires reboot)." -Level "INFO"

        if ($exitCode -ne 0 -and $exitCode -ne 3010) {
            Write-Log "Read&Write MSI installation returned a non-success exit code. Exiting." -Level "ERROR"
            Pop-Location # Restore previous directory
            exit 1 # Indicate failure to the deployment tool
        }
    }
    catch {
        Write-Log "Failed to run the Read&Write MSI installer. Error: $_" -Level "ERROR"
        Pop-Location # Restore previous directory
        exit 1
    }

    Pop-Location # Restore previous directory after successful execution
}
else {
    Write-Log "Read&Write MSI file '$ReadWriteMsiFileName' not found in the extracted directory '$ReadWriteExtractPath'. This indicates a problem with the extraction." -Level "ERROR"
    exit 1
}

# STEP 6: Clean up temporary files
Write-Log "Cleaning up temporary files..." -Level "INFO"
try {
    Remove-Item -Path $ReadWriteDownloadPath -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $ReadWriteExtractPath -Recurse -Force -ErrorAction SilentlyContinue
    Write-Log "Temporary files cleaned up." -Level "INFO"
}
catch {
    Write-Log "Failed to clean up temporary files. Error: $_" -Level "WARN"
}

Write-Log "--- Read&Write 12 Installation Script Completed Successfully ---" -Level "INFO"

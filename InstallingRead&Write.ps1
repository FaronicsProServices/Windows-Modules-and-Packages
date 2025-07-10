# PowerShell Script to Download, Extract, and Install Read&Write 12 US Version

Write-Host "Installing Read&Write... This may take a few minutes depending on internet speed."

# --- Variables ---
$zipUrl = "https://fastdownloads2.texthelp.com/readwrite12/installers/us/setup.zip"
$downloadPath = "$env:TEMP\setup.zip"
$extractPath = "$env:TEMP\ReadWriteInstaller"
$msiFileName = "Setup.msi" # Confirmed from your provided image
$msiInstallArgs = "/i `"$msiFileName`" /qn /norestart" # /qn for quiet, /norestart to prevent immediate reboot

# --- Script Steps ---

# 1. Create extraction folder
Write-Host "Ensuring extraction directory exists: $extractPath"
if (!(Test-Path -Path $extractPath)) {
    try {
        New-Item -ItemType Directory -Path $extractPath -ErrorAction Stop | Out-Null
        Write-Host "Created extraction directory."
    }
    catch {
        Write-Host "ERROR: Failed to create extraction directory '$extractPath'. Error: $_"
        exit 1
    }
}
else {
    Write-Host "Extraction directory already exists."
    # Optional: Clear existing contents if you want a clean extract every time
    # try {
    #     Get-ChildItem -Path $extractPath -Recurse -Force | Remove-Item -ErrorAction SilentlyContinue
    #     Write-Host "Cleared existing contents of extraction directory."
    # } catch {
    #     Write-Host "WARNING: Could not clear existing contents. Error: $_"
    # }
}

# 2. Download using BITS (more reliable in background/system contexts)
Write-Host "Downloading '$zipUrl' to '$downloadPath' using BITS..."
Write-Host "This is a large file (approx. 1 GB) and may take significant time."
try {
    Start-BitsTransfer -Source $zipUrl -Destination $downloadPath -ErrorAction Stop
    Write-Host "Download complete. File saved to: $downloadPath"
}
catch {
    Write-Host "ERROR: Failed to download the file from '$zipUrl'. Error: $_"
    exit 1
}

# 3. Verify Downloaded File (optional but good practice)
if (Test-Path $downloadPath) {
    $fileSize = (Get-Item $downloadPath).Length
    Write-Host "Downloaded file size: $([math]::Round($fileSize / 1GB, 2)) GB"
    if ($fileSize -lt 900MB) { # Basic check for a suspiciously small file
        Write-Host "WARNING: Downloaded file size is unexpectedly small. It might be incomplete or corrupted."
    }
} else {
    Write-Host "ERROR: Downloaded ZIP file does not exist at $downloadPath after download attempt. Exiting."
    exit 1
}

# 4. Extract the zip file
Write-Host "Extracting '$downloadPath' to '$extractPath'..."
Write-Host "This step might also take a while due to the file size."
try {
    Expand-Archive -Path $downloadPath -DestinationPath $extractPath -Force -ErrorAction Stop
    Write-Host "Extraction complete."
}
catch {
    Write-Host "ERROR: Failed to extract the zip file. Error: $_"
    Write-Host "Ensure enough disk space and file integrity. If issues persist, consider using a more robust extractor like 7-Zip."
    exit 1
}

# 5. Install the Application
$msiPath = Join-Path $extractPath $msiFileName

Write-Host "Attempting to install '$msiFileName' from '$msiPath'..."
if (Test-Path $msiPath) {
    Write-Host "Found MSI. Starting silent installation..."
    Write-Host "Command: msiexec.exe /i `"$msiPath`" /qn /norestart"
    Write-Host "Installation may take several minutes..."

    # Change directory to where the MSI is located for proper execution (MSI often needs its associated CAB files)
    Push-Location $extractPath

    try {
        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $msiInstallArgs -Wait -NoNewWindow -PassThru -ErrorAction Stop
        $exitCode = $process.ExitCode
        Write-Host "MSI installer finished with exit code: $exitCode. Common success codes: 0, 3010 (requires reboot)."

        if ($exitCode -ne 0 -and $exitCode -ne 3010) {
            Write-Host "ERROR: MSI installation returned a non-success exit code. Please investigate."
            Pop-Location # Restore previous directory
            exit 1 # Indicate failure to the deployment tool
        }
        Write-Host "MSI installation appears to have completed successfully."
    }
    catch {
        Write-Host "ERROR: Failed to run the MSI installer. Error: $_"
        Pop-Location # Restore previous directory
        exit 1
    }

    Pop-Location # Restore previous directory after successful execution
}
else {
    Write-Host "ERROR: MSI file '$msiFileName' not found at '$msiPath'. This indicates a problem with the extraction or file structure."
    # Optionally, list extracted files to debug:
    # Get-ChildItem -Path $extractPath -Recurse | ForEach-Object { Write-Host " - $($_.FullName)" }
    exit 1
}

# 6. Clean up temporary files (optional but recommended for deployment)
Write-Host "Cleaning up temporary files..."
try {
    Remove-Item -Path $downloadPath -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $extractPath -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "Temporary files cleaned up."
}
catch {
    Write-Host "WARNING: Failed to clean up temporary files. Error: $_"
}

Write-Host "--- Read&Write 12 Installation Script Completed ---"

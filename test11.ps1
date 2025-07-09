# PowerShell Script to Download, Extract, and Install Read&Write 12 US Version (MSI-based)

# --- Configuration ---
$DownloadURL = "https://fastdownloads2.texthelp.com/readwrite12/installers/us/setup.zip"
$DownloadPath = "$env:TEMP\setup.zip"
$ExtractPath = "$env:TEMP\ReadWrite12_Installer"
$MsiFileName = "setup.msi" # <-- UPDATED: Confirmed MSI file name inside the zip

# MSI Silent Install Arguments:
# /i = install
# /qn = quiet (no UI)
# /norestart = do not restart the computer after installation (important for deployments)
$MsiInstallArgs = "/i `"$MsiFileName`" /qn /norestart"

# --- Script Logic ---

Write-Host "Starting Read&Write 12 installation process..."
Write-Host "Note: The installer is large (approx. 1 GB), so this process may take some time."

# 1. Ensure the download and extract directories exist
if (-not (Test-Path $ExtractPath)) {
    try {
        New-Item -ItemType Directory -Path $ExtractPath -ErrorAction Stop | Out-Null
        Write-Host "Created extraction directory: $ExtractPath"
    }
    catch {
        Write-Error "Failed to create extraction directory $ExtractPath. Error: $_"
        exit 1
    }
}

# 2. Download the ZIP file
Write-Host "Downloading $DownloadURL to $DownloadPath..."
try {
    # Set a timeout for Invoke-WebRequest if needed for very slow connections, though default is usually fine for large files.
    # $WebClient = New-Object System.Net.WebClient
    # $WebClient.DownloadFile($DownloadURL, $DownloadPath)
    Invoke-WebRequest -Uri $DownloadURL -OutFile $DownloadPath -ErrorAction Stop
    Write-Host "Download complete."
}
catch {
    Write-Error "Failed to download the file from $DownloadURL. Error: $_"
    exit 1
}

# 3. Extract the ZIP file
Write-Host "Extracting $DownloadPath to $ExtractPath..."
Write-Host "This might take a while due to the large file size..."
try {
    Expand-Archive -LiteralPath $DownloadPath -DestinationPath $ExtractPath -Force -ErrorAction Stop
    Write-Host "Extraction complete."
}
catch {
    Write-Error "Failed to extract the zip file. Error: $_"
    exit 1
}

# 4. Locate and run the MSI installer
$MsiPath = Join-Path $ExtractPath $MsiFileName

if (Test-Path $MsiPath) {
    Write-Host "Found MSI at: $MsiPath"
    Write-Host "Running MSI installer with arguments: msiexec.exe $MsiInstallArgs (relative to MSI path)"
    Write-Host "Installation may take a while..."

    # Change directory to where the MSI is located for proper execution if the MSI has dependencies
    Push-Location $ExtractPath

    try {
        Start-Process -FilePath "msiexec.exe" -ArgumentList $MsiInstallArgs -Wait -NoNewWindow -ErrorAction Stop
        Write-Host "MSI installer finished. Please check application logs or installed programs to verify success."
    }
    catch {
        Write-Error "Failed to run the MSI installer. Error: $_"
        Pop-Location # Restore previous directory in case of error
        exit 1
    }

    Pop-Location # Restore previous directory after successful execution
}
else {
    Write-Error "MSI file '$MsiFileName' not found in the extracted directory '$ExtractPath'. Please verify the \$MsiFileName variable by manually inspecting the zip content."
    exit 1
}

# 5. Clean up (optional but recommended for deployment)
Write-Host "Cleaning up temporary files..."
try {
    Remove-Item -Path $DownloadPath -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $ExtractPath -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "Temporary files cleaned up."
}
catch {
    Write-Warning "Failed to clean up temporary files. Error: $_"
}

Write-Host "Read&Write 12 MSI installation script completed."

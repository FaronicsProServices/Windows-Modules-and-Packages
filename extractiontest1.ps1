# PowerShell Script to Download and Extract a ZIP file from a URL

# --- Configuration ---
# URL of the ZIP file to download
$DownloadURL = "https://fastdownloads2.texthelp.com/readwrite12/installers/us/setup.zip"
# Path where the ZIP file will be downloaded (in the user's temporary directory)
$zipFilePath = "$env:TEMP\setup.zip"
# Directory where the ZIP file contents will be extracted
$extractionDirectory = "$env:TEMP\ReadWrite12_Installer"

# --- Script Logic ---

Write-Host "--- Starting Download and Extraction Process ---"
Write-Host "Current Time: $(Get-Date)"

# 1. Create the extraction directory if it doesn't exist
Write-Host "Ensuring extraction directory exists: $extractionDirectory"
if (!(Test-Path -Path $extractionDirectory)) {
    try {
        New-Item -ItemType Directory -Force -Path $extractionDirectory -ErrorAction Stop | Out-Null
        Write-Host "Created extraction directory."
    }
    catch {
        Write-Host "ERROR: Failed to create extraction directory '$extractionDirectory'. Error: $_"
        exit 1
    }
}
else {
    Write-Host "Extraction directory already exists. Contents might be overwritten if not empty."
}

# 2. Download the ZIP file
Write-Host "Downloading '$DownloadURL' to '$zipFilePath'..."
Write-Host "This is a large file (approx. 1 GB) and may take significant time depending on your network speed."
try {
    Invoke-WebRequest -Uri $DownloadURL -OutFile $zipFilePath -ErrorAction Stop
    Write-Host "Download complete. File saved to: $zipFilePath"
}
catch {
    Write-Host "ERROR: Failed to download the file from '$DownloadURL'. Error: $_"
    exit 1
}

# 3. Verify Downloaded File (optional but good practice)
if (Test-Path $zipFilePath) {
    $fileSize = (Get-Item $zipFilePath).Length
    Write-Host "Downloaded file size: $([math]::Round($fileSize / 1GB, 2)) GB"
    # Basic check for a suspiciously small file
    if ($fileSize -lt 900MB) {
        Write-Host "WARNING: Downloaded file size is unexpectedly small. It might be incomplete or corrupted."
    }
} else {
    Write-Host "ERROR: Downloaded ZIP file does not exist after download attempt. Exiting."
    exit 1
}

# 4. Extract the zip file
Write-Host "Extracting '$zipFilePath' to '$extractionDirectory'..."
Write-Host "This step might also take a while due to the file size."
try {
    # Ensure the extraction directory is empty before extracting to avoid issues with -Force and existing files
    Get-ChildItem -Path $extractionDirectory -Recurse -Force | Remove-Item -ErrorAction SilentlyContinue
    Write-Host "Cleared existing contents of extraction directory (if any)."
    Expand-Archive -Path $zipFilePath -DestinationPath $extractionDirectory -Force -ErrorAction Stop
    Write-Host "Extraction complete."
}
catch {
    Write-Host "ERROR: Failed to extract the zip file. Error: $_"
    Write-Host "Possible causes include: insufficient disk space, corrupted zip file, or issues with Expand-Archive cmdlet for very large files."
    exit 1
}

# --- Installation Part (commented out as requested) ---
# You can uncomment and modify this section later when you're ready to work on installation.

# # Find the setup executable (assuming it's a common name like setup.exe or install.msi)
# $setupExecutable = Get-ChildItem -Path $extractionDirectory -Include "setup.exe", "install.msi", "setup.msi" -Recurse -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName

# # Execute the setup file if found
# if ($setupExecutable) {
#     Write-Host "Executing setup file: $($setupExecutable)"
#     # For MSI files, you typically use msiexec for silent install. Example:
#     # Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$setupExecutable`" /qn /norestart" -Wait -NoNewWindow
#     # For .exe, it depends on its silent switches. Example:
#     # Start-Process -FilePath $setupExecutable -ArgumentList "/S /quiet" -Wait -NoNewWindow
# } else {
#     Write-Host "Setup file not found. Please check the extraction directory."
# }

Write-Host "--- Download and Extraction Process Completed ---"

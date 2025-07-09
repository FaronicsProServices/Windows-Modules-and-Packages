# Display starting message
Write-Output "Installing Read&Write... This may take several minutes. Please wait..."

# Variables
$zipUrl = "https://fastdownloads2.texthelp.com/readwrite12/installers/us/setup.zip"
$sevenZipUrl = "https://www.7-zip.org/a/7z2301-extra.7z" # Contains 7z.exe and 7z.dll
$downloadPath = "$env:TEMP\setup.zip"
$extractPath = "$env:TEMP\ReadWriteExtract"
$sevenZipArchive = "$env:TEMP\7z.7z"
$sevenZipFolder = "$env:TEMP\7zip"

# Create necessary folders
New-Item -ItemType Directory -Force -Path $extractPath | Out-Null
New-Item -ItemType Directory -Force -Path $sevenZipFolder | Out-Null

# Download Read&Write ZIP
Start-BitsTransfer -Source $zipUrl -Destination $downloadPath

# Download 7-Zip portable (Extra version)
Start-BitsTransfer -Source $sevenZipUrl -Destination $sevenZipArchive

# Extract 7-Zip (if 7z.exe not already available)
# Using Windows built-in tools for extracting 7z archive
Expand-Archive -Path $sevenZipArchive -DestinationPath $sevenZipFolder -Force

# Get path to 7z.exe
$sevenZipExe = Get-ChildItem -Path $sevenZipFolder -Recurse -Filter "7z.exe" | Select-Object -First 1

# Verify 7z.exe exists
if (-not $sevenZipExe) {
    Write-Error "7z.exe not found. Cannot continue."
    exit 1
}

# Extract the Read&Write ZIP using 7z
& $sevenZipExe.FullName x $downloadPath "-o$extractPath" -y

# Find installer in extracted files
$installer = Get-ChildItem -Path $extractPath -Recurse | Where-Object { $_.Name -match "setup\.exe" -or $_.Name -like "*.msi" } | Select-Object -First 1

# Run the installer silently
if ($installer) {
    Write-Output "Installing silently..."
    & $installer.FullName /S /quiet /norestart
    Write-Output "✅ Installation completed successfully."
} else {
    Write-Error "❌ Installer not found in the extracted files."
}

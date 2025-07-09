Write-Output "Installing Read&Write... This may take several minutes. Please wait..."

# URLs
$zipUrl = "https://fastdownloads2.texthelp.com/readwrite12/installers/us/setup.zip"
$sevenZipUrl = "https://github.com/jhen0409/7zip-mini/raw/master/7za920.zip"

# Paths
$downloadPath = "$env:TEMP\setup.zip"
$extractPath = "$env:TEMP\ReadWriteExtract"
$sevenZipArchive = "$env:TEMP\7za.zip"
$sevenZipFolder = "$env:TEMP\7zip"

# Create folders
New-Item -ItemType Directory -Force -Path $extractPath | Out-Null
New-Item -ItemType Directory -Force -Path $sevenZipFolder | Out-Null

# Download Read&Write ZIP
Start-BitsTransfer -Source $zipUrl -Destination $downloadPath

# Download 7-Zip portable
Start-BitsTransfer -Source $sevenZipUrl -Destination $sevenZipArchive

# Extract 7-Zip archive (this is a .zip, so Expand-Archive works now)
Expand-Archive -Path $sevenZipArchive -DestinationPath $sevenZipFolder -Force

# Locate 7za.exe (this is the command-line version)
$sevenZipExe = Get-ChildItem -Path $sevenZipFolder -Recurse -Filter "7za.exe" | Select-Object -First 1

# Fail if not found
if (-not $sevenZipExe) {
    Write-Error "❌ 7za.exe not found. Cannot continue."
    exit 1
}

# Extract setup.zip using 7-Zip
& $sevenZipExe.FullName x $downloadPath "-o$extractPath" -y

# Find installer
$installer = Get-ChildItem -Path $extractPath -Recurse | Where-Object { $_.Name -match "setup\.exe" -or $_.Name -like "*.msi" } | Select-Object -First 1

# Run the installer silently
if ($installer) {
    Write-Output "Installing Read&Write silently..."
    & $installer.FullName /S /quiet /norestart
    Write-Output "✅ Installation completed successfully."
} else {
    Write-Error "❌ Installer not found in extracted files."
}

Write-Output "Installing Read&Write... This may take a few minutes depending on internet speed."

# Variables
$zipUrl = "https://fastdownloads2.texthelp.com/readwrite12/installers/us/setup.zip"
$downloadPath = "$env:TEMP\setup.zip"
$extractPath = "$env:TEMP\ReadWriteInstaller"

# Create folder
if (!(Test-Path -Path $extractPath)) {
    New-Item -ItemType Directory -Path $extractPath | Out-Null
}

# Download using BITS (more reliable in background/system contexts)
Start-BitsTransfer -Source $zipUrl -Destination $downloadPath

# Extract
Expand-Archive -Path $downloadPath -DestinationPath $extractPath -Force

# Find installer
$installer = Get-ChildItem -Path $extractPath -Recurse | Where-Object { $_.Name -match "setup\.exe" -or $_.Name -like "*.msi" } | Select-Object -First 1

# Run silent install directly (no Start-Process)
if ($installer) {
    & $installer.FullName /S /quiet /norestart
    Write-Output "Installation completed successfully."
} else {
    Write-Error "Installer not found in the extracted files."
}

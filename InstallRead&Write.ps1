# User message
Write-Host "Installing Read&Write... This may take a few minutes depending on your internet speed and system performance. Please wait..." -ForegroundColor Yellow

# Variables
$zipUrl = "https://fastdownloads2.texthelp.com/readwrite12/installers/us/setup.zip"
$downloadPath = "$env:TEMP\setup.zip"
$extractPath = "$env:TEMP\ReadWriteInstaller"

# Create extract directory if not exists
if (!(Test-Path -Path $extractPath)) {
    New-Item -ItemType Directory -Path $extractPath | Out-Null
}

# Download the ZIP file
Invoke-WebRequest -Uri $zipUrl -OutFile $downloadPath

# Extract the ZIP
Expand-Archive -Path $downloadPath -DestinationPath $extractPath -Force

# Find the installer (e.g., setup.exe or .msi)
$installer = Get-ChildItem -Path $extractPath -Recurse | Where-Object { $_.Name -match "setup\.exe" -or $_.Name -like "*.msi" } | Select-Object -First 1

# Run the installer silently (modify arguments if needed)
if ($installer) {
    Start-Process -FilePath $installer.FullName -ArgumentList "/quiet" -Wait
    Write-Host "Installation completed successfully." -ForegroundColor Green
} else {
    Write-Error "Installer not found in the extracted files."
}

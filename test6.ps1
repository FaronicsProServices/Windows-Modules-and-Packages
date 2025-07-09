Write-Output "Installing Read&Write... Please wait; this may take several minutes."

# URLs
$zipUrl = "https://fastdownloads2.texthelp.com/readwrite12/installers/us/setup.zip"
$sevenZipUrl = "https://downloads.sourceforge.net/project/portableapps/7-Zip%20Portable/25.00/7-ZipPortable_25.00Portable.zip"

# Paths
$downloadPath = "$env:TEMP\setup.zip"
$extractPath = "$env:TEMP\ReadWriteExtract"
$sevenZipArchive = "$env:TEMP\7zipportable.zip"
$sevenZipFolder = "$env:TEMP\7zip"

# Create folders
New-Item -ItemType Directory -Force -Path $extractPath,$sevenZipFolder | Out-Null

# Download files
Invoke-WebRequest -Uri $zipUrl -OutFile $downloadPath
Invoke-WebRequest -Uri $sevenZipUrl -OutFile $sevenZipArchive

# Extract portable 7-Zip
Expand-Archive -Path $sevenZipArchive -DestinationPath $sevenZipFolder -Force

# Find 7za.exe (command-line only)
$sevenZipExe = Get-ChildItem -Path $sevenZipFolder -Recurse -Filter "7za.exe" | Select-Object -First 1
if (-not $sevenZipExe) {
    Write-Error "7za.exe not found – cannot continue."
    exit 1
}

# Extract main installer
& $sevenZipExe.FullName x $downloadPath "-o$extractPath" -y

# Locate installer executable
$installer = Get-ChildItem -Path $extractPath -Recurse |
             Where-Object { $_.Extension -in '.exe','.msi' } |
             Select-Object -First 1

if ($installer) {
    Write-Output "Running installer silently..."
    & $installer.FullName /S /quiet /norestart
    Write-Output "Installation completed successfully."
} else {
    Write-Error "No installer found in the extracted files."
}

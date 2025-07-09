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

# Download Read&Write setup ZIP
Invoke-WebRequest -Uri $zipUrl -OutFile $downloadPath

# Download portable 7-Zip
Invoke-WebRequest -Uri $sevenZipUrl -OutFile $sevenZipArchive

# Extract 7-Zip portable
Expand-Archive -Path $sevenZipArchive -DestinationPath $sevenZipFolder -Force

# Locate 7za.exe (command-line 7-Zip)
$sevenZipExe = Get-ChildItem -Path $sevenZipFolder -Recurse -Filter "7za.exe" | Select-Object -First 1

# Fail if 7za.exe not found
if (-not $sevenZipExe) {
    Write-Error "❌ 7za.exe not found – cannot continue."
    exit 1
}

# Extract the Read&Write ZIP using 7-Zip
& $sevenZipExe.FullName x $downloadPath "-o$extractPath" -y

# Locate setup.exe or .msi inside the extracted files
$installer = Get-ChildItem -Path $extractPath -Recurse | Where-Object {
    $_.Extension -in '.exe', '.msi'
} | Select-Object -First 1

# Run the installer silently
if ($installer) {
    Write-Output "Running installer silently..."
    & $installer.FullName /S /quiet /norestart
    Write-Output "✅ Installation completed successfully."
} else {
    Write-Error "❌ No installer found in the extracted files."
}

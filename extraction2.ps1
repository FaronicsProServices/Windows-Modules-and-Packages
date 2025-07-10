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

# Define variables
$downloadUrl = "https://fastdownloads2.texthelp.com/readwrite12/installers/us/setup.zip"
$tempPath = "$env:TEMP\ReadWriteSetup"
$zipPath = "$tempPath\setup.zip"
$extractPath = "$tempPath\Extracted"

# Create temp directory if it doesn't exist
if (!(Test-Path $tempPath)) {
    New-Item -Path $tempPath -ItemType Directory | Out-Null
}

# Download the ZIP file
Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath

# Extract ZIP contents
Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force

# Install the MSI silently
$msiPath = Join-Path $extractPath "setup.msi"
Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPath`" /qn" -Wait

# Cleanup temp files (optional)
Remove-Item -Path $tempPath -Recurse -Force

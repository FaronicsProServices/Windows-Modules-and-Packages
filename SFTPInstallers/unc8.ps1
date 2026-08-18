# ============================================================
# Freedom Scientific Fusion - Silent Installation
# ============================================================

$ErrorActionPreference = "Stop"

# ------------------------------------------------------------
# Configuration
# ------------------------------------------------------------

$SourceZip = "\\pdc\Share2\Fusion.zip"

$TempPath = Join-Path $env:TEMP "FusionInstall"

$AuthorizationPath = "C:\Program Files\Freedom Scientific\Authorization"

# ------------------------------------------------------------
# Check ZIP exists
# ------------------------------------------------------------

Write-Host "Checking Fusion ZIP..."

if (-not (Test-Path $SourceZip)) {
    Write-Error "Fusion ZIP not found: $SourceZip"
    exit 1
}

# ------------------------------------------------------------
# Create clean temporary folder
# ------------------------------------------------------------

if (Test-Path $TempPath) {
    Remove-Item -Path $TempPath -Recurse -Force
}

New-Item `
    -Path $TempPath `
    -ItemType Directory `
    -Force | Out-Null

# ------------------------------------------------------------
# Extract ZIP
# ------------------------------------------------------------

Write-Host "Extracting Fusion.zip..."

Expand-Archive `
    -Path $SourceZip `
    -DestinationPath $TempPath `
    -Force

# ------------------------------------------------------------
# Automatically find Fusion installer
# ------------------------------------------------------------

Write-Host "Searching for Fusion installer..."

$Installers = Get-ChildItem `
    -Path $TempPath `
    -Filter "*.exe" `
    -Recurse `
    -File

if ($Installers.Count -eq 0) {
    Write-Error "No EXE installer was found inside Fusion.zip."
    exit 2
}

if ($Installers.Count -gt 1) {
    Write-Host "Multiple EXE files found. Searching for the Fusion Offline installer..."

    $Installers = $Installers | Where-Object {
        $_.Name -like "F*Offline*x64.exe"
    }
}

if ($Installers.Count -ne 1) {
    Write-Error "Could not uniquely identify the Fusion installer."
    Write-Host "EXE files found:"

    Get-ChildItem `
        -Path $TempPath `
        -Filter "*.exe" `
        -Recurse `
        -File |
        ForEach-Object {
            Write-Host " - $($_.FullName)"
        }

    exit 3
}

$InstallerPath = $Installers[0].FullName

Write-Host "Fusion installer detected:"
Write-Host $InstallerPath

# ------------------------------------------------------------
# Find Fusion license
# ------------------------------------------------------------

Write-Host "Searching for Fusion.lic..."

$License = Get-ChildItem `
    -Path $TempPath `
    -Filter "Fusion.lic" `
    -Recurse `
    -File |
    Select-Object -First 1

if (-not $License) {
    Write-Error "Fusion.lic was not found inside Fusion.zip."
    exit 4
}

$LicensePath = $License.FullName

Write-Host "License detected:"
Write-Host $LicensePath

# ------------------------------------------------------------
# Install Fusion
# ------------------------------------------------------------

Write-Host "Installing Freedom Scientific Fusion..."

$Process = Start-Process `
    -FilePath $InstallerPath `
    -ArgumentList "/type Silent /DisableExternalServices" `
    -Wait `
    -PassThru

Write-Host "Installer exit code: $($Process.ExitCode)"

if ($Process.ExitCode -ne 0) {
    Write-Error "Fusion installation failed."
    exit $Process.ExitCode
}

Write-Host "Fusion installation completed successfully."

# ------------------------------------------------------------
# Create Authorization folder
# ------------------------------------------------------------

Write-Host "Creating/checking Authorization folder..."

New-Item `
    -Path $AuthorizationPath `
    -ItemType Directory `
    -Force | Out-Null

# ------------------------------------------------------------
# Copy license
# ------------------------------------------------------------

Write-Host "Copying Fusion.lic..."

Copy-Item `
    -Path $LicensePath `
    -Destination (Join-Path $AuthorizationPath "Fusion.lic") `
    -Force

# ------------------------------------------------------------
# Verify license
# ------------------------------------------------------------

$InstalledLicense = Join-Path `
    $AuthorizationPath `
    "Fusion.lic"

if (-not (Test-Path $InstalledLicense)) {
    Write-Error "Fusion.lic was not copied successfully."
    exit 5
}

Write-Host "Fusion.lic installed successfully."

# ------------------------------------------------------------
# Cleanup
# ------------------------------------------------------------

Write-Host "Removing temporary installation files..."

Remove-Item `
    -Path $TempPath `
    -Recurse `
    -Force

Write-Host ""
Write-Host "============================================================"
Write-Host " Freedom Scientific Fusion Installation Successful"
Write-Host "============================================================"

exit 0

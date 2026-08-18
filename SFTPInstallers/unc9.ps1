$ErrorActionPreference = "Stop"

# ============================================================
# Configuration
# ============================================================

$UNCPath = "\\pdc\Share2"
$ZipFile = "\\pdc\Share2\Fusion.zip"

$Username = "FC\df_deploy"
$Password = "Partners@2024"

$LocalZip = Join-Path $env:TEMP "Fusion.zip"
$TempPath = Join-Path $env:TEMP "FusionInstall"

$AuthorizationPath = "C:\Program Files\Freedom Scientific\Authorization"

# ============================================================
# Authenticate to UNC Share
# ============================================================

Write-Host "Connecting to $UNCPath..."

$SecurePassword = ConvertTo-SecureString `
    $Password `
    -AsPlainText `
    -Force

$Credential = New-Object `
    System.Management.Automation.PSCredential(
        $Username,
        $SecurePassword
    )

# Remove any existing connection
cmd.exe /c "net use $UNCPath /delete /y" 2>$null

# Connect to the share
New-PSDrive `
    -Name "FusionShare" `
    -PSProvider FileSystem `
    -Root $UNCPath `
    -Credential $Credential `
    -ErrorAction Stop | Out-Null

Write-Host "Connected successfully."

# ============================================================
# Verify ZIP
# ============================================================

$RemoteZip = "FusionShare:\Fusion.zip"

if (-not (Test-Path $RemoteZip)) {
    Write-Error "Fusion.zip could not be found on the UNC share."
    exit 1
}

# ============================================================
# Copy ZIP locally
# ============================================================

Write-Host "Copying Fusion.zip locally..."

Copy-Item `
    -Path $RemoteZip `
    -Destination $LocalZip `
    -Force

if (-not (Test-Path $LocalZip)) {
    Write-Error "Failed to copy Fusion.zip locally."
    exit 2
}

Write-Host "Fusion.zip copied successfully."

# ============================================================
# Disconnect UNC share
# ============================================================

Remove-PSDrive `
    -Name "FusionShare" `
    -Force

# ============================================================
# Extract ZIP
# ============================================================

Write-Host "Extracting Fusion.zip..."

if (Test-Path $TempPath) {
    Remove-Item `
        -Path $TempPath `
        -Recurse `
        -Force
}

New-Item `
    -Path $TempPath `
    -ItemType Directory `
    -Force | Out-Null

Expand-Archive `
    -Path $LocalZip `
    -DestinationPath $TempPath `
    -Force

# ============================================================
# Automatically find EXE
# ============================================================

Write-Host "Searching for Fusion installer..."

$Installers = Get-ChildItem `
    -Path $TempPath `
    -Filter "*.exe" `
    -Recurse `
    -File

if ($Installers.Count -eq 0) {
    Write-Error "No EXE installer found inside Fusion.zip."
    exit 3
}

# Prefer Fusion Offline x64 installer
$FusionInstaller = $Installers |
    Where-Object {
        $_.Name -like "F*Offline*x64.exe"
    } |
    Select-Object -First 1

if (-not $FusionInstaller) {
    Write-Error "Could not find the Fusion Offline x64 installer."
    exit 4
}

$InstallerPath = $FusionInstaller.FullName

Write-Host "Installer found:"
Write-Host $InstallerPath

# ============================================================
# Automatically find Fusion.lic
# ============================================================

Write-Host "Searching for Fusion.lic..."

$License = Get-ChildItem `
    -Path $TempPath `
    -Filter "Fusion.lic" `
    -Recurse `
    -File |
    Select-Object -First 1

if (-not $License) {
    Write-Error "Fusion.lic was not found."
    exit 5
}

$LicensePath = $License.FullName

Write-Host "License found:"
Write-Host $LicensePath

# ============================================================
# Install Fusion
# ============================================================

Write-Host "Installing Freedom Scientific Fusion..."

$Process = Start-Process `
    -FilePath $InstallerPath `
    -ArgumentList "/type Silent /DisableExternalServices" `
    -Wait `
    -PassThru

Write-Host "Fusion installer exit code: $($Process.ExitCode)"

if ($Process.ExitCode -ne 0) {
    Write-Error "Fusion installation failed."
    exit $Process.ExitCode
}

# ============================================================
# Create Authorization directory
# ============================================================

Write-Host "Creating Authorization directory..."

New-Item `
    -Path $AuthorizationPath `
    -ItemType Directory `
    -Force | Out-Null

# ============================================================
# Copy license
# ============================================================

Write-Host "Copying Fusion.lic..."

Copy-Item `
    -Path $LicensePath `
    -Destination (Join-Path $AuthorizationPath "Fusion.lic") `
    -Force

# ============================================================
# Verify license
# ============================================================

$InstalledLicense = Join-Path `
    $AuthorizationPath `
    "Fusion.lic"

if (-not (Test-Path $InstalledLicense)) {
    Write-Error "Fusion.lic was not copied successfully."
    exit 6
}

Write-Host "Fusion.lic installed successfully."

# ============================================================
# Cleanup
# ============================================================

Remove-Item `
    -Path $LocalZip `
    -Force `
    -ErrorAction SilentlyContinue

Remove-Item `
    -Path $TempPath `
    -Recurse `
    -Force `
    -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "=============================================="
Write-Host " Fusion installation completed successfully"
Write-Host "=============================================="

exit 0

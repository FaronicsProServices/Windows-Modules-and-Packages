# ============================================================
# Freedom Scientific Fusion - Automated Deployment
# ============================================================

$ErrorActionPreference = "Stop"

# ------------------------------------------------------------
# Configuration
# ------------------------------------------------------------

$Server       = "pdc"
$Share        = "Share2"
$RemoteZip    = "\\pdc\Share2\Fusion.zip"

$Username     = "FC\df_deploy"
$Password     = "Partners@2024"

$LocalZip     = Join-Path $env:TEMP "Fusion.zip"
$ExtractPath  = Join-Path $env:TEMP "FusionInstall"

$AuthorizationPath = "C:\Program Files\Freedom Scientific\Authorization"

# ------------------------------------------------------------
# Cleanup previous files
# ------------------------------------------------------------

Remove-Item $LocalZip -Force -ErrorAction SilentlyContinue
Remove-Item $ExtractPath -Recurse -Force -ErrorAction SilentlyContinue

# ------------------------------------------------------------
# Connect to file server using deployment account
# ------------------------------------------------------------

Write-Host "Connecting to \\$Server\$Share..."

$SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force

$Credential = New-Object System.Management.Automation.PSCredential(
    $Username,
    $SecurePassword
)

# Remove existing connection to PDC if one exists.
# Error is intentionally ignored.
cmd.exe /c "net use \\$Server /delete /y" 2>$null | Out-Null

# Create authenticated SMB connection
cmd.exe /c "net use \\$Server\$Share /user:$Username `"$Password`" /persistent:no"

if ($LASTEXITCODE -ne 0) {
    Write-Error "Unable to authenticate to \\$Server\$Share using $Username."
    exit 1
}

Write-Host "Connected successfully."

# ------------------------------------------------------------
# Copy ZIP locally
# ------------------------------------------------------------

Write-Host "Copying Fusion.zip locally..."

Copy-Item `
    -Path $RemoteZip `
    -Destination $LocalZip `
    -Force

if (-not (Test-Path $LocalZip)) {
    cmd.exe /c "net use \\$Server\$Share /delete /y" 2>$null | Out-Null

    Write-Error "Failed to copy Fusion.zip."
    exit 2
}

Write-Host "Fusion.zip copied successfully."

# ------------------------------------------------------------
# Disconnect from PDC
# ------------------------------------------------------------

cmd.exe /c "net use \\$Server\$Share /delete /y" 2>$null | Out-Null

Write-Host "Disconnected from file server."

# ------------------------------------------------------------
# Extract ZIP
# ------------------------------------------------------------

Write-Host "Extracting Fusion.zip..."

New-Item `
    -Path $ExtractPath `
    -ItemType Directory `
    -Force | Out-Null

Expand-Archive `
    -Path $LocalZip `
    -DestinationPath $ExtractPath `
    -Force

Write-Host "Extraction completed."

# ------------------------------------------------------------
# Automatically locate Fusion installer
# ------------------------------------------------------------

Write-Host "Searching for Fusion installer..."

$Installer = Get-ChildItem `
    -Path $ExtractPath `
    -Recurse `
    -Filter "*.exe" `
    -File |
    Where-Object {
        $_.Name -like "F*Offline*x64.exe"
    } |
    Select-Object -First 1

if (-not $Installer) {
    Write-Error "Fusion Offline x64 installer was not found."
    exit 3
}

Write-Host "Installer detected:"
Write-Host $Installer.FullName

# ------------------------------------------------------------
# Automatically locate license
# ------------------------------------------------------------

Write-Host "Searching for Fusion.lic..."

$License = Get-ChildItem `
    -Path $ExtractPath `
    -Recurse `
    -Filter "Fusion.lic" `
    -File |
    Select-Object -First 1

if (-not $License) {
    Write-Error "Fusion.lic was not found."
    exit 4
}

Write-Host "License detected:"
Write-Host $License.FullName

# ------------------------------------------------------------
# Install Fusion
# ------------------------------------------------------------

Write-Host "Installing Freedom Scientific Fusion..."

$Process = Start-Process `
    -FilePath $Installer.FullName `
    -ArgumentList "/type Silent /DisableExternalServices" `
    -Wait `
    -PassThru

Write-Host "Fusion installer exit code: $($Process.ExitCode)"

if ($Process.ExitCode -ne 0) {
    Write-Error "Fusion installation failed."
    exit $Process.ExitCode
}

Write-Host "Fusion installed successfully."

# ------------------------------------------------------------
# Create Authorization directory
# ------------------------------------------------------------

New-Item `
    -Path $AuthorizationPath `
    -ItemType Directory `
    -Force | Out-Null

# ------------------------------------------------------------
# Copy license
# ------------------------------------------------------------

Write-Host "Installing Fusion license..."

Copy-Item `
    -Path $License.FullName `
    -Destination "$AuthorizationPath\Fusion.lic" `
    -Force

if (-not (Test-Path "$AuthorizationPath\Fusion.lic")) {
    Write-Error "Fusion license installation failed."
    exit 5
}

Write-Host "Fusion license installed successfully."

# ------------------------------------------------------------
# Cleanup
# ------------------------------------------------------------

Remove-Item `
    -Path $LocalZip `
    -Force `
    -ErrorAction SilentlyContinue

Remove-Item `
    -Path $ExtractPath `
    -Recurse `
    -Force `
    -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "==============================================="
Write-Host " Fusion Deployment Completed Successfully"
Write-Host "==============================================="

exit 0

param(
    [Parameter(Mandatory = $true)]
    [string]$SftpLink,

    [Parameter(Mandatory = $true)]
    [string]$SftpUser,

    [Parameter(Mandatory = $true)]
    [string]$SftpPassword,

    [Parameter(Mandatory = $true)]
    [string]$LicenseFileName
)

$ErrorActionPreference = "Stop"

# ============================================================
# Paths
# ============================================================

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition

$WinSCPDll = Join-Path $ScriptRoot "WinSCPnet.dll"

$WorkFolder    = Join-Path $env:TEMP "FusionSFTP"
$ZipFile       = Join-Path $WorkFolder "Fusion.zip"
$ExtractFolder = Join-Path $WorkFolder "Extracted"

$LicenseDestination = "C:\Program Files\Freedom Scientific\Authorization"

# ============================================================
# Logging
# ============================================================

function Write-Log {
    param([string]$Message)

    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
}

Write-Log "============================================================"
Write-Log "Fusion SFTP License Deployment"
Write-Log "============================================================"

# ============================================================
# Check WinSCP DLL
# ============================================================

if (-not (Test-Path $WinSCPDll)) {
    throw "WinSCPnet.dll not found: $WinSCPDll"
}

Write-Log "WinSCPnet.dll found."

# ============================================================
# Prepare working directory
# ============================================================

if (Test-Path $WorkFolder) {
    Remove-Item -Path $WorkFolder -Recurse -Force
}

New-Item -Path $WorkFolder -ItemType Directory -Force | Out-Null
New-Item -Path $ExtractFolder -ItemType Directory -Force | Out-Null

# ============================================================
# Parse SFTP URL
# ============================================================

$Uri = [System.Uri]$SftpLink

if ($Uri.Scheme -ne "sftp") {
    throw "SftpLink must start with sftp://"
}

$RemotePath = $Uri.AbsolutePath

# Convert //Share/Fusion.zip to /Share/Fusion.zip
$RemotePath = "/" + $RemotePath.TrimStart("/")

Write-Log "SFTP Host: $($Uri.Host)"
Write-Log "SFTP Port: $($Uri.Port)"
Write-Log "Remote File: $RemotePath"

# ============================================================
# Load WinSCP
# ============================================================

Add-Type -Path $WinSCPDll

# ============================================================
# SFTP Connection
# ============================================================

$SessionOptions = New-Object WinSCP.SessionOptions

$SessionOptions.Protocol   = [WinSCP.Protocol]::Sftp
$SessionOptions.HostName   = $Uri.Host
$SessionOptions.PortNumber = $Uri.Port
$SessionOptions.UserName   = $SftpUser
$SessionOptions.Password   = $SftpPassword

# Accept server host key
# For production, replace this with the actual fingerprint.
$SessionOptions.GiveUpSecurityAndAcceptAnySshHostKey = $true

$Session = New-Object WinSCP.Session

try {

    Write-Log "Connecting to SFTP server..."

    $Session.Open($SessionOptions)

    Write-Log "SFTP connection successful."

    # ========================================================
    # Download Fusion.zip
    # ========================================================

    Write-Log "Downloading Fusion.zip..."

    $TransferResult = $Session.GetFiles(
        $RemotePath,
        $ZipFile,
        $false
    )

    $TransferResult.Check()

    Write-Log "Fusion.zip downloaded successfully."

}
finally {

    if ($Session) {
        $Session.Dispose()
    }

    Write-Log "SFTP connection closed."
}

# ============================================================
# Verify ZIP
# ============================================================

if (-not (Test-Path $ZipFile)) {
    throw "Fusion.zip was not downloaded."
}

Write-Log "ZIP file verified."

# ============================================================
# Extract Fusion.zip
# ============================================================

Write-Log "Extracting Fusion.zip..."

Expand-Archive `
    -Path $ZipFile `
    -DestinationPath $ExtractFolder `
    -Force

Write-Log "Fusion.zip extracted successfully."

# ============================================================
# Find Fusion.lic
# ============================================================

$LicenseFile = Get-ChildItem `
    -Path $ExtractFolder `
    -Filter $LicenseFileName `
    -File `
    -Recurse `
    -ErrorAction SilentlyContinue |
    Select-Object -First 1

if (-not $LicenseFile) {
    throw "$LicenseFileName was not found inside Fusion.zip."
}

Write-Log "$LicenseFileName found."
Write-Log "Source: $($LicenseFile.FullName)"

# ============================================================
# Create destination folder
# ============================================================

if (-not (Test-Path $LicenseDestination)) {

    Write-Log "Creating authorization folder..."

    New-Item `
        -Path $LicenseDestination `
        -ItemType Directory `
        -Force | Out-Null
}

# ============================================================
# Copy license file
# ============================================================

Push-Location $LicenseFile.DirectoryName

try {

    Copy-Item -Path ".\Fusion.lic" -Destination "C:\Program Files\Freedom Scientific\Authorization" -Force -Recurse

}
finally {

    Pop-Location

}

Write-Log "License copied successfully."

# ============================================================
# Verify
# ============================================================

$InstalledLicense = Join-Path `
    $LicenseDestination `
    $LicenseFileName

if (-not (Test-Path $InstalledLicense)) {
    throw "License verification failed: $InstalledLicense"
}

Write-Log "License verified at:"
Write-Log $InstalledLicense

# ============================================================
# Cleanup
# ============================================================

Write-Log "Cleaning temporary files..."

if (Test-Path $WorkFolder) {
    Remove-Item `
        -Path $WorkFolder `
        -Recurse `
        -Force
}

Write-Log "============================================================"
Write-Log "Fusion license deployment completed successfully."
Write-Log "============================================================"

exit 0

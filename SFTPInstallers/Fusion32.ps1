Requires -Version 5.1

[CmdletBinding()]
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
# CONFIGURATION
# ============================================================

$WorkDir = "C:\Temp\FusionDeployment"

$WinScpDir = Join-Path $WorkDir "WinSCP"

$ExtractDir = Join-Path $WorkDir "ExtractedFiles"

$FusionZip = Join-Path $WorkDir "Fusion.zip"

$WinScpExe = Join-Path $WinScpDir "WinSCP.exe"

$WinScpDll = Join-Path $WinScpDir "WinSCPnet.dll"

$DeploymentLog = "$env:WINDIR\Temp\Fusion2026-Deployment.log"

$FusionLogPath = "$env:WINDIR\Temp\Fusion2026-Install.log"

$LicenseDestination =
    "C:\Program Files\Freedom Scientific\Authorization"

$GitHubBaseUrl =
    "https://raw.githubusercontent.com/FaronicsProServices/Windows-Modules-and-Packages/main/SFTPInstallers"


# ============================================================
# LOGGING
# ============================================================

function Write-Log {

    param(
        [string]$Message
    )

    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    $LogLine = "[$Timestamp] $Message"

    Write-Host $LogLine

    Add-Content `
        -Path $DeploymentLog `
        -Value $LogLine
}


function Stop-WithError {

    param(
        [string]$Message
    )

    Write-Log "ERROR: $Message"

    exit 1
}


# ============================================================
# START
# ============================================================

Write-Log "============================================================"
Write-Log "Fusion 2026 SFTP Deployment"
Write-Log "============================================================"


# ============================================================
# ADMINISTRATOR CHECK
# ============================================================

try {

    $CurrentIdentity =
        [Security.Principal.WindowsIdentity]::GetCurrent()

    $Principal =
        New-Object Security.Principal.WindowsPrincipal(
            $CurrentIdentity
        )

    if (-not $Principal.IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator
    )) {

        Stop-WithError "This script must be run as Administrator."
    }

    Write-Log "Administrator privileges confirmed."

}
catch {

    Stop-WithError `
        "Unable to verify administrator privileges: $($_.Exception.Message)"
}


# ============================================================
# PREPARE WORKING DIRECTORIES
# ============================================================

try {

    Write-Log "Creating temporary working directories..."

    if (Test-Path -LiteralPath $WorkDir) {

        Remove-Item `
            -Path $WorkDir `
            -Recurse `
            -Force `
            -ErrorAction SilentlyContinue
    }

    New-Item `
        -Path $WorkDir `
        -ItemType Directory `
        -Force |
        Out-Null

    New-Item `
        -Path $WinScpDir `
        -ItemType Directory `
        -Force |
        Out-Null

    New-Item `
        -Path $ExtractDir `
        -ItemType Directory `
        -Force |
        Out-Null

}
catch {

    Stop-WithError `
        "Failed to create working directories: $($_.Exception.Message)"
}


# ============================================================
# PARSE SFTP URL
# ============================================================

try {

    Write-Log "Parsing SFTP URL..."

    $Uri = [System.Uri]$SftpLink

    if ($Uri.Scheme -ne "sftp") {

        throw "The supplied URL is not an SFTP URL."
    }

    $SftpServer = $Uri.Host

    if ($Uri.Port -gt 0) {
        $SftpPort = $Uri.Port
    }
    else {
        $SftpPort = 22
    }

    $RemoteFile =
        $Uri.AbsolutePath -replace '^//', '/'

    Write-Log "SFTP Server : $SftpServer"
    Write-Log "SFTP Port   : $SftpPort"
    Write-Log "Remote File : $RemoteFile"

}
catch {

    Stop-WithError `
        "Failed to parse SFTP URL: $($_.Exception.Message)"
}


# ============================================================
# DOWNLOAD WINSCP
# ============================================================

try {

    Write-Log "Downloading WinSCP components..."

    [Net.ServicePointManager]::SecurityProtocol =
        [Net.SecurityProtocolType]::Tls12

    Invoke-WebRequest `
        -Uri "$GitHubBaseUrl/WinSCP.exe" `
        -OutFile $WinScpExe `
        -UseBasicParsing

    Invoke-WebRequest `
        -Uri "$GitHubBaseUrl/WinSCPnet.dll" `
        -OutFile $WinScpDll `
        -UseBasicParsing

    if (-not (Test-Path -LiteralPath $WinScpExe)) {

        throw "WinSCP.exe was not downloaded."
    }

    if (-not (Test-Path -LiteralPath $WinScpDll)) {

        throw "WinSCPnet.dll was not downloaded."
    }

    Write-Log "WinSCP components downloaded successfully."

}
catch {

    Stop-WithError `
        "Failed to download WinSCP components: $($_.Exception.Message)"
}


# ============================================================
# LOAD WINSCP.NET
# ============================================================

try {

    Write-Log "Loading WinSCP .NET assembly..."

    Add-Type `
        -Path $WinScpDll

    Write-Log "WinSCP .NET assembly loaded successfully."

}
catch {

    Stop-WithError `
        "Failed to load WinSCPnet.dll: $($_.Exception.Message)"
}


# ============================================================
# CONNECT TO SFTP
# ============================================================

$Session = $null

try {

    Write-Log "Connecting to SFTP server..."

    $SessionOptions =
        New-Object WinSCP.SessionOptions

    $SessionOptions.Protocol =
        [WinSCP.Protocol]::Sftp

    $SessionOptions.HostName =
        $SftpServer

    $SessionOptions.PortNumber =
        $SftpPort

    $SessionOptions.UserName =
        $SftpUser

    $SessionOptions.Password =
        $SftpPassword

    # Accept SFTP host key
    $SessionOptions.SshHostKeyPolicy =
        [WinSCP.SshHostKeyPolicy]::GiveUpSecurityAndAcceptAny

    $Session =
        New-Object WinSCP.Session

    $Session.ExecutablePath =
        $WinScpExe

    $Session.Open(
        $SessionOptions
    )

    Write-Log "SFTP connection established."

    # ========================================================
    # DOWNLOAD FUSION ZIP
    # ========================================================

    Write-Log "Downloading Fusion package..."
    Write-Log "Remote file: $RemoteFile"

    $TransferResult =
        $Session.GetFiles(
            $RemoteFile,
            $FusionZip,
            $false
        )

    $TransferResult.Check()

    if (-not (Test-Path -LiteralPath $FusionZip)) {

        throw "Fusion.zip was not downloaded."
    }

    $ZipSize =
        (Get-Item -LiteralPath $FusionZip).Length

    Write-Log "Fusion ZIP downloaded successfully."
    Write-Log "Downloaded size: $ZipSize bytes"

}
catch {

    Stop-WithError `
        "SFTP download failed: $($_.Exception.Message)"
}
finally {

    if ($Session) {
        $Session.Dispose()
    }
}
# ------------------------------------------------------------
# Verify ZIP
# ------------------------------------------------------------

if (-not (Test-Path $ZipFile)) {
    throw "Fusion.zip was not found after download."
}

Write-Host "Downloaded file verified." -ForegroundColor Green

# ------------------------------------------------------------
# Extract ZIP
# ------------------------------------------------------------

Write-Host "Extracting Fusion.zip..." -ForegroundColor Cyan

Expand-Archive `
    -Path $ZipFile `
    -DestinationPath $ExtractDir `
    -Force

Write-Host "Extraction completed." -ForegroundColor Green

# ------------------------------------------------------------
# Find ANY EXE installer
# ------------------------------------------------------------

Write-Host "Scanning extracted files for EXE installer..." -ForegroundColor Cyan

$Executables = Get-ChildItem `
    -Path $ExtractDir `
    -Filter "*.exe" `
    -Recurse `
    -File

if (-not $Executables) {
    throw "No .exe file was found in the extracted package."
}

Write-Host ""
Write-Host "EXE file(s) found:" -ForegroundColor Cyan

foreach ($Exe in $Executables) {
    Write-Host "  $($Exe.FullName)" -ForegroundColor Gray
}

# Use the first EXE found
$Installer = $Executables | Select-Object -First 1

Write-Host ""
Write-Host "Installer selected:" -ForegroundColor Green
Write-Host $Installer.FullName -ForegroundColor Green

# ------------------------------------------------------------
# Find Fusion License
# ------------------------------------------------------------

Write-Host ""
Write-Host "Searching for Fusion.lic..." -ForegroundColor Cyan

$License = Get-ChildItem `
    -Path $ExtractDir `
    -Filter "Fusion.lic" `
    -Recurse `
    -File |
    Select-Object -First 1

if (-not $License) {
    throw "Fusion.lic was not found in the extracted package."
}

Write-Host "License found:" -ForegroundColor Green
Write-Host $License.FullName -ForegroundColor Green

# ------------------------------------------------------------
# Install Freedom Scientific Fusion
# ------------------------------------------------------------

Write-Host ""
Write-Host "============================================" -ForegroundColor Yellow
Write-Host "Installing Freedom Scientific Fusion..." -ForegroundColor Yellow
Write-Host "============================================" -ForegroundColor Yellow

Set-Location $Installer.DirectoryName

# YOUR EXACT INSTALL COMMAND
Start-Process `
    -FilePath $Installer.FullName `
    -ArgumentList "/type Silent /DisableExternalServices" `
    -Wait

Write-Host "Fusion installer process completed." -ForegroundColor Green

# ------------------------------------------------------------
# Copy License
# ------------------------------------------------------------

Write-Host ""
Write-Host "Copying Fusion license..." -ForegroundColor Yellow

if (-not (Test-Path $LicenseDestination)) {
    New-Item `
        -Path $LicenseDestination `
        -ItemType Directory `
        -Force | Out-Null
}

Copy-Item `
    -Path $License.FullName `
    -Destination $LicenseDestination `
    -Force `
    -Recurse

Write-Host "License copied successfully." -ForegroundColor Green

# ------------------------------------------------------------
# Verify License
# ------------------------------------------------------------

$InstalledLicense = Join-Path `
    $LicenseDestination `
    "Fusion.lic"

if (-not (Test-Path $InstalledLicense)) {
    throw "Fusion.lic verification failed."
}

Write-Host "License verification successful." -ForegroundColor Green

# ------------------------------------------------------------
# Cleanup
# ------------------------------------------------------------

Write-Host ""
Write-Host "Cleaning up temporary files..." -ForegroundColor Cyan

Set-Location C:\

Remove-Item `
    -Path $WorkDir `
    -Recurse `
    -Force

Write-Host "Temporary files removed." -ForegroundColor Green

# ------------------------------------------------------------
# SUCCESS
# ------------------------------------------------------------

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host " FUSION DEPLOYMENT SUCCESSFUL" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host "Fusion installed."
Write-Host "Fusion.lic copied and verified."
Write-Host "Temporary files cleaned."
Write-Host "============================================" -ForegroundColor Green

#requires -version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SftpLink,

    [Parameter(Mandatory = $true)]
    [string]$SftpUser,

    [Parameter(Mandatory = $true)]
    [string]$SftpPassword,

    [Parameter(Mandatory = $true)]
    [string]$LicenseFileName,

    [Parameter(Mandatory = $false)]
    [string]$GitHubBaseUrl = "https://raw.githubusercontent.com/FaronicsProServices/Windows-Modules-and-Packages/main/SFTPInstallers"
)

$ErrorActionPreference = "Stop"

# ------------------------------------------------------------
# Configuration
# ------------------------------------------------------------

$WorkDir = "C:\Temp\FusionDeployment"
$LocalZipPath = Join-Path $WorkDir "FusionPackage.zip"
$ExtractPath = Join-Path $WorkDir "ExtractedFiles"
$WinScpFolder = Join-Path $WorkDir "WinSCP"

$LicenseDestination = "C:\Program Files\Freedom Scientific\Authorization"

$InstallLog = "$env:WINDIR\Temp\Fusion2026-Install.log"

$WinScpExe = Join-Path $WinScpFolder "WinSCP.exe"
$WinScpDll = Join-Path $WinScpFolder "WinSCPnet.dll"

# ------------------------------------------------------------
# Functions
# ------------------------------------------------------------

function Write-Log {
    param(
        [string]$Message
    )

    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
}

function Stop-WithError {
    param(
        [string]$Message
    )

    Write-Error $Message
    exit 1
}

# ------------------------------------------------------------
# Parse SFTP URL
# ------------------------------------------------------------

try {
    Write-Log "Parsing SFTP URL..."

    $Uri = [System.Uri]$SftpLink

    if ($Uri.Scheme -ne "sftp") {
        throw "The supplied link is not an SFTP URL."
    }

    $SftpServer = $Uri.Host
    $SftpPort = if ($Uri.Port -gt 0) { $Uri.Port } else { 22 }

    # Converts:
    # sftp://10.0.1.215:22//Share/Fusion.zip
    #
    # Into:
    # /Share/Fusion.zip

    $RemoteZipPath = $Uri.AbsolutePath -replace '^//', '/'

    Write-Log "SFTP Server : $SftpServer"
    Write-Log "SFTP Port   : $SftpPort"
    Write-Log "Remote File : $RemoteZipPath"
}
catch {
    Stop-WithError "Invalid SFTP URL: $($_.Exception.Message)"
}

# ------------------------------------------------------------
# Create working directories
# ------------------------------------------------------------

try {
    Write-Log "Creating temporary working directories..."

    if (Test-Path $WorkDir) {
        Remove-Item $WorkDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
    New-Item -ItemType Directory -Path $WinScpFolder -Force | Out-Null
    New-Item -ItemType Directory -Path $ExtractPath -Force | Out-Null
}
catch {
    Stop-WithError "Failed to create working directories: $($_.Exception.Message)"
}

# ------------------------------------------------------------
# Download WinSCP.exe and WinSCPnet.dll from GitHub
# ------------------------------------------------------------

try {
    Write-Log "Downloading WinSCP components..."

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    Invoke-WebRequest `
        -Uri "$GitHubBaseUrl/WinSCP.exe" `
        -OutFile $WinScpExe `
        -UseBasicParsing

    Invoke-WebRequest `
        -Uri "$GitHubBaseUrl/WinSCPnet.dll" `
        -OutFile $WinScpDll `
        -UseBasicParsing

    if (-not (Test-Path $WinScpExe)) {
        throw "WinSCP.exe was not downloaded."
    }

    if (-not (Test-Path $WinScpDll)) {
        throw "WinSCPnet.dll was not downloaded."
    }

    Write-Log "WinSCP components downloaded successfully."
}
catch {
    Stop-WithError "Failed to download WinSCP components: $($_.Exception.Message)"
}

# ------------------------------------------------------------
# Load WinSCP .NET assembly
# ------------------------------------------------------------

try {
    Write-Log "Loading WinSCP .NET assembly..."

    Add-Type -Path $WinScpDll

    Write-Log "WinSCP .NET assembly loaded successfully."
}
catch {
    Stop-WithError "Failed to load WinSCPnet.dll: $($_.Exception.Message)"
}

# ------------------------------------------------------------
# Download Fusion ZIP from SFTP
# ------------------------------------------------------------

$Session = $null

try {
    Write-Log "Connecting to SFTP server..."

    $SessionOptions = New-Object WinSCP.SessionOptions

    $SessionOptions.Protocol = [WinSCP.Protocol]::Sftp
    $SessionOptions.HostName = $SftpServer
    $SessionOptions.PortNumber = $SftpPort
    $SessionOptions.UserName = $SftpUser
    $SessionOptions.Password = $SftpPassword

    # Accept the server host key.
    # For production environments, using a known host key fingerprint
    # is more secure than accepting any host key.
    $SessionOptions.GiveUpSecurityAndAcceptAnySshHostKey = $true

    $Session = New-Object WinSCP.Session

    $Session.Open($SessionOptions)

    Write-Log "SFTP connection established."

    Write-Log "Downloading Fusion package..."

    $TransferResult = $Session.GetFiles(
        $RemoteZipPath,
        $LocalZipPath,
        $false
    )

    $TransferResult.Check()

    if (-not (Test-Path $LocalZipPath)) {
        throw "Fusion ZIP was not downloaded."
    }

    Write-Log "Fusion ZIP downloaded successfully."
}
catch {
    Stop-WithError "Failed to download Fusion ZIP from SFTP: $($_.Exception.Message)"
}
finally {
    if ($Session) {
        $Session.Dispose()
    }
}

# ------------------------------------------------------------
# Extract Fusion ZIP
# ------------------------------------------------------------

try {
    Write-Log "Extracting Fusion package..."

    Expand-Archive `
        -Path $LocalZipPath `
        -DestinationPath $ExtractPath `
        -Force

    Write-Log "Fusion package extracted successfully."
}
catch {
    Stop-WithError "Failed to extract Fusion ZIP: $($_.Exception.Message)"
}

# ------------------------------------------------------------
# Locate Fusion installer
# ------------------------------------------------------------

try {
    Write-Log "Searching for Fusion installer..."

    $InstallerCandidates = Get-ChildItem `
        -Path $ExtractPath `
        -Filter "*.exe" `
        -Recurse `
        -File

    if (-not $InstallerCandidates) {
        throw "No EXE installer was found inside Fusion.zip."
    }

    # Ignore any WinSCP executables if they happen to exist in the package.
    $InstallerCandidates = $InstallerCandidates |
        Where-Object {
            $_.Name -notin @(
                "WinSCP.exe"
            )
        }

    if (-not $InstallerCandidates) {
        throw "No valid Fusion installer was found inside Fusion.zip."
    }

    # Prefer an installer containing Fusion in its filename.
    $FusionInstaller = $InstallerCandidates |
        Where-Object {
            $_.Name -match "Fusion"
        } |
        Select-Object -First 1

    # If no Fusion-named EXE exists, use the first remaining EXE.
    if (-not $FusionInstaller) {
        $FusionInstaller = $InstallerCandidates | Select-Object -First 1
    }

    Write-Log "Installer found: $($FusionInstaller.FullName)"
}
catch {
    Stop-WithError "Failed to locate Fusion installer: $($_.Exception.Message)"
}

# ------------------------------------------------------------
# Install Fusion silently
# ------------------------------------------------------------

try {
    Write-Log "Starting Fusion silent installation..."

    if (Test-Path $InstallLog) {
        Remove-Item $InstallLog -Force -ErrorAction SilentlyContinue
    }

    # Exact requested installation switches:
    #
    # /Type Silent
    # /DisableExternalAIServices
    # /Log "%WINDIR%\Temp\Fusion2026-Install.log"

    $InstallerArguments = @(
        "/Type"
        "Silent"
        "/DisableExternalAIServices"
        "/Log"
        "`"$InstallLog`""
    )

    Write-Log "Running Fusion installer..."

    $Process = Start-Process `
        -FilePath $FusionInstaller.FullName `
        -ArgumentList $InstallerArguments `
        -Wait `
        -PassThru

    Write-Log "Fusion installer exit code: $($Process.ExitCode)"

    if ($Process.ExitCode -ne 0) {

        if (Test-Path $InstallLog) {
            Write-Log "Fusion installation log detected at:"
            Write-Log $InstallLog
        }

        throw "Fusion installation failed with exit code $($Process.ExitCode)."
    }

    Write-Log "Fusion installation completed successfully."
}
catch {
    Stop-WithError "Fusion installation failed: $($_.Exception.Message)"
}

# ------------------------------------------------------------
# Verify Freedom Scientific installation
# ------------------------------------------------------------

try {
    Write-Log "Verifying Fusion installation..."

    $FreedomScientificPath = Join-Path `
        $env:ProgramFiles `
        "Freedom Scientific"

    $MaxAttempts = 24
    $Attempt = 0
    $Installed = $false

    while (-not $Installed -and $Attempt -lt $MaxAttempts) {

        if (Test-Path $FreedomScientificPath) {

            $InstalledItems = Get-ChildItem `
                -Path $FreedomScientificPath `
                -ErrorAction SilentlyContinue

            if ($InstalledItems) {
                $Installed = $true
                break
            }
        }

        Start-Sleep -Seconds 5
        $Attempt++
    }

    if (-not $Installed) {
        throw "Freedom Scientific installation directory was not detected."
    }

    Write-Log "Fusion installation verified."
}
catch {
    Stop-WithError "Fusion installation verification failed: $($_.Exception.Message)"
}

# ------------------------------------------------------------
# Locate license file
# ------------------------------------------------------------

try {
    Write-Log "Searching for license file: $LicenseFileName"

    $LicenseFile = Get-ChildItem `
        -Path $ExtractPath `
        -Filter $LicenseFileName `
        -Recurse `
        -File |
        Select-Object -First 1

    if (-not $LicenseFile) {
        throw "License file '$LicenseFileName' was not found inside Fusion.zip."
    }

    Write-Log "License file found: $($LicenseFile.FullName)"
}
catch {
    Stop-WithError "Failed to locate license file: $($_.Exception.Message)"
}

# ------------------------------------------------------------
# Copy license file
# ------------------------------------------------------------

try {
    Write-Log "Creating license destination directory..."

    if (-not (Test-Path $LicenseDestination)) {
        New-Item `
            -ItemType Directory `
            -Path $LicenseDestination `
            -Force |
            Out-Null
    }

    Write-Log "Copying license file..."

    Copy-Item `
        -Path $LicenseFile.FullName `
        -Destination (Join-Path $LicenseDestination $LicenseFileName) `
        -Force

    $InstalledLicense = Join-Path `
        $LicenseDestination `
        $LicenseFileName

    if (-not (Test-Path $InstalledLicense)) {
        throw "License file could not be verified after copying."
    }

    Write-Log "License successfully copied to:"
    Write-Log $InstalledLicense
}
catch {
    Stop-WithError "Failed to copy license file: $($_.Exception.Message)"
}

# ------------------------------------------------------------
# Cleanup
# ------------------------------------------------------------

try {
    Write-Log "Cleaning up temporary deployment files..."

    if (Test-Path $WorkDir) {
        Remove-Item `
            -Path $WorkDir `
            -Recurse `
            -Force `
            -ErrorAction SilentlyContinue
    }

    Write-Log "Cleanup completed."
}
catch {
    Write-Log "Warning: Temporary files could not be completely removed."
}

# ------------------------------------------------------------
# Complete
# ------------------------------------------------------------

Write-Log "=============================================="
Write-Log "Fusion deployment completed successfully."
Write-Log "=============================================="

exit 0

#Requires -RunAsAdministrator

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

$WorkRoot = "C:\Temp\FusionDeployment"

$WinScpDirectory = Join-Path $WorkRoot "WinSCP"

$ZipFile = Join-Path $WorkRoot "Fusion.zip"

$ExtractPath = Join-Path $WorkRoot "ExtractedFiles"

$WinScpExe = Join-Path $WinScpDirectory "WinSCP.exe"

$WinScpDll = Join-Path $WinScpDirectory "WinSCPnet.dll"

$InstallerLog = Join-Path `
    $env:WINDIR `
    "Temp\Fusion2026-Install.log"

$DeploymentLog = Join-Path `
    $env:WINDIR `
    "Temp\Fusion2026-Deployment.log"

$LicenseDestinationDirectory =
    "C:\Program Files\Freedom Scientific\Authorization"

$LicenseDestination =
    Join-Path `
        $LicenseDestinationDirectory `
        $LicenseFileName


# ============================================================
# LOGGING
# ============================================================

function Write-Log {

    param(
        [string]$Message
    )

    $Timestamp =
        Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    $Line =
        "[$Timestamp] $Message"

    Write-Host $Line

    Add-Content `
        -Path $DeploymentLog `
        -Value $Line
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
Write-Log "Fusion 2026 SFTP Deployment Started"
Write-Log "============================================================"


# ============================================================
# ADMIN CHECK
# ============================================================

$CurrentIdentity =
    [Security.Principal.WindowsIdentity]::GetCurrent()

$Principal =
    New-Object `
        Security.Principal.WindowsPrincipal(
            $CurrentIdentity
        )

if (-not $Principal.IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)) {

    Stop-WithError `
        "The script must be run as Administrator."
}

Write-Log "Administrator privileges confirmed."


# ============================================================
# CREATE WORKING DIRECTORIES
# ============================================================

Write-Log "Creating temporary working directories..."

if (Test-Path $WorkRoot) {

    Remove-Item `
        -Path $WorkRoot `
        -Recurse `
        -Force `
        -ErrorAction SilentlyContinue
}

New-Item `
    -Path $WorkRoot `
    -ItemType Directory `
    -Force |
    Out-Null

New-Item `
    -Path $WinScpDirectory `
    -ItemType Directory `
    -Force |
    Out-Null

New-Item `
    -Path $ExtractPath `
    -ItemType Directory `
    -Force |
    Out-Null


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

    # Convert:
    #
    # sftp://10.0.1.215:22//Share/Fusion.zip
    #
    # to:
    #
    # /Share/Fusion.zip

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

    $GitHubBase =
        "https://raw.githubusercontent.com/FaronicsProServices/Windows-Modules-and-Packages/main/SFTPInstallers"

    Invoke-WebRequest `
        -Uri "$GitHubBase/WinSCP.exe" `
        -OutFile $WinScpExe `
        -UseBasicParsing

    Invoke-WebRequest `
        -Uri "$GitHubBase/WinSCPnet.dll" `
        -OutFile $WinScpDll `
        -UseBasicParsing

    if (-not (Test-Path $WinScpExe)) {

        throw "WinSCP.exe download failed."
    }

    if (-not (Test-Path $WinScpDll)) {

        throw "WinSCPnet.dll download failed."
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

    # Accept server host key.
    # For hardened production deployment, pin the
    # server's SSH host-key fingerprint instead.

    $SessionOptions.GiveUpSecurityAndAcceptAnySshHostKey =
        $true

    $Session =
        New-Object WinSCP.Session

    $Session.Open(
        $SessionOptions
    )

    Write-Log "SFTP connection established."


    # ========================================================
    # DOWNLOAD FUSION ZIP
    # ========================================================

    Write-Log "Downloading Fusion package..."
    Write-Log "Remote: $RemoteFile"
    Write-Log "Local : $ZipFile"

    $TransferResult =
        $Session.GetFiles(
            $RemoteFile,
            $ZipFile,
            $false
        )

    $TransferResult.Check()

    if (-not (Test-Path $ZipFile)) {

        throw "Fusion.zip was not downloaded."
    }

    $ZipSize =
        (Get-Item $ZipFile).Length

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


# ============================================================
# EXTRACT FUSION ZIP
# ============================================================

try {

    Write-Log "Extracting Fusion package..."

    Expand-Archive `
        -LiteralPath $ZipFile `
        -DestinationPath $ExtractPath `
        -Force

    Write-Log "Fusion package extracted successfully."
}

catch {

    Stop-WithError `
        "Failed to extract Fusion.zip: $($_.Exception.Message)"
}


# ============================================================
# FIND FUSION INSTALLER
# ============================================================

try {

    Write-Log "Searching for Fusion installer..."

    $FusionInstaller =
        Get-ChildItem `
            -Path $ExtractPath `
            -Filter "*.exe" `
            -File `
            -Recurse `
            -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -match `
                "^F2026.*-Offline-x64\.exe$"
        } |
        Select-Object -First 1


    if (-not $FusionInstaller) {

        throw `
            "Fusion 2026 Offline x64 installer was not found."
    }

    Write-Log "Installer found:"
    Write-Log $FusionInstaller.FullName
}

catch {

    Stop-WithError `
        "Failed to locate Fusion installer: $($_.Exception.Message)"
}


# ============================================================
# INSTALL FUSION
# ============================================================

try {

    Write-Log "Starting Fusion silent installation..."

    if (Test-Path $InstallerLog) {

        Remove-Item `
            -Path $InstallerLog `
            -Force `
            -ErrorAction SilentlyContinue
    }


    # ========================================================
    # ONLY FUSION SWITCHES REQUESTED
    #
    # /Type Silent
    # /DisableExternalServices
    # /Log "%WINDIR%\Temp\Fusion2026-Install.log"
    # ========================================================

    $InstallerArguments =
        '/Type Silent /DisableExternalServices /Log "%WINDIR%\Temp\Fusion2026-Install.log"'


    Write-Log "Fusion installer switches:"
    Write-Log $InstallerArguments


    $Process =
        Start-Process `
            -FilePath $FusionInstaller.FullName `
            -ArgumentList $InstallerArguments `
            -Wait `
            -PassThru `
            -WindowStyle Hidden


    Write-Log `
        "Fusion installer exit code: $($Process.ExitCode)"


    if ($Process.ExitCode -ne 0) {

        Write-Log `
            "Fusion installer log: $InstallerLog"

        Stop-WithError `
            "Fusion installation failed with exit code $($Process.ExitCode)."
    }

    Write-Log `
        "Fusion 2026 installation completed successfully."
}

catch {

    Stop-WithError `
        "Fusion installation failed: $($_.Exception.Message)"
}


# ============================================================
# WAIT FOR INSTALLER FINALIZATION
# ============================================================

Write-Log "Waiting for installation finalization..."

Start-Sleep -Seconds 5


# ============================================================
# FIND LICENSE INSIDE EXTRACTED ZIP
# ============================================================

Write-Log "============================================================"
Write-Log "Searching extracted package for license file..."
Write-Log "License filename: $LicenseFileName"
Write-Log "============================================================"


$LicenseFiles =
    Get-ChildItem `
        -Path $ExtractPath `
        -File `
        -Recurse `
        -ErrorAction SilentlyContinue |
    Where-Object {
        $_.Name -ieq $LicenseFileName
    }


if (-not $LicenseFiles) {

    Write-Log `
        "ERROR: $LicenseFileName was not found inside Fusion.zip."

    Write-Log "Searching for all .lic files..."

    $AllLicenses =
        Get-ChildItem `
            -Path $ExtractPath `
            -Filter "*.lic" `
            -File `
            -Recurse `
            -ErrorAction SilentlyContinue

    foreach ($Lic in $AllLicenses) {

        Write-Log `
            "Found .lic file: $($Lic.FullName)"
    }

    Stop-WithError `
        "License file '$LicenseFileName' was not found in the downloaded Fusion package."
}


# Use the first exact filename match.

$LicenseSource =
    $LicenseFiles |
    Select-Object -First 1


Write-Log "License file found:"
Write-Log $LicenseSource.FullName


# ============================================================
# CREATE AUTHORIZATION DIRECTORY
# ============================================================

Write-Log "Creating Authorization directory..."

if (-not (Test-Path $LicenseDestinationDirectory)) {

    New-Item `
        -Path $LicenseDestinationDirectory `
        -ItemType Directory `
        -Force |
        Out-Null
}

Write-Log "Authorization directory:"
Write-Log $LicenseDestinationDirectory


# ============================================================
# COPY LICENSE
# ============================================================

try {

    Write-Log "Copying license file..."

    Copy-Item `
        -LiteralPath $LicenseSource.FullName `
        -Destination $LicenseDestination `
        -Force

    Write-Log "License copy operation completed."
}

catch {

    Stop-WithError `
        "Failed to copy license file: $($_.Exception.Message)"
}


# ============================================================
# VERIFY LICENSE
# ============================================================

if (-not (Test-Path -LiteralPath $LicenseDestination)) {

    Stop-WithError `
        "License verification failed. File was not found at $LicenseDestination"
}


$InstalledLicense =
    Get-Item `
        -LiteralPath $LicenseDestination


Write-Log "============================================================"
Write-Log "LICENSE INSTALLATION SUCCESSFUL"
Write-Log "============================================================"
Write-Log "License: $($InstalledLicense.FullName)"
Write-Log "Size   : $($InstalledLicense.Length) bytes"


# ============================================================
# CLEANUP
# ============================================================

Write-Log "Cleaning temporary deployment files..."

try {

    Remove-Item `
        -Path $WorkRoot `
        -Recurse `
        -Force `
        -ErrorAction SilentlyContinue

    Write-Log "Temporary deployment files cleaned."
}

catch {

    Write-Log `
        "WARNING: Temporary deployment directory could not be completely removed."
}


# ============================================================
# FINAL SUCCESS
# ============================================================

Write-Log "============================================================"
Write-Log "FUSION 2026 DEPLOYMENT SUCCESSFUL"
Write-Log "============================================================"
Write-Log "Fusion installed successfully."
Write-Log "License installed successfully."
Write-Log "License location:"
Write-Log $LicenseDestination
Write-Log "Fusion installer log:"
Write-Log $InstallerLog
Write-Log "============================================================"

exit 0

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

$InstallLog = Join-Path $env:WINDIR "Temp\Fusion2026-Install.log"

$LicenseDestination = "C:\Program Files\Freedom Scientific\Authorization"

# GitHub location supplied for WinSCP components
$GitHubRawBase = (
    "https" + "://raw.githubusercontent.com" +
    "/FaronicsProServices/Windows-Modules-and-Packages" +
    "/main/SFTPInstallers"
)


# ============================================================
# LOGGING
# ============================================================

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

    Write-Log "ERROR: $Message"
    exit 1
}


# ============================================================
# ADMINISTRATOR CHECK
# ============================================================

try {

    $Identity = [Security.Principal.WindowsIdentity]::GetCurrent()

    $Principal = New-Object Security.Principal.WindowsPrincipal(
        $Identity
    )

    if (-not $Principal.IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator
    )) {

        Stop-WithError "The script must be executed as Administrator."
    }

}
catch {

    Stop-WithError "Unable to verify administrator privileges: $($_.Exception.Message)"
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

    # Converts:
    #
    # sftp://10.0.1.215:22//Share/Fusion.zip
    #
    # into:
    #
    # /Share/Fusion.zip

    $RemoteFile = $Uri.AbsolutePath -replace '^//', '/'

    Write-Log "SFTP Server : $SftpServer"
    Write-Log "SFTP Port   : $SftpPort"
    Write-Log "Remote File : $RemoteFile"

}
catch {

    Stop-WithError "Failed to parse SFTP URL: $($_.Exception.Message)"
}


# ============================================================
# CREATE WORKING DIRECTORIES
# ============================================================

try {

    Write-Log "Creating temporary working directories..."

    if (Test-Path $WorkDir) {

        Remove-Item `
            -Path $WorkDir `
            -Recurse `
            -Force `
            -ErrorAction SilentlyContinue
    }

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
# DOWNLOAD WINSCP COMPONENTS
# ============================================================

try {

    Write-Log "Downloading WinSCP components..."

    [Net.ServicePointManager]::SecurityProtocol = `
        [Net.SecurityProtocolType]::Tls12

    Invoke-WebRequest `
        -Uri "$GitHubRawBase/WinSCP.exe" `
        -OutFile $WinScpExe `
        -UseBasicParsing

    Invoke-WebRequest `
        -Uri "$GitHubRawBase/WinSCPnet.dll" `
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

    Stop-WithError `
        "Failed to download WinSCP components: $($_.Exception.Message)"
}


# ============================================================
# LOAD WINSCP .NET ASSEMBLY
# ============================================================

try {

    Write-Log "Loading WinSCP .NET assembly..."

    Add-Type -Path $WinScpDll

    Write-Log "WinSCP .NET assembly loaded successfully."

}
catch {

    Stop-WithError `
        "Failed to load WinSCPnet.dll: $($_.Exception.Message)"
}


# ============================================================
# CONNECT TO SFTP AND DOWNLOAD FUSION.ZIP
# ============================================================

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
    #
    # For production security, replacing this with the
    # server's specific SSH host-key fingerprint is recommended.

    $SessionOptions.GiveUpSecurityAndAcceptAnySshHostKey = $true

    $Session = New-Object WinSCP.Session

    $Session.Open($SessionOptions)

    Write-Log "SFTP connection established."

    Write-Log "Downloading Fusion package..."

    $TransferResult = $Session.GetFiles(
        $RemoteFile,
        $FusionZip,
        $false
    )

    $TransferResult.Check()

    if (-not (Test-Path $FusionZip)) {

        throw "Fusion.zip was not downloaded."
    }

    Write-Log "Fusion ZIP downloaded successfully."

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
# EXTRACT FUSION PACKAGE
# ============================================================

try {

    Write-Log "Extracting Fusion package..."

    Expand-Archive `
        -Path $FusionZip `
        -DestinationPath $ExtractDir `
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

    $Installers = Get-ChildItem `
        -Path $ExtractDir `
        -Filter "*.exe" `
        -Recurse `
        -File

    $FusionInstaller = $Installers |
        Where-Object {
            $_.Name -match "^F2026.*-Offline-x64\.exe$"
        } |
        Select-Object -First 1

    if (-not $FusionInstaller) {

        throw "Fusion 2026 Offline x64 installer was not found."
    }

    Write-Log "Installer found:"
    Write-Log $FusionInstaller.FullName

}
catch {

    Stop-WithError `
        "Failed to locate Fusion installer: $($_.Exception.Message)"
}


# ============================================================
# FUSION SILENT INSTALLATION
# ============================================================

try {

    Write-Log "Starting Fusion silent installation..."

    # Remove previous Fusion installation log if present.
    if (Test-Path $InstallLog) {

        Remove-Item `
            -Path $InstallLog `
            -Force `
            -ErrorAction SilentlyContinue
    }

    # ========================================================
    # ONLY INSTALLER SWITCHES REQUESTED
    #
    # /Type Silent
    # /DisableExternalServices
    # /Log "%WINDIR%\Temp\Fusion2026-Install.log"
    # ========================================================

    $InstallerArguments = @(
        "/Type"
        "Silent"

        "/DisableExternalServices"

        "/Log"
        "`"$InstallLog`""
    )

    Write-Log "Running Fusion installer with:"
    Write-Log "/Type Silent /DisableExternalServices /Log `"$InstallLog`""

    $Process = Start-Process `
        -FilePath $FusionInstaller.FullName `
        -ArgumentList $InstallerArguments `
        -Wait `
        -PassThru

    $ExitCode = $Process.ExitCode

    Write-Log "Fusion installer exit code: $ExitCode"

    if ($ExitCode -ne 0) {

        if (Test-Path $InstallLog) {

            Write-Log "Fusion installation log:"
            Write-Log $InstallLog
        }

        throw "Fusion installer failed with exit code $ExitCode."
    }

    Write-Log "Fusion installation completed successfully."

}
catch {

    Stop-WithError `
        "Fusion installation failed: $($_.Exception.Message)"
}


# ============================================================
# VERIFY LICENSE FILE
# ============================================================

try {

    Write-Log "Searching for license file: $LicenseFileName"

    $LicenseFile = Get-ChildItem `
        -Path $ExtractDir `
        -Filter $LicenseFileName `
        -Recurse `
        -File |
        Select-Object -First 1

    if (-not $LicenseFile) {

        throw `
            "License file '$LicenseFileName' was not found inside Fusion.zip."
    }

    Write-Log "License file found:"
    Write-Log $LicenseFile.FullName

}
catch {

    Stop-WithError `
        "Failed to locate license file: $($_.Exception.Message)"
}


# ============================================================
# COPY LICENSE FILE
# ============================================================

try {

    Write-Log "Creating license directory..."

    if (-not (Test-Path $LicenseDestination)) {

        New-Item `
            -Path $LicenseDestination `
            -ItemType Directory `
            -Force |
            Out-Null
    }

    $DestinationLicense = Join-Path `
        $LicenseDestination `
        $LicenseFileName

    Write-Log "Copying license file..."

    Copy-Item `
        -Path $LicenseFile.FullName `
        -Destination $DestinationLicense `
        -Force

    if (-not (Test-Path $DestinationLicense)) {

        throw "License file was not found after copying."
    }

    Write-Log "License file successfully copied to:"
    Write-Log $DestinationLicense

}
catch {

    Stop-WithError `
        "Failed to copy license file: $($_.Exception.Message)"
}


# ============================================================
# VERIFY INSTALLATION
# ============================================================

try {

    Write-Log "Verifying Fusion installation..."

    $FusionInstalled = $false

    $InstalledProducts = Get-ItemProperty `
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" ,
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" `
        -ErrorAction SilentlyContinue |
    Where-Object {
        $_.DisplayName -match "Fusion"
    }

    if ($InstalledProducts) {

        $FusionInstalled = $true

        Write-Log "Fusion installation detected."

        $InstalledProducts |
            Select-Object DisplayName, DisplayVersion |
            ForEach-Object {
                Write-Log "$($_.DisplayName) - $($_.DisplayVersion)"
            }
    }

    if (-not $FusionInstalled) {

        Write-Log "WARNING: Fusion was not detected through the uninstall registry."

        # Do not fail the deployment solely on this check because
        # the Fusion bootstrapper may register its components
        # differently depending on installation state.
    }

}
catch {

    Write-Log "WARNING: Fusion verification encountered an error."
}


# ============================================================
# VERIFY LICENSE
# ============================================================

if (Test-Path $DestinationLicense) {

    Write-Log "License verification successful."

}
else {

    Stop-WithError `
        "License verification failed."
}


# ============================================================
# CLEANUP
# ============================================================

try {

    Write-Log "Cleaning temporary deployment files..."

    if (Test-Path $WorkDir) {

        Remove-Item `
            -Path $WorkDir `
            -Recurse `
            -Force `
            -ErrorAction SilentlyContinue
    }

    Write-Log "Temporary files cleaned successfully."

}
catch {

    Write-Log "WARNING: Temporary files could not be completely removed."
}


# ============================================================
# SUCCESS
# ============================================================

Write-Log ""
Write-Log "============================================================"
Write-Log " Fusion 2026 deployment completed successfully."
Write-Log "============================================================"
Write-Log ""

exit 0

#Requires -Version 5.1

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

$DeploymentLog = Join-Path `
    $env:WINDIR `
    "Temp\Fusion2026-Deployment.log"

$FusionInstallLog = Join-Path `
    $env:WINDIR `
    "Temp\Fusion2026-Install.log"

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
        New-Object `
            Security.Principal.WindowsPrincipal(
                $CurrentIdentity
            )

    if (-not $Principal.IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator
    )) {

        Stop-WithError `
            "This script must be run as Administrator."
    }

    Write-Log "Administrator privileges confirmed."

}
catch {

    Stop-WithError `
        "Unable to verify administrator privileges: $($_.Exception.Message)"
}


# ============================================================
# CREATE WORKING DIRECTORIES
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

    # Converts:
    # sftp://10.0.1.215:22//Share/Fusion.zip
    #
    # to:
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
# DOWNLOAD WINSCP COMPONENTS
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
# LOAD WINSCP .NET ASSEMBLY
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
# CONNECT TO SFTP AND DOWNLOAD FUSION.ZIP
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

    # Accept any SSH host key.
    # This matches the configuration used during testing.
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
    # DOWNLOAD FUSION PACKAGE
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


# ============================================================
# EXTRACT FUSION ZIP
# ============================================================

try {

    Write-Log "Extracting Fusion package..."

    Expand-Archive `
        -LiteralPath $FusionZip `
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

    $FusionInstaller =
        Get-ChildItem `
            -Path $ExtractDir `
            -Filter "*.exe" `
            -File `
            -Recurse `
            -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -match "^F2026.*-Offline-x64\.exe$"
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
# INSTALL FUSION 2026
# ============================================================

try {

    Write-Log "Starting Fusion silent installation..."

    if (Test-Path -LiteralPath $FusionInstallLog) {

        Remove-Item `
            -Path $FusionInstallLog `
            -Force `
            -ErrorAction SilentlyContinue
    }

    # ONLY THE REQUESTED FUSION SWITCHES
    $InstallerArguments =
        '/Type Silent /DisableExternalServices /Log "%WINDIR%\Temp\Fusion2026-Install.log"'

    Write-Log "Installer switches:"
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
            "Fusion installer log: $FusionInstallLog"

        Stop-WithError `
            "Fusion installation failed with exit code $($Process.ExitCode)."
    }

    Write-Log "Fusion 2026 installation completed successfully."

}
catch {

    Stop-WithError `
        "Fusion installation failed: $($_.Exception.Message)"
}


# ============================================================
# WAIT FOR INSTALLATION FINALIZATION
# ============================================================

Write-Log "Waiting for installation finalization..."

Start-Sleep -Seconds 5


# ============================================================
# FIND LICENSE FILE
# ============================================================

try {

    Write-Log "Searching for license file: $LicenseFileName"

    $LicenseSource =
        Get-ChildItem `
            -Path $ExtractDir `
            -Filter $LicenseFileName `
            -File `
            -Recurse `
            -ErrorAction SilentlyContinue |
        Select-Object -First 1

    if (-not $LicenseSource) {

        Write-Log "License file was not found."

        $LicFiles =
            Get-ChildItem `
                -Path $ExtractDir `
                -Filter "*.lic" `
                -File `
                -Recurse `
                -ErrorAction SilentlyContinue

        foreach ($LicFile in $LicFiles) {

            Write-Log `
                "Found license candidate: $($LicFile.FullName)"
        }

        Stop-WithError `
            "License file '$LicenseFileName' was not found inside Fusion.zip."
    }

    Write-Log "License file found:"
    Write-Log $LicenseSource.FullName

}
catch {

    Stop-WithError `
        "Failed to locate license file: $($_.Exception.Message)"
}


# ============================================================
# COPY LICENSE FILE
# ============================================================

try {

    Write-Log "Creating Authorization directory..."

    New-Item `
        -Path $LicenseDestination `
        -ItemType Directory `
        -Force |
        Out-Null

    Write-Log "Copying license file..."

    # Same approach as the manager's known-working script.
    Copy-Item `
        -Path $LicenseSource.FullName `
        -Destination $LicenseDestination `
        -Force `
        -Recurse

    $InstalledLicense =
        Join-Path `
            $LicenseDestination `
            $LicenseFileName

    if (-not (Test-Path -LiteralPath $InstalledLicense)) {

        throw `
            "License file was not found at destination after copy."
    }

    $LicenseInfo =
        Get-Item `
            -LiteralPath $InstalledLicense

    Write-Log "License copied successfully."
    Write-Log "License path: $($LicenseInfo.FullName)"
    Write-Log "License size: $($LicenseInfo.Length) bytes"

}
catch {

    Stop-WithError `
        "Failed to copy license file: $($_.Exception.Message)"
}


# ============================================================
# CLEANUP
# ============================================================

try {

    Write-Log "Cleaning temporary deployment files..."

    if (Test-Path -LiteralPath $WorkDir) {

        Remove-Item `
            -Path $WorkDir `
            -Recurse `
            -Force `
            -ErrorAction SilentlyContinue
    }

    Write-Log "Temporary deployment files cleaned."

}
catch {

    Write-Log `
        "WARNING: Temporary deployment files could not be completely removed."
}


# ============================================================
# FINAL SUCCESS
# ============================================================

Write-Log ""
Write-Log "============================================================"
Write-Log "FUSION 2026 DEPLOYMENT SUCCESSFUL"
Write-Log "============================================================"
Write-Log "Fusion installation : SUCCESS"
Write-Log "License installation: SUCCESS"
Write-Log "License location    : $LicenseDestination\$LicenseFileName"
Write-Log "Fusion install log  : $FusionInstallLog"
Write-Log "Deployment log      : $DeploymentLog"
Write-Log "============================================================"

exit 0

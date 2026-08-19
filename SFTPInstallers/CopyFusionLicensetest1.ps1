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
Write-Log "Fusion 2026 SFTP License Deployment"
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


# ============================================================
# VERIFY ZIP
# ============================================================

try {

    if (-not (Test-Path -LiteralPath $FusionZip)) {

        throw "Fusion.zip was not found after download."
    }

    Write-Log "Downloaded Fusion.zip verified."

}
catch {

    Stop-WithError `
        "Fusion ZIP verification failed: $($_.Exception.Message)"
}


# ============================================================
# EXTRACT ZIP
# ============================================================

try {

    Write-Log "Extracting Fusion.zip..."

    Expand-Archive `
        -Path $FusionZip `
        -DestinationPath $ExtractDir `
        -Force

    Write-Log "Fusion.zip extraction completed."

}
catch {

    Stop-WithError `
        "Failed to extract Fusion.zip: $($_.Exception.Message)"
}


# ============================================================
# FIND LICENSE FILE
# ============================================================

try {

    Write-Log "Searching extracted files for $LicenseFileName..."

    $LicenseFile =
        Get-ChildItem `
            -Path $ExtractDir `
            -Filter $LicenseFileName `
            -File `
            -Recurse `
            -ErrorAction SilentlyContinue |
            Select-Object -First 1

    if (-not $LicenseFile) {

        throw `
            "$LicenseFileName was not found inside the extracted Fusion package."
    }

    Write-Log "License file found:"
    Write-Log $LicenseFile.FullName

}
catch {

    Stop-WithError `
        "Failed to locate license file: $($_.Exception.Message)"
}


# ============================================================
# CREATE LICENSE DESTINATION
# ============================================================

try {

    if (-not (Test-Path -LiteralPath $LicenseDestination)) {

        Write-Log "Creating license destination folder..."

        New-Item `
            -Path $LicenseDestination `
            -ItemType Directory `
            -Force |
            Out-Null
    }

    Write-Log "License destination confirmed:"
    Write-Log $LicenseDestination

}
catch {

    Stop-WithError `
        "Failed to create license destination: $($_.Exception.Message)"
}


# ============================================================
# COPY LICENSE FILE
# ============================================================

try {

    Write-Log "Copying $LicenseFileName..."

    Push-Location $LicenseFile.DirectoryName

    try {

        Copy-Item -Path ".\Fusion.lic" -Destination "C:\Program Files\Freedom Scientific\Authorization" -Force -Recurse

    }
    finally {

        Pop-Location
    }

    Write-Log "License file copied successfully."

}
catch {

    Stop-WithError `
        "Failed to copy license file: $($_.Exception.Message)"
}


# ============================================================
# VERIFY LICENSE
# ============================================================

try {

    $InstalledLicense =
        Join-Path `
            $LicenseDestination `
            $LicenseFileName

    if (-not (Test-Path -LiteralPath $InstalledLicense)) {

        throw `
            "License verification failed. File not found at $InstalledLicense"
    }

    Write-Log "License installation verified."
    Write-Log "Installed license: $InstalledLicense"

}
catch {

    Stop-WithError `
        "License verification failed: $($_.Exception.Message)"
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
            -Force
    }

    Write-Log "Temporary files cleaned successfully."

}
catch {

    Write-Log `
        "WARNING: Cleanup failed: $($_.Exception.Message)"
}


# ============================================================
# COMPLETE
# ============================================================

Write-Log "============================================================"
Write-Log "Fusion 2026 SFTP License Deployment Completed Successfully"
Write-Log "============================================================"

exit 0

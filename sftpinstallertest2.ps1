[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SftpLink,

    [Parameter(Mandatory = $true)]
    [string]$SftpUser,

    [Parameter(Mandatory = $true)]
    [string]$SftpPassword,

    [Parameter(Mandatory = $false)]
    [int]$SftpPort = 22,

    [Parameter(Mandatory = $false)]
    [string]$LicenseFileName = "Fusion.lic",

    [Parameter(Mandatory = $false)]
    [string]$SshHostKeyFingerprint = "*"
)

# Parse the SFTP link (supports domain names, IP addresses, and optional inline ports)
try {
    $uri = [System.Uri]$SftpLink
    $SftpServer = $uri.Host
    $RemoteZipPath = $uri.AbsolutePath -replace '^//','/'
    
    # Automatically extract port from URL if specified
    if ($uri.Port -and $uri.Port -ne -1) {
        $SftpPort = $uri.Port
    }
}
catch {
    Write-Error "Invalid SFTP link format provided: $_"
    exit 1
}

# Define local staging paths
$workDir = "C:\Temp\FusionDeployment"
$localZipPath = Join-Path $workDir "FusionPackage.zip"
$extractPath = Join-Path $workDir "ExtractedFiles"
$licenseDestDir = "C:\Program Files\Freedom Scientific\Authorization"
$winscpFolder = Join-Path $workDir "WinSCP"

if (-not (Test-Path $workDir)) {
    New-Item -ItemType Directory -Path $workDir -Force | Out-Null
}

# Automatically download WinSCP Portable so the script requires no external DLL packaging
$winscpDllPath = Join-Path $winscpFolder "WinSCPnet.dll"
if (-not (Test-Path $winscpDllPath)) {
    Write-Host "Downloading SFTP client libraries..."
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $winscpZip = Join-Path $workDir "WinSCP.zip"
    Invoke-WebRequest -Uri "https://winscp.net/download/WinSCP-6.3.2-Portable.zip" -OutFile $winscpZip
    Expand-Archive -Path $winscpZip -DestinationPath $winscpFolder -Force
    Remove-Item $winscpZip -Force
}

# 1. Download the ZIP file from SFTP
try {
    Write-Host "Connecting to SFTP server ($SftpServer on port $SftpPort)..."
    Add-Type -Path "$winscpFolder\WinSCPnet.dll"
    
    $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol = [WinSCP.Protocol]::Sftp
        HostName = $SftpServer
        PortNumber = $SftpPort
        UserName = $SftpUser
        Password = $SftpPassword
        SshHostKeyFingerprint = $SshHostKeyFingerprint
    }

    $session = New-Object WinSCP.Session
    try {
        $session.Open($sessionOptions)
        $transferResult = $session.GetFiles($RemoteZipPath, $localZipPath)
        $transferResult.Check()
        Write-Host "ZIP file successfully downloaded from SFTP."
    }
    finally {
        $session.Dispose()
    }
}
catch {
    Write-Error "Failed to download ZIP file via SFTP: $_"
    exit 1
}

# 2. Extract the ZIP file
try {
    Write-Host "Extracting ZIP file..."
    if (Test-Path $extractPath) {
        Remove-Item $extractPath -Recurse -Force
    }
    Expand-Archive -Path $localZipPath -DestinationPath $extractPath -Force
    Write-Host "Extraction completed successfully."
}
catch {
    Write-Error "Failed to extract ZIP package: $_"
    exit 1
}

# 3. Locate and install Fusion 2026 silently
try {
    Write-Host "Searching for Fusion .exe installer..."
    $exeFile = Get-ChildItem -Path $extractPath -Filter "*.exe" -Recurse | Select-Object -First 1

    if (-not $exeFile) {
        throw "No executable (.exe) file found inside the extracted package."
    }

    Write-Host "Installing $($exeFile.Name) silently..."
    Start-Process -FilePath $exeFile.FullName -ArgumentList "/type Silent /DisableExternalServices" -Wait
    Write-Host "Installation completed."
}
catch {
    Write-Error "Error during installation: $_"
    exit 1
}

# 4. Locate and copy the license file to the authorization directory
try {
    Write-Host "Locating license file ($LicenseFileName)..."
    $licFile = Get-ChildItem -Path $extractPath -Filter $LicenseFileName -Recurse | Select-Object -First 1

    if (-not $licFile) {
        throw "License file '$LicenseFileName' not found inside the extracted package."
    }

    if (-not (Test-Path $licenseDestDir)) {
        New-Item -ItemType Directory -Path $licenseDestDir -Force | Out-Null
    }

    Copy-Item -Path $licFile.FullName -Destination $licenseDestDir -Force
    Write-Host "License file successfully copied to $licenseDestDir."
}
catch {
    Write-Error "Failed to copy license file: $_"
    exit 1
}

# Cleanup staging files
Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Deployment workflow finished successfully."

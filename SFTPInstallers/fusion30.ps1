# ============================================================
# Freedom Scientific Fusion - RMM Deployment
# ============================================================

param (
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
# Configuration
# ============================================================

$WorkDir = "C:\Windows\Temp\FusionInstall"

$WinSCPDir = "$WorkDir\WinSCP"
$WinSCPExe = "$WinSCPDir\WinSCP.exe"
$WinSCPDll = "$WinSCPDir\WinSCPnet.dll"

$ZipFile = "$WorkDir\Fusion.zip"
$ExtractDir = "$WorkDir\Extracted"

$LicenseDestination = "C:\Program Files\Freedom Scientific\Authorization"

# GitHub
$GitHubWinSCPExe = "https://raw.githubusercontent.com/FaronicsProServices/Windows-Modules-and-Packages/main/SFTPInstallers/WinSCP.exe"

$GitHubWinSCPDll = "https://raw.githubusercontent.com/FaronicsProServices/Windows-Modules-and-Packages/main/SFTPInstallers/WinSCPnet.dll"

# ============================================================
# Main
# ============================================================

try {

    Write-Host "============================================"
    Write-Host " Freedom Scientific Fusion Deployment"
    Write-Host "============================================"
    Write-Host ""

    # --------------------------------------------------------
    # Validate parameters
    # --------------------------------------------------------

    if ([string]::IsNullOrWhiteSpace($SftpLink)) {
        throw "SftpLink was not supplied."
    }

    if ([string]::IsNullOrWhiteSpace($SftpUser)) {
        throw "SftpUser was not supplied."
    }

    if ([string]::IsNullOrWhiteSpace($SftpPassword)) {
        throw "SftpPassword was not supplied."
    }

    if ([string]::IsNullOrWhiteSpace($LicenseFileName)) {
        throw "LicenseFileName was not supplied."
    }

    # --------------------------------------------------------
    # Parse SFTP URL
    # --------------------------------------------------------

    Write-Host "Processing SFTP link..."

    $SftpUri = [System.Uri]$SftpLink

    if ($SftpUri.Scheme -ne "sftp") {
        throw "SftpLink must use the sftp:// protocol."
    }

    $SftpHost = $SftpUri.Host
    $SftpPort = $SftpUri.Port

    if ($SftpPort -eq -1) {
        $SftpPort = 22
    }

    $RemoteFile = $SftpUri.AbsolutePath

    # Remove duplicate leading slash
    $RemoteFile = "/" + $RemoteFile.TrimStart("/")

    Write-Host "SFTP Host   : $SftpHost"
    Write-Host "SFTP Port   : $SftpPort"
    Write-Host "Remote File : $RemoteFile"
    Write-Host "License     : $LicenseFileName"
    Write-Host ""

    # --------------------------------------------------------
    # Prepare working directory
    # --------------------------------------------------------

    Write-Host "Preparing working directory..."

    if (Test-Path $WorkDir) {
        Remove-Item `
            -Path $WorkDir `
            -Recurse `
            -Force
    }

    New-Item `
        -Path $WorkDir `
        -ItemType Directory `
        -Force | Out-Null

    New-Item `
        -Path $WinSCPDir `
        -ItemType Directory `
        -Force | Out-Null

    New-Item `
        -Path $ExtractDir `
        -ItemType Directory `
        -Force | Out-Null

    Write-Host "Working directory ready."
    Write-Host ""

    # --------------------------------------------------------
    # Download WinSCP.exe
    # --------------------------------------------------------

    Write-Host "Downloading WinSCP.exe from GitHub..."

    Invoke-WebRequest `
        -Uri $GitHubWinSCPExe `
        -OutFile $WinSCPExe `
        -UseBasicParsing

    if (-not (Test-Path $WinSCPExe)) {
        throw "WinSCP.exe download failed."
    }

    Write-Host "WinSCP.exe downloaded successfully."
    Write-Host ""

    # --------------------------------------------------------
    # Download WinSCPnet.dll
    # --------------------------------------------------------

    Write-Host "Downloading WinSCPnet.dll from GitHub..."

    Invoke-WebRequest `
        -Uri $GitHubWinSCPDll `
        -OutFile $WinSCPDll `
        -UseBasicParsing

    if (-not (Test-Path $WinSCPDll)) {
        throw "WinSCPnet.dll download failed."
    }

    Write-Host "WinSCPnet.dll downloaded successfully."
    Write-Host ""

    # --------------------------------------------------------
    # Load WinSCP
    # --------------------------------------------------------

    Write-Host "Loading WinSCP .NET assembly..."

    Add-Type -Path $WinSCPDll

    Write-Host "WinSCP loaded successfully."
    Write-Host ""

    # --------------------------------------------------------
    # Connect to SFTP
    # --------------------------------------------------------

    Write-Host "Connecting to SFTP..."

    $SessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol = [WinSCP.Protocol]::Sftp
        HostName = $SftpHost
        PortNumber = $SftpPort
        UserName = $SftpUser
        Password = $SftpPassword

        # Internal environment
        GiveUpSecurityAndAcceptAnySshHostKey = $true

        # Use downloaded WinSCP.exe
        ExecutablePath = $WinSCPExe
    }

    $Session = New-Object WinSCP.Session

    try {

        $Session.Open($SessionOptions)

        Write-Host "SFTP connection successful."
        Write-Host ""

        # ----------------------------------------------------
        # Download Fusion ZIP
        # ----------------------------------------------------

        Write-Host "Downloading Fusion package..."

        $TransferResult = $Session.GetFiles(
            $RemoteFile,
            $ZipFile
        )

        $TransferResult.Check()

        Write-Host "Fusion package downloaded successfully."
        Write-Host ""

    }
    finally {

        $Session.Dispose()

    }

    # --------------------------------------------------------
    # Verify ZIP
    # --------------------------------------------------------

    if (-not (Test-Path $ZipFile)) {
        throw "Fusion ZIP was not found after download."
    }

    Write-Host "Fusion ZIP verified."
    Write-Host ""

    # --------------------------------------------------------
    # Extract ZIP
    # --------------------------------------------------------

    Write-Host "Extracting Fusion package..."

    Expand-Archive `
        -Path $ZipFile `
        -DestinationPath $ExtractDir `
        -Force

    Write-Host "Extraction completed."
    Write-Host ""

    # --------------------------------------------------------
    # Find ANY EXE
    # --------------------------------------------------------

    Write-Host "Scanning extracted package for EXE..."

    $Executables = Get-ChildItem `
        -Path $ExtractDir `
        -Filter "*.exe" `
        -Recurse `
        -File

    if (-not $Executables) {
        throw "No EXE file was found in the extracted package."
    }

    Write-Host ""
    Write-Host "EXE file(s) found:"

    foreach ($Exe in $Executables) {
        Write-Host "  $($Exe.FullName)"
    }

    # Select first EXE
    $Installer = $Executables | Select-Object -First 1

    Write-Host ""
    Write-Host "Installer selected:"
    Write-Host $Installer.FullName
    Write-Host ""

    # --------------------------------------------------------
    # Find License
    # --------------------------------------------------------

    Write-Host "Searching for license: $LicenseFileName"

    $License = Get-ChildItem `
        -Path $ExtractDir `
        -Filter $LicenseFileName `
        -Recurse `
        -File |
        Select-Object -First 1

    if (-not $License) {
        throw "$LicenseFileName was not found in the extracted package."
    }

    Write-Host "License found:"
    Write-Host $License.FullName
    Write-Host ""

    # --------------------------------------------------------
    # Install Fusion
    # --------------------------------------------------------

    Write-Host "============================================"
    Write-Host " Installing Freedom Scientific Fusion"
    Write-Host "============================================"
    Write-Host ""

    Set-Location $Installer.DirectoryName

    # EXACT USER-PROVIDED INSTALL COMMAND
    Start-Process `
        -FilePath $Installer.FullName `
        -ArgumentList "/type Silent /DisableExternalServices" `
        -Wait

    Write-Host ""
    Write-Host "Fusion installer process completed."
    Write-Host ""

    # --------------------------------------------------------
    # Copy License
    # --------------------------------------------------------

    Write-Host "Copying license..."

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

    Write-Host "License copied successfully."
    Write-Host ""

    # --------------------------------------------------------
    # Verify License
    # --------------------------------------------------------

    $InstalledLicense = Join-Path `
        $LicenseDestination `
        $LicenseFileName

    if (-not (Test-Path $InstalledLicense)) {
        throw "License verification failed: $InstalledLicense"
    }

    Write-Host "License verification successful."
    Write-Host ""

    # --------------------------------------------------------
    # Cleanup
    # --------------------------------------------------------

    Write-Host "Cleaning up temporary files..."

    Set-Location C:\

    Remove-Item `
        -Path $WorkDir `
        -Recurse `
        -Force

    Write-Host "Temporary files removed."
    Write-Host ""

    # --------------------------------------------------------
    # SUCCESS
    # --------------------------------------------------------

    Write-Host "============================================"
    Write-Host " FUSION DEPLOYMENT SUCCESSFUL"
    Write-Host "============================================"
    Write-Host "Fusion installed."
    Write-Host "$LicenseFileName copied and verified."
    Write-Host "Temporary files cleaned."
    Write-Host "============================================"

    exit 0

}
catch {

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Red
    Write-Host " FUSION DEPLOYMENT FAILED" -ForegroundColor Red
    Write-Host "============================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Error:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Red

    exit 1
}

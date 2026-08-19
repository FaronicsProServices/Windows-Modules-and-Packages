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

try {

    # ========================================================
    # Configuration
    # ========================================================

    $WorkDir = "C:\Windows\Temp\FusionInstall"

    $WinSCPDir = "$WorkDir\WinSCP"
    $WinSCPExe = "$WinSCPDir\WinSCP.exe"
    $WinSCPDll = "$WinSCPDir\WinSCPnet.dll"

    $ZipFile = "$WorkDir\Fusion.zip"
    $ExtractDir = "$WorkDir\Extracted"

    $LicenseDestination = "C:\Program Files\Freedom Scientific\Authorization"

    # GitHub locations
    $GitHubWinSCPExe = "https://raw.githubusercontent.com/FaronicsProServices/Windows-Modules-and-Packages/main/SFTPInstallers/WinSCP.exe"

    $GitHubWinSCPDll = "https://raw.githubusercontent.com/FaronicsProServices/Windows-Modules-and-Packages/main/SFTPInstallers/WinSCPnet.dll"


    # ========================================================
    # Header
    # ========================================================

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host " Freedom Scientific Fusion Deployment" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ""


    # ========================================================
    # Process SFTP Link
    # ========================================================

    Write-Host "Processing SFTP link..." -ForegroundColor Cyan

    $SftpUri = [System.Uri]$SftpLink

    if ($SftpUri.Scheme -ne "sftp") {
        throw "SftpLink must use the sftp:// protocol."
    }

    $SftpHost = $SftpUri.Host

    if ($SftpUri.Port -eq -1) {
        $SftpPort = 22
    }
    else {
        $SftpPort = $SftpUri.Port
    }

    $RemoteFile = "/" + $SftpUri.AbsolutePath.TrimStart("/")

    Write-Host "SFTP Host   : $SftpHost"
    Write-Host "SFTP Port   : $SftpPort"
    Write-Host "Remote File : $RemoteFile"
    Write-Host "License     : $LicenseFileName"
    Write-Host ""


    # ========================================================
    # Prepare Working Directory
    # ========================================================

    Write-Host "Preparing working directory..." -ForegroundColor Cyan

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

    Write-Host "Working directory ready." -ForegroundColor Green
    Write-Host ""


    # ========================================================
    # Download WinSCP.exe from GitHub
    # ========================================================

    Write-Host "Downloading WinSCP.exe from GitHub..." -ForegroundColor Cyan

    Invoke-WebRequest `
        -Uri $GitHubWinSCPExe `
        -OutFile $WinSCPExe `
        -UseBasicParsing

    if (-not (Test-Path $WinSCPExe)) {
        throw "WinSCP.exe download failed."
    }

    Write-Host "WinSCP.exe downloaded successfully." -ForegroundColor Green
    Write-Host ""


    # ========================================================
    # Download WinSCPnet.dll from GitHub
    # ========================================================

    Write-Host "Downloading WinSCPnet.dll from GitHub..." -ForegroundColor Cyan

    Invoke-WebRequest `
        -Uri $GitHubWinSCPDll `
        -OutFile $WinSCPDll `
        -UseBasicParsing

    if (-not (Test-Path $WinSCPDll)) {
        throw "WinSCPnet.dll download failed."
    }

    Write-Host "WinSCPnet.dll downloaded successfully." -ForegroundColor Green
    Write-Host ""


    # ========================================================
    # Load WinSCP .NET Assembly
    # ========================================================

    Write-Host "Loading WinSCP .NET assembly..." -ForegroundColor Cyan

    Add-Type -Path $WinSCPDll

    Write-Host "WinSCP loaded successfully." -ForegroundColor Green
    Write-Host ""


    # ========================================================
    # Connect to SFTP
    # ========================================================

    Write-Host "Connecting to SFTP..." -ForegroundColor Cyan

    $SessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol = [WinSCP.Protocol]::Sftp
        HostName = $SftpHost
        PortNumber = $SftpPort
        UserName = $SftpUser
        Password = $SftpPassword

        # Internal TEST environment
        GiveUpSecurityAndAcceptAnySshHostKey = $true
    }

    $Session = New-Object WinSCP.Session

    # IMPORTANT:
    # WinSCP.exe is downloaded from GitHub and used from the
    # temporary directory. It is NOT required to be installed.

    $Session.ExecutablePath = $WinSCPExe

    try {

        $Session.Open($SessionOptions)

        Write-Host "SFTP connection successful." -ForegroundColor Green
        Write-Host ""


        # ====================================================
        # Download Fusion.zip
        # ====================================================

        Write-Host "Downloading Fusion.zip..." -ForegroundColor Cyan

        $TransferResult = $Session.GetFiles(
            $RemoteFile,
            $ZipFile
        )

        $TransferResult.Check()

        Write-Host "Fusion.zip downloaded successfully." -ForegroundColor Green
        Write-Host ""

    }
    finally {

        $Session.Dispose()

    }


    # ========================================================
    # Verify ZIP
    # ========================================================

    if (-not (Test-Path $ZipFile)) {
        throw "Fusion.zip was not found after download."
    }

    Write-Host "Downloaded file verified." -ForegroundColor Green
    Write-Host ""


    # ========================================================
    # Extract ZIP
    # ========================================================

    Write-Host "Extracting Fusion.zip..." -ForegroundColor Cyan

    Expand-Archive `
        -Path $ZipFile `
        -DestinationPath $ExtractDir `
        -Force

    Write-Host "Extraction completed." -ForegroundColor Green
    Write-Host ""


    # ========================================================
    # Find ANY EXE Installer
    # ========================================================

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
    Write-Host ""


    # ========================================================
    # Find License
    # ========================================================

    Write-Host "Searching for $LicenseFileName..." -ForegroundColor Cyan

    $License = Get-ChildItem `
        -Path $ExtractDir `
        -Filter $LicenseFileName `
        -Recurse `
        -File |
        Select-Object -First 1

    if (-not $License) {
        throw "$LicenseFileName was not found in the extracted package."
    }

    Write-Host "License found:" -ForegroundColor Green
    Write-Host $License.FullName -ForegroundColor Green
    Write-Host ""


    # ========================================================
    # Install Freedom Scientific Fusion
    # ========================================================

    Write-Host "============================================" -ForegroundColor Yellow
    Write-Host "Installing Freedom Scientific Fusion..." -ForegroundColor Yellow
    Write-Host "============================================" -ForegroundColor Yellow
    Write-Host ""

    Set-Location $Installer.DirectoryName


    # ========================================================
    # YOUR EXACT INSTALL COMMAND
    # ========================================================

    Start-Process `
        -FilePath $Installer.FullName `
        -ArgumentList "/type Silent /DisableExternalServices" `
        -Wait


    Write-Host "Fusion installer process completed." -ForegroundColor Green
    Write-Host ""


    # ========================================================
    # Copy License
    # ========================================================

    Write-Host "Copying Fusion license..." -ForegroundColor Yellow

    if (-not (Test-Path $LicenseDestination)) {

        New-Item `
            -Path $LicenseDestination `
            -ItemType Directory `
            -Force | Out-Null
    }


    # YOUR ORIGINAL LICENSE COPY METHOD
    Copy-Item `
        -Path $License.FullName `
        -Destination $LicenseDestination `
        -Force `
        -Recurse


    Write-Host "License copied successfully." -ForegroundColor Green
    Write-Host ""


    # ========================================================
    # Verify License
    # ========================================================

    $InstalledLicense = Join-Path `
        $LicenseDestination `
        $LicenseFileName

    if (-not (Test-Path $InstalledLicense)) {
        throw "$LicenseFileName verification failed."
    }

    Write-Host "License verification successful." -ForegroundColor Green
    Write-Host ""


    # ========================================================
    # Cleanup
    # ========================================================

    Write-Host "Cleaning up temporary files..." -ForegroundColor Cyan

    Set-Location C:\

    Remove-Item `
        -Path $WorkDir `
        -Recurse `
        -Force

    Write-Host "Temporary files removed." -ForegroundColor Green
    Write-Host ""


    # ========================================================
    # SUCCESS
    # ========================================================

    Write-Host "============================================" -ForegroundColor Green
    Write-Host " FUSION DEPLOYMENT SUCCESSFUL" -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "Fusion installed."
    Write-Host "$LicenseFileName copied and verified."
    Write-Host "Temporary files cleaned."
    Write-Host "============================================" -ForegroundColor Green

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

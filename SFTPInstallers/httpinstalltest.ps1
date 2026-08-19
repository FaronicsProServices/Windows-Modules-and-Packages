param (
    [Parameter(Mandatory = $true)]
    [string]$ZipUrl
)

$ErrorActionPreference = "Stop"

$LogFile = "C:\Temp\Fusion_Install.log"
$WorkDir = "C:\Temp\Fusion_Install"
$ZipFile = Join-Path $WorkDir "Fusion.zip"
$ExtractDir = Join-Path $WorkDir "Extracted"

function Write-Log {
    param ([string]$Message)

    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$Timestamp $Message" | Out-File -FilePath $LogFile -Append -Encoding UTF8
}

try {
    Write-Log "============================================================"
    Write-Log "Fusion 2026 Installation Started"
    Write-Log "============================================================"
    Write-Log "ZIP URL: $ZipUrl"

    # Create working directories
    if (Test-Path $WorkDir) {
        Remove-Item $WorkDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    New-Item -Path $WorkDir -ItemType Directory -Force | Out-Null
    New-Item -Path $ExtractDir -ItemType Directory -Force | Out-Null

    # Download ZIP
    Write-Log "Downloading Fusion ZIP..."

    Invoke-WebRequest `
        -Uri $ZipUrl `
        -OutFile $ZipFile `
        -UseBasicParsing

    if (-not (Test-Path $ZipFile)) {
        throw "Fusion ZIP download failed."
    }

    Write-Log "ZIP downloaded successfully."

    # Extract ZIP
    Write-Log "Extracting Fusion ZIP..."

    Expand-Archive `
        -Path $ZipFile `
        -DestinationPath $ExtractDir `
        -Force

    Write-Log "Extraction completed."

    # Find EXE installer
    Write-Log "Scanning extracted files for EXE installer..."

    $Installer = Get-ChildItem `
        -Path $ExtractDir `
        -Filter "*.exe" `
        -File `
        -Recurse |
        Select-Object -First 1

    if (-not $Installer) {
        throw "No EXE installer was found inside the extracted Fusion ZIP."
    }

    Write-Log "Installer found: $($Installer.FullName)"

    # Install Fusion
    Write-Log "Starting Fusion installation..."

    Start-Process `
        -FilePath $Installer.FullName `
        -ArgumentList "/type Silent /DisableExternalServices" `
        -Wait `
        -NoNewWindow

    Write-Log "Fusion installer completed."

    # Find Fusion.lic
    Write-Log "Searching for Fusion.lic..."

    $LicenseFile = Get-ChildItem `
        -Path $ExtractDir `
        -Filter "Fusion.lic" `
        -File `
        -Recurse |
        Select-Object -First 1

    if (-not $LicenseFile) {
        throw "Fusion.lic was not found inside the extracted ZIP."
    }

    Write-Log "License file found: $($LicenseFile.FullName)"

    # Destination
    $LicenseDestination = "C:\Program Files\Freedom Scientific\Authorization"

    if (-not (Test-Path $LicenseDestination)) {
        Write-Log "Authorization folder does not exist. Creating it..."

        New-Item `
            -Path $LicenseDestination `
            -ItemType Directory `
            -Force | Out-Null
    }

    # Copy license
    Write-Log "Copying Fusion.lic..."

    Copy-Item `
        -Path $LicenseFile.FullName `
        -Destination $LicenseDestination `
        -Force `
        -Recurse

    Write-Log "Fusion.lic copied successfully."

    Write-Log "============================================================"
    Write-Log "Fusion installation completed successfully."
    Write-Log "============================================================"

    # Cleanup
    Remove-Item $ZipFile -Force -ErrorAction SilentlyContinue
    Remove-Item $ExtractDir -Recurse -Force -ErrorAction SilentlyContinue

    exit 0
}
catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    Write-Log "Fusion installation FAILED."

    exit 1
}

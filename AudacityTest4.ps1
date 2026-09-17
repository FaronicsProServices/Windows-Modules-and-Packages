$ErrorActionPreference = "Stop"

# ============================================================
# Audacity 3.x -> Audacity 4.0.0 Upgrade
# RMM / Deep Freeze Cloud
# ============================================================

$AudacityURL = "https://github.com/audacity/audacity/releases/download/Audacity-4.0.0/audacity-win-4.0.0-x86_64.msi"

$WorkDir = "C:\Windows\Temp\AudacityUpgrade"
$MsiPath = "$WorkDir\Audacity-4.0.0.msi"
$LogPath = "$WorkDir\Audacity-4.0.0-MSI.log"

$Audacity4Exe = "C:\Program Files\Audacity 4\bin\Audacity4.exe"

$RegistryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

Write-Output "=========================================="
Write-Output "Audacity 4.0.0 Upgrade"
Write-Output "=========================================="

# ============================================================
# STOP AUDACITY
# ============================================================

Write-Output "Stopping Audacity if running..."

Get-Process -Name "Audacity","Audacity4" -ErrorAction SilentlyContinue |
    Stop-Process -Force -ErrorAction SilentlyContinue

Start-Sleep -Seconds 3

# ============================================================
# CHECK FOR AUDACITY 4.x
# ============================================================

Write-Output "Checking for existing Audacity 4.x..."

$Audacity4Apps = @(
    foreach ($Path in $RegistryPaths) {

        Get-ItemProperty $Path -ErrorAction SilentlyContinue |
            Where-Object {
                $_.DisplayName -like "Audacity*" -and
                $_.DisplayVersion -match "^4\."
            }
    }
)

# If Audacity 4 executable already exists, verify it and exit
# ============================================================

if (Test-Path $Audacity4Exe) {

    $ExistingVersion = (Get-Item $Audacity4Exe).VersionInfo.ProductVersion

    Write-Output "Audacity 4 executable already exists."
    Write-Output "Path: $Audacity4Exe"
    Write-Output "Version: $ExistingVersion"

    if ($ExistingVersion -like "4.0.*") {

        Write-Output "Audacity 4.0.x is already installed."
        Write-Output "Nothing to do."

        exit 0
    }
}

# ============================================================
# REMOVE ONLY AUDACITY 3.x
# ============================================================

Write-Output "Searching for Audacity 3.x..."

$Audacity3Apps = @(
    foreach ($Path in $RegistryPaths) {

        Get-ItemProperty $Path -ErrorAction SilentlyContinue |
            Where-Object {
                $_.DisplayName -like "Audacity*" -and
                $_.DisplayVersion -match "^3\."
            }
    }
)

if ($Audacity3Apps.Count -gt 0) {

    foreach ($App in $Audacity3Apps) {

        Write-Output "------------------------------------------"
        Write-Output "Found old Audacity installation"
        Write-Output "Name: $($App.DisplayName)"
        Write-Output "Version: $($App.DisplayVersion)"
        Write-Output "Install Location: $($App.InstallLocation)"
        Write-Output "------------------------------------------"

        # Audacity 3.x uses Inno Setup
        if ($App.UninstallString -match '"([^"]+unins[^"]+\.exe)"') {

            $Uninstaller = $Matches[1]

            Write-Output "Uninstaller: $Uninstaller"
            Write-Output "Removing Audacity 3.x..."

            $UninstallProcess = Start-Process `
                -FilePath $Uninstaller `
                -ArgumentList "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART" `
                -Wait `
                -PassThru

            Write-Output "Uninstall exit code: $($UninstallProcess.ExitCode)"

            if ($UninstallProcess.ExitCode -ne 0) {

                Write-Error "Audacity 3.x uninstall failed."
                exit $UninstallProcess.ExitCode
            }

            Write-Output "Audacity 3.x removed successfully."
        }
        else {

            Write-Error "Could not locate the Audacity 3.x uninstaller."
            exit 1
        }
    }

}
else {

    Write-Output "Audacity 3.x was not found."
}

# ============================================================
# WAIT FOR UNINSTALL TO FINISH
# ============================================================

Write-Output "Waiting for uninstall to complete..."

Start-Sleep -Seconds 5

# ============================================================
# CREATE WORKING DIRECTORY
# ============================================================

Write-Output "Creating working directory..."

New-Item `
    -Path $WorkDir `
    -ItemType Directory `
    -Force |
    Out-Null

# ============================================================
# DOWNLOAD AUDACITY 4.0.0 MSI
# ============================================================

Write-Output "Downloading Audacity 4.0.0..."

if (Test-Path $MsiPath) {

    Remove-Item $MsiPath -Force
}

Invoke-WebRequest `
    -Uri $AudacityURL `
    -OutFile $MsiPath `
    -UseBasicParsing

if (-not (Test-Path $MsiPath)) {

    Write-Error "Audacity 4.0.0 MSI download failed."
    exit 1
}

$MsiSize = (Get-Item $MsiPath).Length

Write-Output "MSI downloaded successfully."
Write-Output "MSI size: $MsiSize bytes"

# Basic file-size validation
if ($MsiSize -lt 40000000) {

    Write-Error "Downloaded MSI appears to be incomplete or invalid."
    exit 1
}

# ============================================================
# INSTALL AUDACITY 4.0.0
# ============================================================

Write-Output "Installing Audacity 4.0.0..."

$InstallArguments = "/i `"$MsiPath`" /quiet /norestart ALLUSERS=1 /L*v `"$LogPath`""

Write-Output "Running MSI installation..."

$InstallProcess = Start-Process `
    -FilePath "msiexec.exe" `
    -ArgumentList $InstallArguments `
    -Wait `
    -PassThru

$MsiExitCode = $InstallProcess.ExitCode

Write-Output "MSI exit code: $MsiExitCode"

# ============================================================
# HANDLE MSI RESULT
# ============================================================

if (($MsiExitCode -ne 0) -and ($MsiExitCode -ne 3010)) {

    Write-Error "Audacity 4.0.0 installation failed."
    Write-Error "MSI exit code: $MsiExitCode"
    Write-Error "MSI log: $LogPath"

    exit $MsiExitCode
}

if ($MsiExitCode -eq 3010) {

    Write-Output "MSI installation completed successfully."
    Write-Output "A reboot may be required."
}

# ============================================================
# VERIFY AUDACITY 4.0.0
# ============================================================

Write-Output "Verifying Audacity 4.0.0 installation..."

Start-Sleep -Seconds 5

if (-not (Test-Path $Audacity4Exe)) {

    Write-Error "Audacity 4 executable was not found."
    Write-Error "Expected path:"
    Write-Error $Audacity4Exe
    Write-Error "MSI log:"
    Write-Error $LogPath

    exit 1
}

# ============================================================
# GET INSTALLED VERSION
# ============================================================

$InstalledVersion = (Get-Item $Audacity4Exe).VersionInfo.ProductVersion

Write-Output "Audacity executable found."
Write-Output "Path: $Audacity4Exe"
Write-Output "Installed version: $InstalledVersion"

# ============================================================
# VERIFY VERSION
# ============================================================

if ($InstalledVersion -like "4.0.*") {

    Write-Output ""
    Write-Output "=========================================="
    Write-Output "SUCCESS"
    Write-Output "=========================================="
    Write-Output "Audacity 4.0.0 installed successfully."
    Write-Output "Version: $InstalledVersion"
    Write-Output "Path: $Audacity4Exe"
    Write-Output "=========================================="

    exit 0
}
else {

    Write-Error "Audacity was installed, but the version is not 4.0.x."
    Write-Error "Detected version: $InstalledVersion"

    exit 1
}

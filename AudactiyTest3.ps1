$ErrorActionPreference = "Stop"

$Url = "https://github.com/audacity/audacity/releases/download/Audacity-4.0.0/audacity-win-4.0.0-x86_64.msi"

$WorkDir = "C:\Windows\Temp\AudacityUpgrade"
$MsiPath = "$WorkDir\Audacity-4.0.0.msi"
$LogPath = "$WorkDir\Audacity-4.0.0-MSI.log"

$RegistryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

Write-Output "=========================================="
Write-Output "Audacity 4.0.0 Upgrade"
Write-Output "=========================================="

# ------------------------------------------------------------
# STOP AUDACITY
# ------------------------------------------------------------

Write-Output "Stopping Audacity if running..."

Get-Process -Name "Audacity" -ErrorAction SilentlyContinue |
    Stop-Process -Force -ErrorAction SilentlyContinue

Start-Sleep -Seconds 3

# ------------------------------------------------------------
# FUNCTION: FIND AUDACITY 4
# ------------------------------------------------------------

function Get-Audacity4 {

    foreach ($Path in $RegistryPaths) {

        Get-ItemProperty $Path -ErrorAction SilentlyContinue |
            Where-Object {
                $_.DisplayName -like "Audacity*" -and
                $_.DisplayVersion -match "^4\."
            }
    }
}

# ------------------------------------------------------------
# CHECK IF AUDACITY 4 IS ALREADY INSTALLED
# ------------------------------------------------------------

Write-Output "Checking for Audacity 4.x..."

$Audacity4 = @(Get-Audacity4)

if ($Audacity4.Count -gt 0) {

    Write-Output "Audacity 4.x is already installed."

    foreach ($App in $Audacity4) {
        Write-Output "Name: $($App.DisplayName)"
        Write-Output "Version: $($App.DisplayVersion)"
        Write-Output "Install Location: $($App.InstallLocation)"
    }

    $Audacity4Exe = "C:\Program Files\Audacity 4\Audacity.exe"

    if (Test-Path $Audacity4Exe) {

        $Version = (Get-Item $Audacity4Exe).VersionInfo.ProductVersion

        Write-Output "Audacity 4 executable found."
        Write-Output "Installed version: $Version"

        Write-Output "Audacity 4.0.0 is already installed."
        Write-Output "Nothing to do."

        exit 0
    }
}

# ------------------------------------------------------------
# FIND AUDACITY 3.x
# ------------------------------------------------------------

Write-Output "Searching for Audacity 3.x..."

$Audacity3 = @(
    foreach ($Path in $RegistryPaths) {

        Get-ItemProperty $Path -ErrorAction SilentlyContinue |
            Where-Object {
                $_.DisplayName -like "Audacity*" -and
                $_.DisplayVersion -match "^3\."
            }
    }
)

# ------------------------------------------------------------
# REMOVE AUDACITY 3.x
# ------------------------------------------------------------

if ($Audacity3.Count -gt 0) {

    foreach ($App in $Audacity3) {

        Write-Output "Found old Audacity:"
        Write-Output "Name: $($App.DisplayName)"
        Write-Output "Version: $($App.DisplayVersion)"
        Write-Output "Install Location: $($App.InstallLocation)"

        if ($App.UninstallString -match '"([^"]+unins[^"]+\.exe)"') {

            $Uninstaller = $Matches[1]

            Write-Output "Uninstaller: $Uninstaller"
            Write-Output "Removing Audacity 3.x..."

            $Process = Start-Process `
                -FilePath $Uninstaller `
                -ArgumentList "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART" `
                -Wait `
                -PassThru

            Write-Output "Uninstall exit code: $($Process.ExitCode)"

            if ($Process.ExitCode -ne 0) {

                Write-Error "Audacity 3.x uninstall failed."
                exit $Process.ExitCode
            }
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

# ------------------------------------------------------------
# WAIT FOR UNINSTALL
# ------------------------------------------------------------

Write-Output "Waiting for uninstall to complete..."

Start-Sleep -Seconds 5

# ------------------------------------------------------------
# CHECK AGAIN FOR AUDACITY 4
# ------------------------------------------------------------

Write-Output "Checking again for Audacity 4.x..."

$Audacity4 = @(Get-Audacity4)

if ($Audacity4.Count -gt 0) {

    Write-Output "Audacity 4.x is already installed."

    foreach ($App in $Audacity4) {
        Write-Output "Name: $($App.DisplayName)"
        Write-Output "Version: $($App.DisplayVersion)"
    }

    $Audacity4Exe = "C:\Program Files\Audacity 4\Audacity.exe"

    if (Test-Path $Audacity4Exe) {

        $Version = (Get-Item $Audacity4Exe).VersionInfo.ProductVersion

        Write-Output "Audacity 4 executable found."
        Write-Output "Version: $Version"
        Write-Output "Upgrade already completed."

        exit 0
    }
}

# ------------------------------------------------------------
# CREATE WORK DIRECTORY
# ------------------------------------------------------------

New-Item `
    -Path $WorkDir `
    -ItemType Directory `
    -Force |
    Out-Null

# ------------------------------------------------------------
# DOWNLOAD AUDACITY 4.0.0
# ------------------------------------------------------------

Write-Output "Downloading Audacity 4.0.0..."

if (Test-Path $MsiPath) {
    Remove-Item $MsiPath -Force
}

Invoke-WebRequest `
    -Uri $Url `
    -OutFile $MsiPath `
    -UseBasicParsing

if (-not (Test-Path $MsiPath)) {

    Write-Error "Audacity 4.0.0 MSI download failed."
    exit 1
}

$MsiSize = (Get-Item $MsiPath).Length

Write-Output "Downloaded MSI."
Write-Output "Size: $MsiSize bytes"

if ($MsiSize -lt 40000000) {

    Write-Error "Downloaded MSI is too small or incomplete."
    exit 1
}

# ------------------------------------------------------------
# INSTALL AUDACITY 4.0.0
# ------------------------------------------------------------

Write-Output "Installing Audacity 4.0.0..."

$Arguments = "/i `"$MsiPath`" /quiet /norestart ALLUSERS=1 /L*v `"$LogPath`""

Write-Output "Command:"
Write-Output "msiexec.exe $Arguments"

$InstallProcess = Start-Process `
    -FilePath "msiexec.exe" `
    -ArgumentList $Arguments `
    -Wait `
    -PassThru

Write-Output "MSI exit code: $($InstallProcess.ExitCode)"

# ------------------------------------------------------------
# CHECK INSTALL RESULT
# ------------------------------------------------------------

if (($InstallProcess.ExitCode -ne 0) -and
    ($InstallProcess.ExitCode -ne 3010)) {

    Write-Error "Audacity 4.0.0 installation failed."
    Write-Error "MSI Exit Code: $($InstallProcess.ExitCode)"
    Write-Error "MSI Log: $LogPath"

    exit $InstallProcess.ExitCode
}

# ------------------------------------------------------------
# VERIFY INSTALLATION
# ------------------------------------------------------------

Write-Output "Verifying Audacity installation..."

Start-Sleep -Seconds 5

$Audacity4Exe = "C:\Program Files\Audacity 4\Audacity.exe"

if (Test-Path $Audacity4Exe) {

    $Version = (Get-Item $Audacity4Exe).VersionInfo.ProductVersion

    Write-Output "=========================================="
    Write-Output "SUCCESS"
    Write-Output "Audacity installed successfully."
    Write-Output "Version: $Version"
    Write-Output "Path: $Audacity4Exe"
    Write-Output "=========================================="

    exit 0
}
else {

    Write-Error "Audacity.exe was not found after installation."
    Write-Error "MSI Log: $LogPath"

    exit 1
}

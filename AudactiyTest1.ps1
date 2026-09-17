# ============================================================
# Audacity 3.x -> 4.0.0 Upgrade Script
# Designed for RMM / Deep Freeze Cloud SYSTEM deployment
# ============================================================

$ErrorActionPreference = "Stop"

$AudacityURL = "https://github.com/audacity/audacity/releases/download/Audacity-4.0.0/audacity-win-4.0.0-x86_64.msi"

$TempDir = "C:\Windows\Temp\AudacityUpgrade"
$MsiPath = Join-Path $TempDir "audacity-win-4.0.0-x86_64.msi"
$LogPath = Join-Path $TempDir "Audacity-4.0.0-Install.log"

Write-Output "=========================================="
Write-Output "Audacity 4.0.0 Upgrade"
Write-Output "=========================================="

# ------------------------------------------------------------
# Create working directory
# ------------------------------------------------------------

if (-not (Test-Path $TempDir)) {
    New-Item -Path $TempDir -ItemType Directory -Force | Out-Null
}

# ------------------------------------------------------------
# Stop Audacity if currently running
# ------------------------------------------------------------

Write-Output "Checking for running Audacity processes..."

Get-Process -Name "Audacity" -ErrorAction SilentlyContinue | ForEach-Object {
    Write-Output "Stopping Audacity process: $($_.Id)"
    Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
}

Start-Sleep -Seconds 2

# ------------------------------------------------------------
# Find Audacity uninstall entries
# ------------------------------------------------------------

Write-Output "Searching for installed Audacity versions..."

$UninstallPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

$AudacityApps = foreach ($Path in $UninstallPaths) {
    Get-ItemProperty $Path -ErrorAction SilentlyContinue |
        Where-Object {
            $_.DisplayName -like "Audacity*"
        }
}

# ------------------------------------------------------------
# Uninstall old Audacity
# ------------------------------------------------------------

if ($AudacityApps) {

    foreach ($App in $AudacityApps) {

        Write-Output "Found installed application:"
        Write-Output "Name: $($App.DisplayName)"
        Write-Output "Version: $($App.DisplayVersion)"
        Write-Output "Install Location: $($App.InstallLocation)"

        $UninstallString = $App.UninstallString

        if ($UninstallString) {

            Write-Output "Uninstall command:"
            Write-Output $UninstallString

            # Handle quoted executable path
            if ($UninstallString -match '^\s*"([^"]+)"\s*(.*)$') {

                $Uninstaller = $Matches[1]
                $Arguments = $Matches[2]

            }
            else {

                $Parts = $UninstallString -split '\s+', 2

                $Uninstaller = $Parts[0]

                if ($Parts.Count -gt 1) {
                    $Arguments = $Parts[1]
                }
                else {
                    $Arguments = ""
                }
            }

            # Inno Setup silent uninstall
            if ($Uninstaller -match "unins.*\.exe") {

                Write-Output "Detected Inno Setup uninstaller."

                $Arguments = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"

            }

            Write-Output "Running uninstall..."

            $Process = Start-Process `
                -FilePath $Uninstaller `
                -ArgumentList $Arguments `
                -Wait `
                -PassThru

            Write-Output "Uninstaller exit code: $($Process.ExitCode)"

        }
    }

}
else {

    Write-Output "No Audacity uninstall entry found."

    # --------------------------------------------------------
    # Fallback: look for the standard Audacity 3.x uninstaller
    # --------------------------------------------------------

    $PossibleUninstallers = @(
        "C:\Program Files\Audacity\unins000.exe",
        "C:\Program Files (x86)\Audacity\unins000.exe"
    )

    foreach ($Uninstaller in $PossibleUninstallers) {

        if (Test-Path $Uninstaller) {

            Write-Output "Found Audacity uninstaller:"
            Write-Output $Uninstaller

            Write-Output "Running silent uninstall..."

            $Process = Start-Process `
                -FilePath $Uninstaller `
                -ArgumentList "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART" `
                -Wait `
                -PassThru

            Write-Output "Uninstaller exit code: $($Process.ExitCode)"
        }
    }
}

# ------------------------------------------------------------
# Wait for uninstall to finish
# ------------------------------------------------------------

Write-Output "Waiting for uninstall to complete..."

Start-Sleep -Seconds 5

# ------------------------------------------------------------
# Download Audacity 4.0.0 MSI
# ------------------------------------------------------------

Write-Output "Downloading Audacity 4.0.0..."

if (Test-Path $MsiPath) {
    Remove-Item $MsiPath -Force
}

Invoke-WebRequest `
    -Uri $AudacityURL `
    -OutFile $MsiPath `
    -UseBasicParsing

# ------------------------------------------------------------
# Verify MSI download
# ------------------------------------------------------------

if (-not (Test-Path $MsiPath)) {
    Write-Error "Audacity 4.0.0 MSI download failed."
    exit 1
}

$MsiSize = (Get-Item $MsiPath).Length

Write-Output "Downloaded MSI successfully."
Write-Output "MSI size: $MsiSize bytes"

if ($MsiSize -lt 1MB) {
    Write-Error "Downloaded MSI appears to be invalid or incomplete."
    exit 1
}

# ------------------------------------------------------------
# Install Audacity 4.0.0
# ------------------------------------------------------------

Write-Output "Installing Audacity 4.0.0..."

$Arguments = "/i `"$MsiPath`" /qn /norestart /L*v `"$LogPath`""

$InstallProcess = Start-Process `
    -FilePath "msiexec.exe" `
    -ArgumentList $Arguments `
    -Wait `
    -PassThru

Write-Output "MSI exit code: $($InstallProcess.ExitCode)"

# ------------------------------------------------------------
# MSI exit codes
# ------------------------------------------------------------

if ($InstallProcess.ExitCode -eq 0) {

    Write-Output "Audacity 4.0.0 installed successfully."

}
elseif ($InstallProcess.ExitCode -eq 3010) {

    Write-Output "Audacity 4.0.0 installed successfully."
    Write-Output "A reboot is required to complete installation."

}
else {

    Write-Error "Audacity 4.0.0 installation failed."
    Write-Error "MSI Exit Code: $($InstallProcess.ExitCode)"
    Write-Error "Installation log: $LogPath"

    exit $InstallProcess.ExitCode
}

# ------------------------------------------------------------
# Verify Audacity 4 installation
# ------------------------------------------------------------

Start-Sleep -Seconds 3

$Audacity4Path = "C:\Program Files\Audacity 4\Audacity.exe"

if (Test-Path $Audacity4Path) {

    $Version = (Get-Item $Audacity4Path).VersionInfo.ProductVersion

    Write-Output "=========================================="
    Write-Output "Audacity installation verified."
    Write-Output "Installed version: $Version"
    Write-Output "Path: $Audacity4Path"
    Write-Output "=========================================="

}
else {

    Write-Error "Audacity.exe was not found after installation."
    Write-Error "Expected path: $Audacity4Path"

    exit 1
}

exit 0

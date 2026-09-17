$ErrorActionPreference = "Stop"

$Url = "https://github.com/audacity/audacity/releases/download/Audacity-4.0.0/audacity-win-4.0.0-x86_64.msi"
$WorkDir = "C:\Windows\Temp\AudacityUpgrade"
$Msi = "$WorkDir\Audacity-4.0.0.msi"
$Log = "$WorkDir\Audacity-4.0.0-MSI.log"

Write-Output "=========================================="
Write-Output "Audacity 4.0.0 Upgrade"
Write-Output "=========================================="

# Create working directory
New-Item -Path $WorkDir -ItemType Directory -Force | Out-Null

# Stop Audacity
Write-Output "Stopping Audacity if running..."

Get-Process -Name "Audacity","Audacity4" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

Start-Sleep -Seconds 3

# ------------------------------------------------------------
# REMOVE ONLY AUDACITY 3.x
# ------------------------------------------------------------

Write-Output "Searching for Audacity 3.x..."

$RegistryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

$OldAudacity = @(
    foreach ($Path in $RegistryPaths) {
        Get-ItemProperty $Path -ErrorAction SilentlyContinue |
        Where-Object {
            $_.DisplayName -like "Audacity*" -and
            $_.DisplayVersion -match "^3\."
        }
    }
)

if ($OldAudacity.Count -gt 0) {

    foreach ($App in $OldAudacity) {

        Write-Output "Found old Audacity:"
        Write-Output "Name: $($App.DisplayName)"
        Write-Output "Version: $($App.DisplayVersion)"

        # Inno Setup uninstaller
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
        }
        else {
            Write-Output "Audacity 3.x uninstall command not recognized."
        }
    }

}
else {

    Write-Output "Audacity 3.x was not found."
}

# ------------------------------------------------------------
# VERIFY AUDACITY 4 IS NOT BEING REMOVED
# ------------------------------------------------------------

Write-Output "Checking for existing Audacity 4..."

$Audacity4 = @(
    foreach ($Path in $RegistryPaths) {
        Get-ItemProperty $Path -ErrorAction SilentlyContinue |
        Where-Object {
            $_.DisplayName -like "Audacity*" -and
            $_.DisplayVersion -match "^4\."
        }
    }
)

if ($Audacity4.Count -gt 0) {

    Write-Output "Audacity 4.x is already installed."

    foreach ($App in $Audacity4) {
        Write-Output "Detected: $($App.DisplayName) $($App.DisplayVersion)"
    }

    Write-Output "No existing Audacity 4.x installation will be removed."

}

# ------------------------------------------------------------
# DOWNLOAD AUDACITY 4.0.0
# ------------------------------------------------------------

Write-Output "Downloading Audacity 4.0.0..."

if (Test-Path $Msi) {
    Remove-Item $Msi -Force
}

Invoke-WebRequest `
    -Uri $Url `
    -OutFile $Msi `
    -UseBasicParsing

if (-not (Test-Path $Msi)) {
    Write-Error "Audacity MSI download failed."
    exit 1
}

$Size = (Get-Item $Msi).Length

Write-Output "Downloaded MSI size: $Size bytes"

if ($Size -lt 40000000) {
    Write-Error "Downloaded MSI appears to be incomplete."
    exit 1
}

# ------------------------------------------------------------
# INSTALL AUDACITY 4.0.0
# ------------------------------------------------------------

Write-Output "Installing Audacity 4.0.0..."

$Arguments = "/i `"$Msi`" /quiet /norestart ALLUSERS=1 /L*v `"$Log`""

Write-Output "Running:"
Write-Output "msiexec.exe $Arguments"

$InstallProcess = Start-Process `
    -FilePath "msiexec.exe" `
    -ArgumentList $Arguments `
    -Wait `
    -PassThru

Write-Output "MSI exit code: $($InstallProcess.ExitCode)"

# ------------------------------------------------------------
# INSTALL RESULT
# ------------------------------------------------------------

if (($InstallProcess.ExitCode -ne 0) -and ($InstallProcess.ExitCode -ne 3010)) {

    Write-Error "Audacity 4.0.0 installation failed."
    Write-Error "MSI exit code: $($InstallProcess.ExitCode)"
    Write-Error "MSI log: $Log"

    exit $InstallProcess.ExitCode
}

# ------------------------------------------------------------
# VERIFY INSTALLATION
# ------------------------------------------------------------

Write-Output "Verifying Audacity 4.0.0 installation..."

Start-Sleep -Seconds 5

$PossiblePaths = @(
    "C:\Program Files\Audacity 4\Audacity.exe",
    "C:\Program Files\Audacity\Audacity.exe"
)

$AudacityExe = $PossiblePaths |
    Where-Object { Test-Path $_ } |
    Select-Object -First 1

if ($AudacityExe) {

    $InstalledVersion = (Get-Item $AudacityExe).VersionInfo.ProductVersion

    Write-Output "=========================================="
    Write-Output "Audacity installation successful."
    Write-Output "Path: $AudacityExe"
    Write-Output "Version: $InstalledVersion"
    Write-Output "=========================================="

    exit 0
}
else {

    Write-Error "Audacity.exe was not found after installation."
    Write-Error "Check MSI log:"
    Write-Error $Log

    exit 1
}

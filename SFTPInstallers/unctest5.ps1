[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$UncPath,

    [Parameter(Mandatory = $false)]
    [string]$Username,

    [Parameter(Mandatory = $false)]
    [string]$Password,

    [Parameter(Mandatory = $true)]
    [string]$LicenseFileName
)

$ErrorActionPreference = "Stop"
$WorkingDir = Join-Path -Path $env:TEMP -ChildPath "Fusion2026_Install_$([guid]::NewGuid().ToString().Substring(0,8))"
$TargetShare = $null

try {
    # 1. Parse root share (e.g., \\pdc\Share2)
    if ($UncPath -match '^(\\\\[^\\]+\\[^\\]+)') {
        $TargetShare = $matches[1]
    } else {
        $TargetShare = $UncPath
    }

    # 2. Establish session credentials
    if (-not [string]::IsNullOrWhiteSpace($Username) -and -not [string]::IsNullOrWhiteSpace($Password)) {
        Write-Host "Authenticating session to $TargetShare..."
        $netUse = Start-Process -FilePath "net.exe" -ArgumentList @('use', $TargetShare, $Password, "/user:$Username") -Wait -NoNewWindow -PassThru
        if ($netUse.ExitCode -ne 0) {
            throw "Failed to authenticate to $TargetShare (Exit Code: $($netUse.ExitCode))."
        }
    }

    # 3. Determine direct UNC file path to the ZIP
    if ($UncPath -match '\.zip$') {
        $SourceZipPath = $UncPath
    } else {
        $SourceZipPath = "$($TargetShare.TrimEnd('\'))\Fusion.zip"
    }

    # 4. Create local staging directories
    New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null
    $LocalZip = Join-Path -Path $WorkingDir -ChildPath "FusionPackage.zip"
    $ExtractDir = Join-Path -Path $WorkingDir -ChildPath "Extracted"

    Write-Host "Copying $SourceZipPath locally using direct .NET file transfer..."
    [System.IO.File]::Copy($SourceZipPath, $LocalZip, $true)

    # 5. Extract installation files
    Write-Host "Extracting archive contents..."
    Expand-Archive -Path $LocalZip -DestinationPath $ExtractDir -Force

    # 6. Locate installer dynamically (handles build numbers like 2608.3.400)
    $Installer = Get-ChildItem -Path $ExtractDir -Filter "*Offline*.exe" -Recurse | Select-Object -First 1
    if (-not $Installer) {
        $Installer = Get-ChildItem -Path $ExtractDir -Filter "*.exe" -Recurse | Select-Object -First 1
    }
    if (-not $Installer) {
        throw "No installer executable found inside the extracted package."
    }

    # 7. Locate license file
    $LicenseFile = Get-ChildItem -Path $ExtractDir -Filter $LicenseFileName -Recurse | Select-Object -First 1
    if (-not $LicenseFile) {
        throw "License file '$LicenseFileName' was not found inside the ZIP archive."
    }

    # 8. Run Fusion 2026 installer
    Set-Location -Path $Installer.DirectoryName
    Write-Host "Executing installer: $($Installer.FullName)..."
    $installProcess = Start-Process -FilePath $Installer.FullName -ArgumentList "/type Silent /DisableExternalServices" -Wait -PassThru

    if ($installProcess.ExitCode -ne 0) {
        Write-Warning "Installer returned non-zero exit code: $($installProcess.ExitCode)"
    }

    # 9. Deploy license
    Write-Host "Deploying license file ($LicenseFileName)..."
    $DestinationAuthDir = "C:\Program Files\Freedom Scientific\Authorization"
    if (-not (Test-Path -Path $DestinationAuthDir)) {
        New-Item -ItemType Directory -Path $DestinationAuthDir -Force | Out-Null
    }

    Copy-Item -Path $LicenseFile.FullName -Destination $DestinationAuthDir -Force -Recurse

    Write-Host "Fusion 2026 deployment and licensing completed successfully."
}
catch {
    Write-Error "Deployment failed: $_"
    exit 1
}
finally {
    # 10. Cleanup local staging and share connection
    Set-Location -Path $env:SystemRoot
    if (Test-Path -Path $WorkingDir) {
        Remove-Item -Path $WorkingDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    if (-not [string]::IsNullOrWhiteSpace($Username) -and $TargetShare) {
        Start-Process -FilePath "net.exe" -ArgumentList @('use', $TargetShare, '/delete', '/y') -Wait -NoNewWindow -PassThru -ErrorAction SilentlyContinue | Out-Null
    }
}

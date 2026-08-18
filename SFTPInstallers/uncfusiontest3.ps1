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
    # 1. Parse root share (\\server\share)
    if ($UncPath -match '^(\\\\[^\\]+\\[^\\]+)') {
        $TargetShare = $matches[1]
    } else {
        $TargetShare = $UncPath
    }

    # 2. Authenticate to the share
    if (-not [string]::IsNullOrWhiteSpace($Username) -and -not [string]::IsNullOrWhiteSpace($Password)) {
        Write-Host "Authenticating to share: $TargetShare..."
        $netUse = Start-Process -FilePath "net.exe" -ArgumentList @('use', $TargetShare, $Password, "/user:$Username") -Wait -NoNewWindow -PassThru
        if ($netUse.ExitCode -ne 0) {
            throw "Failed to authenticate to $TargetShare (Exit Code: $($netUse.ExitCode))."
        }
    }

    # 3. Create local staging directories
    New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null
    $ExtractDir = Join-Path -Path $WorkingDir -ChildPath "Extracted"
    New-Item -ItemType Directory -Path $ExtractDir -Force | Out-Null

    # 4. Parse source folder and file name for Robocopy
    if ($UncPath -match '\.zip$') {
        $SourceDir = [System.IO.Path]::GetDirectoryName($UncPath)
        $ZipFileName = [System.IO.Path]::GetFileName($UncPath)
    } else {
        $SourceDir = $UncPath
        $ZipFileName = "*.zip"
    }

    Write-Host "Copying package locally from $SourceDir to $WorkingDir..."
    # Robocopy runs directly in the current authenticated process token
    $roboArgs = @($SourceDir, $WorkingDir, $ZipFileName, "/R:3", "/W:2", "/NP", "/NDL")
    $robo = Start-Process -FilePath "robocopy.exe" -ArgumentList $roboArgs -Wait -NoNewWindow -PassThru

    # Robocopy exit codes 0-7 indicate successful copy operations
    if ($robo.ExitCode -gt 7) {
        throw "Robocopy failed to fetch the zip package (Exit Code: $($robo.ExitCode))."
    }

    $DownloadedZip = (Get-ChildItem -Path $WorkingDir -Filter "*.zip" | Select-Object -First 1).FullName
    if (-not $DownloadedZip) {
        throw "Could not locate downloaded ZIP archive in $WorkingDir."
    }

    # 5. Extract installation files
    Write-Host "Extracting installation package..."
    Expand-Archive -Path $DownloadedZip -DestinationPath $ExtractDir -Force

    # 6. Set working location to extracted files
    Set-Location -Path $ExtractDir

    # 7. Install Freedom Scientific Fusion
    Write-Host "Running Fusion 2026 silent installation..."
    Start-Process -FilePath .\F2026.2602.8.400-Offline-x64.exe -ArgumentList "/type Silent /DisableExternalServices" -Wait

    # 8. Copy license file
    Write-Host "Applying license file ($LicenseFileName)..."
    $DestinationAuthDir = "C:\Program Files\Freedom Scientific\Authorization"
    if (-not (Test-Path -Path $DestinationAuthDir)) {
        New-Item -ItemType Directory -Path $DestinationAuthDir -Force | Out-Null
    }

    Copy-Item -Path ".\$LicenseFileName" -Destination $DestinationAuthDir -Force -Recurse

    Write-Host "Fusion 2026 installation and license deployment completed successfully."
}
catch {
    Write-Error "Deployment failed: $_"
    exit 1
}
finally {
    # 9. Cleanup staging artifacts and disconnect share
    Set-Location -Path $env:SystemRoot
    if (Test-Path -Path $WorkingDir) {
        Remove-Item -Path $WorkingDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    if (-not [string]::IsNullOrWhiteSpace($Username) -and $TargetShare) {
        Start-Process -FilePath "net.exe" -ArgumentList @('use', $TargetShare, '/delete', '/y') -Wait -NoNewWindow -PassThru -ErrorAction SilentlyContinue | Out-Null
    }
}

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

try {
    # 1. Establish network share connection if credentials are provided
    if (-not [string]::IsNullOrWhiteSpace($Username) -and -not [string]::IsNullOrWhiteSpace($Password)) {
        Write-Host "Authenticating to UNC path..."
        $netUseArgs = @('use', $UncPath, $Password, "/user:$Username")
        $netUse = Start-Process -FilePath "net.exe" -ArgumentList $netUseArgs -Wait -NoNewWindow -PassThru
        if ($netUse.ExitCode -ne 0) {
            throw "Failed to authenticate to $UncPath (Exit Code: $($netUse.ExitCode))."
        }
    }

    # 2. Locate the ZIP file on the share
    if (Test-Path -Path $UncPath -PathType Leaf) {
        $ZipSource = $UncPath
    } else {
        $ZipSource = (Get-ChildItem -Path $UncPath -Filter "*.zip" | Select-Object -First 1).FullName
        if (-not $ZipSource) {
            throw "No ZIP file found at path: $UncPath"
        }
    }

    # 3. Create clean local staging directory
    New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null
    $LocalZip = Join-Path -Path $WorkingDir -ChildPath "FusionPackage.zip"
    $ExtractDir = Join-Path -Path $WorkingDir -ChildPath "Extracted"

    Write-Host "Copying package locally from $ZipSource..."
    Copy-Item -Path $ZipSource -Destination $LocalZip -Force

    Write-Host "Extracting installation package..."
    Expand-Archive -Path $LocalZip -DestinationPath $ExtractDir -Force

    # 4. Set location to the extracted directory to match .\ relative paths
    Set-Location -Path $ExtractDir

    # 5. Install Freedom Scientific Fusion
    Write-Host "Running Fusion 2026 silent installation..."
    Start-Process -FilePath .\F2026.2602.8.400-Offline-x64.exe -ArgumentList "/type Silent /DisableExternalServices" -Wait

    # 6. Copy license file
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
    # 7. Cleanup staging artifacts and network connection
    Set-Location -Path $env:SystemRoot
    if (Test-Path -Path $WorkingDir) {
        Remove-Item -Path $WorkingDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    if (-not [string]::IsNullOrWhiteSpace($Username)) {
        Start-Process -FilePath "net.exe" -ArgumentList @('use', $UncPath, '/delete', '/y') -Wait -NoNewWindow -PassThru -ErrorAction SilentlyContinue | Out-Null
    }
}

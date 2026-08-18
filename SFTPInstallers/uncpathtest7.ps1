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
$DriveName = "FusionShare"
$DriveCreated = $false

try {
    # 1. Parse root share (\\pdc\Share2)
    if ($UncPath -match '^(\\\\[^\\]+\\[^\\]+)') {
        $TargetShare = $matches[1]
    } else {
        $TargetShare = $UncPath
    }

    # 2. Authenticate and mount share directly within PowerShell
    Write-Host "Mounting PSDrive for $TargetShare..."
    if (-not [string]::IsNullOrWhiteSpace($Username) -and -not [string]::IsNullOrWhiteSpace($Password)) {
        $secPassword = ConvertTo-SecureString $Password -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential ($Username, $secPassword)
        New-PSDrive -Name $DriveName -PSProvider FileSystem -Root $TargetShare -Credential $cred -Scope Global -ErrorAction Stop | Out-Null
    } else {
        New-PSDrive -Name $DriveName -PSProvider FileSystem -Root $TargetShare -Scope Global -ErrorAction Stop | Out-Null
    }
    $DriveCreated = $true

    # 3. Discover the ZIP file on the mounted drive
    Write-Host "Locating ZIP package in share..."
    $SourceZip = Get-ChildItem -Path "$($DriveName):\" -Filter "*.zip" | Select-Object -First 1

    if (-not $SourceZip) {
        $ShareContents = (Get-ChildItem -Path "$($DriveName):\" | Select-Object -ExpandProperty Name) -join ", "
        throw "No ZIP file found on share. Items found: [$ShareContents]"
    }

    Write-Host "Found package: $($SourceZip.Name) ($([math]::Round($SourceZip.Length / 1MB, 2)) MB)"

    # 4. Prepare local staging directory
    New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null
    $LocalZip = Join-Path -Path $WorkingDir -ChildPath $SourceZip.Name
    $ExtractDir = Join-Path -Path $WorkingDir -ChildPath "Extracted"

    Write-Host "Copying package locally to $LocalZip..."
    Copy-Item -Path $SourceZip.FullName -Destination $LocalZip -Force

    # 5. Extract archive
    Write-Host "Extracting archive contents..."
    Expand-Archive -Path $LocalZip -DestinationPath $ExtractDir -Force

    # 6. Locate installer (.exe) and license file inside extracted content
    $Installer = Get-ChildItem -Path $ExtractDir -Filter "*Offline*.exe" -Recurse | Select-Object -First 1
    if (-not $Installer) {
        $Installer = Get-ChildItem -Path $ExtractDir -Filter "*.exe" -Recurse | Select-Object -First 1
    }
    if (-not $Installer) {
        throw "No installer executable (.exe) found inside the extracted ZIP package."
    }

    $LicenseFile = Get-ChildItem -Path $ExtractDir -Filter $LicenseFileName -Recurse | Select-Object -First 1
    if (-not $LicenseFile) {
        throw "License file '$LicenseFileName' not found inside the ZIP package."
    }

    # 7. Execute Fusion 2026 installer
    Set-Location -Path $Installer.DirectoryName
    Write-Host "Installing: $($Installer.Name)..."
    $proc = Start-Process -FilePath $Installer.FullName -ArgumentList "/type Silent /DisableExternalServices" -Wait -PassThru

    Write-Host "Installer finished with Exit Code: $($proc.ExitCode)"

    # 8. Deploy license
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
    # 9. Cleanup staging and disconnect drive
    Set-Location -Path $env:SystemRoot
    if (Test-Path -Path $WorkingDir) {
        Remove-Item -Path $WorkingDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    if ($DriveCreated) {
        Remove-PSDrive -Name $DriveName -Force -ErrorAction SilentlyContinue
    }
}

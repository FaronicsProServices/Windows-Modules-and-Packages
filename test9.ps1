$LogFile = "$env:TEMP\ReadWrite_Install_Log.txt"
Start-Transcript -Path $LogFile -Append

Write-Output "Starting Read&Write automated install..."

# Variables
$zipUrl = "https://fastdownloads2.texthelp.com/readwrite12/installers/us/setup.zip"
$downloadPath = "$env:TEMP\setup.zip"
$extractPath = "$env:TEMP\ReadWriteInstaller"

# Create extract path
New-Item -ItemType Directory -Force -Path $extractPath | Out-Null

# Download ZIP
Invoke-WebRequest -Uri $zipUrl -OutFile $downloadPath -UseBasicParsing

# Extract ZIP
Expand-Archive -Path $downloadPath -DestinationPath $extractPath -Force

# Log all extracted files
Write-Output "`nExtracted files:" 
Get-ChildItem -Path $extractPath -Recurse | ForEach-Object {
    Write-Output $_.FullName
}

# Try to find installer
$installer = Get-ChildItem -Path $extractPath -Recurse | Where-Object {
    $_.Extension -in '.exe', '.msi'
} | Select-Object -First 1

if (-not $installer) {
    Write-Error "❌ No installer found."
    Stop-Transcript
    exit 1
}

# Try multiple silent install options
Write-Output "`nTrying to install silently..."

$installerPath = $installer.FullName
$exitCode = $null
$silentSwitches = @(
    "/S", "/silent", "/quiet", "/VERYSILENT", "/qn"
)

foreach ($switch in $silentSwitches) {
    Write-Output "Attempting: $($installer.Name) $switch"
    try {
        $process = Start-Process -FilePath $installerPath -ArgumentList $switch -PassThru -Wait
        $exitCode = $process.ExitCode
        Write-Output "Exit Code: $exitCode"
        if ($exitCode -eq 0) {
            Write-Output "✅ Installer ran successfully with switch: $switch"
            break
        }
    } catch {
        Write-Output ("❌ Error running installer with '{0}': {1}" -f $switch, $_.Exception.Message)
    }
}

if ($exitCode -ne 0) {
    Write-Error "❌ All silent install attempts failed."
} else {
    Write-Output "✅ Installation may have succeeded. Please verify on client."
}

Stop-Transcript

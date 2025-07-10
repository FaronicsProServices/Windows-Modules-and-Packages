# Read&Write Automated Installation Script
# This script downloads, extracts, and installs Read&Write from the provided URL
# Compatible with PDQ Deploy and Deep Freeze Cloud

param(
    [string]$TempPath = "C:\temp\readwrite_install",
    [string]$LogPath = "C:\temp\readwrite_install.log",
    [switch]$Verbose = $false
)

# Function to write logs
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Write-Host $logEntry
    Add-Content -Path $LogPath -Value $logEntry -Force
}

# Function to test internet connectivity
function Test-InternetConnection {
    try {
        $response = Invoke-WebRequest -Uri "https://fastdownloads2.texthelp.com/readwrite12/installers/us/setup.zip" -Method Head -TimeoutSec 10 -UseBasicParsing
        return $true
    }
    catch {
        return $false
    }
}

# Function to download file with retry logic
function Download-FileWithRetry {
    param(
        [string]$Url,
        [string]$OutputPath,
        [int]$MaxRetries = 3
    )
    
    $attempt = 0
    while ($attempt -lt $MaxRetries) {
        try {
            $attempt++
            Write-Log "Downloading attempt $attempt of $MaxRetries..."
            
            # Use different methods based on PowerShell version
            if ($PSVersionTable.PSVersion.Major -ge 3) {
                Invoke-WebRequest -Uri $Url -OutFile $OutputPath -UseBasicParsing -TimeoutSec 300
            } else {
                $webClient = New-Object System.Net.WebClient
                $webClient.DownloadFile($Url, $OutputPath)
                $webClient.Dispose()
            }
            
            if (Test-Path $OutputPath) {
                Write-Log "Download successful"
                return $true
            }
        }
        catch {
            Write-Log "Download attempt $attempt failed: $($_.Exception.Message)" "ERROR"
            if ($attempt -eq $MaxRetries) {
                return $false
            }
            Start-Sleep -Seconds 5
        }
    }
    return $false
}

# Function to extract ZIP file
function Extract-ZipFile {
    param(
        [string]$ZipPath,
        [string]$ExtractPath
    )
    
    try {
        # Method 1: Use .NET Framework (Works on older systems)
        if ($PSVersionTable.PSVersion.Major -lt 5) {
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $ExtractPath)
        }
        # Method 2: Use PowerShell 5+ Expand-Archive
        else {
            Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force
        }
        
        Write-Log "Extraction successful"
        return $true
    }
    catch {
        Write-Log "Extraction failed: $($_.Exception.Message)" "ERROR"
        
        # Fallback method using Shell.Application
        try {
            Write-Log "Trying fallback extraction method..."
            $shell = New-Object -ComObject Shell.Application
            $zip = $shell.NameSpace($ZipPath)
            $destination = $shell.NameSpace($ExtractPath)
            
            if ($zip -ne $null) {
                $destination.CopyHere($zip.Items(), 4)
                Write-Log "Fallback extraction successful"
                return $true
            }
        }
        catch {
            Write-Log "Fallback extraction also failed: $($_.Exception.Message)" "ERROR"
        }
        
        return $false
    }
}

# Function to find MSI file
function Find-MSIFile {
    param([string]$SearchPath)
    
    $msiFiles = Get-ChildItem -Path $SearchPath -Filter "*.msi" -Recurse
    if ($msiFiles.Count -gt 0) {
        return $msiFiles[0].FullName
    }
    return $null
}

# Function to install MSI
function Install-MSI {
    param(
        [string]$MSIPath,
        [string]$InstallArgs = "/quiet /norestart"
    )
    
    try {
        Write-Log "Starting MSI installation: $MSIPath"
        
        # Check if MSI exists
        if (!(Test-Path $MSIPath)) {
            Write-Log "MSI file not found: $MSIPath" "ERROR"
            return $false
        }
        
        # Build msiexec command
        $arguments = "/i `"$MSIPath`" $InstallArgs"
        
        Write-Log "Executing: msiexec.exe $arguments"
        
        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $arguments -Wait -PassThru -NoNewWindow
        
        Write-Log "MSI installation completed with exit code: $($process.ExitCode)"
        
        # Check common exit codes
        switch ($process.ExitCode) {
            0 { Write-Log "Installation successful"; return $true }
            1641 { Write-Log "Installation successful (restart required)"; return $true }
            3010 { Write-Log "Installation successful (restart required)"; return $true }
            1618 { Write-Log "Another installation is already in progress" "ERROR"; return $false }
            1603 { Write-Log "Fatal error during installation" "ERROR"; return $false }
            1619 { Write-Log "Installation package could not be opened" "ERROR"; return $false }
            1633 { Write-Log "This installation package is not supported on this platform" "ERROR"; return $false }
            default { Write-Log "Installation failed with exit code: $($process.ExitCode)" "ERROR"; return $false }
        }
    }
    catch {
        Write-Log "MSI installation error: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Function to cleanup temporary files
function Cleanup-TempFiles {
    param([string]$CleanupPath)
    
    try {
        if (Test-Path $CleanupPath) {
            Write-Log "Cleaning up temporary files..."
            Remove-Item -Path $CleanupPath -Recurse -Force
            Write-Log "Cleanup completed"
        }
    }
    catch {
        Write-Log "Cleanup failed: $($_.Exception.Message)" "WARNING"
    }
}

# Function to check if Read&Write is already installed
function Test-ReadWriteInstalled {
    try {
        # Check registry for installed programs
        $uninstallKeys = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        
        foreach ($key in $uninstallKeys) {
            $programs = Get-ItemProperty $key -ErrorAction SilentlyContinue
            foreach ($program in $programs) {
                if ($program.DisplayName -like "*Read*Write*" -or $program.DisplayName -like "*TextHelp*") {
                    Write-Log "Found existing installation: $($program.DisplayName) - Version: $($program.DisplayVersion)"
                    return $true
                }
            }
        }
        
        return $false
    }
    catch {
        Write-Log "Error checking for existing installation: $($_.Exception.Message)" "WARNING"
        return $false
    }
}

# Main execution starts here
Write-Log "=== Read&Write Automated Installation Started ==="
Write-Log "Script running on: $env:COMPUTERNAME"
Write-Log "Current user: $env:USERNAME"
Write-Log "PowerShell version: $($PSVersionTable.PSVersion.ToString())"

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
Write-Log "Running as administrator: $isAdmin"

# Variables
$downloadUrl = "https://fastdownloads2.texthelp.com/readwrite12/installers/us/setup.zip"
$zipFile = Join-Path $TempPath "setup.zip"
$extractPath = Join-Path $TempPath "extracted"

# Check if already installed
if (Test-ReadWriteInstalled) {
    Write-Log "Read&Write appears to be already installed. Continuing with installation anyway..." "WARNING"
}

# Step 1: Check internet connectivity
Write-Log "Checking internet connectivity..."
if (!(Test-InternetConnection)) {
    Write-Log "No internet connection available. Installation cannot proceed." "ERROR"
    exit 1
}

# Step 2: Create temporary directory
Write-Log "Creating temporary directory: $TempPath"
try {
    if (!(Test-Path $TempPath)) {
        New-Item -ItemType Directory -Path $TempPath -Force | Out-Null
    }
    if (!(Test-Path $extractPath)) {
        New-Item -ItemType Directory -Path $extractPath -Force | Out-Null
    }
}
catch {
    Write-Log "Failed to create temporary directory: $($_.Exception.Message)" "ERROR"
    exit 1
}

# Step 3: Download the setup file
Write-Log "Downloading Read&Write setup from: $downloadUrl"
if (!(Download-FileWithRetry -Url $downloadUrl -OutputPath $zipFile)) {
    Write-Log "Failed to download setup file" "ERROR"
    Cleanup-TempFiles -CleanupPath $TempPath
    exit 1
}

# Verify download
$fileSize = (Get-Item $zipFile).Length
Write-Log "Downloaded file size: $fileSize bytes"

# Step 4: Extract the ZIP file
Write-Log "Extracting setup files..."
if (!(Extract-ZipFile -ZipPath $zipFile -ExtractPath $extractPath)) {
    Write-Log "Failed to extract setup files" "ERROR"
    Cleanup-TempFiles -CleanupPath $TempPath
    exit 1
}

# Step 5: Find the MSI file
Write-Log "Searching for MSI installer..."
$msiFile = Find-MSIFile -SearchPath $extractPath

if ($msiFile -eq $null) {
    Write-Log "MSI file not found in extracted files" "ERROR"
    
    # List all files for debugging
    Write-Log "Files found in extraction:"
    Get-ChildItem -Path $extractPath -Recurse | ForEach-Object {
        Write-Log "  $($_.FullName)"
    }
    
    Cleanup-TempFiles -CleanupPath $TempPath
    exit 1
}

Write-Log "Found MSI installer: $msiFile"

# Step 6: Install the application
Write-Log "Starting Read&Write installation..."
$installSuccess = Install-MSI -MSIPath $msiFile -InstallArgs "/quiet /norestart ALLUSERS=1"

# Step 7: Verify installation
if ($installSuccess) {
    Write-Log "Waiting 10 seconds for installation to complete..."
    Start-Sleep -Seconds 10
    
    if (Test-ReadWriteInstalled) {
        Write-Log "=== Read&Write installation completed successfully ===" "SUCCESS"
        $exitCode = 0
    } else {
        Write-Log "Installation may have failed - software not detected in registry" "WARNING"
        $exitCode = 0  # Don't fail completely as some detection issues are normal
    }
} else {
    Write-Log "=== Read&Write installation failed ===" "ERROR"
    $exitCode = 1
}

# Step 8: Cleanup
Cleanup-TempFiles -CleanupPath $TempPath

Write-Log "=== Installation script completed with exit code: $exitCode ==="
exit $exitCode

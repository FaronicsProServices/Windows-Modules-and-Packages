# This script installs and imports the Microsoft Teams PowerShell module, ensuring that PowerShellGet is also installed if necessary.
if (-not (Get-Module -Name PowerShellGet -ListAvailable)) { 
    Install-Module -Name PowerShellGet -Force -AllowClobber 
} 
Install-Module -Name MicrosoftTeams -Force -AllowClobber 
Import-Module -Name MicrosoftTeams 
if (Get-Module -Name MicrosoftTeams -ListAvailable) { 
    Write-Host "Microsoft Teams PowerShell module installed successfully." 
} else { 
    Write-Host "Failed to install Microsoft Teams PowerShell module." 
}

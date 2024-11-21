# This script checks if the Microsoft Teams PowerShell module is installed and uninstalls it if present.
if (Get-Module -Name MicrosoftTeams -ListAvailable) { 
    Write-Host "Microsoft Teams PowerShell module is installed. Proceeding with uninstallation..."  
    Uninstall-Module MicrosoftTeams -Allversions  -Force 
    Write-Host "Microsoft Teams PowerShell module has been successfully uninstalled." 
} else { 
    Write-Host "Microsoft Teams PowerShell module is not installed." 
}

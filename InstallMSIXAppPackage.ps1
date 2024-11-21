# This command installs an MSIX app package on the system without requiring a license agreement.
DISM /Online /Add-ProvisionedAppxPackage /PackagePath:"<path_of_the_MSIX_app>" /SkipLicense

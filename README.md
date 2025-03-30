# Microsoft 365 Apps Deployment Toolkit

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)](https://github.com/PowerShell/PowerShell)
[![Microsoft 365](https://img.shields.io/badge/Microsoft_365-compatible-brightgreen.svg)](https://www.microsoft.com/microsoft-365)
[![Maintained](https://img.shields.io/badge/Maintained-yes-green.svg)](https://github.com/sargeschultz11/M365-Deployment-Toolkit/graphs/commit-activity)
[![GitHub release](https://img.shields.io/github/release/sargeschultz11/M365-Deployment-Toolkit.svg)](https://GitHub.com/sargeschultz11/M365-Deployment-Toolkit/releases/)
[![GitHub issues](https://img.shields.io/github/issues/sargeschultz11/M365-Deployment-Toolkit.svg)](https://GitHub.com/sargeschultz11/M365-Deployment-Toolkit/issues/)
[![GitHub contributors](https://img.shields.io/github/contributors/sargeschultz11/M365-Deployment-Toolkit.svg)](https://GitHub.com/sargeschultz11/M365-Deployment-Toolkit/graphs/contributors/)
[![made-with-bash](https://img.shields.io/badge/Made%20with-PowerShell-1f425f.svg)](https://www.microsoft.com/powershell)

## Overview

This toolkit provides an automated solution for deploying Microsoft 365 Apps, Visio, and Project in managed environments. It combines the power of the Office Deployment Tool (ODT) with PowerShell scripting to provide a reliable, consistent installation experience across multiple Microsoft products.

Whether you're an MSP managing multiple clients, an IT professional in an enterprise environment, or deploying via Microsoft Intune, this toolkit simplifies the deployment process and handles common challenges like removing pre-installed consumer Office versions.

## Features

- Automatically downloads the latest Office Deployment Tool
- Includes ready-to-use XML configurations for Microsoft 365 Apps, Visio, and Project
- **Consumer Office Removal**: Automatically detects and removes pre-installed consumer versions of Office
- Comprehensive detection of all Microsoft 365 product IDs across multiple languages
- Provides detailed logging for troubleshooting
- Handles the entire installation process including cleanup
- Supports system restart option after installation

## Package Contents

```
M365-Intune-Deployment-Toolkit/
├── Install-Microsoft365Apps.ps1   # Main PowerShell installation script
├── config/                        # Configuration files directory
│   ├── install-office365.xml      # Microsoft 365 Apps installation configuration
│   ├── install-visio.xml          # Visio installation configuration
│   ├── install-project.xml        # Project installation configuration
│   ├── uninstall-office365.xml    # Office uninstallation configuration
│   ├── uninstall-visio.xml        # Visio uninstallation configuration
│   ├── uninstall-project.xml      # Project uninstallation configuration
│   └── uninstall-consumer.xml     # Consumer Office uninstallation configuration
├── assets/                        # Assets for Company Portal app listings
│   ├── office365-icon.png         # Microsoft 365 Apps icon
│   ├── visio-icon.png             # Visio icon
│   ├── project-icon.png           # Project icon
│   └── microsoft365-logo.png      # Microsoft 365 logo
└── README.md                      # Documentation
```

## Quick Start

This toolkit is designed to be immediately usable without modifications:

1. Download the current release from the releases page, fork, or clone this repo
2. Choose your deployment method:
   - Run directly on systems that need Microsoft 365 Apps
   - Deploy through your RMM tool or management platform
   - Package for Intune or SCCM deployment
   - Include in your client onboarding automation

The toolkit includes:
- A comprehensive PowerShell script with robust product detection
- Pre-configured XML files for common Microsoft 365 products
- Consumer Office detection and removal capabilities


## Requirements

- Windows 10 1809 or later (64-bit)
- Administrator privileges
- Internet connection to download the Office Deployment Tool
- Approximately 4GB of free disk space

## Customization

### Modifying Office Configuration

The included XML files are ready to use, but you can (and I strongly recommend you do) customize them for your organization's needs:

- **install-office365.xml**: Microsoft 365 Apps for Business installation configuration
- **install-visio.xml**: Visio installation configuration
- **install-project.xml**: Project installation configuration
- **uninstall-office365.xml**: Office uninstallation configuration
- **uninstall-visio.xml**: Visio uninstallation configuration
- **uninstall-project.xml**: Project uninstallation configuration
- **uninstall-consumer.xml**: Consumer Office uninstallation configuration

You can create custom configurations using the [Microsoft 365 Apps Admin Center](https://config.office.com/officeSettings/configurations).

### Script Parameters

The PowerShell script supports several parameters that can be modified:

- **ConfigXMLPath**: Path to the XML configuration file
- **OfficeInstallDownloadPath**: Temporary directory for installation files
- **Restart**: Switch to enable automatic restart after installation
- **RemoveConsumerOffice**: Switch to enable detection and removal of pre-installed consumer Office products

## Deployment Options

### For MSPs and IT Professionals

This toolkit can be used in various deployment scenarios:

1. **Manual Deployment**: Run the script directly on client machines
2. **RMM Tool Deployment**: Deploy via your RMM platform of choice (ConnectWise Automate, Datto RMM, NinjaOne, etc.)
3. **SCCM/ConfigMgr Deployment**: Package as an application in Configuration Manager
4. **Group Policy Deployment**: Deploy via startup/login scripts or scheduled tasks

The script accepts parameters allowing for flexible deployment options and silent operation.

### For Intune Deployment

1. Package this entire directory using the Microsoft Win32 Content Prep Tool
2. Upload the .intunewin file to Intune
3. Configure the application with these settings:

> **Company Portal Assets**: The `assets` folder contains icons and description content you can use when publishing your app in the Company Portal. This helps provide a professional and consistent appearance for end-users.
   - **Standard installation command:** 
   
   ```powershell
   PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "Install-Microsoft365Apps.ps1" -ConfigXMLPath "config\install-office365.xml"
   ```

   - **Installation command with consumer Office removal:** 
   
   ```powershell
   PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "Install-Microsoft365Apps.ps1" -ConfigXMLPath "config\install-office365.xml" -RemoveConsumerOffice
   ```

   - **Uninstall command:** 
   
   ```powershell
   PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "Install-Microsoft365Apps.ps1" -ConfigXMLPath "config\uninstall-office365.xml"
   ```

   - **For Visio installation:**
   
   ```powershell
   PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "Install-Microsoft365Apps.ps1" -ConfigXMLPath "config\install-visio.xml"
   ```

   - **For Project installation:**
   
   ```powershell
   PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "Install-Microsoft365Apps.ps1" -ConfigXMLPath "config\install-project.xml"
   ```

   - **Detection method:** 
      - Registry key for Microsoft 365 Apps
      - **Rule Type:** `Registry`
      - **Key Path:** `HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365BusinessRetail - en-us`
      - **Value Name:** `DisplayName`
      - **Detection Method:** `Value exists`
      - **Associated with a 32-bit app on 64-bit clients:** `No`
      
      > **Important Note:** You must update the Key Path to match the specific product you are installing. For example:
      > - For Visio: `HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\VisioProRetail - en-us`
      > - For Project: `HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ProjectProRetail - en-us`
      > 
      > Also note that the language code (`en-us` in the examples) should match the language you're deploying.

## Product Detection

The script includes a comprehensive `Test-ProductInstalled` function that detects:
- Microsoft 365 Apps (various SKUs)
- Office 2019, 2021, and 2024 products (Retail & Volume)
- Visio (all editions)
- Project (all editions)
- Installations in multiple languages
- Various registry locations for complete coverage

## Consumer Office Detection and Removal

The script now includes functionality to detect and remove pre-installed consumer versions of Office that often come on OEM Windows installations. This helps prevent conflicts and ensures a clean installation environment - a common challenge for MSPs and IT departments dealing with new devices.

The consumer Office detection and removal functionality:

- Detects Home & Student, Personal, and OEM versions of Office
- Identifies Microsoft Store Office apps
- Automatically removes them using the Office Deployment Tool
- Provides detailed logging of the removal process
- Continues with installation even if some components can't be fully removed

To use this feature, simply add the `-RemoveConsumerOffice` switch to your installation command:

```powershell
# Example for direct execution or RMM deployment
.\Install-Microsoft365Apps.ps1 -ConfigXMLPath "config\install-office365.xml" -RemoveConsumerOffice

# Example with additional options for automated deployment
.\Install-Microsoft365Apps.ps1 -ConfigXMLPath "config\install-office365.xml" -RemoveConsumerOffice -OfficeInstallDownloadPath "C:\Temp\OfficeInstall"
```

This feature is particularly valuable for:
- Client onboarding scenarios
- New device deployments
- Standardizing environments across multiple clients
- Transitioning from consumer to business Microsoft 365 subscriptions

## Troubleshooting

Installation logs are saved in the system temp directory with the naming pattern:
`Office365Install_YYYYMMDD_HHMMSS.log`

These logs contain detailed information about each step of the installation process and any errors that may have occurred.

## Integration and Branding

The `assets` folder contains resources you can use in various deployment scenarios:

- **Icons**: High-quality application icons for Microsoft 365 Apps, Visio, and Project
- **Description Files**: Ready-to-use description text for each application

These assets help create a more professional and consistent experience for end users, whether you're:

- Creating documentation for clients
- Building a self-service portal
- Configuring Intune Company Portal listings
- Adding applications to SCCM Software Center
- Creating custom deployment notifications

## Getting Started with This Repository

### Using Releases

For those who want to quickly deploy without modifications:

1. Navigate to the "Releases" section of this repository
2. Download the latest release package which contains all necessary files
3. Use as-is or replace the XML files with your customized versions
4. Package with the Intune Win32 Content Prep Tool and deploy

### Forking the Repository

For organizations wanting to maintain their own version:

1. Fork this repository to your organization's GitHub account
2. Customize the XML files and scripts for your specific environment
3. Maintain your fork with your specific organizational requirements
4. Pull updates from the main repository as needed

## Contributing

Contributions to improve this package are welcome! Here's how to contribute:

1. Fork the repository
2. Create a feature branch 
3. Make your changes and test thoroughly
4. Commit your changes 
5. Push to the branch 
6. Create a new Pull Request

All contributions should focus on improving reliability, adding features, or enhancing documentation.

## Maintenance and Lifecycle Policy

* Included product IDs will remain in the script as long as they are supported by Microsoft
* When a Microsoft product reaches End of Life (EOL) and is no longer supported by Microsoft, its product ID will be removed from the script in a future update
* Major updates typically coincide with new Office release cycles
* Critical bugs will be addressed as quickly as possible

## License

This package is licensed under the MIT License - see the LICENSE file for details

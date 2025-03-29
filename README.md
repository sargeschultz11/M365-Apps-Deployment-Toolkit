# Microsoft 365 Apps Deployment Package for Intune

## Overview

This package provides an automated solution for deploying Microsoft 365 Apps, Visio, and Project via Microsoft Intune as Win32 applications. It combines the power of the Office Deployment Tool (ODT) with PowerShell scripting to provide a reliable, consistent installation experience across multiple Microsoft products.

## Features

- Automatically downloads the latest Office Deployment Tool
- Includes ready-to-use XML configurations for Microsoft 365 Apps, Visio, and Project
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
│   └── uninstall-office365.xml    # Office uninstallation configuration
├── assets/                        # Assets for Company Portal app listings
│   ├── office365-icon.png         # Microsoft 365 Apps icon
│   ├── visio-icon.png             # Visio icon
│   ├── project-icon.png           # Project icon
│   └── microsoft365-logo.png      # Microsoft 365 logo
└── README.md                      # Documentation
```

## Quick Start

This package is designed to be immediately usable without modifications:

1. Download the current release from the releases page, fork, or clone this repo
2. Package the entire contents using the Microsoft Win32 Content Prep Tool
3. Upload to Intune and deploy

The package includes:
- A comprehensive PowerShell script with robust product detection
- Pre-configured XML files for common Microsoft 365 products


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

You can create custom configurations using the [Microsoft 365 Apps Admin Center](https://config.office.com/officeSettings/configurations).

### Script Parameters

The PowerShell script supports several parameters that can be modified:

- **ConfigXMLPath**: Path to the XML configuration file
- **OfficeInstallDownloadPath**: Temporary directory for installation files
- **Restart**: Switch to enable automatic restart after installation

## Deployment in Intune

1. Package this entire directory using the Microsoft Win32 Content Prep Tool
2. Upload the .intunewin file to Intune
3. Configure the application with these settings:

> **Company Portal Assets**: The `assets` folder contains icons and description content you can use when publishing your app in the Company Portal. This helps provide a professional and consistent appearance for end-users.
   - **Install command:** 
   
   ```powershell
   PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "Install-Microsoft365Apps.ps1" -ConfigXMLPath "config\install-office365.xml"
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

The script includes a comprehensive `Test-OfficeInstalled` function that detects:
- Microsoft 365 Apps (various SKUs)
- Office 2019, 2021, and 2024 products (Retail & Volume)
- Visio (all editions)
- Project (all editions)
- Installations in multiple languages
- Various registry locations for complete coverage

## Troubleshooting

Installation logs are saved in the system temp directory with the naming pattern:
`Office365Install_YYYYMMDD_HHMMSS.log`

These logs contain detailed information about each step of the installation process and any errors that may have occurred.

## Company Portal Integration

The `assets` folder contains resources you can use to enhance the appearance of your Microsoft 365 apps in the Intune Company Portal:

- **Icons**: High-quality application icons for Microsoft 365 Apps, Visio, and Project
- **Description Files**: Ready-to-use description text for each application



These assets help create a more professional and consistent experience for end users accessing applications through the Company Portal.

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
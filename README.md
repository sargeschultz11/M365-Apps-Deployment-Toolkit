# 🚀 Microsoft 365 Apps Deployment Toolkit

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)](https://github.com/PowerShell/PowerShell)
[![Microsoft 365](https://img.shields.io/badge/Microsoft_365-compatible-brightgreen.svg)](https://www.microsoft.com/microsoft-365)
[![Maintained](https://img.shields.io/badge/Maintained-yes-green.svg)](https://github.com/sargeschultz11/M365-Apps-Deployment-Toolkit/graphs/commit-activity)
[![GitHub release](https://img.shields.io/github/release/sargeschultz11/M365-Apps-Deployment-Toolkit.svg)](https://GitHub.com/sargeschultz11/M365-Apps-Deployment-Toolkit/releases/)
[![GitHub issues](https://img.shields.io/github/issues/sargeschultz11/M365-Apps-Deployment-Toolkit.svg)](https://GitHub.com/sargeschultz11/M365-Apps-Deployment-Toolkit/issues/)
[![GitHub contributors](https://img.shields.io/github/contributors/sargeschultz11/M365-Apps-Deployment-Toolkit.svg)](https://GitHub.com/sargeschultz11/M365-Apps-Deployment-Toolkit/graphs/contributors/)
[![made-with-bash](https://img.shields.io/badge/Made%20with-PowerShell-1f425f.svg)](https://www.microsoft.com/powershell)
[![Version](https://img.shields.io/badge/Version-1.3.1-brightgreen.svg)]()

## 📋 Overview

This toolkit provides an automated solution for deploying Microsoft 365 Apps, Visio, and Project in managed environments. It combines the power of the Office Deployment Tool (ODT) with PowerShell scripting to provide a reliable, consistent installation experience across multiple Microsoft products.

Whether you're an MSP managing multiple clients, an IT professional in an enterprise environment, or deploying via Microsoft Intune, this toolkit simplifies the deployment process and handles common challenges like removing pre-installed consumer Office versions.

## ✨ Features

- 🔄 Automatically downloads the latest Office Deployment Tool
- 📝 Includes ready-to-use XML configurations with detailed documentation and comments for Microsoft 365 Apps, Visio, and Project
- 📘 **Enhanced XML Files**: Fully documented configuration files with clear sections and explanatory comments for easier customization
- 🧹 **Consumer Office Removal**: Automatically detects and removes pre-installed consumer versions of Office
- 🔍 **Comprehensive Detection**: Identifies Microsoft 365 products, Office 2010-2024, Visio, and Project across multiple languages
- 🔧 **Multiple Installation Options**: Supports standard installation, forcing installation, uninstalling existing products, detection-only mode, and more
- 🔎 **Microsoft Store Apps Detection**: Identifies Office apps installed from the Microsoft Store
- 📊 Provides detailed logging for troubleshooting
- 🧰 Handles the entire installation process including cleanup
- 🔄 Supports system restart option after installation

## 📦 Package Contents

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

## 🚀 Quick Start

This toolkit is designed to be immediately usable without modifications:

1. Download the current release from the releases page, fork, or clone this repo
2. Choose your deployment method:
   - ⚡ Run directly on systems that need Microsoft 365 Apps
   - 🔄 Deploy through your RMM tool or management platform
   - 📱 Package for Intune or SCCM deployment
   - 🔁 Include in your client onboarding automation

The toolkit includes:
- 💻 A comprehensive PowerShell script with robust product detection
- 📄 Pre-configured XML files with detailed comments for common Microsoft 365 products
- 🧹 Consumer Office detection and removal capabilities


## 📋 Requirements

- 💻 Windows 10 1809 or later (64-bit)
- 🔑 Administrator privileges
- 🌐 Internet connection to download the Office Deployment Tool
- 💾 Approximately 4GB of free disk space

## ⚙️ Customization

### 📝 Modifying Office Configuration

The included XML files are ready to use, but you can (and I strongly recommend you do) customize them for your organization's needs:

- **install-office365.xml**: Microsoft 365 Apps for Business installation configuration
- **install-visio.xml**: Visio installation configuration
- **install-project.xml**: Project installation configuration
- **uninstall-office365.xml**: Office uninstallation configuration
- **uninstall-visio.xml**: Visio uninstallation configuration
- **uninstall-project.xml**: Project uninstallation configuration
- **uninstall-consumer.xml**: Consumer Office uninstallation configuration

Each XML file now includes detailed comments explaining all available options and their purpose, making customization much easier even for those less familiar with ODT configurations.

You can create custom configurations using the [Microsoft 365 Apps Admin Center](https://config.office.com/officeSettings/configurations).

### 🔧 Script Parameters

The PowerShell script supports several parameters that can be modified:

- **ConfigXMLPath**: Path to the XML configuration file
- **OfficeInstallDownloadPath**: Temporary directory for installation files
- **Uninstall**: Switch to be used for standard uninstallation process
- **Restart**: Switch to enable automatic restart after installation
- **RemoveConsumerOffice**: Switch to enable detection and removal of pre-installed consumer Office products
- **Force**: Switch to install even if Office is already detected on the system
- **UninstallExisting**: Switch to remove existing Office products before installation
- **SkipIfInstalled**: Switch to skip installation if any Office product is detected (same as default behavior, but makes intent explicit)
- **DetectOnly**: Switch to only detect Office products without installation (useful for inventory)

## 🚀 Deployment Options

### 💼 For MSPs and IT Professionals

This toolkit can be used in various deployment scenarios:

1. 🖱️ **Manual Deployment**: Run the script directly on client machines
2. 🔄 **RMM Tool Deployment**: Deploy via your RMM platform of choice (ConnectWise Automate, Datto RMM, NinjaOne, etc.)
3. 🔧 **SCCM/ConfigMgr Deployment**: Package as an application in Configuration Manager
4. 🔐 **Group Policy Deployment**: Deploy via startup/login scripts or scheduled tasks

The script accepts parameters allowing for flexible deployment options and silent operation.

### 📱 For Intune Deployment

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

   - **Installation command that forces installation even if Office is already installed:**
   
   ```powershell
   PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "Install-Microsoft365Apps.ps1" -ConfigXMLPath "config\install-office365.xml" -Force
   ```

   - **Installation command that uninstalls existing Office products first:**
   
   ```powershell
   PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "Install-Microsoft365Apps.ps1" -ConfigXMLPath "config\install-office365.xml" -UninstallExisting
   ```

   - **Installation command for detection only:**
   
   ```powershell
   PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "Install-Microsoft365Apps.ps1" -DetectOnly
   ```

   - **Uninstall command:** 
   
   ```powershell
   PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "Install-Microsoft365Apps.ps1" -ConfigXMLPath "config\uninstall-office365.xml" -Uninstall
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

## 🔍 Product Detection

The script includes a comprehensive `Test-ProductInstalled` function that detects:

- ✅ Microsoft 365 Apps (various SKUs)
- ✅ Office 2010, 2013, 2016, 2019, 2021, and 2024 products (Retail & Volume)
- ✅ Visio 2010-2024 (Standard and Professional editions)
- ✅ Project 2010-2024 (Standard and Professional editions)
- ✅ Installations in multiple languages
- ✅ Various registry locations for complete coverage
- ✅ Click-to-Run installations via configuration keys
- ✅ Office application executable paths
- ✅ Office install root registry keys
- ✅ Generic detection through uninstall registry entries

The detection logic uses multiple methods to ensure the most accurate results, including:
1. 🔑 Product-specific uninstall registry keys
2. 🔄 Click-to-Run configuration detection
3. 📂 Office application executable paths
4. 📋 Office install root registry keys
5. 🔎 Generic search through uninstall registry entries

## 🧹 Consumer Office Detection and Removal

The script includes enhanced functionality to detect and remove pre-installed consumer versions of Office that often come on OEM Windows installations. This helps prevent conflicts and ensures a clean installation environment - a common challenge for MSPs and IT departments dealing with new devices.

The consumer Office detection and removal functionality:

- 🔍 Detects Home & Student, Personal, and OEM versions of Office
- 🔎 Identifies Microsoft Store Office apps
- 🗑️ Automatically removes them using the Office Deployment Tool
- 📝 Provides detailed logging of the removal process
- ➡️ Continues with installation even if some components can't be fully removed

To use this feature, simply add the `-RemoveConsumerOffice` switch to your installation command:

```powershell
# Example for direct execution or RMM deployment
.\Install-Microsoft365Apps.ps1 -ConfigXMLPath "config\install-office365.xml" -RemoveConsumerOffice

# Example with additional options for automated deployment
.\Install-Microsoft365Apps.ps1 -ConfigXMLPath "config\install-office365.xml" -RemoveConsumerOffice -OfficeInstallDownloadPath "C:\Temp\OfficeInstall"
```

This feature is particularly valuable for:
- 🆕 Client onboarding scenarios
- 💻 New device deployments
- 🔄 Standardizing environments across multiple clients
- 🔀 Transitioning from consumer to business Microsoft 365 subscriptions

## 🔧 Installation Options

The script supports several installation modes to handle different scenarios:

- ✅ **Default behavior**: Checks if Office is installed and skips installation if detected
- 🔨 **Force Installation**: Uses the `-Force` parameter to explicitly force installation even if Office is already detected
- 🗑️ **Uninstall Existing**: Uses the `-UninstallExisting` parameter to remove existing Office products before installation
- ⏭️ **Skip If Installed**: Uses the `-SkipIfInstalled` parameter to explicitly skip installation if Office is detected (same as default behavior, but makes intent explicit)
- 🔍 **Detection Only**: Uses the `-DetectOnly` parameter to only check for Office products without installing (useful for inventory)

Example usage:

```powershell
# Only detect Office products and report findings (useful for inventory)
.\Install-Microsoft365Apps.ps1 -DetectOnly

# Uninstall existing Office products before installing new ones
.\Install-Microsoft365Apps.ps1 -UninstallExisting -ConfigXMLPath "config\install-office365.xml"

# Force installation even if Office is already installed
.\Install-Microsoft365Apps.ps1 -Force -ConfigXMLPath "config\install-office365.xml"

# Skip installation if Office is already detected (default behavior made explicit)
.\Install-Microsoft365Apps.ps1 -SkipIfInstalled -ConfigXMLPath "config\install-office365.xml"
```

## 🔧 Troubleshooting

Installation logs are saved in the system temp directory with the naming pattern:
`Microsoft365Install_YYYYMMDD_HHMMSS.log`

These logs contain detailed information about each step of the installation process and any errors that may have occurred.

## 🎨 Integration and Branding

The `assets` folder contains resources you can use in various deployment scenarios:

- 🖼️ **Icons**: High-quality application icons for Microsoft 365 Apps, Visio, and Project
- 📄 **Description Files**: Ready-to-use description text for each application

These assets help create a more professional and consistent experience for end users, whether you're:

- 📚 Creating documentation for clients
- 🖥️ Building a self-service portal
- 📱 Configuring Intune Company Portal listings
- 🧰 Adding applications to SCCM Software Center
- 📨 Creating custom deployment notifications

## 🚀 Getting Started with This Repository

### 📦 Using Releases

For those who want to quickly deploy without modifications:

1. Navigate to the "Releases" section of this repository
2. Download the latest release package which contains all necessary files
3. Use as-is or replace the XML files with your customized versions
4. Package with the Intune Win32 Content Prep Tool and deploy

### 🍴 Forking the Repository

For organizations wanting to maintain their own version:

1. Fork this repository to your organization's GitHub account
2. Customize the XML files and scripts for your specific environment
3. Maintain your fork with your specific organizational requirements
4. Pull updates from the main repository as needed

## 👥 Contributing

Contributions to improve this package are welcome! Here's how to contribute:

1. 🍴 Fork the repository
2. 🌿 Create a feature branch 
3. 🔧 Make your changes and test thoroughly
4. 💾 Commit your changes 
5. 📤 Push to the branch 
6. 📩 Create a new Pull Request

All contributions should focus on improving reliability, adding features, or enhancing documentation.


## 📜 License

This package is licensed under the MIT License - see the LICENSE file for details
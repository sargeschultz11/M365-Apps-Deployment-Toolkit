# 🚀 Microsoft 365 Apps Deployment Toolkit

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)](https://github.com/PowerShell/PowerShell)
[![Microsoft 365](https://img.shields.io/badge/Microsoft_365-compatible-brightgreen.svg)](https://www.microsoft.com/microsoft-365)
[![Maintained](https://img.shields.io/badge/Maintained-yes-green.svg)](https://github.com/sargeschultz11/M365-Apps-Deployment-Toolkit/graphs/commit-activity)
[![GitHub release](https://img.shields.io/github/release/sargeschultz11/M365-Apps-Deployment-Toolkit.svg)](https://GitHub.com/sargeschultz11/M365-Apps-Deployment-Toolkit/releases/)
[![GitHub issues](https://img.shields.io/github/issues/sargeschultz11/M365-Apps-Deployment-Toolkit.svg)](https://GitHub.com/sargeschultz11/M365-Apps-Deployment-Toolkit/issues/)
[![GitHub contributors](https://img.shields.io/github/contributors/sargeschultz11/M365-Apps-Deployment-Toolkit.svg)](https://GitHub.com/sargeschultz11/M365-Apps-Deployment-Toolkit/graphs/contributors/)
[![made-with-bash](https://img.shields.io/badge/Made%20with-PowerShell-1f425f.svg)](https://www.microsoft.com/powershell)


## 📋 Overview

This toolkit provides an automated solution for deploying Microsoft 365 Apps, Visio, and Project in managed environments. It combines the power of the Office Deployment Tool (ODT) with PowerShell scripting to provide a reliable, consistent installation experience across multiple Microsoft products.

Whether you're an MSP managing multiple clients, an IT professional in an enterprise environment, or deploying via Microsoft Intune, this toolkit simplifies the deployment process and handles common challenges like removing pre-installed consumer Office versions.

**[Visit our comprehensive Wiki for detailed documentation →](https://github.com/sargeschultz11/M365-Apps-Deployment-Toolkit/wiki)**

## ✨ Key Features

- 🔄 Automatically downloads the latest Office Deployment Tool
- 📝 Includes ready-to-use XML configurations with detailed documentation
- 🧹 Automatically detects and removes pre-installed consumer versions of Office
- 🔍 Comprehensive detection for all Microsoft 365 products and versions
- 🔧 Multiple installation options (standard, force, uninstall-existing, detection-only)
- 🔎 Detects Microsoft Store Office apps
- 📊 Detailed logging for troubleshooting
- 🧰 Handles the entire installation process including cleanup

## 📦 Quick Start

1. Download the current release from the [releases page](https://github.com/sargeschultz11/M365-Apps-Deployment-Toolkit/releases), fork, or clone this repo
2. Choose your deployment method:
   - ⚡ [Manual deployment](https://github.com/sargeschultz11/M365-Apps-Deployment-Toolkit/wiki/Manual-Deployment)
   - 🔄 [RMM deployment](https://github.com/sargeschultz11/M365-Apps-Deployment-Toolkit/wiki/RMM-Deployment)
   - 📱 [Intune deployment](https://github.com/sargeschultz11/M365-Apps-Deployment-Toolkit/wiki/Intune-Deployment)
   - 🔁 [SCCM deployment](https://github.com/sargeschultz11/M365-Apps-Deployment-Toolkit/wiki/SCCM-Deployment)

## 🚀 Basic Usage Examples

```powershell
# Standard installation
.\Install-Microsoft365Apps.ps1 -ConfigXMLPath "config\install-office365.xml"

# Installation with consumer Office removal
.\Install-Microsoft365Apps.ps1 -ConfigXMLPath "config\install-office365.xml" -RemoveConsumerOffice

# Only detect Office products (no installation)
.\Install-Microsoft365Apps.ps1 -DetectOnly

# Uninstall existing Office products before installing
.\Install-Microsoft365Apps.ps1 -ConfigXMLPath "config\install-office365.xml" -UninstallExisting

# For Visio installation
.\Install-Microsoft365Apps.ps1 -ConfigXMLPath "config\install-visio.xml"
```

## 📋 Requirements

- 💻 Windows 10 1809 or later (64-bit)
- 🔑 Administrator privileges
- 🌐 Internet connection to download the Office Deployment Tool
- 💾 Approximately 4GB of free disk space

## 📚 Documentation

For detailed information on installation options, customization, deployment scenarios, and troubleshooting, please refer to our [Wiki](https://github.com/sargeschultz11/M365-Apps-Deployment-Toolkit/wiki).

- [Getting Started](https://github.com/sargeschultz11/M365-Apps-Deployment-Toolkit/wiki/Getting-Started)
- [Configuration Guide](https://github.com/sargeschultz11/M365-Apps-Deployment-Toolkit/wiki/Configuration-Guide)
- [Script Parameters](https://github.com/sargeschultz11/M365-Apps-Deployment-Toolkit/wiki/Script-Parameters)
- [Troubleshooting](https://github.com/sargeschultz11/M365-Apps-Deployment-Toolkit/wiki/Troubleshooting-Guide)

## 👥 Contributing

Contributions to improve this package are welcome! Please see our [Contributing Guidelines](CONTRIBUTING.md) for more information.

## 📜 License

This package is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🛡️ Security

For information about security policies and procedures, please see our [Security Policy](SECURITY.md).

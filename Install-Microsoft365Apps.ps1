<#
.SYNOPSIS
    Installs Microsoft 365 applications using the Office Deployment Tool (ODT).

.DESCRIPTION
    This script automates the download and installation of Microsoft 365 applications.
    It first checks if Office is already installed and can skip installation if detected.
    It downloads the latest Office Deployment Tool and uses the provided XML configuration.
    Now includes functionality to detect and remove pre-installed consumer versions of Office.
    Improved to detect older Office versions and offer multiple handling options.

.PARAMETER ConfigXMLPath
    Path to the XML configuration file for Office installation.
    Default: Uses the XML file in the same directory as the script.

.PARAMETER OfficeInstallDownloadPath
    Directory where temporary installation files will be stored.
    Default: A subfolder in the system temp directory.

.PARAMETER Restart
    Switch parameter. If specified, the computer will restart after successful installation.

.PARAMETER RemoveConsumerOffice
    Switch parameter. If specified, the script will detect and remove consumer versions of Office
    that commonly come pre-installed on Windows devices before installing.

.PARAMETER Force
    Switch parameter. If specified, will force installation even if Office is already installed.

.PARAMETER UninstallExisting
    Switch parameter. If specified, will uninstall existing Office products before installation.

.PARAMETER SkipIfInstalled
    Switch parameter. If specified, will skip installation if any Office product is detected.
    This is the default behavior unless -Force or -UninstallExisting is specified.

.PARAMETER DetectOnly
    Switch parameter. If specified, will only detect Office products and exit.
    Useful for inventory or reporting purposes.

.EXAMPLE
    .\Install-Microsoft365Apps.ps1 -ConfigXMLPath "C:\Config\install-office365.xml" -RemoveConsumerOffice
    # Standard installation with removal of pre-installed consumer Office versions

.EXAMPLE
    .\Install-Microsoft365Apps.ps1 -DetectOnly
    # Only detects Office products and reports findings without installation

.EXAMPLE
    .\Install-Microsoft365Apps.ps1 -UninstallExisting -ConfigXMLPath "C:\Config\install-office365.xml"
    # Uninstalls any existing Office products before installing

.EXAMPLE
    .\Install-Microsoft365Apps.ps1 -Force -ConfigXMLPath "C:\Config\install-office365.xml"
    # Forces installation even if Office is already installed

.EXAMPLE
    .\Install-Microsoft365Apps.ps1 -SkipIfInstalled
    # Checks if Office is installed and exits without installing if detected

.NOTES
    Version: 1.2
    Last Updated: March 30, 2025
#>

[CmdletBinding(DefaultParameterSetName='DefaultAction')]
param(
    [Parameter()]
    [String]$ConfigXMLPath = "$PSScriptRoot\config\install-office365.xml",
    
    [Parameter()]
    [String]$OfficeInstallDownloadPath = "$env:TEMP\Office365Install",
    
    [Parameter()]
    [Switch]$Restart,
    
    [Parameter()]
    [Switch]$RemoveConsumerOffice,
    
    [Parameter(ParameterSetName='Force')]
    [Switch]$Force,
    
    [Parameter(ParameterSetName='Uninstall')]
    [Switch]$UninstallExisting,
    
    [Parameter(ParameterSetName='Skip')]
    [Switch]$SkipIfInstalled,
    
    [Parameter(ParameterSetName='DetectOnly')]
    [Switch]$DetectOnly
)

$ErrorActionPreference = 'Stop'
$LogFilePath = Join-Path -Path $env:TEMP -ChildPath "Microsoft365Install_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} catch {
    Write-Warning "TLS 1.2 is not supported on this system. This script may fail!"
}

#######################################################################
#                          HELPER FUNCTIONS                           #
#######################################################################

function Write-Log {
    [CmdletBinding()]
    param (
        [string]$Message,
        [string]$Type = "INFO"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Type] $Message"
    
    switch ($Type) {
        "ERROR" { Write-Host $LogMessage -ForegroundColor Red }
        "WARNING" { Write-Host $LogMessage -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $LogMessage -ForegroundColor Green }
        default { Write-Host $LogMessage }
    }
    
    Add-Content -Path $LogFilePath -Value $LogMessage
}

#======================================================================#
# Validates the script is running with administrator privileges        #
#======================================================================#
function Test-RunningAsAdmin {
    [CmdletBinding()]
    param()
    
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object System.Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

#======================================================================#
# Obtains the latest Office Deployment Tool download URL               #
#======================================================================#
function Get-OfficeDeploymentTool {
    [CmdletBinding()]
    param()
    
    $Uri = 'https://www.microsoft.com/en-us/download/details.aspx?id=49117'
    
    try {
        Write-Log -Message "Retrieving Office Deployment Tool download URL..."
        $MSWebPage = Invoke-WebRequest -Uri $Uri -UseBasicParsing
        $DownloadURL = $MSWebPage.Links | Where-Object { $_.href -like "*officedeploymenttool*.exe" } | 
                       Select-Object -ExpandProperty href -First 1
        
        if (-not $DownloadURL) {
            Write-Log -Message "Unable to find the download link for the Office Deployment Tool." -Type "ERROR"
            exit 1
        }
        Write-Log -Message "Found ODT download URL: $DownloadURL"
        return $DownloadURL
    }
    catch {
        Write-Log -Message "Unable to connect to the Microsoft website: $_" -Type "ERROR"
        exit 1
    }
}

#======================================================================#
# Downloads file from specified URL to the destination path            #
#======================================================================#
function Invoke-Download {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [String]$URL,
        
        [Parameter(Mandatory=$true)]
        [String]$Path
    )
    
    Write-Log -Message "Downloading from '$URL'..."
    
    try {
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $URL -OutFile $Path -UseBasicParsing
        $ProgressPreference = 'Continue'
    }
    catch {
        Write-Log -Message "Download failed: $_" -Type "ERROR"
        if (Test-Path -Path $Path) {
            Remove-Item $Path -Force -ErrorAction SilentlyContinue
        }
        exit 1
    }
    
    if (-not (Test-Path $Path)) {
        Write-Log -Message "File failed to download to '$Path'." -Type "ERROR"
        exit 1
    }
    
    return $Path
}

#####################################################################
#                MICROSOFT PRODUCT DETECTION LOGIC                  #
#####################################################################
function Test-ProductInstalled {
    [CmdletBinding()]
    param(
        [Parameter()]
        [Switch]$ReturnDetectedProducts,
        
        [Parameter()]
        [Switch]$IncludeDetailsInLog
    )
    
    Write-Log -Message "Checking for Microsoft Office, Visio, or Project installation..."
    
    $DetectedProducts = @()
    
    #======================================================================================================================#
    # Array of common Microsoft 365 product IDs                                                                            #
    # Add or remove as necessary                                                                                           #
    # Logs can show installation not found but may be false a positive detection of failure if your ID is not in the array #
    #======================================================================================================================#
    $ProductIDs = @(
        # Microsoft 365 Apps
        "O365BusinessRetail",
        "O365ProPlusRetail",
        "O365BusinessEEANoTeamsRetail",
        "O365ProPlusEEANoTeamsRetail",
        "O365HomePremRetail",
        "O365SmallBusPremRetail",
        
        # Office 2024/2021/2019 Products (Retail & Volume)
        "ProPlus2024Retail", "ProPlus2024Volume", 
        "ProPlus2021Retail", "ProPlus2021Volume",
        "ProPlus2019Retail", "ProPlus2019Volume",
        "Standard2024Retail", "Standard2024Volume",
        "Standard2021Retail", "Standard2021Volume",
        "Standard2019Retail", "Standard2019Volume",
        "StandardSPLA2021Volume",
        
        # Office 2016 Products (Retail & Volume)
        "ProPlus2016Retail", "ProPlus2016Volume",
        "Standard2016Retail", "Standard2016Volume",
        "ProPlusRetail", "ProPlusVolume",
        "StandardRetail", "StandardVolume",
        "HomeBusinessRetail", "HomeStudentRetail",
        "PersonalRetail",
        
        # Office 2013 Products (Retail & Volume)
        "ProPlus2013Retail", "ProPlus2013Volume",
        "Standard2013Retail", "Standard2013Volume",
        "HomeBusinessRetail2013", "HomeStudentRetail2013",
        "PersonalRetail2013",
        
        # Office 2010 Products (Retail & Volume)
        "ProPlus2010Retail", "ProPlus2010Volume",
        "Standard2010Retail", "Standard2010Volume",
        "HomeAndBusiness2010Retail", "HomeAndStudent2010Retail",
        
        # Visio Products (2010-2024)
        "VisioProRetail", "VisioStdRetail",
        "VisioPro2024Retail", "VisioPro2024Volume", "VisioStd2024Retail", "VisioStd2024Volume",
        "VisioPro2021Retail", "VisioPro2021Volume", "VisioStd2021Retail", "VisioStd2021Volume",
        "VisioPro2019Retail", "VisioPro2019Volume", "VisioStd2019Retail", "VisioStd2019Volume",
        "VisioPro2016Retail", "VisioPro2016Volume", "VisioStd2016Retail", "VisioStd2016Volume",
        "VisioPro2013Retail", "VisioPro2013Volume", "VisioStd2013Retail", "VisioStd2013Volume",
        "VisioPro2010Retail", "VisioPro2010Volume", "VisioStd2010Retail", "VisioStd2010Volume",
        
        # Project Products (2010-2024)
        "ProjectProRetail", "ProjectStdRetail",
        "ProjectPro2024Retail", "ProjectPro2024Volume", "ProjectStd2024Retail", "ProjectStd2024Volume",
        "ProjectPro2021Retail", "ProjectPro2021Volume", "ProjectStd2021Retail", "ProjectStd2021Volume",
        "ProjectPro2019Retail", "ProjectPro2019Volume", "ProjectStd2019Retail", "ProjectStd2019Volume",
        "ProjectPro2016Retail", "ProjectPro2016Volume", "ProjectStd2016Retail", "ProjectStd2016Volume",
        "ProjectPro2013Retail", "ProjectPro2013Volume", "ProjectStd2013Retail", "ProjectStd2013Volume",
        "ProjectPro2010Retail", "ProjectPro2010Volume", "ProjectStd2010Retail", "ProjectStd2010Volume"
    )
    
    $CommonLanguages = @("en-us", "en-gb", "fr-fr", "de-de", "es-es", "it-it", "ja-jp", "ko-kr", "pt-br", "ru-ru", "zh-cn", "zh-tw")
    
    #===============================================================#
    # Method 1: Check for product-specific uninstall registry keys  #
    #===============================================================#
    foreach ($ProductID in $ProductIDs) {
        $NeutralKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$ProductID"
        if (Test-Path -Path $NeutralKey) {
            $KeyInfo = Get-ItemProperty -Path $NeutralKey -ErrorAction SilentlyContinue
            $ProductDetails = [PSCustomObject]@{
                ProductID = $ProductID
                DisplayName = $KeyInfo.DisplayName
                Version = $KeyInfo.DisplayVersion
                InstallDate = $KeyInfo.InstallDate
                RegistryKey = $NeutralKey
                DetectionMethod = "Registry Key (Neutral)"
            }
            $DetectedProducts += $ProductDetails
            
            if ($IncludeDetailsInLog) {
                Write-Log -Message "Found Microsoft product: $($KeyInfo.DisplayName) ($ProductID), Version: $($KeyInfo.DisplayVersion)" -Type "SUCCESS"
            } else {
                Write-Log -Message "Found Microsoft product: $($KeyInfo.DisplayName)" -Type "SUCCESS"
            }
        }
        
        foreach ($Language in $CommonLanguages) {
            $LocalizedKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$ProductID - $Language"
            if (Test-Path -Path $LocalizedKey) {
                $KeyInfo = Get-ItemProperty -Path $LocalizedKey -ErrorAction SilentlyContinue
                $ProductDetails = [PSCustomObject]@{
                    ProductID = $ProductID
                    Language = $Language
                    DisplayName = $KeyInfo.DisplayName
                    Version = $KeyInfo.DisplayVersion
                    InstallDate = $KeyInfo.InstallDate
                    RegistryKey = $LocalizedKey
                    DetectionMethod = "Registry Key (Localized)"
                }
                $DetectedProducts += $ProductDetails
                
                if ($IncludeDetailsInLog) {
                    Write-Log -Message "Found Microsoft product: $($KeyInfo.DisplayName) ($ProductID - $Language), Version: $($KeyInfo.DisplayVersion)" -Type "SUCCESS"
                } else {
                    Write-Log -Message "Found Microsoft product: $($KeyInfo.DisplayName)" -Type "SUCCESS"
                }
            }
        }
    }
    
    #===============================================================#
    # Method 2: Check for Click-to-Run configuration                #
    #===============================================================#
    $ClickToRunPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    if (Test-Path -Path $ClickToRunPath) {
        $ClickToRun = Get-ItemProperty -Path $ClickToRunPath -ErrorAction SilentlyContinue
        
        if ($ClickToRun.ProductReleaseIds) {
            $C2RProducts = $ClickToRun.ProductReleaseIds -split ","
            $Platform = if ($ClickToRun.Platform) { $ClickToRun.Platform } else { "Unknown" }
            $VersionToReport = if ($ClickToRun.VersionToReport) { $ClickToRun.VersionToReport } else { "Unknown" }
            
            Write-Log -Message "Found Click-to-Run installation (Version: $VersionToReport, Platform: $Platform) with products:" -Type "SUCCESS"
            
            foreach ($Product in $C2RProducts) {
                $ProductDetails = [PSCustomObject]@{
                    ProductID = $Product
                    DisplayName = "Microsoft Office Click-to-Run Product"
                    Version = $VersionToReport
                    InstallDate = $null
                    Platform = $Platform
                    RegistryKey = $ClickToRunPath
                    DetectionMethod = "Click-to-Run Configuration"
                }
                $DetectedProducts += $ProductDetails
                
                Write-Log -Message "  - $Product" -Type "SUCCESS"
            }
        }
    }
    
    #===============================================================#
    # Method 3: Check for Office application executable paths       #
    #===============================================================#
    $AppPaths = @(
        @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe"; App = "Word" },
        @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Excel.exe"; App = "Excel" },
        @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Outlook.exe"; App = "Outlook" },
        @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Powerpnt.exe"; App = "PowerPoint" },
        @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Visio.exe"; App = "Visio" },
        @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WinProj.exe"; App = "Project" }
    )
    
    foreach ($AppPath in $AppPaths) {
        if (Test-Path -Path $AppPath.Path) {
            $ExePath = (Get-ItemProperty -Path $AppPath.Path -ErrorAction SilentlyContinue).'(Default)'
            if ($ExePath -and (Test-Path -Path $ExePath)) {
                $Version = "Unknown"
                try {
                    $FileVersionInfo = (Get-Item $ExePath).VersionInfo
                    $Version = "$($FileVersionInfo.FileMajorPart).$($FileVersionInfo.FileMinorPart).$($FileVersionInfo.FileBuildPart).$($FileVersionInfo.FilePrivatePart)"
                } catch {
                    # Can't get version info
                }
                
                $ProductDetails = [PSCustomObject]@{
                    ProductID = "Unknown"
                    DisplayName = "Microsoft $($AppPath.App)"
                    Version = $Version
                    InstallDate = $null
                    ExePath = $ExePath
                    RegistryKey = $AppPath.Path
                    DetectionMethod = "App Paths Registry"
                }
                $DetectedProducts += $ProductDetails
                
                Write-Log -Message "Found Microsoft $($AppPath.App) at $ExePath (Version: $Version)" -Type "SUCCESS"
            }
        }
    }
    
    #===============================================================#
    # Method 4: Check for Office install root registry keys         #
    #===============================================================#
    $OfficeVersions = @(
        @{ Version = "16.0"; Name = "Office 2024/2021/2019/2016" },
        @{ Version = "15.0"; Name = "Office 2013" },
        @{ Version = "14.0"; Name = "Office 2010" },
        @{ Version = "12.0"; Name = "Office 2007" }
    )
    
    foreach ($OfficeVersion in $OfficeVersions) {
        $OfficeKeys = @(
            "HKLM:\SOFTWARE\Microsoft\Office\$($OfficeVersion.Version)\Common\InstallRoot",
            "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\$($OfficeVersion.Version)\Common\InstallRoot"
        )
        
        foreach ($OfficeKey in $OfficeKeys) {
            if (Test-Path -Path $OfficeKey) {
                $InstallPath = (Get-ItemProperty -Path $OfficeKey -ErrorAction SilentlyContinue).Path
                if ($InstallPath -and (Test-Path -Path $InstallPath)) {
                    $ProductDetails = [PSCustomObject]@{
                        ProductID = "Unknown"
                        DisplayName = $OfficeVersion.Name
                        Version = $OfficeVersion.Version
                        InstallDate = $null
                        InstallPath = $InstallPath
                        RegistryKey = $OfficeKey
                        DetectionMethod = "Office InstallRoot Registry"
                    }
                    $DetectedProducts += $ProductDetails
                    
                    Write-Log -Message "Found $($OfficeVersion.Name) installation at $InstallPath" -Type "SUCCESS"
                }
            }
        }
    }
    
    #===============================================================#
    # Method 5: Generic search through uninstall registry entries   #
    #===============================================================#
    $UninstallKeys = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    foreach ($Key in $UninstallKeys) {
        $Results = Get-ItemProperty $Key -ErrorAction SilentlyContinue | Where-Object { 
            ($_.DisplayName -like "*Microsoft 365*") -or 
            ($_.DisplayName -like "*Office 365*") -or
            ($_.DisplayName -like "*Microsoft Office*") -or
            ($_.DisplayName -like "*Microsoft Visio*") -or
            ($_.DisplayName -like "*Microsoft Project*") -or
            ($_.UninstallString -like "*OfficeClickToRun.exe*") -or
            ($_.PSChildName -like "*O365*") -or
            ($_.PSChildName -like "*Office16*") -or
            ($_.PSChildName -like "*Visio*") -or
            ($_.PSChildName -like "*Project*")
        }
        
        if ($Results) {
            foreach ($Result in $Results) {
                # Check if this product was already detected by other methods
                $AlreadyDetected = $false
                foreach ($DetectedProduct in $DetectedProducts) {
                    if ($DetectedProduct.RegistryKey -eq $Result.PSPath) {
                        $AlreadyDetected = $true
                        break
                    }
                }
                
                if (-not $AlreadyDetected) {
                    $ProductDetails = [PSCustomObject]@{
                        ProductID = $Result.PSChildName
                        DisplayName = $Result.DisplayName
                        Version = $Result.DisplayVersion
                        InstallDate = $Result.InstallDate
                        UninstallString = $Result.UninstallString
                        RegistryKey = $Result.PSPath
                        DetectionMethod = "Generic Uninstall Registry"
                    }
                    $DetectedProducts += $ProductDetails
                    
                    Write-Log -Message "Found Microsoft product: $($Result.DisplayName) (Version: $($Result.DisplayVersion))" -Type "SUCCESS"
                }
            }
        }
    }
    
    if ($DetectedProducts.Count -eq 0) {
        Write-Log -Message "No Microsoft Office, Visio, or Project installation detected" -Type "WARNING"
        
        if ($ReturnDetectedProducts) {
            return @($false, $DetectedProducts)
        } else {
            return $false
        }
    } else {
        Write-Log -Message "Total Microsoft products detected: $($DetectedProducts.Count)" -Type "SUCCESS"
        
        if ($ReturnDetectedProducts) {
            return @($true, $DetectedProducts)
        } else {
            return $true
        }
    }
}

#####################################################################
#                CONSUMER OFFICE DETECTION AND REMOVAL              #
#####################################################################
function Test-ConsumerOfficeInstalled {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter()]
        [Switch]$ReturnDetectedProducts
    )
    
    Write-Log -Message "Checking for consumer versions of Office..."
    
    $DetectedProducts = @()
    
    # Array of common consumer Office product IDs
    $ConsumerProductIDs = @(
        # Office Home & Student Products
        "HomeBusiness2024Retail", "HomeStudent2024Retail",
        "HomeBusiness2021Retail", "HomeStudent2021Retail",
        "HomeBusiness2019Retail", "HomeStudent2019Retail", 
        "HomeBusinessRetail", "HomeStudentRetail",
        "HomeBusiness2016Retail", "HomeStudent2016Retail",
        "HomeBusiness2013Retail", "HomeStudent2013Retail",
        "HomeAndBusiness2010Retail", "HomeAndStudent2010Retail",
        
        # Personal/Family/Individual Products
        "Personal2024Retail", "Personal2021Retail",
        "Personal2019Retail", "PersonalRetail",
        "Personal2016Retail", "Personal2013Retail",
        
        # Pre-installed/OEM Products
        "HomeStudentVNextRetail",  # Sometimes used for pre-installed versions
        "HomeBusinessVNextRetail",  # Sometimes used for pre-installed versions
        "O365HomePremRetail",      # Office 365 Home
        "MondoRetail",             # Office Mondo (OEM version)
        "O365SmallBusPremRetail"   # Office 365 Small Business Premium
    )
    
    $CommonLanguages = @("en-us", "en-gb", "fr-fr", "de-de", "es-es", "it-it", "ja-jp", "ko-kr", "pt-br", "ru-ru", "zh-cn", "zh-tw")
    
    #===============================================================#
    # Method 1: Check for product-specific uninstall registry keys  #
    #===============================================================#
    foreach ($ProductID in $ConsumerProductIDs) {
        $NeutralKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$ProductID"
        if (Test-Path -Path $NeutralKey) {
            $KeyInfo = Get-ItemProperty -Path $NeutralKey -ErrorAction SilentlyContinue
            $ProductDetails = [PSCustomObject]@{
                ProductID = $ProductID
                DisplayName = $KeyInfo.DisplayName
                Version = $KeyInfo.DisplayVersion
                InstallDate = $KeyInfo.InstallDate
                RegistryKey = $NeutralKey
                DetectionMethod = "Registry Key (Consumer)"
            }
            $DetectedProducts += $ProductDetails
            
            Write-Log -Message "Found consumer Office product: $($KeyInfo.DisplayName) ($ProductID)" -Type "WARNING"
        }
        
        foreach ($Language in $CommonLanguages) {
            $LocalizedKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$ProductID - $Language"
            if (Test-Path -Path $LocalizedKey) {
                $KeyInfo = Get-ItemProperty -Path $LocalizedKey -ErrorAction SilentlyContinue
                $ProductDetails = [PSCustomObject]@{
                    ProductID = $ProductID
                    Language = $Language
                    DisplayName = $KeyInfo.DisplayName
                    Version = $KeyInfo.DisplayVersion
                    InstallDate = $KeyInfo.InstallDate
                    RegistryKey = $LocalizedKey
                    DetectionMethod = "Registry Key (Consumer Localized)"
                }
                $DetectedProducts += $ProductDetails
                
                Write-Log -Message "Found consumer Office product: $($KeyInfo.DisplayName) ($ProductID - $Language)" -Type "WARNING"
            }
        }
    }
    
    #===============================================================#
    # Method 2: Check Click-to-Run configuration for consumer SKUs  #
    #===============================================================#
    $ClickToRunPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    if (Test-Path -Path $ClickToRunPath) {
        $ClickToRun = Get-ItemProperty -Path $ClickToRunPath -ErrorAction SilentlyContinue
        
        if ($ClickToRun.ProductReleaseIds) {
            $products = $ClickToRun.ProductReleaseIds -split ","
            foreach ($product in $products) {
                foreach ($consumerID in $ConsumerProductIDs) {
                    if ($product -like "*$consumerID*") {
                        $ProductDetails = [PSCustomObject]@{
                            ProductID = $product
                            DisplayName = "Microsoft Office Click-to-Run Consumer Product"
                            Version = $ClickToRun.VersionToReport
                            InstallDate = $null
                            Platform = $ClickToRun.Platform
                            RegistryKey = $ClickToRunPath
                            DetectionMethod = "Click-to-Run Configuration (Consumer)"
                        }
                        $DetectedProducts += $ProductDetails
                        
                        Write-Log -Message "Found consumer Office product in Click-to-Run configuration: $product" -Type "WARNING"
                    }
                }
            }
        }
    }
    
    #===============================================================#
    # Method 3: Check for preloaded Microsoft 365 apps              #
    #===============================================================#
    $PreinstalledKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    foreach ($Key in $PreinstalledKeys) {
        $Results = Get-ItemProperty $Key -ErrorAction SilentlyContinue | Where-Object { 
            ($_.DisplayName -like "*Microsoft Office*") -and 
            ($_.DisplayName -notlike "*Microsoft Office 365*") -and
            (
                ($_.DisplayName -like "*Home*") -or 
                ($_.DisplayName -like "*Student*") -or 
                ($_.DisplayName -like "*Personal*") -or
                ($_.DisplayName -like "*Family*")
            )
        }
        
        if ($Results) {
            foreach ($Result in $Results) {
                $ProductDetails = [PSCustomObject]@{
                    ProductID = $Result.PSChildName
                    DisplayName = $Result.DisplayName
                    Version = $Result.DisplayVersion
                    InstallDate = $Result.InstallDate
                    UninstallString = $Result.UninstallString
                    RegistryKey = $Result.PSPath
                    DetectionMethod = "Generic Uninstall Registry (Consumer)"
                }
                $DetectedProducts += $ProductDetails
                
                Write-Log -Message "Found consumer Office product: $($Result.DisplayName) (Version: $($Result.DisplayVersion))" -Type "WARNING"
            }
        }
    }
    
    #===============================================================#
    # Method 4: Check for Microsoft Store Office apps               #
    #===============================================================#
    $AppxPackages = Get-AppxPackage -Name "*Microsoft.Office*" -ErrorAction SilentlyContinue
    if ($AppxPackages) {
        $ConsumerAppx = $AppxPackages | Where-Object { 
            $_.Name -notlike "*Microsoft.Office.Desktop*" -and ($_.Name -like "*Microsoft.Office*")
        }
        
        if ($ConsumerAppx) {
            foreach ($Appx in $ConsumerAppx) {
                $ProductDetails = [PSCustomObject]@{
                    ProductID = $Appx.Name
                    DisplayName = $Appx.Name
                    Version = $Appx.Version
                    InstallDate = $Appx.InstallDate
                    PackageFullName = $Appx.PackageFullName
                    DetectionMethod = "Microsoft Store App"
                }
                $DetectedProducts += $ProductDetails
                
                Write-Log -Message "Found Microsoft Store Office app: $($Appx.Name) (Version: $($Appx.Version))" -Type "WARNING"
            }
        }
    }
    
    if ($DetectedProducts.Count -eq 0) {
        Write-Log -Message "No consumer versions of Office detected"
        
        if ($ReturnDetectedProducts) {
            return @($false, $DetectedProducts)
        } else {
            return $false
        }
    } else {
        Write-Log -Message "Found $($DetectedProducts.Count) consumer Office products" -Type "WARNING"
        
        if ($ReturnDetectedProducts) {
            return @($true, $DetectedProducts)
        } else {
            return $true
        }
    }
}

function Remove-ConsumerOffice {
    [CmdletBinding()]
    param(
        [Parameter()]
        [String]$OfficeInstallDownloadPath = "$env:TEMP\Office365Install",
        
        [Parameter()]
        [String]$UninstallXMLPath = ""
    )
    
    Write-Log -Message "Starting removal of consumer Office products..."
    
    if ([string]::IsNullOrEmpty($UninstallXMLPath)) {
        $UninstallXMLPath = "$OfficeInstallDownloadPath\uninstall-consumer.xml"
        
        # Check if we have a pre-existing uninstall XML in the config directory
        $ConfigUninstallPath = "$PSScriptRoot\config\uninstall-consumer.xml"
        if (Test-Path $ConfigUninstallPath) {
            Write-Log -Message "Using existing uninstall configuration from $ConfigUninstallPath"
            Copy-Item -Path $ConfigUninstallPath -Destination $UninstallXMLPath -Force
        } else {
            # Create the uninstall XML content
            Write-Log -Message "Creating uninstall configuration file at $UninstallXMLPath"
            
            $UninstallXMLContent = @"
<?xml version="1.0" encoding="UTF-8"?>
<Configuration>
 <Remove All="True">
    <Product ID="HomeBusiness2024Retail">
      <Language ID="MatchOS" />
    </Product>
    <Product ID="HomeStudent2024Retail">
      <Language ID="MatchOS" />
    </Product>
    <Product ID="HomeBusiness2021Retail">
      <Language ID="MatchOS" />
    </Product>
    <Product ID="HomeStudent2021Retail">
      <Language ID="MatchOS" />
    </Product>
    <Product ID="HomeBusiness2019Retail">
      <Language ID="MatchOS" />
    </Product>
    <Product ID="HomeBusinessRetail">
      <Language ID="MatchOS" />
    </Product>
    <Product ID="HomeStudent2019Retail">
      <Language ID="MatchOS" />
    </Product>
    <Product ID="HomeStudentRetail">
      <Language ID="MatchOS" />
    </Product>
    <Product ID="Personal2024Retail">
      <Language ID="MatchOS" />
    </Product>
    <Product ID="Personal2021Retail">
      <Language ID="MatchOS" />
    </Product>
    <Product ID="Personal2019Retail">
      <Language ID="MatchOS" />
    </Product>
    <Product ID="PersonalRetail">
      <Language ID="MatchOS" />
    </Product>
    <Product ID="HomeStudentVNextRetail">
      <Language ID="MatchOS" />
    </Product>
    <Product ID="HomeBusinessVNextRetail">
      <Language ID="MatchOS" />
    </Product>
    <Product ID="MondoRetail">
      <Language ID="MatchOS" />
    </Product>
    <Product ID="O365HomePremRetail">
      <Language ID="MatchOS" />
    </Product>
    <Product ID="O365SmallBusPremRetail">
      <Language ID="MatchOS" />
    </Product>
 </Remove>
 <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@
            
            Set-Content -Path $UninstallXMLPath -Value $UninstallXMLContent -Force
        }
    }
    
    if (-not (Test-Path "$OfficeInstallDownloadPath\setup.exe")) {
        Write-Log -Message "Office Deployment Tool not found, downloading..."
        $ODTInstallLink = Get-OfficeDeploymentTool
        $ODTPath = "$OfficeInstallDownloadPath\ODTSetup.exe"
        Invoke-Download -URL $ODTInstallLink -Path $ODTPath
        
        Write-Log -Message "Extracting the Office Deployment Tool..."
        try {
            Start-Process $ODTPath -ArgumentList "/quiet /extract:$OfficeInstallDownloadPath" -Wait -NoNewWindow
            Write-Log -Message "Office Deployment Tool extracted successfully"
        }
        catch {
            Write-Log -Message "Failed to extract Office Deployment Tool: $_" -Type "ERROR"
            return $false
        }
    }
    
    try {
        Write-Log -Message "Running ODT to remove consumer Office products..."
        $setupPath = "$OfficeInstallDownloadPath\setup.exe"
        
        $process = Start-Process -FilePath $setupPath -ArgumentList "/configure `"$UninstallXMLPath`"" -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -ne 0) {
            Write-Log -Message "ODT uninstallation process completed with exit code: $($process.ExitCode)" -Type "WARNING"
        } else {
            Write-Log -Message "ODT uninstallation process completed successfully"
        }
    }
    catch {
        Write-Log -Message "Error during uninstallation: $_" -Type "ERROR"
        return $false
    }
    
    try {
        $AppxPackages = Get-AppxPackage -Name "*Microsoft.Office*" -ErrorAction SilentlyContinue
        if ($AppxPackages) {
            $ConsumerAppx = $AppxPackages | Where-Object { 
                $_.Name -notlike "*Microsoft.Office.Desktop*" -and ($_.Name -like "*Microsoft.Office*")
            }
            
            if ($ConsumerAppx) {
                Write-Log -Message "Removing Microsoft Store Office apps..."
                $ConsumerAppx | ForEach-Object {
                    Write-Log -Message "Removing $($_.Name)..."
                    Remove-AppxPackage -Package $_.PackageFullName -ErrorAction SilentlyContinue
                }
            }
        }
    }
    catch {
        Write-Log -Message "Error removing Microsoft Store Office apps: $_" -Type "WARNING"
    }
    
    if (Test-ConsumerOfficeInstalled) {
        Write-Log -Message "Some consumer Office products may still be installed" -Type "WARNING"
        return $false
    }
    else {
        Write-Log -Message "Successfully removed consumer Office products"
        return $true
    }
}

#####################################################################
#                     UNINSTALL EXISTING PRODUCTS                   #
#####################################################################
function Uninstall-ExistingOfficeProducts {
    [CmdletBinding()]
    param(
        [Parameter()]
        [String]$OfficeInstallDownloadPath = "$env:TEMP\Office365Install"
    )
    
    Write-Log -Message "Starting uninstallation of existing Office products..."
    
    # First check if there's anything to uninstall
    $detectionResult = Test-ProductInstalled -ReturnDetectedProducts
    $isInstalled = $detectionResult[0]
    $detectedProducts = $detectionResult[1]
    
    if (-not $isInstalled) {
        Write-Log -Message "No Office products detected to uninstall"
        return $true
    }
    
    # Get uninstall files from config folder
    $uninstallFiles = @{
        "Office365" = "$PSScriptRoot\config\uninstall-office365.xml"
        "Visio" = "$PSScriptRoot\config\uninstall-visio.xml"
        "Project" = "$PSScriptRoot\config\uninstall-project.xml"
    }
    
    $anyFailures = $false
    
    # Check which products are installed and uninstall them
    $officeInstalled = $detectedProducts | Where-Object { 
        $_.ProductID -like "*O365*" -or 
        $_.ProductID -like "*ProPlus*" -or 
        $_.ProductID -like "*Standard*" -or
        $_.DisplayName -like "*Microsoft 365*" -or
        $_.DisplayName -like "*Office 365*" -or
        $_.DisplayName -like "*Microsoft Office*"
    }
    
    $visioInstalled = $detectedProducts | Where-Object {
        $_.ProductID -like "*Visio*" -or
        $_.DisplayName -like "*Visio*"
    }
    
    $projectInstalled = $detectedProducts | Where-Object {
        $_.ProductID -like "*Project*" -or
        $_.DisplayName -like "*Project*"
    }
    
    if (-not (Test-Path "$OfficeInstallDownloadPath\setup.exe")) {
        Write-Log -Message "Office Deployment Tool not found, downloading..."
        $ODTInstallLink = Get-OfficeDeploymentTool
        $ODTPath = "$OfficeInstallDownloadPath\ODTSetup.exe"
        Invoke-Download -URL $ODTInstallLink -Path $ODTPath
        
        Write-Log -Message "Extracting the Office Deployment Tool..."
        try {
            Start-Process $ODTPath -ArgumentList "/quiet /extract:$OfficeInstallDownloadPath" -Wait -NoNewWindow
            Write-Log -Message "Office Deployment Tool extracted successfully"
        }
        catch {
            Write-Log -Message "Failed to extract Office Deployment Tool: $_" -Type "ERROR"
            return $false
        }
    }
    
    $setupPath = "$OfficeInstallDownloadPath\setup.exe"
    
    # Uninstall Office
    if ($officeInstalled) {
        if (Test-Path $uninstallFiles["Office365"]) {
            Write-Log -Message "Uninstalling Microsoft Office/Microsoft 365 Apps..."
            try {
                $uninstallXMLPath = "$OfficeInstallDownloadPath\uninstall-office.xml"
                Copy-Item -Path $uninstallFiles["Office365"] -Destination $uninstallXMLPath -Force
                
                $process = Start-Process -FilePath $setupPath -ArgumentList "/configure `"$uninstallXMLPath`"" -Wait -PassThru -NoNewWindow
                
                if ($process.ExitCode -ne 0) {
                    Write-Log -Message "Office uninstallation completed with exit code: $($process.ExitCode)" -Type "WARNING"
                    $anyFailures = $true
                } else {
                    Write-Log -Message "Office uninstallation completed successfully"
                }
            }
            catch {
                Write-Log -Message "Error during Office uninstallation: $_" -Type "ERROR"
                $anyFailures = $true
            }
        } else {
            Write-Log -Message "Uninstall configuration for Office not found at $($uninstallFiles["Office365"])" -Type "WARNING"
            $anyFailures = $true
        }
    }
    
    # Uninstall Visio
    if ($visioInstalled) {
        if (Test-Path $uninstallFiles["Visio"]) {
            Write-Log -Message "Uninstalling Microsoft Visio..."
            try {
                $uninstallXMLPath = "$OfficeInstallDownloadPath\uninstall-visio.xml"
                Copy-Item -Path $uninstallFiles["Visio"] -Destination $uninstallXMLPath -Force
                
                $process = Start-Process -FilePath $setupPath -ArgumentList "/configure `"$uninstallXMLPath`"" -Wait -PassThru -NoNewWindow
                
                if ($process.ExitCode -ne 0) {
                    Write-Log -Message "Visio uninstallation completed with exit code: $($process.ExitCode)" -Type "WARNING"
                    $anyFailures = $true
                } else {
                    Write-Log -Message "Visio uninstallation completed successfully"
                }
            }
            catch {
                Write-Log -Message "Error during Visio uninstallation: $_" -Type "ERROR"
                $anyFailures = $true
            }
        } else {
            Write-Log -Message "Uninstall configuration for Visio not found at $($uninstallFiles["Visio"])" -Type "WARNING"
            $anyFailures = $true
        }
    }
    
    # Uninstall Project
    if ($projectInstalled) {
        if (Test-Path $uninstallFiles["Project"]) {
            Write-Log -Message "Uninstalling Microsoft Project..."
            try {
                $uninstallXMLPath = "$OfficeInstallDownloadPath\uninstall-project.xml"
                Copy-Item -Path $uninstallFiles["Project"] -Destination $uninstallXMLPath -Force
                
                $process = Start-Process -FilePath $setupPath -ArgumentList "/configure `"$uninstallXMLPath`"" -Wait -PassThru -NoNewWindow
                
                if ($process.ExitCode -ne 0) {
                    Write-Log -Message "Project uninstallation completed with exit code: $($process.ExitCode)" -Type "WARNING"
                    $anyFailures = $true
                } else {
                    Write-Log -Message "Project uninstallation completed successfully"
                }
            }
            catch {
                Write-Log -Message "Error during Project uninstallation: $_" -Type "ERROR"
                $anyFailures = $true
            }
        } else {
            Write-Log -Message "Uninstall configuration for Project not found at $($uninstallFiles["Project"])" -Type "WARNING"
            $anyFailures = $true
        }
    }
    
    # Check if all products were uninstalled successfully
    $detectionResult = Test-ProductInstalled
    if ($detectionResult) {
        Write-Log -Message "Some Office products remain installed after uninstallation" -Type "WARNING"
        $anyFailures = $true
    } else {
        Write-Log -Message "All Office products successfully uninstalled"
    }
    
    return -not $anyFailures
}

################################################################################
#                          MAIN INSTALLATION PROCESS                           #
################################################################################
Write-Log -Message "Starting Microsoft 365 Apps installation script (Version 1.2)"

#===============================================================================#
# Step 1: Validate prerequisites                                                #
#===============================================================================#
if (-not (Test-RunningAsAdmin)) {
    Write-Log -Message "This script requires administrator privileges" -Type "ERROR"
    exit 1
}

#===============================================================================#
# Step 2: Check existing Office installation (BEFORE downloading anything)      #
#===============================================================================#
Write-Log -Message "Checking for existing Office installation..."
$detectionResult = Test-ProductInstalled -ReturnDetectedProducts -IncludeDetailsInLog
$isOfficeInstalled = $detectionResult[0]
$detectedProducts = $detectionResult[1]

# Handle parameter logic
if ($DetectOnly) {
    Write-Log -Message "Running in detection-only mode. Skipping installation."
    
    if ($isOfficeInstalled) {
        Write-Log -Message "Office products detected. Details:"
        foreach ($product in $detectedProducts) {
            $details = "  - $($product.DisplayName)"
            if ($product.Version) { $details += " (Version: $($product.Version))" }
            Write-Log -Message $details
        }
    } else {
        Write-Log -Message "No Office products detected on this system."
    }
    
    exit 0
}

# Handle skip if installed mode
if ($SkipIfInstalled -and $isOfficeInstalled) {
    Write-Log -Message "Office products are already installed and -SkipIfInstalled specified."
    Write-Log -Message "Skipping installation."
    exit 0
}

# Default behavior (skip if installed unless Force or UninstallExisting is specified)
if ($isOfficeInstalled -and (-not $Force) -and (-not $UninstallExisting)) {
    Write-Log -Message "Office products are already installed and -Force/-UninstallExisting not specified."
    Write-Log -Message "Skipping installation. Use -Force to install anyway or -UninstallExisting to remove existing products first."
    exit 0
}

# At this point, we proceed with installation because:
# 1. No Office is installed, OR
# 2. Force parameter was specified, OR
# 3. UninstallExisting parameter was specified

# If we need to continue with installation, validate the configuration file
if (-not (Test-Path -Path $ConfigXMLPath)) {
    Write-Log -Message "Configuration XML file not found at '$ConfigXMLPath'" -Type "ERROR"
    exit 1
}

if (-not (Test-Path $OfficeInstallDownloadPath)) {
    New-Item -Path $OfficeInstallDownloadPath -ItemType Directory -Force | Out-Null
    Write-Log -Message "Created installation directory at '$OfficeInstallDownloadPath'"
}

try {
    [xml]$xml = Get-Content -Path $ConfigXMLPath
    Write-Log -Message "Successfully loaded configuration XML"
}
catch {
    Write-Log -Message "Invalid XML configuration file: $_" -Type "ERROR"
    exit 1
}

#===============================================================================#
# Step 3: Handle existing Office installation if needed                         #
#===============================================================================#
if ($isOfficeInstalled -and $UninstallExisting) {
    Write-Log -Message "Existing Office products detected and -UninstallExisting specified."
    if (-not (Uninstall-ExistingOfficeProducts -OfficeInstallDownloadPath $OfficeInstallDownloadPath)) {
        Write-Log -Message "Some products could not be uninstalled. Continuing with installation anyway." -Type "WARNING"
    }
}

#===============================================================================#
# Step 4: Check and remove consumer Office products if requested                #
#===============================================================================#
if ($RemoveConsumerOffice) {
    Write-Log -Message "Checking for consumer Office products..."
    if (Test-ConsumerOfficeInstalled) {
        Write-Log -Message "Consumer Office products detected, attempting to remove..."
        if (-not (Remove-ConsumerOffice -OfficeInstallDownloadPath $OfficeInstallDownloadPath)) {
            Write-Log -Message "Unable to completely remove some consumer Office products. Proceeding with installation anyway." -Type "WARNING"
        }
    }
    else {
        Write-Log -Message "No consumer Office products detected, proceeding with installation"
    }
}

#===============================================================================#
# Step 5: Download and extract Office Deployment Tool                           #
#===============================================================================#
Write-Log -Message "Getting Office Deployment Tool download URL..."
$ODTInstallLink = Get-OfficeDeploymentTool

Write-Log -Message "Downloading the Office Deployment Tool..."
$ODTPath = "$OfficeInstallDownloadPath\ODTSetup.exe"
Invoke-Download -URL $ODTInstallLink -Path $ODTPath

Write-Log -Message "Extracting the Office Deployment Tool..."
try {
    Start-Process $ODTPath -ArgumentList "/quiet /extract:$OfficeInstallDownloadPath" -Wait -NoNewWindow
    Write-Log -Message "Office Deployment Tool extracted successfully"
}
catch {
    Write-Log -Message "Failed to extract Office Deployment Tool: $_" -Type "ERROR"
    exit 1
}

#===============================================================================#
# Step 6: Install Microsoft 365 Apps                                            #
#===============================================================================#
Write-Log -Message "Copying configuration XML to installation directory..."
Copy-Item -Path $ConfigXMLPath -Destination "$OfficeInstallDownloadPath\configuration.xml" -Force

Write-Log -Message "Starting Microsoft 365 installation..."
try {
    $setupPath = "$OfficeInstallDownloadPath\setup.exe"
    $configPath = "$OfficeInstallDownloadPath\configuration.xml"
    
    $process = Start-Process -FilePath $setupPath -ArgumentList "/configure `"$configPath`"" -Wait -PassThru -NoNewWindow
    
    if ($process.ExitCode -ne 0) {
        Write-Log -Message "Installation failed with exit code: $($process.ExitCode)" -Type "ERROR"
        exit $process.ExitCode
    }
    
    Write-Log -Message "Microsoft 365 installation process completed with exit code: $($process.ExitCode)"
}
catch {
    Write-Log -Message "Error during installation: $_" -Type "ERROR"
    exit 1
}

#===============================================================================#
# Step 7: Verify installation and cleanup                                       #
#===============================================================================#
Write-Log -Message "Verifying installation..."
if (Test-ProductInstalled) {
    Write-Log -Message "Microsoft 365 installed successfully!" -Type "SUCCESS"
    
    # Only clean up installation files after successful verification
    Write-Log -Message "Cleaning up installation files..."
    Remove-Item -Path $OfficeInstallDownloadPath -Force -Recurse -ErrorAction SilentlyContinue
    
    if ($Restart) {
        Write-Log -Message "Restarting computer in 60 seconds..."
        Start-Process shutdown.exe -ArgumentList "-r -t 60" -NoNewWindow
    }
    
    exit 0
}
else {
    Write-Log -Message "Microsoft 365 was not detected after installation! Verification failed." -Type "WARNING"
    Write-Log -Message "Leaving installation files at $OfficeInstallDownloadPath for troubleshooting." -Type "WARNING"
    exit 1
}
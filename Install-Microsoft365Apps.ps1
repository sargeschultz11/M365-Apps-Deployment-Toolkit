<#
.SYNOPSIS
    Installs Microsoft 365 applications using the Office Deployment Tool (ODT).

.DESCRIPTION
    This script automates the download and installation of Microsoft 365 applications.
    It downloads the latest Office Deployment Tool and uses the provided XML configuration.

.PARAMETER ConfigXMLPath
    Path to the XML configuration file for Office installation.
    Default: Uses the XML file in the same directory as the script.

.PARAMETER OfficeInstallDownloadPath
    Directory where temporary installation files will be stored.
    Default: A subfolder in the system temp directory.

.PARAMETER Restart
    Switch parameter. If specified, the computer will restart after successful installation.

.EXAMPLE
    .\Install-Microsoft365Apps.ps1 -ConfigXMLPath "C:\Config\install-office365.xml"

.NOTES
    Version: 1.0
    Last Updated: March 28, 2025
#>

[CmdletBinding()]
param(
    [Parameter()]
    [String]$ConfigXMLPath = "$PSScriptRoot\config\install-office365.xml",
    
    [Parameter()]
    [String]$OfficeInstallDownloadPath = "$env:TEMP\Office365Install",
    
    [Parameter()]
    [Switch]$Restart
)

$ErrorActionPreference = 'Stop'
$LogFilePath = Join-Path -Path $env:TEMP -ChildPath "Office365Install_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} catch {
    Write-Warning "TLS 1.2 is not supported on this system. This script may fail!"
}

#######################################################################
#                          HELPER FUNCTIONS                           #
#######################################################################

function Write-Log {
    param (
        [string]$Message,
        [string]$Type = "INFO"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Type] $Message"
    
    switch ($Type) {
        "ERROR" { Write-Host $LogMessage -ForegroundColor Red }
        "WARNING" { Write-Host $LogMessage -ForegroundColor Yellow }
        default { Write-Host $LogMessage }
    }
    
    Add-Content -Path $LogFilePath -Value $LogMessage
}

#======================================================================#
# Validates the script is running with administrator privileges        #
#======================================================================#
function Test-RunningAsAdmin {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object System.Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

#======================================================================#
# Obtains the latest Office Deployment Tool download URL               #
#======================================================================#
function Get-OfficeDeploymentTool {
    $Uri = 'https://www.microsoft.com/en-us/download/details.aspx?id=49117'
    
    try {
        $MSWebPage = Invoke-WebRequest -Uri $Uri -UseBasicParsing
        $DownloadURL = $MSWebPage.Links | Where-Object { $_.href -like "*officedeploymenttool*.exe" } | 
                       Select-Object -ExpandProperty href -First 1
        
        if (-not $DownloadURL) {
            Write-Log -Message "Unable to find the download link for the Office Deployment Tool." -Type "ERROR"
            exit 1
        }
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
    param()
    
    Write-Log -Message "Checking for Microsoft Office, Visio, or Project installation..."
    
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
        
        # Office 2019/2021/2024 Products
        "ProPlus2019Retail", "ProPlus2019Volume",
        "ProPlus2021Retail", "ProPlus2021Volume",
        "ProPlus2024Retail", "ProPlus2024Volume",
        "Standard2019Retail", "Standard2019Volume",
        "Standard2021Retail", "Standard2021Volume",
        "StandardSPLA2021Volume",
        "Standard2024Retail", "Standard2024Volume",
        
        # Visio Products
        "VisioProRetail", 
        "VisioPro2019Retail", "VisioPro2019Volume",
        "VisioPro2021Retail", "VisioPro2021Volume",
        "VisioPro2024Retail", "VisioPro2024Volume",
        "VisioStdRetail", 
        "VisioStd2019Retail", "VisioStd2019Volume",
        "VisioStd2021Retail", "VisioStd2021Volume",
        "VisioStd2024Retail", "VisioStd2024Volume",
        
        # Project Products
        "ProjectProRetail",
        "ProjectPro2019Retail", "ProjectPro2019Volume",
        "ProjectPro2021Retail", "ProjectPro2021Volume",
        "ProjectPro2024Retail", "ProjectPro2024Volume",
        "ProjectStdRetail",
        "ProjectStd2019Retail", "ProjectStd2019Volume",
        "ProjectStd2021Retail", "ProjectStd2021Volume",
        "ProjectStd2024Retail", "ProjectStd2024Volume"
    )
    
    $CommonLanguages = @("en-us", "en-gb", "fr-fr", "de-de", "es-es", "it-it", "ja-jp", "ko-kr", "pt-br", "ru-ru", "zh-cn", "zh-tw")
    
    #===============================================================#
    # Method 1: Check for product-specific uninstall registry keys  #
    #===============================================================#
    foreach ($ProductID in $ProductIDs) {
        $NeutralKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$ProductID"
        if (Test-Path -Path $NeutralKey) {
            $KeyInfo = Get-ItemProperty -Path $NeutralKey -ErrorAction SilentlyContinue
            Write-Log -Message "Found Microsoft product at $NeutralKey with display name: $($KeyInfo.DisplayName)"
            return $true
        }
        
        foreach ($Language in $CommonLanguages) {
            $LocalizedKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$ProductID - $Language"
            if (Test-Path -Path $LocalizedKey) {
                $KeyInfo = Get-ItemProperty -Path $LocalizedKey -ErrorAction SilentlyContinue
                Write-Log -Message "Found Microsoft product at $LocalizedKey with display name: $($KeyInfo.DisplayName)"
                return $true
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
            Write-Log -Message "Found Click-to-Run installation with products: $($ClickToRun.ProductReleaseIds)"
            return $true
        }
    }
    
    #===============================================================#
    # Method 3: Check for application executable paths              #
    #===============================================================#
    $AppPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Excel.exe",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Outlook.exe",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Powerpnt.exe",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Visio.exe",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WinProj.exe"
    )
    
    foreach ($AppPath in $AppPaths) {
        if (Test-Path -Path $AppPath) {
            $ExePath = (Get-ItemProperty -Path $AppPath -ErrorAction SilentlyContinue).'(Default)'
            if ($ExePath -and (Test-Path -Path $ExePath)) {
                Write-Log -Message "Found application at $ExePath"
                return $true
            }
        }
    }
    
    #===============================================================#
    # Method 4: Check for Office install root registry keys         #
    #===============================================================#
    $OfficeKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\InstallRoot",
        "HKLM:\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Common\InstallRoot"
    )
    
    foreach ($OfficeKey in $OfficeKeys) {
        if (Test-Path -Path $OfficeKey) {
            $InstallPath = (Get-ItemProperty -Path $OfficeKey -ErrorAction SilentlyContinue).Path
            if ($InstallPath -and (Test-Path -Path $InstallPath)) {
                Write-Log -Message "Found installation path at $InstallPath"
                return $true
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
            Write-Log -Message "Found installation in uninstall registry: $($Results.DisplayName)"
            return $true
        }
    }
    
    Write-Log -Message "No Microsoft Office, Visio, or Project installation detected" -Type "WARNING"
    return $false
}

################################################################################
#                          MAIN INSTALLATION PROCESS                           #
################################################################################
Write-Log -Message "Starting Microsoft 365 Apps installation script"

#===============================================================================#
# Step 1: Validate prerequisites                                                #
#===============================================================================#
if (-not (Test-RunningAsAdmin)) {
    Write-Log -Message "This script requires administrator privileges" -Type "ERROR"
    exit 1
}

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
# Step 2: Download and extract Office Deployment Tool                           #
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
# Step 3: Install Microsoft 365 Apps                                            #
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
# Step 4: Cleanup and finalization                                              #
#===============================================================================#
Write-Log -Message "Cleaning up installation files..."
Remove-Item -Path $OfficeInstallDownloadPath -Force -Recurse -ErrorAction SilentlyContinue

if (Test-ProductInstalled) {
    Write-Log -Message "Microsoft 365 installed successfully!"
    
    if ($Restart) {
        Write-Log -Message "Restarting computer in 60 seconds..."
        Start-Process shutdown.exe -ArgumentList "-r -t 60" -NoNewWindow
    }
    
    exit 0
}
else {
    Write-Log -Message "Microsoft 365 was not detected after installation!" -Type "WARNING"
    
    exit 0
}
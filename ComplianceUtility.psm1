#Requires -Version 5.1

# Copyright (c) Claus Schiroky
# Licensed under the MIT License

<# Variables #>
$Global:strVersion = "4.0.0" <# Version #>
$Global:strDefaultWindowTitle = $Host.UI.RawUI.WindowTitle <# Caching window title #>
$Global:host.UI.RawUI.WindowTitle = "Compliance Utility ($Global:strVersion)" <# Set window title #>
$Global:bolMenuCollectExtended = $false <# Variable for COLLECT menu handling #>
$script:RequiredModuleAvailability = @{} <# Cache local prerequisite checks for the current session #>
$Global:bolComingFromMenu = $false <# Variable for menu handling inside functions #>
$Global:FormatEnumerationLimit = -1 <# Variable to show full Format-List for arrays #>

Function fncInitialize{

    <# Variable for user log path #>
    $Global:strTempFolder = $env:TEMP
    $Global:strUserLogPath = Join-Path -Path $Global:strTempFolder -ChildPath "ComplianceUtility"

    <# Variable for elevated permission check #>
    $currentWindowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $currentWindowsPrincipal = [System.Security.Principal.WindowsPrincipal]::new($currentWindowsIdentity)
    $Global:bolRunningPrivileged = $currentWindowsPrincipal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)

    <# Variable for Windows version #>
    $Global:strOSVersion = (Get-CimInstance Win32_OperatingSystem).Caption

    <# Variable to detect and log Microsoft 365 #>
    $Private:strOfficeC2RPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"    

    <# Logging Windows edition and version #>
    fncLogging -strLogFunction "fncInitialize" -strLogDescription "OS edition" -strLogValue $Global:strOSVersion
    fncLogging -strLogFunction "fncInitialize" -strLogDescription "OS version" -strLogValue $([System.Environment]::OSVersion.Version)

    <# Logging Windows codepage#>
    fncLogging -strLogFunction "fncInitialize" -strLogDescription "Console input encoding" -strLogValue "$([System.Console]::InputEncoding.EncodingName) (CodePage $([System.Console]::InputEncoding.CodePage))"
    fncLogging -strLogFunction "fncInitialize" -strLogDescription "Console output encoding" -strLogValue "$([System.Console]::OutputEncoding.EncodingName) (CodePage $([System.Console]::OutputEncoding.CodePage))"
    fncLogging -strLogFunction "fncInitialize" -strLogDescription "Windows default encoding" -strLogValue "$([System.Text.Encoding]::Default.EncodingName) (CodePage $([System.Text.Encoding]::Default.CodePage))"
 
    <# Logging: Default entries for Windows #>
    fncLogging -strLogFunction "fncInitialize" -strLogDescription "OS 64-Bit" -strLogValue $([System.Environment]::Is64BitOperatingSystem) <# Architecture #>
    fncLogging -strLogFunction "fncInitialize" -strLogDescription "Module version" -strLogValue "$Global:strVersion" <# Module version #>
    fncLogging -strLogFunction "fncInitialize" -strLogDescription "Username" -strLogValue $([System.Environment]::UserName) <# Username #>
    fncLogging -strLogFunction "fncInitialize" -strLogDescription "Machine name" -strLogValue $([System.Environment]::MachineName) <# Machine name #>
    fncLogging -strLogFunction "fncInitialize" -strLogDescription "PowerShell Host" -strLogValue $($Host.Name.ToString()) <# PowerShell host #>
    fncLogging -strLogFunction "fncInitialize" -strLogDescription "PowerShell Version" -strLogValue $($Host.Version.ToString()) <# PowerShell version #>
    fncLogging -strLogFunction "fncInitialize" -strLogDescription "PowerShell Edition" -strLogValue $($PSVersionTable.PSEdition.ToString()) <# PowerShell edition #>
    fncLogging -strLogFunction "fncInitialize" -strLogDescription "PowerShell Current culture" -strLogValue $($Host.CurrentCulture.ToString()) <# PowerShell culture #>
    fncLogging -strLogFunction "fncInitialize" -strLogDescription "PowerShell Current UI culture" -strLogValue $($Host.CurrentUICulture.ToString()) <# PowerShell UI culture #>
    fncLogging -strLogFunction "fncInitialize" -strLogDescription "Running privileged" -strLogValue $Global:bolRunningPrivileged <# Administrative privileges #>

    <# Detect and log Microsoft 365 #>
    If ($(Test-Path -Path $Private:strOfficeC2RPath) -eq $true) {

        $Private:objOfficeC2R = Get-ItemProperty -Path $Private:strOfficeC2RPath -ErrorAction SilentlyContinue

        If ($null -ne $Private:objOfficeC2R) {

            $Private:objOfficeChannels = @{
                "492350f6-3a01-4f97-b9c0-c7c6ddf67d60" = "Current Channel"
                "55336b82-a18d-4dd6-b5f6-9e5095c314a6" = "Monthly Enterprise Channel"
                "7ffbc6bf-bc32-4f92-8982-f9dd17fd3114" = "Semi-Annual Enterprise Channel"
                "b8f9b850-328d-4355-9145-c59439a0c4cf" = "Semi-Annual Enterprise Channel (Preview)"
                "5440fd1f-7ecb-4221-8110-145efaa6372f" = "Current Channel (Preview)"
                "64256afe-f5d9-4f86-8936-8840a6a4f5be" = "Beta Channel"
            }

            $Private:strOfficeChannelGuid = $null
            $Private:strOfficeChannelName = "Not determinable"

            If (-not [String]::IsNullOrWhiteSpace($Private:objOfficeC2R.UpdateChannel)) {

                $Private:strOfficeChannelGuid = ($Private:objOfficeC2R.UpdateChannel -split "/")[-1].ToLowerInvariant()

                If ($Private:objOfficeChannels.ContainsKey($Private:strOfficeChannelGuid)) {
                    $Private:strOfficeChannelName = $Private:objOfficeChannels[$Private:strOfficeChannelGuid]
                }
                Else {
                    $Private:strOfficeChannelName = "Unknown ($Private:strOfficeChannelGuid)"
                }

            }

            <# Logging M365 #>
            fncLogging -strLogFunction "fncInitialize" -strLogDescription "M365 installed" -strLogValue $true
            fncLogging -strLogFunction "fncInitialize" -strLogDescription "M365 version" -strLogValue $Private:objOfficeC2R.VersionToReport
            fncLogging -strLogFunction "fncInitialize" -strLogDescription "M365 architecture" -strLogValue $Private:objOfficeC2R.Platform
            fncLogging -strLogFunction "fncInitialize" -strLogDescription "M365 channel" -strLogValue $Private:strOfficeChannelName
            fncLogging -strLogFunction "fncInitialize" -strLogDescription "M365 products" -strLogValue $($Private:objOfficeC2R.ProductReleaseIds -join ", ")

        }
        Else {

            <# Logging #>
            fncLogging -strLogFunction "fncInitialize" -strLogDescription "M365 installed" -strLogValue "Not determinable"
        }

    }
    Else {

        <# Logging #>
        fncLogging -strLogFunction "fncInitialize" -strLogDescription "M365 installed" -strLogValue $false

    }

    <# Variable to check for unsupported PowerShell #>
    $Global:bolSupportedPowerShell | Out-Null
    $Global:bolSupportedPowerShell = $true

    <# Detect PowerShell Destkop 5.1 #>
    If ($PSVersionTable.PSEdition.ToString() -eq "Desktop" -and [Version]::new($PSVersionTable.PSVersion.Major, $PSVersionTable.PSVersion.Minor) -ne [Version]::new("5.1")) {

        <# Set unsupported PowerShell #>
        $Global:bolSupportedPowerShell = $false

    }

    <# Detect PowerShell Core 7.6 (or less) #>
    If ($PSVersionTable.PSEdition.ToString() -eq "Core" -and [Version]::new($PSVersionTable.PSVersion.Major, $PSVersionTable.PSVersion.Minor) -lt [Version]::new("7.6")) {

        <# Set unsupported PowerShell #>
        $Global:bolSupportedPowerShell = $false

    }

    <# Check for supported PowerShell #>
    If ($Global:bolSupportedPowerShell -eq $false) {

        <# Logging #>
        fncLogging -strLogFunction "fncInitialize" -strLogDescription "Supported PowerShell version" -strLogValue $false

        <# Output #>
        Write-ColoredOutput Red "ATTENTION: The version of PowerShell that is required by the Compliance Utility does not match the currently running version of PowerShell $($PSVersionTable.PSVersion).`n"

        <# Set back window title to default #>
        $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

        <# Exit #>
        Return

    }

    <# Release variable #>
    $Global:bolSupportedPowerShell = $null

}

<# Core definitions #>
Function ComplianceUtility {

    <#
    .SYNOPSIS
        The Compliance Utility is a powerful tool that helps troubleshoot and diagnose sensitivity labels, policies, settings and more. Whether you need to fix issues or reset configurations, this tool gives you everything you need.

        .DESCRIPTION
        Have you ever used the sensitivity button in a Microsoft 365 app or applied a sensitivity label by right-clicking on a file? If so, you've either used the Office's built-in labeling experience or the Purview Information Protection labeling client. If something is not working as expected with your DLP policies, sensitivity labels or you don't see any labels at all the Compliance Utility will help you.

        INTERNET ACCESS
        The Compliance Utility uses additional sources from the Internet to make its functionality fully available.
        
        WARNING: Unexpected errors may occur, and some features may be limited, if there is no connection to the Internet.

    .NOTES
        MIT LICENSE

        Copyright (c) Claus Schiroky.

        Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

        The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

        THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

        VERSION
        4.0.0

        CREATE DATE
        21/08/2026

        AUTHOR
        Claus Schiroky

        HOMEPAGE
        https://compliance-utility.com

        COPYRIGHT
        Copyright (c) Claus Schiroky.

    .PARAMETER Information
        This shows syntax, description and version information.

    .PARAMETER License
        This displays the MIT License.

    .PARAMETER Help
        This opens the online manual.

    .PARAMETER Reset
        IMPORTANT: Before you proceed with this option, please close all open applications.

        This option removes all relevant policies, labels and settings.

        Valid arguments are: "Default", or "Silent".

        Note:
        - Reset with the default argument will not reset all settings, but only user-specific settings if you run PowerShell with user privileges. This is sufficient in most cases to reset Microsoft 365 apps, while a complete reset is useful for all other applications.
        - If you want a complete reset, you need to run the Compliance Utility in an administrative PowerShell window as a user with local administrative privileges.

        Default:

        When you run PowerShell with user privileges, this argument removes all relevant policies, labels and settings:

        ComplianceUtility -Reset Default

        With the above command the following registry keys are cleaned up:

        [HKCU:\SOFTWARE\Classes\Local Settings\Software\Microsoft\MSIPC]
        [HKCU:\SOFTWARE\Classes\Local Settings\Software\Microsoft\AIPMigration]
        [HKCU:\SOFTWARE\Classes\Microsoft.IPViewerChildMenu]
        [HKCU:\SOFTWARE\Microsoft\Cloud\Office]
        [HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\DRM]
        [HKCU:\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\DRM]
        [HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\DRM]
        [HKCU:\SOFTWARE\Microsoft\XPSViewer\Common\DRM]
        [HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Identity]
        [HKCU:\SOFTWARE\Microsoft\MSIP]
        [HKCU:\SOFTWARE\Microsoft\MSOIdentityCRL]
        [HKCR:\AllFilesystemObjects\shell\Microsoft.Azip.Inspect]
        [HKCR:\AllFilesystemObjects\shell\Microsoft.Azip.RightClick]

        The DRMEncryptProperty and OpenXMLEncryptProperty registry settings are purged of the following keys:

        [HKCU:\SOFTWARE\Policies\Microsoft\Cloud\Office\16.0\Common\Security]
        [HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security]
        [HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Security]

        The UseOfficeForLabelling (Use the Sensitivity feature in Office to apply and view sensitivity labels) and AIPException (Use the Azure Information Protection add-in for sensitivity labeling) registry setting is purged of the following keys:

        [HKCU:\SOFTWARE\Policies\Microsoft\Cloud\Office\16.0\Common\Security\Labels]
        [HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security\Labels]
        [HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Security\Labels]

        The following file system folders are cleaned up as well:

        %LOCALAPPDATA%\Microsoft\Word\MIPSDK\mip
        %LOCALAPPDATA%\Microsoft\Excel\MIPSDK\mip
        %LOCALAPPDATA%\Microsoft\PowerPoint\MIPSDK\mip
        %LOCALAPPDATA%\Microsoft\Outlook\MIPSDK\mip
        %LOCALAPPDATA%\Microsoft\Outlook\MIPSDKPDF\mip
        %LOCALAPPDATA%\Microsoft\OneNote\MIPSDK\mip
        %localappdata%\Microsoft\Office\MIPSDK\mip
        %LOCALAPPDATA%\Microsoft\Office\DLP
        %LOCALAPPDATA%\Microsoft\Office\CLP
        %TEMP%\Diagnostics
        %LOCALAPPDATA%\Microsoft\MSIP
        %LOCALAPPDATA%\Microsoft\MSIPC
        %LOCALAPPDATA%\Microsoft\DRM

        The Clear-AIPAuthentication cmdlet is used to reset user settings, if a Purview Information Protection labeling client installation is found.

        When you run the Compliance Utility in an administrative PowerShell window as a user with local administrative privileges, the following registry keys are cleaned up in addition:

        [HKLM:\SOFTWARE\Wow6432Node\Microsoft\MSIPC]
        [HKLM:\SOFTWARE\Microsoft\MSIPC]
        [HKLM:\SOFTWARE\Microsoft\MSDRM]
        [HKLM:\SOFTWARE\Wow6432Node\Microsoft\MSDRM]
        [HKLM:\SOFTWARE\WOW6432Node\Microsoft\MSIP]

        Silent:

        This command line parameter argument does the same as "-Reset Default", but does not print any output - unless an error occurs when attempting to reset:

        ComplianceUtility -Reset Silent

        If a silent reset triggers an error, you can use the additional parameter "-Verbose" to find out more about the cause of the error:

        ComplianceUtility -Reset Silent -Verbose

        You can also review the Script.log file for errors of silent reset.

    .PARAMETER RecordProblem
        IMPORTANT: Before you proceed with this option, please close all open applications.

        As a first step, this parameter activates the required logging and then prompts you to reproduce the problem. While you’re doing so, the Compliance Utility collects and records data. Once you have reproduced the problem, all collected files will be stored into the default logs folder ('%temp%\ComplianceUtility'). Every time you call this option, a new unique subfolder will be created in the logs-folder that reflects the date and time when it was created.

        In the event that you accidentally close the PowerShell window while logging is enabled, the Compliance Utility disables logging the next time you start it.

        Note:
        - Neither CAPI2 or AIP event logs nor filter drivers are recorded if the Compliance Utility is not run in an administrative PowerShell window as a user with local administrative privileges.

    .PARAMETER CollectAIPServiceConfiguration
        This parameter collects your AIP service configuration information (e.g. SuperUsers or OnboardingControlPolicy, etc.) by using the AIPService module.

        The results are written to the log file AIPServiceConfiguration.log in the subfolder "Collect" of the Logs folder.

        Note:
        - You must run the Compliance Utility in an administrative PowerShell window as a user with local administrative privileges to continue with this option. Please contact your administrator if necessary.
        - You need to know your Microsoft 365 global administrator account information to proceed, as you will be asked for your credentials.
        - The AIPService module does not yet support PowerShell 7.x. Therefore, unexpected errors may occur because the AIPService module is executed in compatibility mode in PowerShell 7.x.

    .PARAMETER CollectProtectionTemplates
        This parameter collects protection templates of your tenant by using the AIPService module.

        The results are written to the log files ProtectionTemplates.xml and ProtectionTemplateDetails.xml in the subfolder "Collect\ProtectionTemplates" of the Logs folder, and an export of each protection template (.xml) into the subfolder "ProtectionTemplatesBackup".

        TIP: You can use this feature to create a backup copy of your protection templates.

        Note:
        - You must run the Compliance Utility in an administrative PowerShell window as a user with local administrative privileges to continue with this option. Please contact your administrator if necessary.
        - You need to know your Microsoft 365 global administrator account information to proceed, as you will be asked for your credentials.
        - The AIPService module does not yet support PowerShell 7.x. Therefore, unexpected errors may occur because the AIPService module is executed in compatibility mode in PowerShell 7.x.

    .PARAMETER CollectEndpointURLs
        This parameter collects important endpoint URLs. The URLs are taken from your local registry or your tenant's AIP service configuration information (by using the AIPService module), and extended by additional relevant URLs.

        In a first step, this parameter is used to check whether you can access the URL. In a second step, the issuer of the corresponding certificate of the URL is collected. This process is represented by an output with the Tenant Id, Endpoint name, URL, and Issuer of the certificate. For example:

        --------------------------------------------------
        Tenant Id: 48fc03bd-c84b-44ac-b7761b7-a4c5eefd5ac1
        --------------------------------------------------

        Endpoint: UnifiedLabelingDistributionPointUrl
        URL:      https://dataservice.protection.outlook.com
        Issuer:   CN=Microsoft Azure RSA TLS Issuing CA 08, O=Microsoft Corporation, C=US

        In addition, results are written into log file EndpointURLs.log in the subfolder "Collect" of the Logs folder.

        Note:
        - You must run the Compliance Utility in an administrative PowerShell window as a user with local administrative privileges to continue with this option, if the corresponding Microsoft 365 app is not bootstraped. Please contact your administrator if necessary.
        - You need to know your Microsoft 365 global administrator account information to proceed, as you will be asked for your credentials.
        - The AIPService module does not yet support PowerShell 7.x. Therefore, unexpected errors may occur because the AIPService module is executed in compatibility mode in PowerShell 7.x.

    .PARAMETER CollectLabelsAndPolicies
        This parameter collects Information Protection labels, policies (with detailled actions and rules), auto-label policies and rules from your Microsoft Purview compliance portal by using the Exchange Online PowerShell module.

        The results are written to the log files Labels.xml, LabelsDetailedActions.xml, LabelPolicies.xml, LabelRules.xml, AutoLabelPolicies.xml and AutoLabelRules.xml in the subfolder "Collect\LabelsAndPolicies" of the Logs folder, and you can also have a CLP subfolder with the Office CLP policy.

        TIP: You can use the resulting log file to create exact copies of the label and policy settings for troubleshooting purposes, e.g. in test environments.

        Note:
        - You must run the Compliance Utility in an administrative PowerShell window as a user with local administrative privileges to continue with this option. Please contact your administrator if necessary.
        - You need to know your Microsoft 365 global administrator account information to proceed with this option, as you will be asked for your credentials.
        - The Microsoft Exchange Online Management module is required. Run ComplianceUtility -UpdateModules to install missing or update outdated modules.

    .PARAMETER CollectDLPRulesAndPolicies
        This parameter collects DLP rules and policies, sensitive information type details, rule packages, keyword dictionaries and exact data match schemas from the Microsoft Purview compliance portal by using the Exchange Online PowerShell module.

        The results are written to the log files DlpPolicy.xml, DlpRule.xml, DlpPolicyDistributionStatus.xml, DlpSensitiveInformationType.xml, DlpSensitiveInformationTypeRulePackage.xml, DlpKeywordDictionary.xml and DlpEdmSchema.xml in the subfolder "Collect\DLPRulesAndPolicies" of the Logs folder.

        Note:
        - You must run the Compliance Utility in an administrative PowerShell window as a user with local administrative privileges to continue with this option. Please contact your administrator if necessary.
        - You need to know your Microsoft 365 global administrator account information to proceed with this option, as you will be asked for your credentials.
        - The Microsoft Exchange Online Management module is required. Run ComplianceUtility -UpdateModules to install missing or update outdated modules.

    .PARAMETER CollectUserLicenseDetails
        This parameter collects the user license details by the Graph PowerShell module.

        The results are written to the log file UserLicenseDetails.log in the subfolder "Collect" of the Logs folder.

        Note:
        - You must log in with the corresponding Microsoft 365 user account for which you want to check the license details.
        - The Microsoft Graph PowerShell modules are required. Run ComplianceUtility -UpdateModules to install missing or update outdated modules.

    .PARAMETER CollectExchangeIRMConfiguration
        This parameter collects your Exchange IRM configuration information by using the Exchange Online PowerShell module.

        The results are written to the log file ExchangeIRMConfiguration.log in the subfolder "Collect" of the Logs folder.

        Note:
        - You must run the Compliance Utility in an administrative PowerShell window as a user with local administrative privileges to continue with this option. Please contact your administrator if necessary.
        - You need to know your Exchange online administrator account information to proceed, as you will be asked for your credentials.
        - The Microsoft Exchange Online Management module is required. Run ComplianceUtility -UpdateModules to install missing or update outdated modules.

    .PARAMETER CompressLogs
        This command line parameter should always be used at the very end of a scenario.

        This parameter compresses all collected log files and folders into a .zip archive, and the corresponding file is saved to your desktop. In addition, the default logs folder ('%temp%\ComplianceUtility') is cleaned.

    .PARAMETER UpdateModules
        Checks the required PowerShell modules (AIPService, ExchangeOnlineManagement, and Microsoft Graph).
        It installs missing modules and updates outdated ones from the PowerShell Gallery.

    .PARAMETER Menu
        This will start the Compliance Utility with the default menu.

    .EXAMPLE
        ComplianceUtility -Information
        This shows syntax and description.

    .EXAMPLE
        ComplianceUtility -License
        This displays the MIT License.

    .EXAMPLE
        ComplianceUtility -Help
        This parameter opens the online manual.

    .EXAMPLE
        ComplianceUtility -Reset Default
        This parameter removes all relevant policies, labels and settings.

    .EXAMPLE
        ComplianceUtility -Reset Silent
        This parameter removes all relevant policies, labels and settings without any output.

    .EXAMPLE
        ComplianceUtility -RecordProblem
        This parameter cleans up existing MSIP/MSIPC log folders, then it removes all relevant policies, labels and settings, and starts recording data.

    .EXAMPLE
        ComplianceUtility -CollectAIPServiceConfiguration
        This parameter collects AIP service configuration information of your tenant.

    .EXAMPLE
        ComplianceUtility -CollectProtectionTemplates
        This parameter collects protection templates of your tenant.

    .EXAMPLE
        ComplianceUtility -CollectLabelsAndPolicies
        This parameter collects the labels and policy definitions from the Microsoft Purview compliance portal.

    .EXAMPLE
        ComplianceUtility -CollectEndpointURLs
        This parameter collects important enpoint URLs.

    .EXAMPLE
        ComplianceUtility -CollectDLPRulesAndPolicies
        This parameter collects DLP rules and policies from the Microsoft Purview compliance portal.

    .EXAMPLE
        ComplianceUtility -CollectUserLicenseDetails
        This parameter collects the user license details by Microsoft Graph.

    .EXAMPLE
        ComplianceUtility -CollectExchangeIRMConfiguration
        This parameter collects your Exchange IRM configuration information.

    .EXAMPLE
        ComplianceUtility -UpdateModules
        This installs missing and updates outdated PowerShell modules required by the Compliance Utility.

    .EXAMPLE
        ComplianceUtility -CompressLogs
        This parameter compress all collected logs files into a .zip archive, and the corresponding path and file name is displayed.

    .EXAMPLE
        ComplianceUtility -Reset Default -RecordProblem -CompressLogs
        This parameter removes all relevant policies, labels and settings, starts recording data, and compress all collected logs files to a .zip archive on the desktop.

    .EXAMPLE
        ComplianceUtility -Menu
        This will start the Compliance Utility with the default menu.

    .LINK
        https://compliance-utility.com

    #>

    <# Binding for parameters #>
    [CmdletBinding (
        HelpURI = "https://github.com/schiroky/ComplianceUtility/blob/main/Manuals/4.0.0/Manual.md", <# URL for online manual #>
        PositionalBinding = $false, <# None-positional parameters #>
        DefaultParameterSetName = "Menu" <# Default start parameter #>
    )]
    [Alias("CompUtil")]

    <# Parameter definitions #>
    Param (

        <# Information #>
        [Alias("i")]
        [Parameter(ParameterSetName = "Information")]
        [switch]$Information,

        <# License #>
        [Alias("m")]
        [Parameter(ParameterSetName = "License")]
        [switch]$License,

        <# Help #>
        [Alias("h")]
        [Parameter(ParameterSetName = "Help")]
        [switch]$Help,

        <# Reset #>
        [Alias("r")]
        [Parameter(ParameterSetName = "Reset and logging")]
        [ValidateSet("Default", "Silent")]
        [string]$Reset="Default",

        <# RecordProblem #>
        [Alias("p")]
        [parameter(ParameterSetName = "Reset and logging")]
        [switch]$RecordProblem,

        <# CollectAIPServiceConfiguration #>
        [Alias("a")]
        [Parameter(ParameterSetName = "Reset and logging")]
        [switch]$CollectAIPServiceConfiguration,

        <# CollectProtectionTemplates #>
        [Alias("t")]
        [Parameter(ParameterSetName = "Reset and logging")]
        [switch]$CollectProtectionTemplates,

        <# CollectEndpointURLs #>
        [Alias("e")]
        [Parameter(ParameterSetName = "Reset and logging")]
        [switch]$CollectEndpointURLs,

        <# CollectLabelsAndPolicies #>
        [Alias("l")]
        [Parameter(ParameterSetName = "Reset and logging")]
        [switch]$CollectLabelsAndPolicies,

        <# CollectDLPPoliciesAndRules #>
        [Alias("d")]
        [Parameter(ParameterSetName = "Reset and logging")]
        [switch]$CollectDLPRulesAndPolicies,

        <# CollectUserLicenseDetails #>
        [Alias("u")]
        [Parameter(ParameterSetName = "Reset and logging")]
        [switch]$CollectUserLicenseDetails,

        <# CollectExchangeIRMConfiguration #>
        [Alias("g")]
        [Parameter(ParameterSetName = "Reset and logging")]
        [switch]$CollectExchangeIRMConfiguration,

        <# CompressLogs #>
        [Alias("z")]
        [Parameter(ParameterSetName = "Reset and logging")]
        [switch]$CompressLogs,

        <# UpdateModules #>
        [Alias("q")]
        [Parameter(ParameterSetName = "Update modules")]
        [switch]$UpdateModules,

        <# Menu #>
        [Parameter(ParameterSetName = "Menu")]
        [switch]$Menu

    )

    <# Actions for Information #>
    If ($PsCmdlet.ParameterSetName -eq "Information") {

        <# Call Information #>
        fncInformation

        <# Logging #>
        fncLogging -strLogFunction "ComplianceUtility" -strLogDescription "INFORMATION" -strLogValue "Proceeded"

    }

    <# Actions for License #>
    If ($PSBoundParameters.ContainsKey("License")) {

        <# Call License #>
        fncLicense

        <# Logging #>
        fncLogging -strLogFunction "ComplianceUtility" -strLogDescription "LICENSE" -strLogValue "Proceeded"

    }

    <# Actions for Help #>
    If ($PSBoundParameters.ContainsKey("Help")) {

        <# Call Help #>
        fncHelp

        <# Logging #>
        fncLogging -strLogFunction "ComplianceUtility" -strLogDescription "HELP" -strLogValue "Proceeded"

    }

    <# Actions for Reset #>
    If ($PSBoundParameters.ContainsKey("Reset")) {

        <# Logging #>
        fncLogging -strLogFunction "ComplianceUtility" -strLogDescription "Parameter Help" -strLogValue "Triggered"

        <# Call Reset #>
        fncReset -strResetMethod $Reset

    }

    <# Actions for RecordProblem #>
    If ($PSBoundParameters.ContainsKey("RecordProblem")) {

        <# Logging #>
        fncLogging -strLogFunction "ComplianceUtility" -strLogDescription "Parameter RecordProblem" -strLogValue "Triggered"

        <# Call RecordProblem #>
        fncRecordProblem

    }

    <# Actions CollectAIPServiceConfiguration #>
    If ($PSBoundParameters.ContainsKey("CollectAIPServiceConfiguration")) {

        <# Logging #>
        fncLogging -strLogFunction "ComplianceUtility" -strLogDescription "Parameter CollectAIPServiceConfiguration" -strLogValue "Triggered"

        <# Call CollectAIPServiceConfiguration #>
        fncCollectAIPServiceConfiguration

    }

    <# Actions for CollectProtectionTemplates #>
    If ($PSBoundParameters.ContainsKey("CollectProtectionTemplates")) {

        <# Logging #>
        fncLogging -strLogFunction "ComplianceUtility" -strLogDescription "Parameter CollectProtectionTemplates" -strLogValue "Triggered"

        <# Call CollectProtectionTemplates #>
        fncCollectProtectionTemplates

    }

    <# Actions for CollectUserLicenseDetails #>
    If ($PSBoundParameters.ContainsKey("CollectUserLicenseDetails")) {

        <# Logging #>
        fncLogging -strLogFunction "ComplianceUtility" -strLogDescription "Parameter CollectUserLicenseDetails" -strLogValue "Triggered"

        <# Call CollectUserLicenseDetails #>
        fncCollectUserLiceneseDetails

    }

    <# Actions for CollectLabelsAndPolicies #>
    If ($PSBoundParameters.ContainsKey("CollectLabelsAndPolicies")) {

        <# Logging #>
        fncLogging -strLogFunction "ComplianceUtility" -strLogDescription "Parameter CollectLabelsAndPolicies" -strLogValue "Triggered"

        <# Call CollectLabelsAndPolicies #>
        fncCollectLabelsAndPolicies

    }

    <# Actions for CollectEndpointURLs #>
    If ($PSBoundParameters.ContainsKey("CollectEndpointURLs")) {

        <# Logging #>
        fncLogging -strLogFunction "ComplianceUtility" -strLogDescription "Parameter CollectEndpointsURLs" -strLogValue "Triggered"

        <# Call CollectEndpointURLs #>
        fncCollectEndpointURLs

    }

    <# Actions for CollectDLPRulesAndPolicies #>
    If ($PSBoundParameters.ContainsKey("CollectDLPRulesAndPolicies")) {

        <# Logging #>
        fncLogging -strLogFunction "ComplianceUtility" -strLogDescription "Parameter CollectDLPRulesAndPolicies" -strLogValue "Triggered"

        <# Call CollectDLPRulesAndPolicies #>
        fncCollectDLPRulesAndPolicies

    }

    <# Actions for CollectExchangeIRMConfiguration #>
    If ($PSBoundParameters.ContainsKey("CollectExchangeIRMConfiguration")) {

        <# Logging #>
        fncLogging -strLogFunction "ComplianceUtility" -strLogDescription "Parameter CollectExchangeIRMConfiguration" -strLogValue "Triggered"

        <# Call CollectExchangeIRMConfiguration #>
        fncCollectExchangeIRMConfiguration

    }

    <# Actions for CompressLogs #>
    If ($PSBoundParameters.ContainsKey("CompressLogs")) {

        <# Logging #>
        fncLogging -strLogFunction "ComplianceUtility" -strLogDescription "Parameter CompressLogs" -strLogValue "Triggered"

        <# Call CompressLogs #>
        fncCompressLogs

        <# Set back window title to default #>
        $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

        <# Exit #>
        Return

    }

    <# Actions for update moduls #>
    If ($PSBoundParameters.ContainsKey("UpdateModules")) {

        <# Logging #>
        fncLogging -strLogFunction "ComplianceUtility" -strLogDescription "Parameter UpdateModules" -strLogValue "Triggered"

        <# Call UpdateModules #>
        fncUpdateModules

        <# Set back window title to default #>
        $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

        <# Exit #>
        Return

    }

    <# Actions for ShowMenu #>
    If ($PsCmdlet.ParameterSetName -eq "Menu") {

        <# Logging #>
        fncLogging -strLogFunction "ComplianceUtility" -strLogDescription "MENU" -strLogValue "Triggered"

        <# Call ShowMenu #>
        fncShowMenu

    }

}

Function fncLogging ($strLogFunction, $strLogDescription, $strLogValue) {

    <# Detect/create UserLogPath #>
    If ($(Test-Path -Path $Global:strUserLogPath) -Eq $false) {

        New-Item -ItemType Directory -Force -Path $Global:strUserLogPath | Out-Null <# Define UserLogPath #>

    }

    <# Output #>
    Write-Verbose "$(Get-Date -UFormat "%Y-%m-%d"), $(Get-Date -UFormat "%H:%M"), $strLogFunction, $strLogDescription, $strLogValue"

    <# Append output #>
    Write-Verbose "$(Get-Date -UFormat "%Y-%m-%d"), $(Get-Date -UFormat "%H:%M"), $strLogFunction, $strLogDescription, $strLogValue" -ErrorAction SilentlyContinue -Verbose 4>> "$Global:strUserLogPath\Script.log"

}

Function fncInformation {

    <# Logging #>
    fncLogging -strLogFunction "fncInformation" -strLogDescription "INFORMATION" -strLogValue "Called"

    <# Command line actions #>
    If ($Global:bolComingFromMenu -eq $false) {

        <# Call Information #>
        Get-Help -Verbose:$false ComplianceUtility

    }

    <# Menu Actions #>
    If ($Global:bolComingFromMenu -eq $true) {

        <# Output #>
        Write-Output "NAME:`nComplianceUtility`n`nDESCRIPTION:`nThe Compliance Utility is a powerful tool that helps troubleshoot and diagnose sensitivity labels, policies, settings and more. Whether you need to fix issues or reset configurations, this tool gives you everything you need.`n`nVERSION:`n$Global:strVersion`n`nAUTHOR:`nClaus Schiroky`n`nHOMEPAGE:`nhttps://compliance-utility.com`n`nCOPYRIGHT:`nCopyright (c) Claus Schiroky.`n"

    }

}

Function fncLicense {

    <# Logging #>
    fncLogging -strLogFunction "fncLicense" -strLogDescription "LICENSE" -strLogValue "Called"

    <# Output #>
    Write-Output "MIT License`n`nCopyright (c) Claus Schiroky.`n`nPermission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the `"Software`"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:`n`nThe above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.`n`nTHE SOFTWARE IS PROVIDED `"AS IS`", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.`n"

}

Function fncHelp {

    <# No internet message #>
    $Private:strNoOnlineHelp = "ATTENTION: The online manual cannot be accessed.`nEither the website (github.com) is unavailable or there is no internet connection.`n`nNote:`n`n- Please use the command line help by entering the command:`nGet-Help ComplianceUtility -Detailed"

    <# Check internet connection #>
    If ($(fncTestInternetAccess "github.com") -Eq $true) {

        <# Open manual #>
        Start-Process "https://github.com/schiroky/ComplianceUtility/blob/main/Manuals/4.0.0/Manual.md"

        <# Logging #>
        fncLogging -strLogFunction "fncHelp" -strLogDescription "HELP" -strLogValue "Called"

    }
    Else { <# Offline actions #>

        <# Output #>
        Write-ColoredOutput Red $Private:strNoOnlineHelp

        <# Logging #>
        fncLogging -strLogFunction "fncHelp" -strLogDescription "Help" -strLogValue "No internet connection"

    }

}

Function Write-ColoredOutput($Private:ForegroundColor) {

    <# Variables #>
    $Private:TextColor = $Global:host.UI.RawUI.ForegroundColor <# Backup current text color #>
    $Global:host.UI.RawUI.ForegroundColor = $Private:ForegroundColor <# Set text color #>

    <# Output #>
    If ($Private:args) {
        Write-Output $Private:args
    }
    Else {
        $Private:input | Write-Output
    }

    <# Set back color #>
    $Global:host.UI.RawUI.ForegroundColor = $Private:TextColor

}

<# Detect and delete a registry setting #>
Function fncDeleteRegistrySetting ($strRegistryKey, $strRegistrySetting) {

    <# Try to remove registry setting #>
    Try {

        <# Set registry setting variable #>
        $strItemExists = Get-ItemProperty $strRegistryKey $strRegistrySetting -ErrorAction SilentlyContinue

        <# Remove registry setting only if it exists #>
        If (-not ($Null -eq $strItemExists) -or ($strItemExists.Length -eq 0)) {

            <# Remove registry setting #>
            Remove-ItemProperty -Path $strRegistryKey -Name $strRegistrySetting -Force -ErrorAction Stop

            <# Logging #>
            fncLogging -strLogFunction "fncDeleteRegistrySetting" -strLogDescription "$strRegistrySetting ($strRegistryKey)" -strLogValue "Removed"

        }

    }
    Catch {

        <# Silently ignore if setting does not exist #>
        Write-Verbose "Registry setting not found or cannot be removed: $strRegistrySetting"

    }

}

Function fncReset ($strResetMethod) {

    <# ShowMenu/Silent actions #>
    If ($strResetMethod -notmatch "Silent") {

        <# Output #>
        Write-Output "RESET:"
        Write-ColoredOutput Red "IMPORTANT: Before you proceed with this option, please close all open applications."
        $Private:ReadHost = Read-Host "Only if the above is true, please press [Y]es to continue, or [N]o to cancel"

        <# Cancel/no actions #>
        If ($Private:ReadHost -eq "N") {

            <# Logging #>
            fncLogging -strLogFunction "fncReset" -strLogDescription "RESET Default" -strLogValue "Canceled"

            <# Command line actions #>
            If ($Global:bolComingFromMenu -eq $false) {

                <# Set back window title #>
                $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

                <# Exit #>
                Return

            }

            <# ShowMenu actions #>
            If ($Global:bolComingFromMenu -eq $true) {

                <# Clear console #>
                Clear-Host

                <# Call ShowMenu #>
                fncShowMenu

            }

        }

        <# Logging #>
        fncLogging -strLogFunction "fncReset" -strLogDescription "RESET Default" -strLogValue "Initiated"

        <# Output #>
        Write-Output "Resetting..."

    }

    <# Silent actions #>
    If ($strResetMethod -match "Silent") {

        <# Logging #>
        fncLogging -strLogFunction "fncReset" -strLogDescription "RESET Silent" -strLogValue "Initiated"

    }

    <# "Yes"/Silent actions (by reset default) #>
    If ($Private:ReadHost -eq "Y" -or ($strResetMethod -match "Silent")) {

        <# Clean user keys #>
        fncDeleteItem "HKCU:\SOFTWARE\Classes\Local Settings\Software\Microsoft\MSIPC"
        fncDeleteItem "HKCU:\SOFTWARE\Classes\Local Settings\Software\Microsoft\AIPMigration"
        fncDeleteItem "HKCU:\SOFTWARE\Classes\Microsoft.IPViewerChildMenu"
        fncDeleteItem "HKCU:\SOFTWARE\Microsoft\Cloud\Office"
        fncDeleteItem "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\DRM"
        fncDeleteItem "HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\DRM"
        fncDeleteItem "HKCU:\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\DRM"
        fncDeleteItem "HKCU:\SOFTWARE\Microsoft\XPSViewer\Common\DRM"
        fncDeleteItem "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Identity"
        fncDeleteItem "HKCU:\SOFTWARE\Microsoft\MSIP"
        fncDeleteItem "HKCU:\SOFTWARE\Microsoft\MSOIdentityCRL"

        <# Clean registry settings #>
        fncDeleteRegistrySetting -strRegistryKey "HKCU:\SOFTWARE\Policies\Microsoft\Cloud\Office\16.0\Common\Security\Labels" -strRegistrySetting "UseOfficeForLabelling"
        fncDeleteRegistrySetting -strRegistryKey "HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security\Labels" -strRegistrySetting "UseOfficeForLabelling"
        fncDeleteRegistrySetting -strRegistryKey "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Security\Labels" -strRegistrySetting "UseOfficeForLabelling"
        fncDeleteRegistrySetting -strRegistryKey "HKCU:\SOFTWARE\Policies\Microsoft\Cloud\Office\16.0\Common\Security\Labels" -strRegistrySetting "AIPException"
        fncDeleteRegistrySetting -strRegistryKey "HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security\Labels" -strRegistrySetting "AIPException"
        fncDeleteRegistrySetting -strRegistryKey "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Security\Labels" -strRegistrySetting "AIPException"
        fncDeleteRegistrySetting -strRegistryKey "HKCU:\SOFTWARE\Policies\Microsoft\Cloud\Office\16.0\Common\Security" -strRegistrySetting "OpenXMLEncryptProperty"
        fncDeleteRegistrySetting -strRegistryKey "HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security" -strRegistrySetting "OpenXMLEncryptProperty"
        fncDeleteRegistrySetting -strRegistryKey "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Security" -strRegistrySetting "OpenXMLEncryptProperty"
        fncDeleteRegistrySetting -strRegistryKey "HKCU:\SOFTWARE\Policies\Microsoft\Cloud\Office\16.0\Common\Security" -strRegistrySetting "DRMEncryptProperty"
        fncDeleteRegistrySetting -strRegistryKey "HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security" -strRegistrySetting "DRMEncryptProperty"
        fncDeleteRegistrySetting -strRegistryKey "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Security" -strRegistrySetting "DRMEncryptProperty"

        <# Clean client classes keys #>
        fncDeleteItem "HKCR:\AllFilesystemObjects\shell\Microsoft.Azip.Inspect"
        fncDeleteItem "HKCR:\AllFilesystemObjects\shell\Microsoft.Azip.RightClick"

        <# Clean client folders #>
        fncDeleteItem "$env:LOCALAPPDATA\Microsoft\Word\MIPSDK\mip"
        fncDeleteItem "$env:LOCALAPPDATA\Microsoft\Excel\MIPSDK\mip"
        fncDeleteItem "$env:LOCALAPPDATA\Microsoft\PowerPoint\MIPSDK\mip"
        fncDeleteItem "$env:LOCALAPPDATA\Microsoft\Outlook\MIPSDK\mip"
        fncDeleteItem "$env:LOCALAPPDATA\Microsoft\Outlook\MIPSDKPDF\mip"
        fncDeleteItem "$env:LOCALAPPDATA\Microsoft\OneNote\MIPSDK\mip"
        fncDeleteItem "$env:LOCALAPPDATA\Microsoft\Office\MIPSDK\mip"
        fncDeleteItem "$env:LOCALAPPDATA\Microsoft\Office\DLP"
        fncDeleteItem "$env:LOCALAPPDATA\Microsoft\Office\CLP"
        fncDeleteItem "$env:TEMP\Diagnostics"
        fncDeleteItem "$env:LOCALAPPDATA\Microsoft\MSIP"
        fncDeleteItem "$env:LOCALAPPDATA\Microsoft\MSIPC"
        fncDeleteItem "$env:LOCALAPPDATA\Microsoft\DRM"

        <# Administrative reset actions #>
        If ($Global:bolRunningPrivileged -eq $true) {

            # Clean machine keys #>
            fncDeleteItem "HKLM:\SOFTWARE\Wow6432Node\Microsoft\MSIPC"
            fncDeleteItem "HKLM:\SOFTWARE\Microsoft\MSIPC"
            fncDeleteItem "HKLM:\SOFTWARE\Microsoft\MSDRM"
            fncDeleteItem "HKLM:\SOFTWARE\Wow6432Node\Microsoft\MSDRM"
            fncDeleteItem "HKLM:\SOFTWARE\WOW6432Node\Microsoft\MSIP"
            fncDeleteItem "HKLM:\SOFTWARE\Microsoft\PolicyManager\AdmxDefault" <# Intune MDM Policies #>

        }

        <# Actions on PowerShell Core for compatibility mode #>
        If ($PSVersionTable.PSEdition.ToString() -eq "Core") {

            # Remove if an existing installation was found #>
            If (Get-Module -Name AzureInformationProtection -ListAvailable -ErrorAction SilentlyContinue -WarningAction SilentlyContinue) {

                <# Remove AzureInformationProtection module #>
                Remove-Module -Name AzureInformationProtection -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

                <# Import module in compatiblity mode #>
                Import-Module -Name AzureInformationProtection -UseWindowsPowerShell -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

                <# Logging #>
                fncLogging -strLogFunction "fncReset" -strLogDescription "AzureInformationProtection compatiblity mode" -strLogValue $true

            }

            # Remove if an existing installation was found #>
            If (Get-Module -Name PurviewInformationProtection -ListAvailable -ErrorAction SilentlyContinue -WarningAction SilentlyContinue) {

                <# Remove UnifiedLabelingSupportTool module #>
                Remove-Module -Name PurviewInformationProtection -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

                <# Import module in compatiblity mode #>
                Import-Module -Name PurviewInformationProtection -UseWindowsPowerShell -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

                <# Logging #>
                fncLogging -strLogFunction "fncReset" -strLogDescription "PurviewInformationProtection compatiblity mode" -strLogValue $true

            }

        }

        <# Clear user settings #>
        If (Get-Module -ListAvailable -Name AzureInformationProtection, PurviewInformationProtection) { <# Check for AIP/PIP client #>

            <# Clear user settings #>
            Clear-AIPAuthentication -ErrorAction SilentlyContinue | Out-Null

            <# Logging #>
            fncLogging -strLogFunction "fncReset" -strLogDescription "AIPAuthentication" -strLogValue "Cleared"

        }

        <# Default command line/menu actions #>
        If ($strResetMethod -notmatch "Silent") {

            <# Output #>
            Write-ColoredOutput Green "RESET: Proceeded.`n"

            <# Logging #>
            fncLogging -strLogFunction "fncReset" -strLogDescription "RESET Default" -strLogValue "Proceeded"

        }

        <# Silent command line actions #>
        If ($strResetMethod -match "Silent") {

            <# Logging #>
            fncLogging -strLogFunction "fncReset" -strLogDescription "RESET Silent" -strLogValue "Proceeded"

        }

    }
    Else { <# Any key actions #>

        <# Logging #>
        fncLogging -strLogFunction "fncReset" -strLogDescription "RESET" -strLogValue "Canceled"

        <# Command line actions #>
        If ($Global:bolComingFromMenu -eq $false) {

            <# Set back window title #>
            $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

            <# Exit #>
            Return

        }

        <# Menu actions #>
        If ($Global:bolComingFromMenu -eq $true) {

            <# Clear console #>
            Clear-Host

            <# Call ShowMenu #>
            fncShowMenu

            Return

        }

    }

}

Function fncDeleteItem ($Private:objItem) {

    <# Detect key, file or folder #>
    If ($(Test-Path -Path $Private:objItem) -Eq $true) {

        <# Try to delete item/folder #>
        Try {

            <# Delete folder/registry key #>
            Get-ChildItem -Path $Private:objItem -Exclude "Telemetry", "powershell.exe", "powershell" -Force | Remove-Item -Recurse -Force -ErrorAction Stop | Out-Null

            <# Logging #>
            fncLogging -strLogFunction "fncDeleteItem" -strLogDescription "Item deleted" -strLogValue $Private:objItem

        }
        Catch [System.IO.IOException] { <# Actions if files or folders cannot be accessed, because they are locked/used by another process <#>

            <# Output #>
            Write-ColoredOutput Red "WARNING: Some items or folders are still used by another process.`nIMPORTANT: Please close all applications, restart the PowerShell session (or restart machine) and try again."

            <# Logging #>
            fncLogging -strLogFunction "fncDeleteItem" -strLogDescription "Item locked" -strLogValue $Private:objItem
            fncLogging -strLogFunction "fncDeleteItem" -strLogDescription "RESET" -strLogValue "ERROR: RESET failed"

            <# Release variable #>
            $Private:objItem = $null

            <# ShowMenu actions #>
            If ($Global:bolComingFromMenu -eq $false) {

                <# Output #>
                Write-ColoredOutput Red "RESET: Failed.`n"

                <# Set back window title #>
                $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

                <# Exit #>
                Return

            }
            <# ShowMenu actions #>
            If ($Global:bolComingFromMenu -eq $true) {

                <# Output #>
                Write-ColoredOutput Red "RESET: Failed.`n"

                <# Call Pause #>
                fncPause

                <# Call ShowMenu #>
                fncShowMenu

            }

        }

    }

    <# Release variable #>
    $Private:objItem = $null

}

Function fncCopyItem ($Private:objItem, $Private:strDestination, $Private:strFileName) {

    <# Try to copy item/s #>
    Try {

        <# Detect path and copy item #>
        If ($(Test-Path -Path $Private:objItem) -Eq $true) {

            <# Copy item #>
            Copy-Item -Path $Private:objItem -Destination $Private:strDestination -Recurse -Force -ErrorAction Stop | Out-Null

            <# Logging #>
            fncLogging -strLogFunction "fncCopyItem" -strLogDescription "Item copied" -strLogValue $Private:strFileName

        }

    }
    Catch [System.IO.IOException] { <# Action if file cannot be accessed #>

        <# Detect path for individual Logging. Caused by PowerShell Telemetry #>
        If ($Private:objItem -like "*MSIP") {

            <# Logging #>
            fncLogging -strLogFunction "fncCopyItem" -strLogDescription "Item locked" -strLogValue "ERROR: \MSIP"

        }
        Else {

            <# Logging #>
            fncLogging -strLogFunction "fncCopyItem" -strLogDescription "Item locked" -strLogValue "ERROR: "$Private:objItem

        }

        <# Release variables #>
        $Private:objItem = $null
        $Private:strDestination = $null

    }

    <# Release variables #>
    $Private:objItem = $null
    $Private:strDestination = $null

}

Function fncTestInternetAccess ($Private:strURL) {

    <# Test internet access #>
    If ($(Test-Connection $Private:strURL -Count 1 -Quiet) -Eq $true) {

        <# Logging #>
        fncLogging -strLogFunction "fncTestInternetAccess" -strLogDescription "Internet access" -strLogValue $true

        <# Internet access #>
        Return $true

    }
    Else {

        <# No internet access #>
        Return $false

        <# Logging #>
        fncLogging -strLogFunction "fncTestInternetAccess" -strLogDescription "Internet access" -strLogValue $false

    }

    <# Release variable #>
    $Private:strURL = $null

}

Function fncRecordProblem {

    <# Output #>
    Write-Output "RECORD PROBLEM:"
    Write-ColoredOutput Red "IMPORTANT: Before you proceed with this option, please close all open applications."
    $Private:ReadHost = Read-Host "Only if the above is true, please press [Y]es to continue, or [N]o to cancel"

    <# Logging #>
    fncLogging -strLogFunction "fncRecordProblem" -strLogDescription "RECORD PROBLEM" -strLogValue "Initiated"

    <# "Yes"-actions #>
    If ($Private:ReadHost -Eq "Y") {

        <# Detect admin permissions #>
        If ($Global:bolRunningPrivileged -eq $false) {

            <# Logging #>
            Write-ColoredOutput Red "ATTENTION: Please note that neither CAPI2 or AIP event logs nor filter drivers are recorded.`nIf you want a complete record, you must run the Compliance Utility in an administrative PowerShell window as a user with local administrative privileges."

        }

        <# Output #>
        Write-Output "Initializing, please wait..."

        <# Variables for log folder #>
        $Private:strUniqueFolderName = (Get-Date -Verbose:$false -UFormat "%y%m%d-%H%M%S")
        $Global:strUniqueLogFolder = $Global:strUserLogPath.ToString() + "\" +  $Private:strUniqueFolderName.ToString()

        <# Create log folder #>
         New-Item -ItemType Directory -Force -Path $Global:strUniqueLogFolder | Out-Null

        <# Logging #>
        fncLogging "fncRecordProblem" -strLogDescription "New log folder created" -strLogValue $Private:strUniqueFolderName

        <# Call Enablelogging #>
        fncEnableLogging

        <# Output by privileges check #>
        If ($Global:bolRunningPrivileged -eq $false) {

            <# Output with no admin privileges #>
            Write-Output "Recording is now underway for user `"$Env:UserName`"."

        }
        Else {

            <# Output with admin privileges #>
            Write-Output "Recording is now underway for administrator `"$Env:UserName`"."

        }

        <# Output #>
        Write-ColoredOutput Red "IMPORTANT: Now reproduce the problem, but leave this window open."
        Read-Host "After reproducing the problem, close all the applications you were using, return here and press enter to complete the recording."

        <# Output #>
        Write-Output "Collecting logs, please wait...`n"

        <# Call CollectingLogs #>
        fncCollectingLogs

        <# Call Disablelogging #>
        fncDisableLogging

        <# Logging #>
        fncLogging -strLogFunction "fncRecordProblem" -strLogDescription "RECORD PROBLEM" -strLogValue "Proceeded"

        <# Output #>
        Write-Output "Log files: $Global:strUniqueLogFolder"
        Write-ColoredOutput Green "RECORD PROBLEM: Proceeded.`n"

        <# Release variable #>
        $Global:strUniqueLogFolder = $null

    }
    ElseIf ($Private:ReadHost -eq "N") { <# "No"/cancel actions #>

        <# Logging #>
        fncLogging -strLogFunction "fncRecordProblem" -strLogDescription "RECORD PROBLEM" -strLogValue "Canceled"

        <# Command line actions #>
        If ($Global:bolComingFromMenu -eq $false) {

            <# Set back window title #>
            $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

            <# Exit #>
            Return

        }

        <# ShowMenu actions #>
        If ($Global:bolComingFromMenu -eq $true) {

            <# Clear console #>
            Clear-Host

            <# Call ShowMenu #>
            fncShowMenu

            Return

        }

    }
    Else { <# Any key actions #>

        <# Logging #>
        fncLogging -strLogFunction "fncRecordProblem" -strLogDescription "RECORD PROBLEM" -strLogValue "Canceled"

        <# Command line actions #>
        If ($Global:bolComingFromMenu -eq $false) {

            <# Set back window title #>
            $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

            <# Exit #>
            Return

        }

        <# ShowMenu actions #>
        If ($Global:bolComingFromMenu -eq $true) {

            <# Clear console #>
            Clear-Host

            <# Call ShowMenu #>
            fncShowMenu

            Return

        }

    }

    <# Release variable #>
    $Private:ReadHost = $null

}

Function fncEnableLogging {

    <# Logging #>
    fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "Enable logging" -strLogValue "Triggered"

    <# Implement registry key for function fncValidateForActivatedLogging to check whether logging was left enabled (for problem record) #>
    If ($(Test-Path -Path "HKCU:\SOFTWARE\Schiroky\ComplianceUtility") -Eq $false) { <# Check, if path exist (to check for logging enabled), and create it if not #>

        <# Create registry key #>
        New-Item -Path "HKCU:\SOFTWARE\Schiroky\ComplianceUtility" -Force | Out-Null

    }

    <# Implement registry key to check for enabled logging on next start, and rollback settings if necessary #>
    New-ItemProperty -Path "HKCU:\SOFTWARE\Schiroky\ComplianceUtility" -Name "LoggingActivated" -Value $true -PropertyType DWord -Force -ErrorAction SilentlyContinue | Out-Null

    <# Progress bar #>
    Write-Progress -Activity " Enable logging, please wait..." -PercentComplete 0

    <# Check for administrative privileges, and enabling corresponding logs #>
    If ($Global:bolRunningPrivileged -eq $true) {

        <# Progress bar update #>
        Write-Progress -Activity " Enable logging: CAPI2 event logging..." -PercentComplete (100/7 * 1)

        <# Enable CAPI2 event log #>
        Write-Output Y | wevtutil set-log Microsoft-Windows-CAPI2/Operational /enabled:True

        <# Logging #>
        fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "CAPI2 event log" -strLogValue "Enabled"

        <# Clear CAPI2 event log #>
        wevtutil.exe clear-log Microsoft-Windows-CAPI2/Operational

        <# Logging #>
        fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "CAPI2 event log" -strLogValue "Cleared"

    }

    <# Progress bar update #>
    Write-Progress -Activity " Enable logging: Office logging..." -PercentComplete (100/7 * 2)

    <# Enable Office logging for 2016 (16.0) #>
    If ($(Test-Path -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Logging") -Eq $false) {

        <# Create registry key, if does not exist #>
        New-Item -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Logging" -Force | Out-Null

    }

    <# Check for registry key "Logging" (2016 x64) #>
    If ($(Test-Path -Path "HKCU:\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\Logging") -Eq $false) {

        <# Create registry key, if does not exist #>
        New-Item -Path "HKCU:\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\Logging" -Force | Out-Null

    }

    <# Implement/enable Office logging #>
    New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Logging" -Name "EnableLogging" -Value 1 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path "HKCU:\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\Logging" -Name "EnableLogging" -Value 1 -PropertyType DWord -Force | Out-Null

    <# Logging #>
    fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "Office Logging" -strLogValue "Enabled"

    <# Progress bar update #>
    Write-Progress -Activity " Enable logging: Office TCOTrace..." -PercentComplete (100/7 * 3)

    <# <# Check for registry key "Debug" (2016) #>
    If ($(Test-Path -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Debug") -Eq $false) {

        <# Create registry key #>
        New-Item -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Debug" -Force | Out-Null

    }
    <# Enable Office TCOTrace logging #>
    New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Debug" -Name "TCOTrace" -Value 1 -PropertyType DWord -Force | Out-Null

    <# Logging #>
    fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "Office TCOTrace" -strLogValue "Enabled"

    <# Progress bar update #>
    Write-Progress -Activity " Enable logging: Cleaning MSIP/MSIPC logs..." -PercentComplete (100/7 * 4)

    <# Clean MSIP/MSIPC/AIP v2 logs folder #>
    If ($(Test-Path -Path $env:LOCALAPPDATA\Microsoft\MSIP\Logs) -Eq $true) { <# If foler exist #>

        <# Clean MSIP/AIP v1/2 log folder content #>
        Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\MSIP\Logs" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null

        <# Logging #>
        fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "MSIP log folder" -strLogValue "Cleared"

    }

    <# Check for MSIPC folder #>
    If ($(Test-Path -Path $env:LOCALAPPDATA\Microsoft\MSIPC\Logs) -Eq $true) {

        <# Clean MSIPC log folder #>
        Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\MSIPC\Logs" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null

        <# Logging #>
        fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "MSIPC log folder" -strLogValue "Cleared"

    }

    <# Check for MSIP folder #>
    If ($(Test-Path -Path $env:LOCALAPPDATA\Microsoft\MSIP\mip) -Eq $true) {

        <# Clean MIP SDK/AIP v2 log folder #>
        Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\MSIP\mip" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null

        <# Logging #>
        fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "MIP log folder" -strLogValue "Cleared"

    }

    <# Check for DLP folder #>
    If ($(Test-Path -Path $env:LOCALAPPDATA\Microsoft\Office\DLP) -Eq $true) {

        <# Clean Office DLP/MIP log folder #>
        Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Office\DLP" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null

        <# Logging #>
        fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "Office DLP log folder" -strLogValue "Cleared"

    }

    <# Check for Word MIPSDK log folder #>
    If ($(Test-Path -Path $env:LOCALAPPDATA\Microsoft\Word\MIPSDK\mip) -Eq $true) {

        <# Clean Word MIPSDK log folder #>
        Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Word\MIPSDK\mip" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null

        <# Logging #>
        fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "Word MIPSDK log folder" -strLogValue "Cleared"

    }

    <# Check for Excel MIPSDK log folder #>
    If ($(Test-Path -Path $env:LOCALAPPDATA\Microsoft\Excel\MIPSDK\mip) -Eq $true) {

        <# Clean Excel MIPSDK log folder #>
        Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Excel\MIPSDK\mip" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null

        <# Logging #>
        fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "Excel MIPSDK log folder" -strLogValue "Cleared"

    }

    <# Check for PowerPoint MIPSDK log folder #>
    If ($(Test-Path -Path $env:LOCALAPPDATA\Microsoft\PowerPoint\MIPSDK\mip) -Eq $true) {

        <# Clean PowerPoint MIPSDK log folder #>
        Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\PowerPoint\MIPSDK\mip" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null

        <# Logging #>
        fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "PowerPoint MIPSDK log folder" -strLogValue "Cleared"

    }

    <# Check for Outlook MIPSDK log folder #>
    If ($(Test-Path -Path $env:LOCALAPPDATA\Microsoft\Outlook\MIPSDK\mip) -Eq $true) {

        <# Clean Outlook MIPSDK log folder #>
        Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Outlook\MIPSDK\mip" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null

        <# Logging #>
        fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "Outlook MIPSDK log folder" -strLogValue "Cleared"

    }

    <# Check for Outlook MIPSDKPDF log folder #>
    If ($(Test-Path -Path $env:LOCALAPPDATA\Microsoft\Outlook\MIPSDKPDF\mip) -Eq $true) {

        <# Clean Outlook MIPSDKPDF log folder #>
        Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Outlook\MIPSDKPDF\mip" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null

        <# Logging #>
        fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "Outlook MIPSDKPDF log folder" -strLogValue "Cleared"

    }

    <# Check for OneNote MIPSDK log folder #>
    If ($(Test-Path -Path $env:LOCALAPPDATA\Microsoft\OneNote\MIPSDK\mip) -Eq $true) {

        <# Clean OneNote MIPSDK log folder #>
        Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\OneNote\MIPSDK\mip" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null

        <# Logging #>
        fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "OneNote MIPSDK log folder" -strLogValue "Cleared"

    }

    <# Check for Office MIPSDK log folder #>
    If ($(Test-Path -Path $env:LOCALAPPDATA\Microsoft\Office\MIPSDK\mip) -Eq $true) {

        <# Clean Office MIPSDK log folder #>
        Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Office\MIPSDK\mip" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null

        <# Logging #>
        fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "Office MIPSDK log folder" -strLogValue "Cleared"

    }

    <# Detect Diagnostic folder #>
    If ($(Test-Path -Path $env:TEMP\Diagnostics) -Eq $true) {

        <# Clean Office Diagnostics folder #>
        Remove-Item -Path "$env:TEMP\Diagnostics" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null

        <# Logging #>
        fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "Office Diagnostics log folder" -strLogValue "Cleared"

    }

    <# Progress bar update #>
    Write-Progress -Activity " Enable logging: Flushing DNS..." -PercentComplete (100/7 * 5)

    <# Flush DNS #>
    ipconfig.exe /flushdns | Out-Null

    <# Logging #>
    fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "Flush DNS" -strLogValue "Called"

    <# Progress bar update #>
    Write-Progress -Activity " Enable logging: Starting PSR..." -PercentComplete (100/7 * 6)

    <# Start PSR #>
    psr.exe /gui 0 /start /output "$Global:strUniqueLogFolder\ProblemSteps.zip"

    <# Logging #>
    fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "PSR" -strLogValue "Started"

    <# Clean temp folder for office.log (TCOTrace) #>
    If ($(Test-Path $Global:strTempFolder"\office.log") -Eq $true) {

        <# Remove file office.log #>
        Remove-Item -Path "$Global:strTempFolder\office.log" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null

        <# Logging #>
        fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "Office TCOTrace temp file" -strLogValue "Cleared"

    }

    <# Clean temp folder for office log (machine name) #>
    If ($(Test-Path "$Global:strTempFolder\$([System.Environment]::MachineName)*.log") -Eq $true) {

        <# Remove file office.log #>
        Remove-Item -Path "$Global:strTempFolder\$([System.Environment]::MachineName)*.log" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null

        <# Logging #>
        fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "Office log temp file" -strLogValue "Cleared"

    }

    <# Progress bar update #>
    Write-Progress -Activity "  Logging enabled" -Completed

    <# Logging #>
    fncLogging -strLogFunction "fncEnableLogging" -strLogDescription "Enable logging" -strLogValue "Proceeded"

}

Function fncDisableLogging {

    <# Logging #>
    fncLogging -strLogFunction "fncDisableLogging" -strLogDescription "Disable logging" -strLogValue "Triggered"

    <# Progress bar #>
    Write-Progress -Activity " Disable logging, please wait..." -PercentComplete 0

    <# Check for administrative privileges, and enabling admininistrative actions #>
    If ($Global:bolRunningPrivileged -eq $true) {

        <# Progress bar update #>
        Write-Progress -Activity " Disable logging: CAPI2 event log..." -PercentComplete (100/5 * 1)

        <# Disable CAPI2 event log #>
        wevtutil.exe set-log Microsoft-Windows-CAPI2/Operational /enabled:false

        <# Logging #>
        fncLogging -strLogFunction "fncDisableLogging" -strLogDescription "CAPI2 event log" -strLogValue "Disabled"

    }

    <# Progress bar update #>
    Write-Progress -Activity " Disable logging: Office logging..." -PercentComplete (100/5 * 2)

    <# Disable Office logging for  2016 (16.0) #>
    fncDeleteItem "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Logging"
    fncDeleteItem "HKCU:\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\Logging"

    <# Logging #>
    fncLogging -strLogFunction "fncDisableLogging" -strLogDescription "Office Logging" -strLogValue "Disabled"

    <# Progress bar update #>
    Write-Progress -Activity " Disable logging: Office TCOTrace..." -PercentComplete (100/5 * 3)

    <# Disable Office TCOTrace logging #>
    fncDeleteItem "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Debug"

    <# Logging #>
    fncLogging -strLogFunction "fncDisableLogging" -strLogDescription "Office TCOTrace" -strLogValue "Disabled"

    <# Progress bar update #>
    Write-Progress -Activity " Disable logging: PSR..." -PercentComplete (100/5 * 4)

    <# Stop PSR #>
    psr.exe /stop

    <# Logging #>
    fncLogging -strLogFunction "fncDisableLogging" -strLogDescription "PSR" -strLogValue "Disabled"

    <# Implement registry key for fncValidateForActivatedLogging to check whether logging was left enabled (for problem record) #>
    If ($(Test-Path -Path "HKCU:\SOFTWARE\Schiroky\ComplianceUtility") -Eq $false) { <# Detect/create path to check for logging enabled #>

        <# Create registry key #>
        New-Item -Path "HKCU:\SOFTWARE\Schiroky\ComplianceUtility" -Force | Out-Null

    }

    <# Implement registry key to check for enabled logging on next start, and rollback settings if necessary #>
    New-ItemProperty -Path "HKCU:\SOFTWARE\Schiroky\ComplianceUtility" -Name "LoggingActivated" -Value $false -PropertyType DWord -Force -ErrorAction SilentlyContinue | Out-Null

    <# Progress bar update #>
    Write-Progress -Activity " Logging disabled" -Completed

    <# Logging #>
    fncLogging -strLogFunction "fncDisableLogging" -strLogDescription "Disable logging" -strLogValue "Proceeded"

}

<# Check whether logging (for problem record) was left enabled #>
Function fncValidateForActivatedLogging {

    <# Read registry key to check for enabled logging. Used in fncEnableLogging, and fncDisableLogging #>
    If ((Get-ItemProperty "HKCU:\SOFTWARE\Schiroky\ComplianceUtility" -Name LoggingActivated -ErrorAction SilentlyContinue).LoggingActivated -eq $true) {

         <# Logging #>
        fncLogging -strLogFunction "fncValidateForActivatedLogging" -strLogDescription "Disable logging" -strLogValue "Initiated"

        <# Call DisableLogging #>
        fncDisableLogging

    }

}

Function fncCollectingLogs {

    <# Logging #>
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Collecting logs" -strLogValue "Triggered"

    <# Progress bar #>
    Write-Progress -Activity " Collecting logs, please wait..." -PercentComplete 0

    <# Collecting system information #>
    Get-ComputerInfo > "$Global:strUniqueLogFolder\SystemInformation.log"

    <# Collecting device join status #>
    dsregcmd.exe /status | Out-File "$Global:strUniqueLogFolder\SystemInformation.log" -Append

    <# Logging #>
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export system information" -strLogValue "SystemInformation.log"

    <# Check for administrative permissons, and enabling admininistrative actions #>
    If ($Global:bolRunningPrivileged -eq $true) {

        <# Progress bar update #>
        Write-Progress -Activity " Collecting logs: CAPI2 event log..." -PercentComplete (100/26 * 1)

        <# Export CAPI2 event log #>
        wevtutil.exe export-log Microsoft-Windows-CAPI2/Operational "$Global:strUniqueLogFolder\CAPI2.evtx" /overwrite:true

        <# Logging #>
        fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export CAPI2 event log" -strLogValue "CAPI2.evtx"

        <# Progress bar update #>
        Write-Progress -Activity " Collecting logs: Azure Information Protection event log..." -PercentComplete (100/26 * 2)

        <# Actions when AIP event log exist #>
        If ([System.Diagnostics.EventLog]::Exists("Azure Information Protection") -Eq $true) {

            <# Export AIP event log #>
            wevtutil.exe export-log "Azure Information Protection" "$Global:strUniqueLogFolder\AIP.evtx" /overwrite:true

            <# Logging #>
            fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export AIP event log" -strLogValue "AIP.evtx"

        }

        <# Progress bar update #>
        Write-Progress -Activity " Collecting logs: Purview Information Protection event log..." -PercentComplete (100/27 * 3)

        <# Actions when 'PIP' event log exist #>
        If ([System.Diagnostics.EventLog]::Exists("Microsoft Purview Information Protection") -Eq $true) {

            <# Export AIP event log #>
            wevtutil.exe export-log "Microsoft Purview Information Protection" "$Global:strUniqueLogFolder\PIP.evtx" /overwrite:true

            <# Logging #>
            fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export PIP event log" -strLogValue "PIP.evtx"

        }

        <# Progress bar update #>
        Write-Progress -Activity " Collecting logs: Filter drivers..." -PercentComplete (100/26 * 4)

        <# Export filter drivers #>
        fltmc.exe filters > "$Global:strUniqueLogFolder\Filters.log"

        <# Logging #>
        fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export filter drivers" -strLogValue "Filters.log"

    }

    <# Progress bar update #>
    Write-Progress -Activity " Collecting logs: PSR recording..." -PercentComplete (100/26 * 5)

    <# Stop PSR #>
    psr.exe /stop

    <# Logging #>
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "PSR" -strLogValue "Stopped"
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export PSR" -strLogValue "ProblemSteps.zip"

    <# Progress bar update #>
    Write-Progress -Activity " Collecting logs: Application event log..." -PercentComplete (100/26 * 6)

    <# Export Application event log #>
    wevtutil.exe export-log Application "$Global:strUniqueLogFolder\Application.evtx" /overwrite:true

    <# Logging #>
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export Application event log" -strLogValue "Application.evtx"

    <# Progress bar update #>
    Write-Progress -Activity " Collecting logs: System event log..." -PercentComplete (100/26 * 7)

    <# Export System event log #>
    wevtutil.exe export-log System "$Global:strUniqueLogFolder\System.evtx" /overwrite:true

    <# Logging #>
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export System event log" -strLogValue "System.evtx"

    <# Progress bar update #>
     Write-Progress -Activity " Collecting logs: Office log files..." -PercentComplete (100/26 * 8)

    <# Check for Office log path and create it, if it not exist #>
    If ($(Test-Path -Path "$Global:strUniqueLogFolder\Office") -Eq $false) {

        <# Create Office log folder #>
        New-Item -ItemType Directory -Force -Path "$Global:strUniqueLogFolder\Office" | Out-Null

        <# Check for Office CLP path #>
        If ($(Test-Path -Path $env:LOCALAPPDATA\Microsoft\Office\CLP) -Eq $true) {

            <# Perform action only, if the CLP folder contain files (Note: Afer a RESET this folder is empty). #>
            If (((Get-ChildItem -LiteralPath $env:LOCALAPPDATA\Microsoft\Office\CLP -File -Force | Select-Object -First 1 | Measure-Object).Count -ne 0)) {

                <# Compress label and policy xml files into zip file (overwrites) #>
                Compress-Archive -Path $env:LOCALAPPDATA\Microsoft\Office\CLP"\*" -DestinationPath "$Global:strUniqueLogFolder\Office\OfficeCLP" -Force -ErrorAction SilentlyContinue

                <# Logging #>
                fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export Office CLP" -strLogValue "\Office\OfficeCLP.zip"

            }

        }

        <# Check for Office DLP path #>
        If ($(Test-Path -Path "$Global:strUniqueLogFolder\Office\DLP") -Eq $true) {

            <# Perform action only, if the DLP folder contain files (Note: Afer a RESET this folder is empty). #>
            If (((Get-ChildItem -LiteralPath $env:LOCALAPPDATA\Microsoft\Office\DLP -File -Force | Select-Object -First 1 | Measure-Object).Count -ne 0)) {

                <# Compress DLP folder content into zip file (overwrites) #>
                Compress-Archive -Path $env:LOCALAPPDATA\Microsoft\Office\DLP"\*" -DestinationPath "$Global:strUniqueLogFolder\Office\OfficeDLP" -Force -ErrorAction SilentlyContinue

                <# Logging #>
                fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export Office DLP" -strLogValue "\Office\OfficeDLP.zip"

            }

        }

        <# Check for Outlook MIPSDKPDF path #>
        If ($(Test-Path -Path "$Global:strUniqueLogFolder\Outlook\MIPSDKPDF\mip") -Eq $true) {

            <# Perform action only, if the MIPSDKPDF folder contain files (Note: Afer a RESET this folder is empty). #>
            If (((Get-ChildItem -LiteralPath $env:LOCALAPPDATA\Microsoft\Outlook\MIPSDKPDF\mip -File -Force | Select-Object -First 1 | Measure-Object).Count -ne 0)) {

                <# Compress MIPSDKPDF folder content into zip file (overwrites) #>
                Compress-Archive -Path $env:LOCALAPPDATA\Microsoft\Office\Outlook\MIPSDKPDF"\*" -DestinationPath "$Global:strUniqueLogFolder\Office\MIPSDKPDF-Outlook" -Force -ErrorAction SilentlyContinue

                <# Logging #>
                fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export Outlook MIPSDKPDF" -strLogValue "\Office\MIPSDKPDF-Outlook.zip"

            }

        }

    }

    <# Define array for MIPSDK apps/folders #>
    $Private:arrMIPSDKApps = "Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Office"
    $Private:strMipPathItem

    <# Loop though array and collect MIPSDK logs #>
    ForEach ($Private:strMipPathItem in $Private:arrMIPSDKApps) {

        <# Check for each App MIPSDK log path, and collect log files #>
        If ($(Test-Path -Path $env:LOCALAPPDATA\Microsoft\$Private:strMipPathItem\MIPSDK\mip) -Eq $true) {

            <# Collect MIPSDK log folder only, if the folder contains files (Note: Afer a RESET this folder is empty). #>
            If (((Get-ChildItem -LiteralPath $env:LOCALAPPDATA\Microsoft\$Private:strMipPathItem\MIPSDK\mip -File -Force | Select-Object -First 1 | Measure-Object).Count -ne 0)) {

                <# Compress MIPSDK\mip content to .zip file (overwrites) #>
                Compress-Archive -Path $env:LOCALAPPDATA\Microsoft\$Private:strMipPathItem\MIPSDK\mip"\*" -DestinationPath "$Global:strUniqueLogFolder\Office\MIPSDK-$Private:strMipPathItem.zip" -Force -ErrorAction SilentlyContinue

                <# Logging #>
                fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export $Private:strMipPathItem MIPSDK logs" -strLogValue "\Office\MIPDSK-$Private:strMipPathItem.zip"

            }

        }

    }

    <# Releasing MIPSDK variables/array #>
    $Private:arrMIPSDKApps = $null
    $Private:strMipPathItem = $null

    <# Copy Office Diagnostics folder from temp folder to Office logs folder #>
    fncCopyItem $env:TEMP\Diagnostics "$Global:strUniqueLogFolder\Office" "Diagnostics\*"

    <# Logging #>
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export Office Diagnostics logs" -strLogValue "\Office\Diagnostics"

    <# Copy office log files from temp folder to logs folder #>
    fncCopyItem $Global:strTempFolder"\office.log" "$Global:strUniqueLogFolder\Office\office.log" "office.log"

    <# Logging #>
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export Office log" -strLogValue "office.log"

    <# Copy Office logging files for 2016 (16.0) to logs folder #>
    fncCopyItem "$Global:strTempFolder\$([System.Environment]::MachineName)*.log" "$Global:strUniqueLogFolder\Office" "Office\$([System.Environment]::MachineName)*.log"

    <# Logging #>
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export Office log" -strLogValue "\Office"

    <# Clean Office log files from temp folder #>
    fncDeleteItem "$Global:strTempFolder\$([System.Environment]::MachineName)*.log"
    fncDeleteItem "$Global:strTempFolder\Office.log"

    # Progress bar update #>
    Write-Progress -Activity " Collecting logs: AIP/PIP/Office Diagnostics logs folders..." -PercentComplete (100/26 * 9)

    <# Remember default progress bar status: 'Continue' #>
    $Private:strOriginalPreference = $Global:ProgressPreference
    $Global:ProgressPreference = "SilentlyContinue" <# Hiding progress bar #>

    <# Export AIP logs folder #>
    If (Get-Module -ListAvailable -Name AzureInformationProtection, PurviewInformationProtection) {

        <# Check for AIP #>
        If (Get-Module -ListAvailable -Name AzureInformationProtection){

            <# Logging AIP client version #>
            fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "AIP client version" -strLogValue $((Get-Module -ListAvailable -Name AzureInformationProtection).Version).ToString()

            <# Actions on PowerShell Core (7.x) for compatibility mode #>
            If ($PSVersionTable.PSEdition.ToString() -eq "Core") {

                <# Remove AzureInformationProtection module, because it's not compatible with PowerShell Core (7.x) #>
                Remove-Module -Name AzureInformationProtection -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

                <# Import AzureInformationProtection module in compatiblity mode #>
                Import-Module -Name AzureInformationProtection -UseWindowsPowerShell -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

                <# Logging #>
                fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "AzureInformationProtection compatiblity mode" -strLogValue $true

            }

        }

        <# Check for PIP #>
        If (Get-Module -ListAvailable -Name PurviewInformationProtection){

            <# Logging: PIP client version #>
            fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "PIP client version" -strLogValue $((Get-Module -ListAvailable -Name PurviewInformationProtection).Version).ToString()

            <# Actions on PowerShell Core (7.x) for compatibility mode #>
            If ($PSVersionTable.PSEdition.ToString() -eq "Core") {

                <# Remove AzureInformationProtection module, because it's not compatible with PowerShell Core (7.x) #>
                Remove-Module -Name PurviewInformationProtection -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

                <# Import AzureInformationProtection module in compatiblity mode #>
                Import-Module -Name PurviewInformationProtection -UseWindowsPowerShell -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

                <# Logging #>
                fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "PurviewInformationProtection compatiblity mode" -strLogValue $true

            }

        }

        <# Try to export log folders with authentication; fails without #>
        Try {

            <# Export AIP log folders #>
            Export-AIPLogs -FileName "$Global:strUniqueLogFolder\AIPLogs.zip" | Out-Null

            <# Logging #>
            fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export AIP Log folders" -strLogValue $true

        }
        Catch{ <# Actions without authentication #>

            <# Clear authentication #>
            Clear-AIPAuthentication -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

            <# Output #>
            Write-Output "Please authenticate with your user credentials to retrieve your AIP/PIP log folders."

            Try {

                <# Authenticate for accessing logs #>
                Set-AIPAuthentication -ErrorAction Stop | Out-Null

                <# Export AIP log folders #>
                Export-AIPLogs -FileName "$Global:strUniqueLogFolder\AIPLogs.zip" -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null

                <# Logging #>
                fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export AIP Log folders" -strLogValue $true

            }
            Catch {

                $Private:strAIPLogExportError = $_.Exception.Message

                If ($Private:strAIPLogExportError -match "(?i)cancel(?:ed|led)?") {
                    Write-ColoredOutput Yellow "AIP/PIP log folder collection was canceled by the user. Continuing with the remaining logs.`n"
                    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export AIP Log folders" -strLogValue "Canceled by user"
                }
                Else {
                    Write-ColoredOutput Red "AIP/PIP log folder collection failed: $Private:strAIPLogExportError`n"
                    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export AIP Log folders" -strLogValue "Failed: $Private:strAIPLogExportError"
                }

                $Private:strAIPLogExportError = $null

            }

        }

    }
    Else {<# Action without any AIP/PIP client #>

        <# Logging: If no AIP/PIP client is installed #>
        fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "AIP/PIP client installed" -strLogValue $false

        <# Export Office DLP content to logs folder #>
        fncCopyItem $env:LOCALAPPDATA\Microsoft\Office\DLP "$Global:strUniqueLogFolder\Office" "DLP\*"

        <# Logging #>
        fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export Office DLP content" -strLogValue "\Office\DLP"

        <# Export Office Diagnostics content to logs folder #>
        fncCopyItem $env:TEMP\Diagnostics "$Global:strUniqueLogFolder\Office" "Diagnostics\*"

        <# Logging #>
        fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export Office Diagnostics content" -strLogValue "\Office"

        <# Export MSIP/MSIPC content to logs folder #>
        fncCopyItem $env:LOCALAPPDATA\Microsoft\MSIP "$Global:strUniqueLogFolder\MSIP" "MSIP\*"

        <# Logging #>
        fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export MSIP content" -strLogValue "\MSIP"

        <# Copy files to logs folder #>
        fncCopyItem $env:LOCALAPPDATA\Microsoft\MSIPC "$Global:strUniqueLogFolder\MSIPC" "MSIPC\*"

        <# Logging #>
        fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export MSIPC content" -strLogValue "\MSIPC"

    }

    <# Set back progress bar to previous setting #>
    $Global:ProgressPreference = $Private:strOriginalPreference

    <# Progress bar update #>
    Write-Progress -Activity " Collecting logs: WinHTTP..." -PercentComplete (100/26 * 10)

    <# Export WinHTTP #>
    netsh.exe winhttp show proxy > "$Global:strUniqueLogFolder\WinHTTP.log"

    <# Logging #>
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export WinHTTP" -strLogValue "WinHTTP.log"

    <# Progress bar update #>
    Write-Progress -Activity " Collecting logs: WinHTTP (WoW6432)..." -PercentComplete (100/26 * 11)

    <# Export WinHTTP_WoW6432 (only 64-bit OS) #>
    If ((Get-CimInstance Win32_OperatingSystem  -Verbose:$false).OSArchitecture -eq "64-bit") {

        & $env:WINDIR\SysWOW64\netsh.exe winhttp show proxy > "$Global:strUniqueLogFolder\WinHTTP_WoW6432.log"

        <# Logging #>
        fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export WinHTTP_WoW6432" -strLogValue "WinHTTP_WoW6432.log"

    }

    <# Export AutoConfigURL #>
    If ((Get-ItemProperty "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\" -Name AutoConfigURL -ErrorAction SilentlyContinue).AutoConfigURL) {

        <# Progress bar update #>
        Write-Progress -Activity " Collecting logs: AutoConfigURL..." -PercentComplete (100/26 * 12)

        <# Logging #>
        fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export IE AutoConfigURL" -strLogValue "AutoConfigURL.log" <# Windows version and release ID #>

        <# Export AutoConfigURL #>
        Get-ItemProperty "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings" | Select-Object AutoConfigURL > "$Global:strUniqueLogFolder\AutoConfigURL.log"

    }

    <# Progress bar update #>
    Write-Progress -Activity " Collecting logs: Machine certificates..." -PercentComplete (100/26 * 13)

    <# Export machine certificates #>
    certutil.exe -silent -store my > "$Global:strUniqueLogFolder\CertMachine.log"

    <# Logging #>
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export machine certificates" -strLogValue "CertMachine.log"

    <# Progress bar update #>
    Write-Progress -Activity " Collecting logs: User certificates..." -PercentComplete (100/26 * 14)

    <# Export user certificates #>
    certutil.exe -silent -user -store my > "$Global:strUniqueLogFolder\CertUser.log"

    <# Logging #>
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export user certificates" -strLogValue "CertUser.log"

    <# Progress bar update #>
    Write-Progress -Activity " Collecting logs: Credentials information..." -PercentComplete (100/26 * 15)

    <# Export Credential Manager data #>
    cmdkey.exe /list > "$Global:strUniqueLogFolder\CredMan.log"

    <# Logging #>
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export Credential Manager" -strLogValue "CredMan.log"

    <# Progress bar update #>
    Write-Progress -Activity " Collecting logs: IP configuration..." -PercentComplete (100/26 * 16)

    <# Export IP configuration #>
    ipconfig.exe /all > "$Global:strUniqueLogFolder\IPConfigAll.log"

    <# Logging #>
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export ipconfig" -strLogValue "IPConfigAll.log"

    <# Progress bar update #>
    Write-Progress -Activity " Collecting logs: DNS..." -PercentComplete (100/26 * 17)

    <# Export DNS configuration  #>
    ipconfig.exe /displaydns > "$Global:strUniqueLogFolder\WinIPConfig.txt" | Out-Null

    <# Logging #>
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export DNS" -strLogValue "WinIPConfig.txt"

    <# Progress bar update #>
    Write-Progress -Activity " Collecting logs: Environment information..." -PercentComplete (100/26 * 18)

    <# Export environment variables #>
    Get-ChildItem Env: | Out-File "$Global:strUniqueLogFolder\EnvVar.log"

    <# Logging #>
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export environment variables" -strLogValue "EnvVar.log"

    <# Progress bar update #>
    Write-Progress -Activity " Collecting logs: Group policy report..." -PercentComplete (100/26 * 19)

    <# Export group policies #>
    gpresult /f /h "$Global:strUniqueLogFolder\Gpresult.htm" | Out-Null

    <# Logging #>
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export group policy report" -strLogValue "Gpresult.htm"

    <# Progress bar update #>
    Write-Progress -Activity " Collecting logs: Time zone information..." -PercentComplete (100/26 * 20)

    <# Export timezone offse #>
    (Get-Timezone).BaseUTCOffset.Hours | Out-File "$Global:strUniqueLogFolder\BaseUTCOffset.log"

    <# Logging #>
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export timezone offset" -strLogValue "BaseUTCOffset.log"

    <# Progress bar update #>
    Write-Progress -Activity " Collecting logs: Tasklist..." -PercentComplete (100/26 * 21)

    <# Export Tasklist #>
    Tasklist.exe /svc > "$Global:strUniqueLogFolder\Tasklist.log"

    <# Logging #>
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export Tasklist" -strLogValue "Tasklist.log"

    <# Progress bar update #>
    Write-Progress -Activity " Collecting logs: Programs and Features..." -PercentComplete (100/26 * 22)

    <# Export Programs and Features (32) #>
    If ($(Test-Path -Path "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall") -Eq $true) {

        <# Programs32 #>
        Get-ItemProperty "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | Export-Csv -Path "$Global:strUniqueLogFolder\Programs32.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction SilentlyContinue

        <# Logging #>
        fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export Programs (x86)" -strLogValue "Programs32.csv"

    }

    <# Export Programs and Features (64) #>
    Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | Export-Csv -Path  "$Global:strUniqueLogFolder\Programs64.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction SilentlyContinue

    <# Logging #>
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export Programs (x64)" -strLogValue "Programs64.csv"

    <# Progress bar update #>
    Write-Progress -Activity " Collecting logs: Scheduled Tasks..." -PercentComplete (100/26 * 23)

    <# Array to collect Scheduled Tasks data #>
    [System.Collections.ArrayList]$Private:arrScheduledTasks = @()

    <# Variable for task data #>
    $Private:strAllTasks = Get-ScheduledTask

    <# Looping trouth all Scheduled Tasks #>
    ForEach ($Private:strTask in $Private:strAllTasks) {

        <# Variable to collect task details #>
        $Private:strTaskInfo = $Private:strTask | Get-ScheduledTaskInfo

        <# Collecing data when NextRunTime is not empty #>
        If ( -not ($Null -eq $Private:strTaskInfo.NextRunTime)){
            $Private:arrScheduledTasks.Add([PSCustomObject]@{
                TaskName    = $Private:strTask.TaskName
                TaskPath    = $Private:strTask.TaskPath
                NextRunTime = $Private:strTaskInfo.NextRunTime
                State       = $Private:strTask.State}) | Out-Null
        }

    }

    <# Export Scheduled Tasks #>
    $Private:arrScheduledTasks | Sort-Object -Property 'NextRunTime' | Export-Csv -Path "$Global:strUniqueLogFolder\ScheduledTasks.csv" -NoTypeInformation -Delimiter ";" -Encoding UTF8

    <# Logging #>
     fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export Scheduled Tasks" -strLogValue "ScheduledTasks.csv"

    <# Progress bar update #>
    Write-Progress -Activity " Collecting logs: AIP registry keys..." -PercentComplete (100/26 * 24)

    <# Export AIP plugin Adobe Acrobat RMS logs #>
    If ($(Test-Path -Path $env:LOCALAPPDATA\Microsoft\RMSLocalStorage\MIP\logs) -Eq $true) {

        <# Progress bar update #>
        Write-Progress -Activity " Collecting logs: Adobe logs..." -PercentComplete (100/26 * 25)

        <# Export MSIP/MSIPC content to logs folder #>
        fncCopyItem $env:LOCALAPPDATA\Microsoft\RMSLocalStorage\MIP\logs "$Global:strUniqueLogFolder\Adobe\LOCALAPPDATA" "Adobe\LOCALAPPDATA\*"

        <# Logging #>
        fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export Adobe logs" -strLogValue "\Adobe"

    }

    <# Export AIP plugin Adobe Acrobat RMS logs #>
    If ($(Test-Path -Path $env:USERPROFILE\appdata\locallow\Microsoft\RMSLocalStorage\mip\logs) -Eq $true) {

        <# Export MSIP/MSIPC content to logs folder #>
        fncCopyItem $env:USERPROFILE\appdata\locallow\Microsoft\RMSLocalStorage\mip\logs "$Global:strUniqueLogFolder\Adobe\USERPROFILE" "Adobe\USERPROFILE\*"

        <# Logging #>
        fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export Adobe logs" -strLogValue "\Adobe"

    }

    <# Export several registry keys: Define an array and feeding it with related registry keys #>
    $Private:arrRegistryKeys = "HKLM:\Software\Classes\MSIP.ExcelAddin",
                               "HKLM:\Software\Classes\MSIP.WordAddin",
                               "HKLM:\SOFTWARE\Classes\MSIP.PowerPointAddin",
                               "HKLM:\SOFTWARE\Classes\MSIP.OutlookAddin",
                               "HKLM:\SOFTWARE\Classes\AllFileSystemObjects\shell\Microsoft.Azip.RightClick",
                               "HKLM:\SOFTWARE\Microsoft\MSIPC",
                               "HKLM:\SOFTWARE\Microsoft\Office\Word\Addins",
                               "HKLM:\SOFTWARE\Microsoft\Office\Excel\Addins",
                               "HKLM:\SOFTWARE\Microsoft\Office\PowerPoint\Addins",
                               "HKLM:\SOFTWARE\Microsoft\Office\Outlook\Addins",
                               "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\SOFTWARE\Microsoft\Office\Word\Addins",
                               "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\SOFTWARE\Microsoft\Office\Excel\Addins",
                               "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\SOFTWARE\Microsoft\Office\PowerPoint\Addins",
                               "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\SOFTWARE\Microsoft\Office\Outlook\Addins",
                               "HKLM:\SOFTWARE\WOW6432Node\Microsoft\MSIPC",
                               "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\Word\Addins",
                               "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\Excel\Addins",
                               "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\PowerPoint\Addins",
                               "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\Outlook\Addins",
                               "HKLM:\SOFTWARE\Policies",
                               "HKLM:\SOFTWARE\Microsoft\PolicyManager\AdmxDefault",
                               "HKCU:\SOFTWARE\Microsoft\Policies",
                               "HKCU:\SOFTWARE\Microsoft\MSIP",
                               "HKCU:\Software\Microsoft\Office\16.0\Common\Identity",
                               "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Internet",
                               "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\DRM",
                               "HKCU:\SOFTWARE\Microsoft\Office\Word\Addins",
                               "HKCU:\SOFTWARE\Microsoft\Office\Excel\Addins",
                               "HKCU:\SOFTWARE\Microsoft\Office\PowerPoint\Addins",
                               "HKCU:\SOFTWARE\Microsoft\Office\Outlook\Addins",
                               "HKCU:\SOFTWARE\Microsoft\Office\16.0\Word\Resiliency",
                               "HKCU:\SOFTWARE\Microsoft\Office\16.0\Excel\Resiliency",
                               "HKCU:\SOFTWARE\Microsoft\Office\16.0\PowerPoint\Resiliency",
                               "HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Resiliency",
                               "HKCU:\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\MSIPC",
                               "HKCR:\MSIP.ExcelAddin",
                               "HKCR:\MSIP.WordAddin",
                               "HKCR:\MSIP.PowerPointAddin",
                               "HKCR:\MSIP.OutlookAddin",
                               "HKCR:\Local Settings\SOFTWARE\Microsoft\MSIPC",
                               "HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\DRM",
                               "HKCU:\SOFTWARE\Policies\Microsoft\Cloud\Office\16.0\Common\Security",
                               "HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security",
                               "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Security",
                               "HKCU:\Software\Microsoft\Office\16.0\Common\Licensing\CurrentSkuIdAggregationForApp",
                               "HKCU:\Software\Microsoft\Office\16.0\Common\Licensing\LastKnownC2RProductReleaseId"

    <# Define variable for reg values #>
    $Private:strRegValue

    <# Loop though array and cache to a temp file #>
    ForEach ($Private:strRegValue in $Private:arrRegistryKeys) {

        If ($(Test-Path -Path $Private:strRegValue) -Eq $true) {

            $Private:strTempFile = $Private:strTempFile + 1
            & REG EXPORT $Private:strRegValue.Replace(":", $null) "$Global:strTempFolder\$Private:strTempFile.reg" /Y | Out-Null <# Remove the ":" to export (replace) #>

        }

    }

    <# Releasing variable for reg values #>
    $Private:strRegValue = $null

    <# Insert first information; create log file #>
    "Windows Registry Editor Version 5.00" | Set-Content "$Global:strUniqueLogFolder\Registry.log"

    <# Read data from cached temp file, and add it to the logfile #>
    (Get-Content "$Global:strTempFolder\*.reg" | Where-Object {$_ -ne "Windows Registry Editor Version 5.00"} | Add-Content "$Global:strUniqueLogFolder\Registry.log")

    <# Clean temp folder of cached files #>
    Remove-Item "$Global:strTempFolder\*.reg" -Force -ErrorAction SilentlyContinue | Out-Null

    <# Logging #>
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Export AIP registry keys" -strLogValue "Registry.log"

    <# Progress bar update #>
    Write-Progress -Activity " Logs collected" -Completed

    <# Logging #>
    fncLogging -strLogFunction "fncCollectingLogs" -strLogDescription "Collecting logs" -strLogValue "Proceeded"

}

<# Verify that a module required by a collect operation is locally available #>
Function fncTestRequiredModule {

    <# Variables for checking module availability #>
    Param (
        [Parameter(Mandatory = $true)]
        [string]$strModuleName
    )

    <# Checking for module availability in cache, if not available, check locally and cache the result #>
    If (-not $script:RequiredModuleAvailability.ContainsKey($strModuleName)) {
        $script:RequiredModuleAvailability[$strModuleName] = [bool](Get-Module -ListAvailable -Name $strModuleName)
    }

    If ($script:RequiredModuleAvailability[$strModuleName]) {
        Return $true
    }

    <# Output #>
    Write-ColoredOutput Red "ATTENTION: The required PowerShell module '$strModuleName' is not installed.`nRun 'ComplianceUtility -UpdateModules' to install missing modules, then try again.`n"

    <# Logging #>
    fncLogging -strLogFunction "fncTestRequiredModule" -strLogDescription "$strModuleName module" -strLogValue "Not installed"

    Return $false
}

Function fncUpdateModules {

    <# Variables for checking module availability #>
    $Private:RequiredModules = @("AIPService", "ExchangeOnlineManagement", "Microsoft.Graph")
    $Private:UpdateSucceeded = $false
    $Private:ModulesChanged = $false

    <# Output #>
    Write-Output "CHECK AND UPDATE REQUIRED MODULES:"
    Write-Output "Initializing, please wait..."

    <# Logging #>
    fncLogging -strLogFunction "fncUpdateModules" -strLogDescription "UPDATE MODULES" -strLogValue "Initiated"

    <# Try and check for required modules, install missing and update outdated modules #>
    Try {

        If (-not (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue)) {

            Register-PSRepository -Default -ErrorAction Stop

        }

        If ($PSVersionTable.PSEdition -eq "Desktop") {

            $Private:NuGetProvider = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue

            If (-not $Private:NuGetProvider -or $Private:NuGetProvider.Version -lt [Version]"2.8.5.208") {

                <# Output #>
                Write-Output "Installing the NuGet package provider, please wait..."
                
                Install-PackageProvider -Name NuGet -MinimumVersion "2.8.5.208" -ForceBootstrap -Scope CurrentUser -ErrorAction Stop | Out-Null
                $Private:ModulesChanged = $true

            }
        }

        ForEach ($Private:ModuleName in $Private:RequiredModules) {

            $Private:LocalModule = Get-Module -ListAvailable -Name $Private:ModuleName | Sort-Object Version -Descending | Select-Object -First 1

            <# Query PowerShell Gallery exactly once for each module #>
            $Private:OnlineModule = Find-Module -Name $Private:ModuleName -Repository PSGallery -ErrorAction Stop

            If (-not $Private:LocalModule) {

                <# Output #>
                Write-Output "Installing $Private:ModuleName $($Private:OnlineModule.Version), please wait..."

                Install-Module -Name $Private:ModuleName -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                $script:RequiredModuleAvailability[$Private:ModuleName] = $true
                $Private:ModulesChanged = $true

                <# Logging #>
                fncLogging -strLogFunction "fncUpdateModules" -strLogDescription "$Private:ModuleName module" -strLogValue "Installed: $($Private:OnlineModule.Version)"

            }
            ElseIf ($Private:OnlineModule.Version -gt $Private:LocalModule.Version) {

                <# Output #>
                Write-Output "Updating $Private:ModuleName from $($Private:LocalModule.Version) to $($Private:OnlineModule.Version), please wait..."

                Update-Module -Name $Private:ModuleName -Force -ErrorAction Stop
                $Private:ModulesChanged = $true

                <# Logging #>
                fncLogging -strLogFunction "fncUpdateModules" -strLogDescription "$Private:ModuleName module" -strLogValue "Updated: $($Private:LocalModule.Version) -> $($Private:OnlineModule.Version)"

            }
            Else {

                <# Output #>
                Write-ColoredOutput Yellow "$Private:ModuleName is up to date: $($Private:LocalModule.Version)"

                <# Logging #>
                fncLogging -strLogFunction "fncUpdateModules" -strLogDescription "$Private:ModuleName module" -strLogValue "Current: $($Private:LocalModule.Version)"

            }
        }

        <# Output #>
        Write-Output "`nAll required modules have been checked."

        <# Logging #>
        fncLogging -strLogFunction "fncUpdateModules" -strLogDescription "UPDATE MODULES" -strLogValue "Proceeded"

        $Private:UpdateSucceeded = $true

    }
    Catch {

        $Private:UpdateError = $_.Exception.ToString()

        <# Output #>
        Write-ColoredOutput Red "ATTENTION: Required modules could not be updated.`n$Private:UpdateError`n"

        <# Logging #>
        fncLogging -strLogFunction "fncUpdateModules" -strLogDescription "UPDATE MODULES" -strLogValue "Failed: $Private:UpdateError"

    }

    If ($Private:UpdateSucceeded -and $Private:ModulesChanged) {

        $Private:PowerShellExecutable = Join-Path -Path $PSHOME -ChildPath $(If ($PSVersionTable.PSEdition -eq "Core") { "pwsh.exe" } Else { "powershell.exe" })
        $Private:ModulePath = Join-Path -Path $PSScriptRoot -ChildPath "ComplianceUtility.psm1"
        $Private:RestartCommand = "`$Host.UI.RawUI.BackgroundColor = 'Black'; Clear-Host; Import-Module -Name '$Private:ModulePath' -Force; ComplianceUtility -Menu"

        <# Output #>
        Write-Output "The required modules have been installed or updated."
        Write-ColoredOutput Yellow "The Compliance Utility will now be restarted."
        fncPause

        $Private:RestartParameters = @{
            FilePath = $Private:PowerShellExecutable
            ArgumentList = @("-NoExit", "-Command", $Private:RestartCommand)
        }

        If ($Global:bolRunningPrivileged -eq $true) {
            $Private:RestartParameters.Verb = "RunAs"
        }

        Start-Process @Private:RestartParameters | Out-Null

        <# Close the current session after starting the updated session #>
        Exit

    }
    ElseIf ($Private:UpdateSucceeded) {

        <# Output #>
        Write-ColoredOutput Green "All required modules are already installed and up to date.`n"

    }
}

Function fncCollectAIPServiceConfiguration {

    <# Output #>
    Write-Output "COLLECT AIP SERVICE CONFIGURATION:"

    <# Check if not running as administrator #>
    If ($Global:bolRunningPrivileged -ne $true) {

        <# Output #>
        Write-ColoredOutput Red "ATTENTION: You must run the Compliance Utility in an administrative PowerShell window as a user with local administrative privileges to continue with this option.`nCOLLECT AIP SERVICE CONFIGURATION: Failed.`n"

        <# Command line actions #>
        If ($Global:bolComingFromMenu -eq $false) {

            <# Set back window title #>
            $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

            <# Release variable (updates active) #>
            $Global:bolSkipRequiredUpdates = $false

            <# Exit #>
            Return

        }

        <# ShowMenu actions #>
        If ($Global:bolComingFromMenu -eq $true) {

            <# Call Pause #>
            fncPause

            <# Clear console #>
            Clear-Host

            <# Call ShowMenu #>
            fncShowMenu

            Return

        }

    }

    <# Logging #>
    fncLogging -strLogFunction "fncCollectAIPServiceConfiguration" -strLogDescription "COLLECT AIP SERVICE CONFIGURATION" -strLogValue "Initiated"

    If (-not (fncTestRequiredModule -strModuleName "AIPService")) { Return }

    <# Output #>
    Write-Output "Connecting to AIPService..."

    <# Actions on PowerShell Core (7.x) for compatibility mode #>
    If ($PSVersionTable.PSEdition.ToString() -eq "Core") {

        <# Remove AIPService module, because it's not yet compatible with PowerShell Core (7.x) #>
        Remove-Module -Name AIPService -Verbose:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

        <# Import AIPService module in compatiblity mode #>
        Import-Module -Name AIPService -UseWindowsPowerShell -Verbose:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

        <# Logging #>
        fncLogging -strLogFunction "fncCollectAIPServiceConfiguration" -strLogDescription "AIPService compatiblity mode" -strLogValue $true

    }

    If (-not (Get-Command -Name Connect-AIPService -ErrorAction SilentlyContinue)) {
        Write-ColoredOutput Red "ATTENTION: The required cmdlet 'Connect-AIPService' is not available.`nRun 'ComplianceUtility -UpdateModules' to install or update the required module, then restart the Compliance Utility.`n"
        fncLogging -strLogFunction "fncCollectAIPServiceConfiguration" -strLogDescription "AIPService module" -strLogValue "Connect-AIPService unavailable"
        Return
    }

    Try {

        <# Connect/logon to AIPService #>
        Connect-AIPService -Verbose:$false -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null

        <# Output #>
        Write-Output "AIPService connected."

        <# Logging #>
        fncLogging -strLogFunction "fncCollectAIPServiceConfiguration" -strLogDescription "AIPService connected" -strLogValue $true

    }
    Catch { <# Action if AIPService connection failed #>

        $Private:strAIPServiceConnectionError = $_.Exception.Message

        <# Logging #>
        fncLogging -strLogFunction "fncCollectAIPServiceConfiguration" -strLogDescription "AIPService connected" -strLogValue $false
        fncLogging -strLogFunction "fncCollectAIPServiceConfiguration" -strLogDescription "COLLECT AIP SERVICE CONFIGURATION" -strLogValue "Login failed: $Private:strAIPServiceConnectionError"

        <# Output #>
        Write-ColoredOutput Red "COLLECT AIP SERVICE CONFIGURATION: Login failed: $Private:strAIPServiceConnectionError`n"

        <# Action if function was called from command line #>
        If ($Global:bolComingFromMenu -eq $false) {

            <# Set back window title to default #>
            $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

            <# Release global variable back to default (updates active) #>
            $Global:bolSkipRequiredUpdates = $false

            <# Exit #>
            Return

        }

        $Private:strAIPServiceConnectionError = $null
        Return

    }

    <# Check if COLLECT folder exist and create it, if it not exist #>
    If ($(Test-Path -Path $Global:strUserLogPath"\Collect") -Eq $false) {

        New-Item -ItemType Directory -Force -Path $Global:strUserLogPath"\Collect" | Out-Null <# Define Collect path #>

    }

    <# Check for existing AIPService log file and create it, if it not exist #>
    If ($(Test-Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log") -Eq $false) {

        <# Create AIPService logging file #>
        Out-File -FilePath $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Encoding UTF8 -Append -Force

    }

    <# Output #>
    Write-Output "Collecting AIP service configuration, please wait..."

    <# Check for existing AIPService logging file, and extend it if it exist #>
    If ($(Test-Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log") -Eq $true) { <# Exporting AIP service configuration and output result: #>

        <# Timestamp #>
        $Private:Timestamp = (Get-Date -Verbose:$false -UFormat "%y%m%d-%H%M%S") <# Filling private variable #>
        Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("Date/Timestamp                            : " + $Private:Timestamp) <# Extend log file #>
        Write-ColoredOutput Yellow "Date/Timestamp                            : $Private:Timestamp" <# Output #>
        $Private:Timestamp = $null <# Releasing variable #>

        <# AIPService Module version #>
        $Private:AIPServiceModule = (Get-Module -Verbose:$false -ListAvailable -Name AIPService).Version <# Filling private variable #>
        Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("Module version                            : $Private:AIPServiceModule") <# Extend log file #>
        Write-ColoredOutput Yellow "Module version                            : $Private:AIPServiceModule" <# Output #>
        $Private:AIPServiceModule = $null <# Releasing variable #>

        <# Retrieve the AIP service configuration once and write its values directly to the log #>
        Get-AipServiceConfiguration | ForEach-Object {
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("BPOSId                                    : " + $_.BPOSId)
            Write-ColoredOutput Yellow ("BPOSId                                    : " + $_.BPOSId)
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("RightsManagementServiceId                 : " + $_.RightsManagementServiceId)
            Write-ColoredOutput Yellow ("RightsManagementServiceId                 : " + $_.RightsManagementServiceId)
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("LicensingIntranetDistributionPointUrl     : " + $_.LicensingIntranetDistributionPointUrl)
            Write-ColoredOutput Yellow ("LicensingIntranetDistributionPointUrl     : " + $_.LicensingIntranetDistributionPointUrl)
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("LicensingExtranetDistributionPointUrl     : " + $_.LicensingExtranetDistributionPointUrl)
            Write-ColoredOutput Yellow ("LicensingExtranetDistributionPointUrl     : " + $_.LicensingExtranetDistributionPointUrl)
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("CertificationIntranetDistributionPointUrl : " + $_.CertificationIntranetDistributionPointUrl)
            Write-ColoredOutput Yellow ("CertificationIntranetDistributionPointUrl : " + $_.CertificationIntranetDistributionPointUrl)
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("CertificationExtranetDistributionPointUrl : " + $_.CertificationExtranetDistributionPointUrl)
            Write-ColoredOutput Yellow ("CertificationExtranetDistributionPointUrl : " + $_.CertificationExtranetDistributionPointUrl)
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("AdminConnectionUrl                        : " + $_.AdminConnectionUrl)
            Write-ColoredOutput Yellow ("AdminConnectionUrl                        : " + $_.AdminConnectionUrl)
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("AdminV2ConnectionUrl                      : " + $_.AdminV2ConnectionUrl)
            Write-ColoredOutput Yellow ("AdminV2ConnectionUrl                      : " + $_.AdminV2ConnectionUrl)
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("OnPremiseDomainName                       : " + $_.OnPremiseDomainName)
            Write-ColoredOutput Yellow ("OnPremiseDomainName                       : " + $_.OnPremiseDomainName)
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("Keys                                      : " + $_.Keys)
            Write-ColoredOutput Yellow ("Keys                                      : " + $_.Keys)
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("CurrentLicensorCertificateGuid            : " + $_.CurrentLicensorCertificateGuid)
            Write-ColoredOutput Yellow ("CurrentLicensorCertificateGuid            : " + $_.CurrentLicensorCertificateGuid)
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("Template IDs                              : " + $_.Templates)
            Write-ColoredOutput Yellow ("Template IDs                              : " + $_.Templates)
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("FunctionalState                           : " + $_.FunctionalState)
            Write-ColoredOutput Yellow ("FunctionalState                           : " + $_.FunctionalState)
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("SuperUsersEnabled                         : " + $_.SuperUsersEnabled)
            Write-ColoredOutput Yellow ("SuperUsersEnabled                         : " + $_.SuperUsersEnabled)
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("SuperUsers                                : " + $_.SuperUsers)
            Write-ColoredOutput Yellow ("SuperUsers                                : " + $_.SuperUsers)
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("AdminRoleMembers                          : " + $_.AdminRoleMembers)
            Write-ColoredOutput Yellow ("AdminRoleMembers                          : " + $_.AdminRoleMembers)
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("KeyRolloverCount                          : " + $_.KeyRolloverCount)
            Write-ColoredOutput Yellow ("KeyRolloverCount                          : " + $_.KeyRolloverCount)
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("ProvisioningDate                          : " + $_.ProvisioningDate)
            Write-ColoredOutput Yellow ("ProvisioningDate                          : " + $_.ProvisioningDate)
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("IPCv3ServiceFunctionalState               : " + $_.IPCv3ServiceFunctionalState)
            Write-ColoredOutput Yellow ("IPCv3ServiceFunctionalState               : " + $_.IPCv3ServiceFunctionalState)
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("DevicePlatformState                       : " + $_.DevicePlatformState)
            Write-ColoredOutput Yellow ("DevicePlatformState                       : " + $_.DevicePlatformState)
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("FciEnabledForConnectorAuthorization       : " + $_.FciEnabledForConnectorAuthorization)
            Write-ColoredOutput Yellow ("FciEnabledForConnectorAuthorization       : " + $_.FciEnabledForConnectorAuthorization)
        }

        <# AipServiceDocumentTrackingFeature #>
        Get-AipServiceDocumentTrackingFeature | ForEach-Object {
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("AipServiceDocumentTrackingFeature         : " + $_)
            Write-ColoredOutput Yellow ("AipServiceDocumentTrackingFeature         : " + $_)
        }

        <# AipServiceOnboardingControlPolicy #>
        Get-AipServiceOnboardingControlPolicy | ForEach-Object {
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("AipServiceOnboardingControlPolicy         : {[UseRmsUserLicense, " + $_.UseRmsUserLicense + "], [SecurityGroupObjectId, " + $_.SecurityGroupObjectId + "], [Scope, " + $_.Scope + "]}")
            Write-ColoredOutput Yellow ("AipServiceOnboardingControlPolicy         : {[UseRmsUserLicense, " + $_.UseRmsUserLicense + "], [SecurityGroupObjectId, " + $_.SecurityGroupObjectId + "], [Scope, " + $_.Scope + "]}")
        }

        <# AipServiceDoNotTrackUserGroup #>
        $AipServiceDoNotTrackUserGroups = @(Get-AipServiceDoNotTrackUserGroup)
        If ($AipServiceDoNotTrackUserGroups.Count -eq 0) {
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value "AipServiceDoNotTrackUserGroup             :"
            Write-ColoredOutput Yellow "AipServiceDoNotTrackUserGroup             :"
        }
        Else {
            $AipServiceDoNotTrackUserGroups | ForEach-Object {
                Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("AipServiceDoNotTrackUserGroup             : " + $_)
                Write-ColoredOutput Yellow ("AipServiceDoNotTrackUserGroup             : " + $_)
            }
        }

        <# AipServiceRoleBasedAdministrator #>
        $AipServiceRoleBasedAdministrators = @(Get-AipServiceRoleBasedAdministrator)
        If ($AipServiceRoleBasedAdministrators.Count -eq 0) {
            Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value "AipServiceRoleBasedAdministrator          :"
            Write-ColoredOutput Yellow "AipServiceRoleBasedAdministrator          :"
        }
        Else {
            $AipServiceRoleBasedAdministrators | ForEach-Object {
                Add-Content -Path $Global:strUserLogPath"\Collect\AIPServiceConfiguration.log" -Value ("AipServiceRoleBasedAdministrator          : " + $_)
                Write-ColoredOutput Yellow ("AipServiceRoleBasedAdministrator          : " + $_)
            }
        }

    }

    <# Disconnect from AIPService #>
    Disconnect-AIPService | Out-Null

    <# Output #>
    Write-Output "AIPService disconnected.`n"

    <# Logging #>
    fncLogging -strLogFunction "fncCollectAIPServiceConfiguration" -strLogDescription "AIPService disconnected" -strLogValue $true
    fncLogging -strLogFunction "fncCollectAIPServiceConfiguration" -strLogDescription "Export AIP service configuration" -strLogValue "AIPServiceConfiguration.log"
    fncLogging -strLogFunction "fncCollectAIPServiceConfiguration" -strLogDescription "COLLECT AIP SERVICE CONFIGURATION" -strLogValue "Proceeded"

    <# Output #>
    Write-Output "Log file: $Global:strUserLogPath\Collect\AIPServiceConfiguration.log"
    Write-ColoredOutput Green "COLLECT AIP SERVICE CONFIGURATION: Proceeded.`n"

    <# Action if function was called from command line #>
    If ($Global:bolComingFromMenu -eq $false) {

        <# Set back window title to default #>
        $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

        <# Release global variable back to default (updates active) #>
        $Global:bolSkipRequiredUpdates = $false

        <# Exit #>
        Return

    }

}

Function fncCollectProtectionTemplates {

    <# Output #>
    Write-Output "COLLECT PROTECTION TEMPLATES:"

    <# Check if not running as administrator #>
    If ($Global:bolRunningPrivileged -ne $true) {

        <# Output #>
        Write-ColoredOutput Red "ATTENTION: You must run the Compliance Utility in an administrative PowerShell window as a user with local administrative privileges to continue with this option.`nCOLLECT PROTECTION TEMPLATES: Failed.`n"

        <# Action if function was called from command line #>
        If ($Global:bolComingFromMenu -eq $false) {

            <# Set back window title to default #>
            $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

            <# Release global variable back to default (updates active) #>
            $Global:bolSkipRequiredUpdates = $false

            <# Exit #>
            Return

        }

        <# Action if function was called from the menu #>
        If ($Global:bolComingFromMenu -eq $true) {

            <# Call Pause #>
            fncPause

            <# Clear console #>
            Clear-Host

            <# Call ShowMenu #>
            fncShowMenu

            Return

        }

    }

    <# Logging #>
    fncLogging -strLogFunction "fncCollecProtectionTemplates" -strLogDescription "COLLECT PROTECTION TEMPLATES" -strLogValue "Initiated"

    If (-not (fncTestRequiredModule -strModuleName "AIPService")) { Return }

    <# Output #>
    Write-Output "Connecting to AIPService..."

    <# Actions on PowerShell Core (7.x) for compatibility mode #>
    If ($PSVersionTable.PSEdition.ToString() -eq "Core") {

        <# Remove AIPService module, because it's not yet compatible with PowerShell Core (7.x) #>
        Remove-Module -Name AIPService -Verbose:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

        <# Import AIPService module in compatiblity mode #>
        Import-Module -Name AIPService -UseWindowsPowerShell -Verbose:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

        <# Logging #>
        fncLogging -strLogFunction "fncCollectProtectionTemplates" -strLogDescription "AIPService compatiblity mode" -strLogValue $true

    }

    If (-not (Get-Command -Name Connect-AIPService -ErrorAction SilentlyContinue)) {
        Write-ColoredOutput Red "ATTENTION: The required cmdlet 'Connect-AIPService' is not available.`nRun 'ComplianceUtility -UpdateModules' to install or update the required module, then restart the Compliance Utility.`n"
        fncLogging -strLogFunction "fncCollectProtectionTemplates" -strLogDescription "AIPService module" -strLogValue "Connect-AIPService unavailable"
        Return
    }

    <# Connect/logon to AIPService #>
    Try {

        Connect-AIPService -Verbose:$false -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null

        <# Output #>
        Write-Output "AIPService connected."

        <# Logging #>
        fncLogging -strLogFunction "fncCollectProtectionTemplates" -strLogDescription "AIPService connected" -strLogValue $true

    }
    Catch { <# Action if AIPService connection failed #>

        $Private:strAIPServiceConnectionError = $_.Exception.Message

        <# Logging #>
        fncLogging -strLogFunction "fncCollectProtectionTemplates" -strLogDescription "AIPService connected" -strLogValue $false
        fncLogging -strLogFunction "fncCollectProtectionTemplates" -strLogDescription "COLLECT PROTECTION TEMPLATES" -strLogValue "Login failed: $Private:strAIPServiceConnectionError"

        <# Output #>
        Write-ColoredOutput Red "COLLECT PROTECTION TEMPLATES: Login failed: $Private:strAIPServiceConnectionError`n"

        <# Action if function was called from command line #>
        If ($Global:bolComingFromMenu -eq $false) {

            <# Set back window title to default #>
            $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

            <# Release global variable back to default (updates active) #>
            $Global:bolSkipRequiredUpdates = $false

            <# Exit #>
            Return

        }

        $Private:strAIPServiceConnectionError = $null
        Return

    }

    <# Check if COLLECT folder exist and create it, if not #>
    If ($(Test-Path -Path $Global:strUserLogPath"\Collect\ProtectionTemplates") -Eq $false) {

        New-Item -ItemType Directory -Force -Path $Global:strUserLogPath"\Collect\ProtectionTemplates" | Out-Null <# Define Collect path #>

    }

    <# Output #>
    Write-Output "Collecting protection templates, please wait..."

    <# Check for existing folder #>
    If ($(Test-Path $Global:strUserLogPath"\Collect\ProtectionTemplates") -Eq $true) {

        <# Exporting protection templates #>
        Get-AipServiceConfiguration -WarningAction SilentlyContinue | Select-Object -ExpandProperty Templates | Export-Clixml -Path $Global:strUserLogPath"\Collect\ProtectionTemplates\ProtectionTemplates.xml" | Out-Null
        Get-AIPServicetemplate -WarningAction SilentlyContinue | Format-List * | Export-Clixml -Path $Global:strUserLogPath"\Collect\ProtectionTemplates\ProtectionTemplateDetails.xml" | Out-Null

        <# Logging #>
        fncLogging -strLogFunction "fncCollectProtectionTemplates" -strLogDescription "Export protection templates" -strLogValue "ProtectionTemplates.xml"
        fncLogging -strLogFunction "fncCollectProtectionTemplates" -strLogDescription "Export protection templates" -strLogValue "ProtectionTemplateDetails.xml"

    }

    <# Check if COLLECT\ProtectionTemplates folder exist and create it, if not #>
    If ($(Test-Path -Path $Global:strUserLogPath"\Collect\ProtectionTemplates\ProtectionTemplatesBackup") -Eq $false) {

        New-Item -ItemType Directory -Force -Path $Global:strUserLogPath"\Collect\ProtectionTemplates\ProtectionTemplatesBackup" | Out-Null <# Define Service Templates path #>

    }

    <# Detect Protection Template IDs for backup #>
    ForEach ($Private:ProtectionTemplate in (Get-AIPServicetemplate).TemplateID) {

        <# Backup Service Template to XML #>
        Export-AipServiceTemplate -Path $Global:strUserLogPath"\Collect\ProtectionTemplates\ProtectionTemplatesBackup\$Private:ProtectionTemplate.xml" -TemplateId $Private:ProtectionTemplate -Force

        <# Logging #>
        fncLogging -strLogFunction "fncCollectProtectionTemplates" -strLogDescription "Protection template exported" -strLogValue "$Private:ProtectionTemplate.xml"

    }

    <# Disconnect from AIPService #>
    Disconnect-AIPService | Out-Null

    <# Output #>
    Write-Output "AIPService disconnected.`n"

    <# Logging #>
    fncLogging -strLogFunction "fncCollectProtectionTemplates" -strLogDescription "AIPService disconnected" -strLogValue $true
    fncLogging -strLogFunction "fncCollectProtectionTemplates" -strLogDescription "COLLECT PROTECTION TEMPLATES" -strLogValue "Proceeded"

    <# Output #>
    Write-Output "Protection templates: $Global:strUserLogPath\Collect\ProtectionTemplates\ProtectionTemplatesBackup"
    Write-Output "Logs folder: $Global:strUserLogPath\Collect\ProtectionTemplates"
    Write-ColoredOutput Green "COLLECT PROTECTION TEMPLATES: Proceeded.`n"

    <# Action if function was called from command line #>
    If ($Global:bolComingFromMenu -eq $false) {

        <# Set back window title to default #>
        $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

        <# Release global variable back to default (updates active) #>
        $Global:bolSkipRequiredUpdates = $false

        <# Exit #>
        Return

    }

}

Function fncCollectLabelsAndPolicies {

     <# Output #>
    Write-Output "COLLECT LABELS AND POLICIES:"

    <# Check if not running as administrator #>
    If ($Global:bolRunningPrivileged -ne $true) {

        <# Output #>
        Write-ColoredOutput Red "ATTENTION: You must run the Compliance Utility in an administrative PowerShell window as a user with local administrative privileges to continue with this option.`nCOLLECT LABELS AND POLICIES: Failed.`n"

        <# Action if function was called from command line #>
        If ($Global:bolComingFromMenu -eq $false) {

            <# Set back window title to default #>
            $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

            <# Release global variable back to default (updates active) #>
            $Global:bolSkipRequiredUpdates = $false

            <# Exit #>
            Return

        }

        <# Action if function was called from the menu #>
        If ($Global:bolComingFromMenu -eq $true) {

            <# Call Pause #>
            fncPause

            <# Clear console #>
            Clear-Host

            <# Call ShowMenu #>
            fncShowMenu

            Return

        }

    }

    <# Logging #>
    fncLogging -strLogFunction "fncCollectLabelsAndPolicies" -strLogDescription "COLLECT LABELS AND POLICIES" -strLogValue "Initiated"

    If (-not (fncTestRequiredModule -strModuleName "ExchangeOnlineManagement")) { Return }

    If (-not (Get-Command -Name Connect-IPPSSession -ErrorAction SilentlyContinue)) {
        Write-ColoredOutput Red "ATTENTION: The required cmdlet 'Connect-IPPSSession' is not available.`nRun 'ComplianceUtility -UpdateModules' to install or update the required module, then restart the Compliance Utility.`n"
        fncLogging -strLogFunction "fncCollectLabelsAndPolicies" -strLogDescription "ExchangeOnlineManagement module" -strLogValue "Connect-IPPSSession unavailable"
        Return
    }

    <# Logging #>
    fncLogging -strLogFunction "fncCollectLabelsAndPolicies" -strLogDescription "ExchangeOnlineManagement module version" -strLogValue (Get-Module -Verbose:$false -ListAvailable -Name ExchangeOnlineManagement).Version

    <# Output #>
    Write-Output "Connecting to Microsoft Purview compliance portal..."

    <# Remember default progress bar status: "Continue" #>
    $Private:strOriginalPreference = $Global:ProgressPreference
    $Global:ProgressPreference = "SilentlyContinue" <# Hiding progress bar #>

    <# Try to connect/logon #>
    Try {

        <# Connect to Microsoft Purview compliance portal #>
        $Private:RequiredPurviewCommands = @(
            "Get-Label"
            "Get-LabelPolicy"
            "Get-LabelPolicyRule"
            "Get-AutoSensitivityLabelPolicy"
            "Get-AutoSensitivityLabelRule"
        )
        Connect-IPPSSession -CommandName $Private:RequiredPurviewCommands -ShowBanner:$false -Verbose:$false -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null

    }
    Catch { <# Catch for any error #>

        $Private:PurviewConnectionError = $_.Exception.Message

        <# Logging #>
        fncLogging -strLogFunction "fncCollectLabelsAndPolicies" -strLogDescription "Microsoft Purview compliance portal connected" -strLogValue $false
        fncLogging -strLogFunction "fncCollectLabelsAndPolicies" -strLogDescription "Microsoft Purview compliance portal" -strLogValue "Login failed: $Private:PurviewConnectionError"

        <# Output #>
        Write-ColoredOutput Red "COLLECT LABELS AND POLICIES: Login failed: $Private:PurviewConnectionError`n"

        $Global:ProgressPreference = $Private:strOriginalPreference

        <# Action if function was called from the menu #>
        If ($Global:bolComingFromMenu -eq $true) {

            <# Call Pause #>
            fncPause

            <# Clear console #>
            Clear-Host

            <# Call ShowMenu #>
            fncShowMenu

            Return

        }

        <# Action if function was called from command line #>
        If ($Global:bolComingFromMenu -eq $false) {

            <# Release global variable back to default (updates active) #>
            $Global:bolSkipRequiredUpdates = $false

            <# Set back window title to default #>
            $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

            <# Interrupt, because of missing internet connection #>
            Return

        }

        Return

    }

    <# Output #>
    Write-Output "Microsoft Purview compliance portal connected."

    <# Logging #>
    fncLogging -strLogFunction "fncCollectLabelsAndPolicies" -strLogDescription "Microsoft Purview compliance portal connected" -strLogValue $true

    <# Output #>
    Write-Output "Collecting labels and policies. Please wait, this may take a while..."

    <# Detect/create COLLECT\LabelsAndPolicies folder #>
    If ($(Test-Path -Path $Global:strUserLogPath"\Collect\LabelsAndPolicies") -Eq $false) {

        New-Item -ItemType Directory -Force -Path $Global:strUserLogPath"\Collect\LabelsAndPolicies" | Out-Null <# Create folder #>

    }

    <# Check for existing LabelsAndPolicies folder #>
    If ($(Test-Path $Global:strUserLogPath"\Collect\LabelsAndPolicies") -Eq $true) {

        <# Collect labels #>
        Get-Label -WarningAction SilentlyContinue | Export-Clixml -Path $Global:strUserLogPath"\Collect\LabelsAndPolicies\Labels.xml" | Out-Null
        fncLogging -strLogFunction "fncCollectLabelsAndPolicies" -strLogDescription "Export labels and policies" -strLogValue "Labels.xml" <# Logging #>

        <# Collect labels with details #>
        Get-Label -IncludeDetailedLabelActions -WarningAction SilentlyContinue | Export-Clixml -Path $Global:strUserLogPath"\Collect\LabelsAndPolicies\LabelsDetailedActions.xml" | Out-Null
        fncLogging -strLogFunction "fncCollectLabelsAndPolicies" -strLogDescription "Export labels and policies" -strLogValue "LabelsDetailedActions.xml" <# Logging #>

        <# Collect policies #>
        Get-LabelPolicy -WarningAction SilentlyContinue | ForEach-Object {Get-LabelPolicy -Identity $_.Identity} | Export-Clixml -Path $Global:strUserLogPath"\Collect\LabelsAndPolicies\LabelPolicies.xml" | Out-Null
        fncLogging -strLogFunction "fncCollectLabelsAndPolicies" -strLogDescription "Export labels and policies" -strLogValue "LabelPolicies.xml" <# Logging #>

        <# Collect rules #>
        Get-LabelPolicy -WarningAction SilentlyContinue | ForEach-Object {Get-LabelPolicyRule -Policy $_.Identity} | Export-Clixml -Path $Global:strUserLogPath"\Collect\LabelsAndPolicies\LabelRules.xml" | Out-Null
        fncLogging -strLogFunction "fncCollectLabelsAndPolicies" -strLogDescription "Export labels and policies" -strLogValue "LabelRules.xml" <# Logging #>

        <# Auto-labeling cmdlets are only exposed when the connected account has the required Purview permissions #>
        $Private:GetAutoLabelPolicyCommand = Get-Command -Name Get-AutoSensitivityLabelPolicy -ErrorAction SilentlyContinue
        $Private:GetAutoLabelRuleCommand = Get-Command -Name Get-AutoSensitivityLabelRule -ErrorAction SilentlyContinue

        If ($Private:GetAutoLabelPolicyCommand -and $Private:GetAutoLabelRuleCommand) {

            <# Collect auto-labeling policies #>
            $Private:AutoLabelPolicies = @(Get-AutoSensitivityLabelPolicy -WarningAction SilentlyContinue)
            $Private:AutoLabelPolicies | ForEach-Object {Get-AutoSensitivityLabelPolicy -Identity $_.Identity} | Export-Clixml -Path $Global:strUserLogPath"\Collect\LabelsAndPolicies\AutoLabelPolicies.xml" | Out-Null
            fncLogging -strLogFunction "fncCollectLabelsAndPolicies" -strLogDescription "Export labels and policies" -strLogValue "AutoLabelPolicies.xml" <# Logging #>

            <# Collect rules for every auto-labeling policy #>
            $Private:AutoLabelPolicies | ForEach-Object {Get-AutoSensitivityLabelRule -Policy $_.Identity} | Export-Clixml -Path $Global:strUserLogPath"\Collect\LabelsAndPolicies\AutoLabelRules.xml" | Out-Null
            fncLogging -strLogFunction "fncCollectLabelsAndPolicies" -strLogDescription "Export labels and policies" -strLogValue "AutoLabelRules.xml" <# Logging #>

        }
        Else {

            $Private:MissingAutoLabelCommands = @(
                If (-not $Private:GetAutoLabelPolicyCommand) { "Get-AutoSensitivityLabelPolicy" }
                If (-not $Private:GetAutoLabelRuleCommand) { "Get-AutoSensitivityLabelRule" }
            ) -join ", "

            <# Logging #>
            fncLogging -strLogFunction "fncCollectLabelsAndPolicies" -strLogDescription "Auto-labeling policies skipped; unavailable Purview cmdlets" -strLogValue $Private:MissingAutoLabelCommands

        }

    }

    <# Logging #>
    fncLogging -strLogFunction "fncCollectLabelsAndPolicies" -strLogDescription "Export labels and policies folder" -strLogValue "\Collect\LabelsAndPolicies"

    <# Check if Office CLP folder exist #>
    If ($(Test-Path -Path $env:LOCALAPPDATA\Microsoft\Office\CLP) -Eq $true) {

        <# Perform action only, if the CLP folder contain files (Note: Afer a RESET this folder is empty). #>
        If (((Get-ChildItem -LiteralPath $env:LOCALAPPDATA\Microsoft\Office\CLP -File -Force | Select-Object -First 1 | Measure-Object).Count -ne 0)) {

            <# Copy CLP Office policy folder content #>
            fncCopyItem $env:LOCALAPPDATA\Microsoft\Office\CLP $Global:strUserLogPath"\Collect\LabelsAndPolicies" "CLP\*"

            <# Private variable for unique logging/output with CLP #>
            $Private:CLPPolicy = $true

            <# Logging #>
            fncLogging -strLogFunction "fncCollectLabelsAndPolicies" -strLogDescription "Export Office CLP folder" -strLogValue "\Collect\LabelsAndPolicies\CLP"

        }

    }

    <# Set back progress bar #>
    $Global:ProgressPreference = $Private:strOriginalPreference

    <# logging based on existence of CLP folder #>
    If ($Private:CLPPolicy -Eq $true) {

        <# Output #>
        Write-Output "`nLog folder: $Global:strUserLogPath\Collect\LabelsAndPolicies"
        Write-Output "Office CLP policy folder: $Global:strUserLogPath\Collect\LabelsAndPolicies\CLP"

    }
    Else {

        <# Output #>
        Write-Output "`nLog folder: $Global:strUserLogPath\Collect\LabelsAndPolicies"

    }

    <# Release private variable #>
    $Private:CLPPolicy = $null

    <# Disconnect Exchange Online #>
    Disconnect-ExchangeOnline -Confirm:$false

    <# Logging #>
    fncLogging -strLogFunction "fncCollectLabelsAndPolicies" -strLogDescription "Microsoft Purview compliance portal disconnected" -strLogValue $true
    fncLogging -strLogFunction "fncCollectLabelsAndPolicies" -strLogDescription "COLLECT LABELS AND POLICIES" -strLogValue "Proceeded"

    <# Output #>
    Write-Output "Microsoft Purview compliance portal disconnected."
    Write-ColoredOutput Green "COLLECT LABELS AND POLICIES: Proceeded.`n"

    <# Action if function was called from command line #>
    If ($Global:bolComingFromMenu -eq $false) {

        <# Release global variable back to default (updates active) #>
        $Global:bolSkipRequiredUpdates = $false

        <# Set back window title to default #>
        $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

        <# Interrupt, because of missing internet connection #>
        Return

    }

}

Function fncCollectEndpointURLs {

    <# Logging #>
    fncLogging -strLogFunction "fncCollectEndpointURLs" -strLogDescription "COLLECT ENDPOINT URLs" -strLogValue "Initiated"

    <# Output #>
    Write-Output "COLLECT ENDPOINT URLs:"

    <# Define and fill variables with static URLs #>
    $Private:MyUnifiedLabelingDistributionPointUrl = "dataservice.protection.outlook.com"
    $Private:MyTelemetryDistributionPointUrl = "self.events.data.microsoft.com"

    <# Define and fill variable with date/time for unique log folder #>
    $Private:MyTimestamp = (Get-Date -Verbose:$false -UFormat "%y%m%d-%H%M%S")
    $Private:strCertLogPath = "$Global:strUserLogPath\Collect\EndpointURLs"

    <# Function to check if "EndpointURLs"-folder and log file exist #>
    Function fncCreateLogFileAndFolder ($Private:strCertLogPath) {

        <# Check if "EndpointURLs"-folder exist and create it, if not #>
        If ($(Test-Path -Path $Private:strCertLogPath) -Eq $false) {

            New-Item -ItemType Directory -Force -Path $Private:strCertLogPath | Out-Null <# Define EndpointURLs path #>

        }

        <# Check for existing EndpointURLs.log file and create it, if it not exist #>
        If ($(Test-Path $Global:strUserLogPath"\Collect\EndpointURLs.log") -Eq $false) {

            Out-File -FilePath $Global:strUserLogPath"\Collect\EndpointURLs.log" -Encoding UTF8 -Append -Force

        }

    }

    <# Return the host portion of an endpoint URL without relying on a fixed URL length #>
    Function fncGetEndpointHost ($strEndpointUrl, $strFallbackHost) {

        If (-not [string]::IsNullOrWhiteSpace([string]$strEndpointUrl)) {
            Try {
                $Private:EndpointUri = [System.Uri]$strEndpointUrl
                If (-not [string]::IsNullOrWhiteSpace($Private:EndpointUri.Host)) {
                    Return $Private:EndpointUri.Host
                }
            }
            Catch {
                <# The value is not a valid absolute URI; use the fallback below #>
                Write-Verbose "Endpoint URL '$strEndpointUrl' is not a valid URI; using the fallback host."
            }
        }

        If (-not [string]::IsNullOrWhiteSpace([string]$strFallbackHost)) {
            Return $strFallbackHost
        }

        Return $null
    }

    <# Check for COLLECT Endpoints URLs [MSIPC] if bootstrap was done/running with user privileges/reading URLs from registry #>
    If ($(Test-Path -Path "HKCU:\Software\Classes\Local Settings\Software\Microsoft\MSIPC") -Eq $true) {

        <# Output #>
        Write-Output "Verifying endpoint URLs...`n"

        <# Check if "EndpointURLs"-folder and log file exist and create it, if not #>
        fncCreateLogFileAndFolder $Private:strCertLogPath

        <# Read URLs from registry #>
        Get-ChildItem -Path "HKCU:\Software\Classes\Local Settings\Software\Microsoft\MSIPC" | ForEach-Object {

            <# Read Tenant Id #>
            $Private:strMainKey = $_.PSChildName

            <# Actions if it's about ".aadrm.com", but not about "discover.aadrm.com" #>
            If ($Private:strMainKey -like "*.aadrm.com" -and $Private:strMainKey -notmatch "discover.aadrm.com") {

                <# Private variabel definition for Tenant Id string #>
                $Private:strTenantId = ($Private:strMainKey -split "\.")[0]

                <# Output #>
                Write-ColoredOutput Magenta "-------------------------------------------------`nTenant Id:  $Private:strTenantId`n-------------------------------------------------`n"

                <# Create Tenant Id as first log entry #>
                Add-Content -Path $Global:strUserLogPath"\Collect\EndpointURLs.log" -Value "-----------------------------------------------`nTenant Id: $Private:strTenantId`n-----------------------------------------------"

                <# Define and filling variables with URLs #>
                $Private:IdentityConfiguration = Get-ItemProperty "HKCU:\Software\Classes\Local Settings\Software\Microsoft\MSIPC\$Private:strMainKey\Identities" -ErrorAction SilentlyContinue
                $Private:ServerConfiguration = Get-ItemProperty "HKCU:\Software\Classes\Local Settings\Software\Microsoft\MSIPC\$Private:strMainKey\ServerInfo" -ErrorAction SilentlyContinue
                $Private:ServiceDiscoveryHost = fncGetEndpointHost $Private:ServerConfiguration.ServiceDiscoveryUri $Private:strMainKey

                <# Current clients may only expose ServiceDiscoveryUri; use its host as fallback #>
                $Private:MyLicensingIntranetDistributionPointUrl = fncGetEndpointHost $Private:IdentityConfiguration.InternalUrl $Private:ServiceDiscoveryHost
                $Private:MyLicensingExtranetDistributionPointUrl = fncGetEndpointHost $Private:IdentityConfiguration.ExternalUrl $Private:ServiceDiscoveryHost

                <# Define and fill variables: Extending colledted registry key #>
                $Private:MyCertificationDistributionPointUrl = $Private:strMainKey

                <# Create Timestamp #>
                Add-Content -Path $Global:strUserLogPath"\Collect\EndpointURLs.log" -Value ("Date/Timestamp: " + (Get-Date -Verbose:$false -UFormat "$Private:MyTimestamp"))

                <# Add read mode #>
                Add-Content -Path $Global:strUserLogPath"\Collect\EndpointURLs.log" -Value ("Read from registry [MSIPC]:`n")

                <# Call function to verify endpoint and certificate issuer #>
                fncVerifyIssuer -strCertURL $Private:MyLicensingIntranetDistributionPointUrl -strEndpointName "LicensingIntranetDistributionPointUrl" -strLogPath $Private:strCertLogPath
                fncVerifyIssuer -strCertURL $Private:MyLicensingExtranetDistributionPointUrl -strEndpointName "LicensingExtranetDistributionPointUrl" -strLogPath $Private:strCertLogPath
                fncVerifyIssuer -strCertURL $Private:MyCertificationDistributionPointUrl -strEndpointName "CertificationDistributionPointUrl" -strLogPath $Private:strCertLogPath
                fncVerifyIssuer -strCertURL $Private:MyUnifiedLabelingDistributionPointUrl -strEndpointName "UnifiedLabelingDistributionPointUrl" -strLogPath $Private:strCertLogPath
                fncVerifyIssuer -strCertURL $Private:MyTelemetryDistributionPointUrl -strEndpointName "TelemetryDistributionPointUrl" -strLogPath $Private:strCertLogPath

                <# Logging #>
                fncLogging -strLogFunction "fncCollectEndpointURLs" -strLogDescription "Export endpoint URLs" -strLogValue "EndpointURLs.log"
                fncLogging -strLogFunction "fncCollectEndpointURLs" -strLogDescription "COLLECT ENDPOINT URLs" -strLogValue "Proceeded"

            }

        }

        <# Check for COLLECT Endpoints URLs [MSIP] if bootstrap was done/running in "non-admin mode"/reading URLs from registry #>
        If ($(Test-Path -Path "HKCU:\Software\Classes\Local Settings\Software\Microsoft\MSIPC\MSIP") -Eq $true) {

            <# Check if "EndpointURLs"-folder and log file exist and create it, if not #>
            fncCreateLogFileAndFolder $Private:strCertLogPath

            <# Read URLs from registry #>
            Get-ChildItem -Path "HKCU:\Software\Classes\Local Settings\Software\Microsoft\MSIPC\MSIP" | ForEach-Object {

                <# Read Tenant Id #>
                $Private:strMainKey = $_.PSChildName

                <# Actions if it's about ".aadrm.com", but not about "discover.aadrm.com" #>
                If ($Private:strMainKey -like "*.aadrm.com" -and $Private:strMainKey -notmatch "discover.aadrm.com") {

                    <# Private variabel definition for Tenant Id string #>
                    $Private:strTenantId = ($Private:strMainKey -split "\.")[0]

                    <# Output #>
                    Write-ColoredOutput Magenta "------------------------------------------------`nTenant Id:  $Private:strTenantId`n------------------------------------------------`n"

                    <# Create Tenant Id as first log entry #>
                    Add-Content -Path $Global:strUserLogPath"\Collect\EndpointURLs.log" -Value "------------------------------------------------`nTenant Id: $Private:strTenantId`n------------------------------------------------"

                    <# Define and fill variables with URLs #>
                    $Private:IdentityConfiguration = Get-ItemProperty "HKCU:\Software\Classes\Local Settings\Software\Microsoft\MSIPC\MSIP\$Private:strMainKey\Identities" -ErrorAction SilentlyContinue
                    $Private:ServerConfiguration = Get-ItemProperty "HKCU:\Software\Classes\Local Settings\Software\Microsoft\MSIPC\MSIP\$Private:strMainKey\ServerInfo" -ErrorAction SilentlyContinue
                    $Private:ServiceDiscoveryHost = fncGetEndpointHost $Private:ServerConfiguration.ServiceDiscoveryUri $Private:strMainKey

                    <# Current clients may only expose ServiceDiscoveryUri; use its host as fallback #>
                    $Private:MyLicensingIntranetDistributionPointUrl = fncGetEndpointHost $Private:IdentityConfiguration.InternalUrl $Private:ServiceDiscoveryHost
                    $Private:MyLicensingExtranetDistributionPointUrl = fncGetEndpointHost $Private:IdentityConfiguration.ExternalUrl $Private:ServiceDiscoveryHost

                    <# Define and fill variables: Extending colledted registry key #>
                    $Private:MyCertificationDistributionPointUrl = $Private:strMainKey

                    <# Create Timestamp #>
                    Add-Content -Path $Global:strUserLogPath"\Collect\EndpointURLs.log" -Value ("Date/Timestamp: " + (Get-Date -Verbose:$false -UFormat "$Private:MyTimestamp"))

                    <# Add read mode #>
                    Add-Content -Path $Global:strUserLogPath"\Collect\EndpointURLs.log" -Value ("Read from registry [MSIP]:`n")

                    <# Call function to verify endpoint and certificate issuer #>
                    fncVerifyIssuer -strCertURL $Private:MyLicensingIntranetDistributionPointUrl -strEndpointName "LicensingIntranetDistributionPointUrl" -strLogPath $Private:strCertLogPath
                    fncVerifyIssuer -strCertURL $Private:MyLicensingExtranetDistributionPointUrl -strEndpointName "LicensingExtranetDistributionPointUrl" -strLogPath $Private:strCertLogPath
                    fncVerifyIssuer -strCertURL $Private:MyCertificationDistributionPointUrl -strEndpointName "CertificationDistributionPointUrl" -strLogPath $Private:strCertLogPath
                    fncVerifyIssuer -strCertURL $Private:MyUnifiedLabelingDistributionPointUrl -strEndpointName "UnifiedLabelingDistributionPointUrl" -strLogPath $Private:strCertLogPath
                    fncVerifyIssuer -strCertURL $Private:MyTelemetryDistributionPointUrl -strEndpointName "TelemetryDistributionPointUrl" -strLogPath $Private:strCertLogPath

                }

            }

        }

    }
    Else { <# Actions for COLLECT Endpoints URLs, if bootstrap has failed/reading URLs from portal/running administrative #>

        <# Actions if running administrative #>
        If ($Global:bolRunningPrivileged -eq $true) {

            If (-not (fncTestRequiredModule -strModuleName "AIPService")) { Return }

            <# Output #>
            Write-Output "Connecting to AIPService..."

            <# Actions on PowerShell Core (7.x) for compatibility mode #>
            If ($PSVersionTable.PSEdition.ToString() -eq "Core") {

                <# Remove AIPService module, because it's not yet compatible with PowerShell Core (7.x) #>
                Remove-Module -Name AIPService -Verbose:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

                <# Import AIPService module in compatiblity mode #>
                Import-Module -Name AIPService -UseWindowsPowerShell -Verbose:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

                <# Logging #>
                fncLogging -strLogFunction "fncCollectEndpointURLs" -strLogDescription "AIPService compatiblity mode" -strLogValue $true

            }

            If (-not (Get-Command -Name Connect-AIPService -ErrorAction SilentlyContinue)) {
                Write-ColoredOutput Red "ATTENTION: The required cmdlet 'Connect-AIPService' is not available.`nRun 'ComplianceUtility -UpdateModules' to install or update the required module, then restart the Compliance Utility.`n"
                fncLogging -strLogFunction "fncCollectEndpointURLs" -strLogDescription "AIPService module" -strLogValue "Connect-AIPService unavailable"
                Return
            }

            <# Connect/logon to AIPService #>
            Try {

                Connect-AIPService -Verbose:$false -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null

                <# Output #>
                Write-Output "AIPService connected"

                <# Logging #>
                fncLogging -strLogFunction "fncCollectEndpointURLs" -strLogDescription "AIPService connected" -strLogValue $true

            }
            Catch { <# Action if AIPService connection failed #>

                $Private:strAIPServiceConnectionError = $_.Exception.Message

                <# Logging #>
                fncLogging -strLogFunction "fncCollectEndpointURLs" -strLogDescription "AIPService connected" -strLogValue $false
                fncLogging -strLogFunction "fncCollectEndpointURLs" -strLogDescription "COLLECT ENDPOINT URLs" -strLogValue "Login failed: $Private:strAIPServiceConnectionError"

                <# Output #>
                Write-ColoredOutput Red "COLLECT ENDPOINT URLs: Login failed: $Private:strAIPServiceConnectionError`n"

                <# Action if function was called from command line #>
                If ($Global:bolComingFromMenu -eq $false) {

                    <# Set back window title to default #>
                    $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

                    <# Release global variable back to default (updates active) #>
                    $Global:bolSkipRequiredUpdates = $false

                    <# Exit #>
                    Return

                }

                $Private:strAIPServiceConnectionError = $null
                Return

            }

            <# Output #>
            Write-Output "Verifying endpoint URLs...`n"

            <# Retrieve the AIP service configuration once and process it directly #>
            Get-AipServiceConfiguration | ForEach-Object {
                Write-ColoredOutput Magenta ("------------------------------------------------`nTenant Id:  " + $_.RightsManagementServiceId + "`n------------------------------------------------`n")

                fncCreateLogFileAndFolder $Private:strCertLogPath
                Add-Content -Path $Global:strUserLogPath"\Collect\EndpointURLs.log" -Value ("------------------------------------------------`nTenant Id: " + $_.RightsManagementServiceId + "`n------------------------------------------------")
                Add-Content -Path $Global:strUserLogPath"\Collect\EndpointURLs.log" -Value ("Date/Timestamp: " + (Get-Date -Verbose:$false -UFormat "$Private:MyTimestamp"))
                Add-Content -Path $Global:strUserLogPath"\Collect\EndpointURLs.log" -Value ("Read from portal:`n")

                fncVerifyIssuer -strCertURL (fncGetEndpointHost $_.LicensingIntranetDistributionPointUrl $null) -strEndpointName "LicensingIntranetDistributionPointUrl" -strLogPath $Private:strCertLogPath
                fncVerifyIssuer -strCertURL (fncGetEndpointHost $_.LicensingExtranetDistributionPointUrl $null) -strEndpointName "LicensingExtranetDistributionPointUrl" -strLogPath $Private:strCertLogPath
                fncVerifyIssuer -strCertURL (fncGetEndpointHost $_.CertificationExtranetDistributionPointUrl $null) -strEndpointName "CertificationDistributionPointUrl" -strLogPath $Private:strCertLogPath
                fncVerifyIssuer -strCertURL $Private:MyUnifiedLabelingDistributionPointUrl -strEndpointName "UnifiedLabelingDistributionPointUrl" -strLogPath $Private:strCertLogPath
                fncVerifyIssuer -strCertURL $Private:MyTelemetryDistributionPointUrl -strEndpointName "TelemetryDistributionPointUrl" -strLogPath $Private:strCertLogPath
            }

            <# Disconnect from AIPService #>
            Disconnect-AIPService | Out-Null

            <# Output #>
            Write-Output "AIPService disconnected`n"

            <# Logging #>
            fncLogging -strLogFunction "fncCollectEndpointURLs" -strLogDescription "AIPService disconnected" -strLogValue $true
            fncLogging -strLogFunction "fncCollectEndpointURLs" -strLogDescription "Export endpoint URLs" -strLogValue "EndpointURLs.log"
            fncLogging -strLogFunction "fncCollectEndpointURLs" -strLogDescription "COLLECT ENDPOINT URLs" -strLogValue "Proceeded"

            <# Release private variable #>
            $Private:strTenantId = $null

        }
        Else { <# Actions if running with user privileges #>

            <# Logging on PowerShell Desktop (5.1) #>
            If ($PSVersionTable.PSEdition.ToString() -eq "Desktop" -and [Version]::new($PSVersionTable.PSVersion.Major, $PSVersionTable.PSVersion.Minor) -eq [Version]::new("5.1")) {

                <# Output #>
                Write-ColoredOutput Red "ATTENTION: You must run the Compliance Utility in an administrative PowerShell window as a user with local administrative privileges to continue with this option. Alternatively, you can start (bootstrap) any Microsoft 365 app and try again."

            }
            Else { <# Logging on PowerShell 7.x #>

                <# Output #>
                Write-ColoredOutput Red "ATTENTION: You must run the Compliance Utility in an administrative PowerShell window as a user with local administrative privileges to continue with this option."

            }

             <# Output #>
             Write-ColoredOutput Red "COLLECT ENDPOINT URLs: Failed.`n"

            <# Action if function was called from command line #>
            If ($Global:bolComingFromMenu -eq $false) {

                <# Set back window title to default #>
                $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

                <# Release global variable back to default (updates active) #>
                $Global:bolSkipRequiredUpdates = $false

                <# Exit #>
                Return

            }

            <# Action if function was called from the menu #>
            If ($Global:bolComingFromMenu -eq $true) {

                <# Call Pause #>
                fncPause

                <# Clear console #>
                Clear-Host

                <# Call ShowMenu #>
                fncShowMenu

            }

        }

    }

    <# Output #>
    Write-Output "Log file: $Global:strUserLogPath\Collect\EndpointURLs.log"
    Write-ColoredOutput Green "COLLECT ENDPOINT URLs: Proceeded.`n"

    <# Release private variables #>
    $Private:MyLicensingIntranetDistributionPointUrl = $null
    $Private:MyLicensingExtranetDistributionPointUrl = $null
    $Private:MyCertificationDistributionPointUrl = $null
    $Private:MyTimestamp = $null
    $Private:strTenantId = $null
    $Private:strMainKey = $null
    $Private:strCertLogPath = $null

}

Function fncCollectExchangeIRMConfiguration {

    <# Output #>
    Write-Output "COLLECT EXCHANGE IRM CONFIGURATION:"

    <# Check if not running as administrator #>
    If ($Global:bolRunningPrivileged -ne $true) {

        <# Output #>
        Write-ColoredOutput Red "ATTENTION: You must run the Compliance Utility in an administrative PowerShell window as a user with local administrative privileges to continue with this option.`nCOLLECT EXCHANGE IRM CONFIGURATION: Failed.`n"

        <# Action if function was called from command line #>
        If ($Global:bolComingFromMenu -eq $false) {

            <# Set back window title to default #>
            $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

            <# Release global variable back to default (updates active) #>
            $Global:bolSkipRequiredUpdates = $false

            <# Exit #>
            Return

        }

        <# Action if function was called from the menu #>
        If ($Global:bolComingFromMenu -eq $true) {

            <# Call Pause #>
            fncPause

            <# Clear console #>
            Clear-Host

            <# Call ShowMenu #>
            fncShowMenu

            Return

        }

    }

    <# Logging #>
    fncLogging -strLogFunction "fncCollectExchangeIRMConfiguration" -strLogDescription "COLLECT EXCHANGE IRM CONFIGURATION" -strLogValue "Initiated"

    If (-not (fncTestRequiredModule -strModuleName "ExchangeOnlineManagement")) { Return }

    If (-not (Get-Command -Name Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {

        Write-ColoredOutput Red "ATTENTION: The required cmdlet 'Connect-ExchangeOnline' is not available.`nRun 'ComplianceUtility -UpdateModules' to install or update the required module, then restart the Compliance Utility.`n"
        fncLogging -strLogFunction "fncCollectExchangeIRMConfiguration" -strLogDescription "ExchangeOnlineManagement module" -strLogValue "Connect-ExchangeOnline unavailable"
        Return

    }

    <# Output #>
    Write-Output "Connecting to Exchange Online..."

    <# Connect/logon to Exchange Online and treat only a terminating error as a failed login #>
    Try {

        Connect-ExchangeOnline -CommandName Get-IRMConfiguration -ShowBanner:$false -Verbose:$false -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null

        If (-not (Get-Command -Name Get-IRMConfiguration -ErrorAction SilentlyContinue)) {
            Throw "Get-IRMConfiguration is not available for the connected account. Check the Exchange Online RBAC permissions."
        }

        <# Output #>
        Write-Output "Exchange Online connected."

        <# Logging #>
        fncLogging -strLogFunction "fncCollectExchangeIRMConfiguration" -strLogDescription "Exchange Online connected" -strLogValue $true

    }
    Catch { <# Action if the Exchange Online connection failed #>

        $Private:ExchangeOnlineError = $_.Exception.Message

        <# Logging #>
        fncLogging -strLogFunction "fncCollectExchangeIRMConfiguration" -strLogDescription "Exchange Online connected" -strLogValue $false
        fncLogging -strLogFunction "fncCollectExchangeIRMConfiguration" -strLogDescription "COLLECT EXCHANGE IRM CONFIGURATION" -strLogValue "Connection failed: $Private:ExchangeOnlineError"

        <# Output #>
        Write-ColoredOutput Red "COLLECT EXCHANGE IRM CONFIGURATION: Connection failed: $Private:ExchangeOnlineError`n"

        <# Action if function was called from command line #>
        If ($Global:bolComingFromMenu -eq $false) {

            <# Set back window title to default #>
            $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

            <# Release global variable back to default (updates active) #>
            $Global:bolSkipRequiredUpdates = $false

            <# Exit #>
            Return

        }

        <# Action if function was called from the menu #>
        If ($Global:bolComingFromMenu -eq $true) {

            <# Call Pause #>
            fncPause

            <# Clear console #>
            Clear-Host

            <# Call ShowMenu #>
            fncShowMenu

            Return

        }

    }

    <# Check if COLLECT folder exist and create it, if not #>
    If ($(Test-Path -Path $Global:strUserLogPath"\Collect\ExchangeIRMConfiguration") -Eq $false) {

        New-Item -ItemType Directory -Force -Path $Global:strUserLogPath"\Collect\ExchangeIRMConfiguration" | Out-Null <# Define Collect path #>

    }

    <# Output #>
    Write-Output "Collecting Exchange IRM configuration, please wait..."

    $Private:IRMCollectionSucceeded = $true

    <# Check for existing folder #>
    If ($(Test-Path $Global:strUserLogPath"\Collect\ExchangeIRMConfiguration") -Eq $true) {

        Try {

        <# Collect Exchange IRM configuration once for console and file output #>
        $Private:IRMConfiguration = Get-IRMConfiguration -WarningAction SilentlyContinue -ErrorAction Stop

        If ($null -eq $Private:IRMConfiguration) {
            Throw "Get-IRMConfiguration returned no configuration data."
        }
        $Private:IRMConfigurationLogPath = Join-Path -Path $Global:strUserLogPath -ChildPath "Collect\ExchangeIRMConfiguration\IRMConfiguration.log"

        <# Export Exchange IRM configuration as CLIXML #>
        $Private:IRMConfiguration | Export-Clixml -Path $Global:strUserLogPath"\Collect\ExchangeIRMConfiguration\IRMconfiguration.xml" -Force -ErrorAction Stop | Out-Null

        <# Create a readable output similar to Collect AIP Service Configuration #>
        $Private:IRMTimestamp = Get-Date -Verbose:$false -UFormat "%y%m%d-%H%M%S"
        $Private:ExchangeOnlineModuleVersion = (Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1).Version
        $Private:IRMOutputLines = @(
            ("{0,-42}: {1}" -f "Date/Timestamp", $Private:IRMTimestamp)
            ("{0,-42}: {1}" -f "Module version", $Private:ExchangeOnlineModuleVersion)
        )

        ForEach ($Private:IRMProperty in $Private:IRMConfiguration.PSObject.Properties) {
            $Private:IRMPropertyValue = $Private:IRMProperty.Value

            If ($Private:IRMPropertyValue -is [System.Collections.IEnumerable] -and $Private:IRMPropertyValue -isnot [string]) {
                $Private:IRMPropertyValue = @($Private:IRMPropertyValue | ForEach-Object { [string]$_ }) -join ", "
            }

            $Private:IRMOutputLines += ("{0,-42}: {1}" -f $Private:IRMProperty.Name, $Private:IRMPropertyValue)
        }

        Add-Content -Path $Private:IRMConfigurationLogPath -Value $Private:IRMOutputLines
        Add-Content -Path $Private:IRMConfigurationLogPath -Value ""

        ForEach ($Private:IRMOutputLine in $Private:IRMOutputLines) {
            Write-ColoredOutput Yellow $Private:IRMOutputLine
        }

        Write-Output ""

        <# Logging #>
        fncLogging -strLogFunction "fncCollectExchangeIRMConfiguration" -strLogDescription "Export IRM configuration" -strLogValue "IRMconfiguration.xml"
        fncLogging -strLogFunction "fncCollectExchangeIRMConfiguration" -strLogDescription "Export readable IRM configuration" -strLogValue "IRMConfiguration.log"

        <# Release private variables #>
        $Private:IRMConfiguration = $null
        $Private:IRMConfigurationLogPath = $null
        $Private:IRMTimestamp = $null
        $Private:ExchangeOnlineModuleVersion = $null
        $Private:IRMOutputLines = $null
        $Private:IRMProperty = $null
        $Private:IRMPropertyValue = $null
        $Private:IRMOutputLine = $null

        }
        Catch {
            $Private:IRMCollectionSucceeded = $false
            $Private:IRMCollectionError = $_.Exception.Message
            fncLogging -strLogFunction "fncCollectExchangeIRMConfiguration" -strLogDescription "Collect IRM configuration failed" -strLogValue $Private:IRMCollectionError
            Write-ColoredOutput Red "COLLECT EXCHANGE IRM CONFIGURATION: Collection failed: $Private:IRMCollectionError`n"
        }

    }

    <# Disconnect from Exchange Online #>
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null

    <# Output #>
    Write-Output "Exchange Online disconnected.`n"

    If (-not $Private:IRMCollectionSucceeded) {
        fncLogging -strLogFunction "fncCollectExchangeIRMConfiguration" -strLogDescription "COLLECT EXCHANGE IRM CONFIGURATION" -strLogValue "Failed"

        If ($Global:bolComingFromMenu -eq $false) {
            $Global:bolSkipRequiredUpdates = $false
            $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle
        }

        Return
    }

    <# Logging #>
    fncLogging -strLogFunction "fncCollectExchangeIRMConfiguration" -strLogDescription "Exchange Online disconnected" -strLogValue $true
    fncLogging -strLogFunction "fncCollectExchangeIRMConfiguration" -strLogDescription "COLLECT EXCHANGE IRM CONFIGURATION" -strLogValue "Proceeded"

    <# Output #>
    Write-Output "Log file: $Global:strUserLogPath\Collect\ExchangeIRMConfiguration\IRMConfiguration.log"
    Write-Output "CLIXML file: $Global:strUserLogPath\Collect\ExchangeIRMConfiguration\IRMconfiguration.xml"
    Write-ColoredOutput Green "COLLECT EXCHANGE IRM CONFIGURATION: Proceeded.`n"

    <# Action if function was called from command line #>
    If ($Global:bolComingFromMenu -eq $false) {

        <# Set back window title to default #>
        $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

        <# Release global variable back to default (updates active) #>
        $Global:bolSkipRequiredUpdates = $false

        <# Exit #>
        Return

    }

}

<# Verify certificates issuer #>
Function fncVerifyIssuer ($strCertURL, $strEndpointName, $strLogPath) {

    <# Define variabel for TCP client/SSL stream #>
    $Private:MyClient = $Private:MySSLtream = $null

    <# Do not pass missing endpoint data to TcpClient.Connect #>
    If ([string]::IsNullOrWhiteSpace([string]$strCertURL)) {
        Write-ColoredOutput Yellow "Endpoint: $strEndpointName"
        Write-ColoredOutput Yellow "URL:      Not available"
        Write-ColoredOutput Yellow "Issuer:   Not verified`n"

        If ($(Test-Path $Global:strUserLogPath"\Collect\EndpointURLs.log") -Eq $true) {
            Add-Content -Path $Global:strUserLogPath"\Collect\EndpointURLs.log" -Value "Endpoint: $strEndpointName"
            Add-Content -Path $Global:strUserLogPath"\Collect\EndpointURLs.log" -Value "URL:      Not available"
            Add-Content -Path $Global:strUserLogPath"\Collect\EndpointURLs.log" -Value "Issuer:   Not verified`n"
        }

        Return
    }

    <# Try to verify certificates issuer #>
    Try {

        <# Create TCP client #>
        $Private:MyClient = New-Object System.Net.Sockets.TcpClient
        $Private:MyClient.ReceiveTimeout = 5000
        $Private:MyClient.SendTimeout = 5000
        $Private:MyClient.Connect($strCertURL, 443)

        <# Create SSL stream #>
        $Private:MySSLtream = [System.Net.Security.SslStream]::new($Private:MyClient.GetStream(), $false, {$true}, $null)
        $Private:MySSLtream.AuthenticateAsClient(
            $strCertURL,
            $null, <# No athentication #>
            "Tls, Tls11, Tls12", <# Enabled protocols #>
            $false <# Revocation check #>
        )

        <# Define certificate file conditions #>
        $Private:MyWebCert = $Private:MySSLtream.RemoteCertificate

        <# Export web certificate #>
        $Private:MyCertBinaries = $Private:MyWebCert.Export([Security.Cryptography.X509Certificates.X509ContentType]::Cert)
        [System.IO.File]::WriteAllBytes("$strLogPath\$strEndpointName.ce_", $Private:MyCertBinaries)

        <# Logging #>
        fncLogging -strLogFunction "fncVerifyIssuer" -strLogDescription "Export certificate" -strLogValue "$strEndpointName.ce_"

        <# Feed variable/certificate data with issuer #>
        $Private:MyWebCert = $Private:MyWebCert.Issuer

        <# Output #>
        Write-ColoredOutput Yellow "Endpoint: $strEndpointName"
        Write-ColoredOutput Yellow "URL:      https://$strCertURL"
        Write-ColoredOutput Yellow "Issuer:   $Private:MyWebCert`n"

        <# Check for existing EndpointURLs.log file and extend it, if it exist #>
        If ($(Test-Path $Global:strUserLogPath"\Collect\EndpointURLs.log") -Eq $true) {

            <# Exporting result #>
            Add-Content -Path $Global:strUserLogPath"\Collect\EndpointURLs.log" -Value "Endpoint: $strEndpointName"
            Add-Content -Path $Global:strUserLogPath"\Collect\EndpointURLs.log" -Value "URL:      https://$strCertURL"
            Add-Content -Path $Global:strUserLogPath"\Collect\EndpointURLs.log" -Value "Issuer:   $Private:MyWebCert`n"

        }

    }
    Catch {

        $Private:VerificationError = $_.Exception.Message
        Write-ColoredOutput Yellow "Endpoint: $strEndpointName"
        Write-ColoredOutput Yellow "URL:      https://$strCertURL"
        Write-ColoredOutput Yellow "Issuer:   Verification failed ($Private:VerificationError)`n"

        If ($(Test-Path $Global:strUserLogPath"\Collect\EndpointURLs.log") -Eq $true) {
            Add-Content -Path $Global:strUserLogPath"\Collect\EndpointURLs.log" -Value "Endpoint: $strEndpointName"
            Add-Content -Path $Global:strUserLogPath"\Collect\EndpointURLs.log" -Value "URL:      https://$strCertURL"
            Add-Content -Path $Global:strUserLogPath"\Collect\EndpointURLs.log" -Value "Issuer:   Verification failed ($Private:VerificationError)`n"
        }

        fncLogging -strLogFunction "fncVerifyIssuer" -strLogDescription "Certificate verification failed" -strLogValue $Private:VerificationError

    }
    Finally {

        <# Closing SSL streamt #>
        If ($Private:MySSLtream) {
            $Private:MySSLtream.Close()
        }

        <# Closing TCP client #>
        If ($Private:MyClient) {
            $Private:MyClient.Close()
        }

    }

    <# Release private variables #>
    $Private:MyWebCert = $null
    $Private:MyCertBinaries = $null
    $Private:MySSLtream= $null
    $Private:MyClient = $null
    $Private:VerificationError = $null

}

Function fncCollectDLPRulesAndPolicies {

    <# Output #>
    Write-Output "COLLECT DLP RULES AND POLICIES:"

    <# Check for admin permissions #>
    If ($Global:bolRunningPrivileged -ne $true) {

        <# Output #>
        Write-ColoredOutput Red "ATTENTION: You must run the Compliance Utility in an administrative PowerShell window as a user with local administrative privileges to continue with this option.`nCOLLECT DLP RULES AND POLICIES: Failed.`n"

        <# Action if function was called from command line #>
        If ($Global:bolComingFromMenu -eq $false) {

            <# Set back window title to default #>
            $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

            <# Release global variable back to default (updates active) #>
            $Global:bolSkipRequiredUpdates = $false

            <# Exit #>
            Return

        }

        <# Action if function was called from the menu #>
        If ($Global:bolComingFromMenu -eq $true) {

            <# Call Pause #>
            fncPause

            <# Clear console #>
            Clear-Host

            <# Call ShowMenu #>
            fncShowMenu

        }

    }

    <# Logging #>
    fncLogging -strLogFunction "fncCollectDLPRulesAndPolicies" -strLogDescription "COLLECT DLP RULES AND POLICIES" -strLogValue "Initiated"

    If (-not (fncTestRequiredModule -strModuleName "ExchangeOnlineManagement")) { Return }

    If (-not (Get-Command -Name Connect-IPPSSession -ErrorAction SilentlyContinue)) {
        Write-ColoredOutput Red "ATTENTION: The required cmdlet 'Connect-IPPSSession' is not available.`nRun 'ComplianceUtility -UpdateModules' to install or update the required module, then restart the Compliance Utility.`n"
        fncLogging -strLogFunction "fncCollectDLPRulesAndPolicies" -strLogDescription "ExchangeOnlineManagement module" -strLogValue "Connect-IPPSSession unavailable"
        Return
    }

    <# Logging #>
    fncLogging -strLogFunction "fncCollectDLPRulesAndPolicies" -strLogDescription "ExchangeOnlineManagement module version" -strLogValue (Get-Module -Verbose:$false -ListAvailable -Name ExchangeOnlineManagement).Version

    <# Output #>
    Write-Output "Connecting to Microsoft Purview compliance portal..."

    <# Remember default progress bar status: "Continue" #>
    $Private:strOriginalPreference = $Global:ProgressPreference
    $Global:ProgressPreference = "SilentlyContinue" <# Hiding progress bar #>

    <# Try to connect/logon #>
    Try {

        <# Connect to Microsoft Purview compliance portal #>
        Connect-IPPSSession -ShowBanner:$false -Verbose:$false -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null

    }
    Catch { <# Catch for any error #>

        $Private:PurviewConnectionError = $_.Exception.Message

        <# Logging #>
        fncLogging -strLogFunction "fncCollectDLPRulesAndPolicies" -strLogDescription "Microsoft Purview compliance portal connected" -strLogValue $false
        fncLogging -strLogFunction "fncCollectDLPRulesAndPolicies" -strLogDescription "Microsoft Purview compliance portal" -strLogValue "Login failed: $Private:PurviewConnectionError"

        <# Output #>
        Write-ColoredOutput Red "COLLECT DLP RULES AND POLICIES: Login failed: $Private:PurviewConnectionError`n"

        $Global:ProgressPreference = $Private:strOriginalPreference

        <# Action if function was called from command line #>
        If ($Global:bolComingFromMenu -eq $false) {

            <# Release global variable back to default (updates active) #>
            $Global:bolSkipRequiredUpdates = $false

            <# Set back window title to default #>
            $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

            <# Interrupt, because of missing internet connection #>
            Return

        }

        <# Action if function was called from the menu #>
        If ($Global:bolComingFromMenu -eq $true) {

            <# Call Pause #>
            fncPause

            <# Clear console #>
            Clear-Host

            <# Call ShowMenu #>
            fncShowMenu

            Return

        }

        Return

    }

    <# Output #>
    Write-Output "Microsoft Purview compliance portal connected."

    <# Logging #>
    fncLogging -strLogFunction "fncCollectDLPRulesAndPolicies" -strLogDescription "Microsoft Purview compliance portal connected" -strLogValue $true

    <# Output #>
    Write-Output "Collecting DLP rules and policies. Please wait, this may take a while..."

    <# Check if COLLECT\DLPRulesAndPolicies folder exist #>
    If ($(Test-Path -Path $Global:strUserLogPath"\Collect\DLPRulesAndPolicies") -Eq $false) {

        New-Item -ItemType Directory -Force -Path $Global:strUserLogPath"\Collect\DLPRulesAndPolicies" | Out-Null <# Create folder #>

    }

    <# Collecting DLP policies #>
    Get-DlpCompliancePolicy -WarningAction SilentlyContinue | Export-Clixml -Path $Global:strUserLogPath"\Collect\DLPRulesAndPolicies\DlpPolicy.xml" | Out-Null
    fncLogging -strLogFunction "fncCollectDLPRulesAndPolicies" -strLogDescription "Export DLP rules and policy" -strLogValue "DlpPolicy.xml" <# Logging #>

    <# Collecting DLP rules #>
    Get-DlpComplianceRule -WarningAction SilentlyContinue | Select-Object -Property * -ExcludeProperty SerializationData | Format-List | Export-Clixml -Path $Global:strUserLogPath"\Collect\DLPRulesAndPolicies\DlpRule.xml" | Out-Null
    fncLogging -strLogFunction "fncCollectDLPRulesAndPolicies" -strLogDescription "Export DLP rules and policy" -strLogValue "DlpRule.xml" <# Logging #>

    <# Collecting DLP distribution status #>
    Get-DlpCompliancePolicy | ForEach-Object {Get-DLPcompliancePolicy -Identity $_.Identity -DistributionDetail} | Format-List Name,GUID,Distr* | Export-Clixml -Path $Global:strUserLogPath"\Collect\DLPRulesAndPolicies\DlpPolicyDistributionStatus.xml" | Out-Null
    fncLogging -strLogFunction "fncCollectDLPRulesAndPolicies" -strLogDescription "Export DLP rules and policy" -strLogValue "DlpPolicyDistributionStatus.xml" <# Logging #>

    <# Collecting DLP sensitive information types #>
    Get-DlpSensitiveInformationType | Select-Object -Property * | Export-Clixml -Path $Global:strUserLogPath"\Collect\DLPRulesAndPolicies\DlpSensitiveInformationType.xml" | Out-Null
    fncLogging -strLogFunction "fncCollectDLPRulesAndPolicies" -strLogDescription "Export DLP rules and policy" -strLogValue "DlpSensitiveInformationType.xml" <# Logging #>

    <# Collecting DLP sensitive information type rules #>
    Get-DlpSensitiveInformationTypeRulePackage | Select-Object -Property * | Export-Clixml -Path $Global:strUserLogPath"\Collect\DLPRulesAndPolicies\DlpSensitiveInformationTypeRulePackage.xml" | Out-Null
    fncLogging -strLogFunction "fncCollectDLPRulesAndPolicies" -strLogDescription "Export DLP rules and policy" -strLogValue "DlpSensitiveInformationTypeRulePackage.xml" <# Logging #>

    <# Collecting DLP keyword dictionary #>
    Get-DlpKeywordDictionary | Select-Object -Property * | Format-List | Export-Clixml -Path $Global:strUserLogPath"\Collect\DLPRulesAndPolicies\DlpKeywordDictionary.xml" | Out-Null
    fncLogging -strLogFunction "fncCollectDLPRulesAndPolicies" -strLogDescription "Export DLP rules and policy" -strLogValue "DlpKeywordDictionary.xml" <# Logging #>

    <# Collecting DLP Exact Data Match (EDM) schemas #>
    Get-DlpEdmSchema | Select-Object -Property * | Format-List | Export-Clixml -Path $Global:strUserLogPath"\Collect\DLPRulesAndPolicies\DlpEdmSchema.xml" | Out-Null
    fncLogging -strLogFunction "fncCollectDLPRulesAndPolicies" -strLogDescription "Export DLP rules and policy" -strLogValue "DlpEdmSchema.xml" <# Logging #>

    <# Disconnect Exchange Online #>
    Disconnect-ExchangeOnline -Confirm:$false

    <# Set back progress bar to previous default #>
    $Global:ProgressPreference = $Private:strOriginalPreference

    <# Output #>
    Write-Output "Microsoft Purview compliance portal disconnected."

    <# Logging #>
    fncLogging -strLogFunction "fncCollectDLPRulesAndPolicies" -strLogDescription "Microsoft Purview compliance portal disconnected" -strLogValue $true
    fncLogging -strLogFunction "fncCollectDLPRulesAndPolicies" -strLogDescription "COLLECT DLP RULES AND POLICIES" -strLogValue "Proceeded"

    <# Output #>
    Write-Output "`nLog folder: $Global:strUserLogPath\Collect\DLPRulesAndPolicies"

    <# Output #>
    Write-ColoredOutput Green "COLLECT DLP RULES AND POLICIES: Proceeded.`n"

    <# Action if function was called from command line #>
    If ($Global:bolComingFromMenu -eq $false) {

        <# Release global variable back to default (updates active) #>
        $Global:bolSkipRequiredUpdates = $false

        <# Set back window title to default #>
        $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

        <# Interrupt, because of missing internet connection #>
        Return

    }

}

Function fncCollectUserLiceneseDetails {

    <# Isolate Microsoft Graph from assemblies already loaded in the current PowerShell process #>
    If ($env:COMPLIANCEUTILITY_GRAPH_CHILD -ne "1") {

        $Private:GraphModulePath = $MyInvocation.MyCommand.Module.Path
        If ($PSVersionTable.PSEdition -eq "Core") {
            $Private:GraphPowerShellPath = Join-Path -Path $PSHOME -ChildPath "pwsh.exe"
        }
        Else {
            $Private:GraphPowerShellPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        }

        If ([string]::IsNullOrWhiteSpace($Private:GraphModulePath) -or -not (Test-Path -LiteralPath $Private:GraphModulePath)) {
            Write-ColoredOutput Red "COLLECT USER LICENSE DETAILS: Unable to determine the Compliance Utility module path for the isolated Graph process.`n"
            Return
        }

        $Private:GraphChildCommand = @'

<# Run Microsoft Graph collection in an isolated Windows PowerShell process to prevent module and assembly conflicts #>
$ErrorActionPreference = "Stop"

Try {
    Import-Module -Name $env:COMPLIANCEUTILITY_GRAPH_MODULE_PATH -Force -ErrorAction Stop

    $Global:strUserLogPath = $env:COMPLIANCEUTILITY_GRAPH_LOG_PATH
    $Global:bolComingFromMenu = $false
    $Global:bolSkipRequiredUpdates = $true
    $Global:strDefaultWindowTitle = $Host.UI.RawUI.WindowTitle

    $module = Get-Module -Name ComplianceUtility | Where-Object { $_.Path -eq $env:COMPLIANCEUTILITY_GRAPH_MODULE_PATH } | Select-Object -First 1
    If (-not $module) {
        Throw "The Compliance Utility module could not be loaded in the isolated Graph process."
    }

    & $module { fncCollectUserLiceneseDetails }
}
Catch {
    Write-Error "Isolated Microsoft Graph collection failed: $($_.Exception.Message)"
    Exit 1
}
'@

        $Private:GraphChildEncodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Private:GraphChildCommand))
        $Private:GraphProcessStartInfo = New-Object System.Diagnostics.ProcessStartInfo
        $Private:GraphProcessStartInfo.FileName = $Private:GraphPowerShellPath
        $Private:GraphProcessStartInfo.Arguments = "-NoLogo -NoProfile -ExecutionPolicy RemoteSigned -EncodedCommand $Private:GraphChildEncodedCommand"
        $Private:GraphProcessStartInfo.UseShellExecute = $false
        $Private:GraphProcessStartInfo.CreateNoWindow = $false
        $Private:GraphProcessStartInfo.EnvironmentVariables["COMPLIANCEUTILITY_GRAPH_CHILD"] = "1"
        $Private:GraphProcessStartInfo.EnvironmentVariables["COMPLIANCEUTILITY_GRAPH_MODULE_PATH"] = $Private:GraphModulePath
        $Private:GraphProcessStartInfo.EnvironmentVariables["COMPLIANCEUTILITY_GRAPH_LOG_PATH"] = [string]$Global:strUserLogPath

        Try {
            $Private:GraphProcess = [System.Diagnostics.Process]::Start($Private:GraphProcessStartInfo)
            $Private:GraphProcess.WaitForExit()

            If ($Private:GraphProcess.ExitCode -ne 0) {
                Throw "The isolated Microsoft Graph process exited with code $($Private:GraphProcess.ExitCode)."
            }
        }
        Catch {
            Write-ColoredOutput Red "COLLECT USER LICENSE DETAILS: Isolated Microsoft Graph collection failed: $($_.Exception.Message)`n"
        }

        If ($Global:bolComingFromMenu -eq $true) {
            fncPause
            Clear-Host
            fncShowMenu
        }
        Else {
            $Global:bolSkipRequiredUpdates = $false
            $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle
        }

        Return
    }

    <# Output #>
    Write-Output "COLLECT USER LICENSE DETAILS:"

    <# Logging #>
    fncLogging -strLogFunction "fncCollectUserLiceneseDetails" -strLogDescription "COLLECT USER LICENSE DETAILS" -strLogValue "Initiated"

    If (-not (fncTestRequiredModule -strModuleName "Microsoft.Graph")) { Return }

    <# Logging #>
    fncLogging -strLogFunction "fncCollectUserLiceneseDetails" -strLogDescription "Microsoft Graph PowerShell module version" -strLogValue (Get-Module -Verbose:$false -ListAvailable -Name Microsoft.Graph).Version

    <# Output #>
    Write-Output "Connecting to Microsoft Graph..."

    <# Remember default progress bar status: "Continue" #>
    $Private:strOriginalPreference = $Global:ProgressPreference
    $Global:ProgressPreference = "SilentlyContinue" <# Hiding progress bar #>

    <# Try to connect/logon to Microsoft Graph #>
    Try {

        <# PowerShell 5.1 with the default Restricted policy blocks the signed Graph format data #>
        $Private:OriginalProcessExecutionPolicy = Get-ExecutionPolicy -Scope Process
        $Private:GraphImportPolicyChanged = $false

        Try {
            If ((Get-ExecutionPolicy) -eq "Restricted") {
                Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned -Force -ErrorAction Stop
                $Private:GraphImportPolicyChanged = $true
            }

            <# Import the required Graph submodules explicitly so load errors are caught here #>
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            Import-Module Microsoft.Graph.Users -ErrorAction Stop
            Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
        }
        Finally {
            If ($Private:GraphImportPolicyChanged) {
                Set-ExecutionPolicy -Scope Process -ExecutionPolicy $Private:OriginalProcessExecutionPolicy -Force -ErrorAction SilentlyContinue
            }
        }

        <# Connect/logon to Microsoft Graph #>
        Connect-MgGraph -Scopes "User.Read", "LicenseAssignment.Read.All" -ContextScope Process -Verbose:$false -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null

        <# Resolve the signed-in user explicitly; the Account property in Get-MgContext can be empty #>
        $Private:GraphUser = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/me?`$select=id,userPrincipalName" -ErrorAction Stop

        If ([string]::IsNullOrWhiteSpace([string]$Private:GraphUser.id)) {
            Throw "Microsoft Graph returned no user ID for the signed-in account."
        }

    }
    Catch { <# Catch action for any error that occur on connect/logon #>

        $Private:GraphConnectionError = $_.Exception.Message

        <# Logging #>
        fncLogging -strLogFunction "fncCollectUserLiceneseDetails" -strLogDescription "Microsoft Graph connected" -strLogValue $false
        fncLogging -strLogFunction "fncCollectUserLiceneseDetails" -strLogDescription "Microsoft Graph" -strLogValue "Connection failed: $Private:GraphConnectionError"

        <# Output #>
        Write-ColoredOutput Red "COLLECT USER LICENSE DETAILS: Microsoft Graph connection failed: $Private:GraphConnectionError`n"

        <# Set back progress bar to previous default #>
        $Global:ProgressPreference = $Private:strOriginalPreference

        <# Close a partially established Graph session if the authentication module was loaded #>
        If (Get-Command -Name Disconnect-MgGraph -ErrorAction SilentlyContinue) {
            Disconnect-MgGraph -Verbose:$false -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Out-Null
        }

        <# Action if function was called from the menu #>
        If ($Global:bolComingFromMenu -eq $true) {

            <# Call Pause #>
            fncPause

            <# Clear console #>
            Clear-Host

            <# Call ShowMenu #>
            fncShowMenu

            Return

        }

        <# Action if function was called from command line #>
        If ($Global:bolComingFromMenu -eq $false) {

            <# Release global variable back to default (updates active) #>
            $Global:bolSkipRequiredUpdates = $false

            <# Set back window title to default #>
            $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

            <# Interrupt, because the Microsoft Graph connection failed #>
            Return

        }

        <# Never continue collecting data after a Graph connection or module import failure #>
        Return

    }

    <# Output #>
    Write-Output "Microsoft Graph connected."

    <# Logging #>
    fncLogging -strLogFunction "fncCollectUserLiceneseDetails" -strLogDescription "Microsoft Graph connected" -strLogValue $true

    <# Output #>
    Write-Output "Collecting user license details, please wait..."

    <# Check if COLLECT folder exist and create it, if not #>
    If ($(Test-Path -Path $Global:strUserLogPath"\Collect") -Eq $false) {

        New-Item -ItemType Directory -Force -Path $Global:strUserLogPath"\Collect" | Out-Null <# Define Collect path #>

    }

    <# Check for existing UserLicenseDetails.log file and create it, if it not exist #>
    If ($(Test-Path $Global:strUserLogPath"\Collect\UserLicenseDetails.log") -Eq $false) {

        <# Create DLPRulesAndPolicies.log logging file #>
        Out-File -FilePath $Global:strUserLogPath"\Collect\UserLicenseDetails.log" -Encoding UTF8 -Append -Force

    }

    <# Check for existing UserLicenseDetails.log file and extend it #>
    $Private:GraphCollectionSucceeded = $true

    If ($(Test-Path $Global:strUserLogPath"\Collect\UserLicenseDetails.log") -Eq $true) {

        Try {

        <# Use the user resolved via Microsoft Graph instead of the potentially empty context Account property #>
        $Private:strGraphAccountUPN = $Private:GraphUser.userPrincipalName

        If ([string]::IsNullOrWhiteSpace([string]$Private:strGraphAccountUPN)) {
            $Private:strGraphAccountUPN = $Private:GraphUser.id
        }

        <# Log UPN into log file as seperator #>
        Add-Content -Path $Global:strUserLogPath"\Collect\UserLicenseDetails.log" -Value "ACCOUNT: $Private:strGraphAccountUPN`n"

        <# Collecting user license details #>
        $Private:GraphUserLicenseDetails = @(Get-MgUserLicenseDetail -UserId $Private:GraphUser.id -ErrorAction Stop -WarningAction SilentlyContinue)
        $Private:GraphUserLicenseDetails | Format-Table -AutoSize | Out-File $Global:strUserLogPath"\Collect\UserLicenseDetails.log" -Encoding UTF8 -Append -Force

        <# Collecting user service plan details #>
        $Private:GraphUserLicenseDetails.ServicePlans | Out-File $Global:strUserLogPath"\Collect\UserLicenseDetails.log" -Encoding UTF8 -Append -Force

        <# Collecting subscribed Skus - if required authorization/rule exist #>
        Get-MgSubscribedSku -ErrorAction Stop -WarningAction SilentlyContinue | Format-List | Out-File $Global:strUserLogPath"\Collect\UserLicenseDetails.log" -Encoding UTF8 -Append -Force

        <# Releasing private variable #>
        $Private:strGraphAccountUPN = $null
        $Private:GraphUserLicenseDetails = $null

        }
        Catch {
            $Private:GraphCollectionSucceeded = $false
            $Private:GraphCollectionError = $_.Exception.Message
            fncLogging -strLogFunction "fncCollectUserLiceneseDetails" -strLogDescription "Collect user license details failed" -strLogValue $Private:GraphCollectionError
            Write-ColoredOutput Red "COLLECT USER LICENSE DETAILS: Collection failed: $Private:GraphCollectionError`n"
        }

    }

    <# Disconnect Microsoft Graph #>
    Disconnect-MgGraph -Verbose:$false -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Out-Null

    <# Release Graph connection data #>
    $Private:GraphUser = $null
    $Private:GraphConnectionError = $null

    <# Set back progress bar to previous default #>
    $Global:ProgressPreference = $Private:strOriginalPreference

    If (-not $Private:GraphCollectionSucceeded) {
        fncLogging -strLogFunction "fncCollectUserLiceneseDetails" -strLogDescription "COLLECT USER LICENSE DETAILS" -strLogValue "Failed"
        Return
    }

    <# Output #>
    Write-Output "Microsoft Graph disconnected."

    <# Logging #>
    fncLogging -strLogFunction "fncCollectUserLiceneseDetails" -strLogDescription "Microsoft Graph disconnected" -strLogValue $true
    fncLogging -strLogFunction "fncCollectUserLiceneseDetails" -strLogDescription "Export user license details" -strLogValue "UserLicenseDetails.log"
    fncLogging -strLogFunction "fncCollectUserLiceneseDetails" -strLogDescription "COLLECT USER LICENSE DETAILS" -strLogValue "Proceeded"

    <# Output #>
    Write-Output "`nLog file: $Global:strUserLogPath\Collect\UserLicenseDetails.log"

    <# Output #>
    Write-ColoredOutput Green "COLLECT USER LICENSE DETAILS: Proceeded.`n"

    <# Action if function was called from command line #>
    If ($Global:bolComingFromMenu -eq $false) {

        <# Release global variable back to default (updates active) #>
        $Global:bolSkipRequiredUpdates = $false

        <# Set back window title to default #>
        $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

        <# Interrupt, because of missing internet connection #>
        Return

    }

}

Function fncCompressLogs {

    <# Output #>
    Write-Output "COMPRESS LOGS:`nCompressing logs, please wait...`n"

    <# Define default zip folder path #>
    $Global:strZipSourcePath = $Global:strTempFolder + "\ComplianceUtility"

    <# Logging #>
    fncLogging -strLogFunction "fncCompressLogs" -strLogDescription "COMPRESS LOGS" -strLogValue "Initiated"
    fncLogging -strLogFunction "fncCompressLogs" -strLogDescription "Zip source path" -strLogValue $Global:strZipSourcePath

    <# Compress all files into a .zip file #>
    If ($(Test-Path -Path $Global:strZipSourcePath) -Eq $true) { <# Actions, if path exist #>

        <# Define .zip file name #>
        $Private:strZipFile = "ComplianceUtility (" + $([System.Environment]::USERNAME) + (Get-Date -UFormat "-%H%M%S") + ").zip".ToString()

        <# Define user desktop path #>
        $Private:DesktopPath = [Environment]::GetFolderPath("Desktop")

        <# Logging #>
        fncLogging -strLogFunction "fncCompressLogs" -strLogDescription "Zip destination path" -strLogValue $Private:DesktopPath
        fncLogging -strLogFunction "fncCompressLogs" -strLogDescription "Zip file name" -strLogValue $Private:strZipFile
        fncLogging -strLogFunction "fncCompressLogs" -strLogDescription "COMPRESS LOGS" -strLogValue "Proceeded"

        <# Compress all files and logs into zip file (overwrites) #>
        Compress-Archive -Path $Global:strZipSourcePath"\*" -DestinationPath "$Private:DesktopPath\$Private:strZipFile" -Force -ErrorAction SilentlyContinue

    }

    <# Output #>
    Write-Output "Zip file: $Private:DesktopPath\$Private:strZipFile"

    <# Output #>
    Write-ColoredOutput Green "COMPRESS LOGS: Proceeded.`n"

    <# Clean Logs folders if .zip archive is on the desktop #>
    If ($(Test-Path -Path $Private:DesktopPath\$Private:strZipFile) -Eq $true) { <# Actions, if file exist on desktop #>

        <# Clean Logs folders #>
        Remove-Item "$Global:strZipSourcePath" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null

        <# Logging #>
        fncLogging -strLogFunction "fncCompressLogs" -strLogDescription "Log folders cleaned" -strLogValue $true

    }
    Else{

        <# Logging #>
        fncLogging -strLogFunction "fncCompressLogs" -strLogDescription "Log folders cleaned" -strLogValue $false

    }

    <# Release private variable #>
    $Private:strZipFile = $null
    $Private:DesktopPath = $null

    <# Release global variable #>
    $Global:strZipSourcePath = $null

}

<# Pause menu for message display #>
Function fncPause {

    <# Define and fill variables #>
    $Private:strPauseMessage = "Press any key to continue" <# Pause message #>
    $Private:strValue | Out-Null

    <# Pause the script module with a message #>
    If ($Global:psISE) { <# Actions, if running in PowerShell ISE #>

        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show("$Private:strPauseMessage")

    }
    Else { <# Actions if running in PowerShell command window #>

        <# Output #>
        Write-ColoredOutput Yellow $Private:strPauseMessage
        $Private:strValue = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

    }

    <# Logging #>
    fncLogging -strLogFunction "fncPause" -strLogDescription "PAUSE" -strLogValue "Selected"

}

Function fncShowMenu {

    <# Clear console #>
    Clear-Host

    <# Define variables #>
    $Global:bolComingFromMenu | Out-Null
    $Global:bolSkipRequiredUpdates | Out-Null

    <# Helper variable to control menu handling inside function calls #>
    $Global:bolComingFromMenu = $true

    <# Logging #>
    fncLogging -strLogFunction "fncShowMenu" -strLogDescription "MENU" -strLogValue "Selected"

    <# Menu output #>
    Write-Output "ComplianceUtility:`n"
    Write-ColoredOutput Green "  [I] INFORMATION"
    Write-ColoredOutput Green "  [M] MIT LICENSE"
    Write-ColoredOutput Green "  [H] HELP"
    Write-ColoredOutput Yellow "  [R] RESET"
    Write-ColoredOutput Yellow "  [P] RECORD PROBLEM"
    Write-ColoredOutput Yellow "  [C] COLLECT"
    If (@($Global:bolMenuCollectExtended) -Match $true) {
        Write-ColoredOutput Yellow "   ├──[A] AIP service configuration"
        Write-ColoredOutput Yellow "   ├──[T] Protection templates"
        Write-ColoredOutput Yellow "   ├──[E] Endpoint URLs"
        Write-ColoredOutput Yellow "   ├──[L] Labels and policies"
        Write-ColoredOutput Yellow "   ├──[D] DLP rules and policies"
        Write-ColoredOutput Yellow "   ├──[U] User license details"
        Write-ColoredOutput Yellow "   └──[G] Exchange IRM configuration"
    }
    Write-ColoredOutput Yellow "  [Z] COMPRESS LOGS"
    Write-ColoredOutput Yellow "  [Q] UPDATE REQUIRED MODULES"
    Write-ColoredOutput Green "  [X] EXIT`n"

    <# Define menu selection variable #>
    $Private:intMenuSelection = Read-Host "Please select an option and press enter"

    <# Actions for information menu selected #>
    If ($Private:intMenuSelection -Eq "I") {

        <# Logging #>
        fncLogging -strLogFunction "fncShowMenu" -strLogDescription "INFORMATION" -strLogValue "Selected"

        <# Clear console #>
        Clear-Host

        <# Call information function #>
        fncInformation

        <# Call Pause #>
        fncPause

    }

    <# Actions for License menu selected #>
    If ($Private:intMenuSelection -Eq "M") {

        <# Logging #>
        fncLogging -strLogFunction "fncShowMenu" -strLogDescription "MIT LICENSE" -strLogValue "Selected"

        <# Clear console #>
        Clear-Host

        <# Call License #>
        fncLicense

        <# Call Pause #>
        fncPause
    }

    <# Actions for help menu selected #>
    If ($Private:intMenuSelection -Eq "H") {

        <# Logging #>
        fncLogging -strLogFunction "fncShowMenu" -strLogDescription "HELP" -strLogValue "Selected"

        <# Clear console #>
        Clear-Host

        <# Call Help #>
        fncHelp

    }

    <# Actions for reset menu selected #>
    If ($Private:intMenuSelection -Eq "R") {

        <# Logging #>
        fncLogging -strLogFunction "fncShowMenu" -strLogDescription "RESET" -strLogValue "Selected"

        <# Clear console #>
        Clear-Host

        <# Call Reset #>
        fncReset

        <# Call Pause #>
        fncPause

    }

    <# Actions for record problem menu selected #>
    If ($Private:intMenuSelection -Eq "P") {

        <# Logging #>
        fncLogging -strLogFunction "fncShowMenu" -strLogDescription "RECORD PROBLEM" -strLogValue "Selected"

        <# Clear console #>
        Clear-Host

        <# Call user logging function #>
        fncRecordProblem

        <# Call Pause #>
        fncPause

    }

    <# COLLECT actions #>
    If ($Private:intMenuSelection -Eq "C") {

        <# Menu extenstion  #>
        If (@($Global:bolMenuCollectExtended) -Match $true) {$Global:bolMenuCollectExtended = $false}
        Else {$Global:bolMenuCollectExtended = $true}

    }

    <# Detect Windows #>
    If ([System.Environment]::OSVersion.Platform -eq "Win32NT") {

        <# Service configuration actions #>
        If ($Private:intMenuSelection -Eq "A") {

            <# Logging #>
            fncLogging -strLogFunction "fncShowMenu" -strLogDescription "COLLECT AIP SERVICE CONFIGURATION" -strLogValue "Selected"

            <# Clear console #>
            Clear-Host

            <# Call CollectAIPServiceConfigurationn #>
            fncCollectAIPServiceConfiguration

            <# Call Pause #>
            fncPause

        }

        <# Protection templates actions #>
        If ($Private:intMenuSelection -Eq "T") {

            <# Logging #>
            fncLogging -strLogFunction "fncShowMenu" -strLogDescription "COLLECT PROTECTION TEMPLATES" -strLogValue "Selected"

            <# Clear console #>
            Clear-Host

            <# Call CollectProtectionTemplates #>
            fncCollectProtectionTemplates

            <# Call Pause #>
            fncPause

        }

        <# CollectEndpointURLs actions #>
        If ($Private:intMenuSelection -Eq "E") {

            <# Logging #>
            fncLogging -strLogFunction "fncShowMenu" -strLogDescription "COLLECT ENDPOINT URLs" -strLogValue "Selected"

            <# Clear console #>
            Clear-Host

            <# Call CollectEndpointURLs #>
            fncCollectEndpointURLs

            <# Call Pause #>
            fncPause

        }

    }

    <# labels and policies actions #>
    If ($Private:intMenuSelection -Eq "L") {

        <# Logging #>
        fncLogging -strLogFunction "fncShowMenu" -strLogDescription "COLLECT LABELS AND POLICIES" -strLogValue "Selected"

        <# Clear console #>
        Clear-Host

        <# Call CollectLabelsAndPolicies #>
        fncCollectLabelsAndPolicies

        <# Call Pause #>
        fncPause

    }

    <# DLP rules and policies actions #>
    If ($Private:intMenuSelection -Eq "D") {

        <# Logging #>
        fncLogging -strLogFunction "fncShowMenu" -strLogDescription "COLLECT DLP RULES AND POLICIES" -strLogValue "Selected"

        <# Clear console #>
        Clear-Host

        <# Call CollectDLPRulesAndPolicies #>
        fncCollectDLPRulesAndPolicies

        <# Call Pause #>
        fncPause

    }

    <# User license details actions #>
    If ($Private:intMenuSelection -Eq "U") {

        <# Logging #>
        fncLogging -strLogFunction "fncShowMenu" -strLogDescription "COLLECT USER LICENSE DETAILS" -strLogValue "Selected"

        <# Clear console #>
        Clear-Host

        <# Call CollectUserLiceneseDetails #>
        fncCollectUserLiceneseDetails

        <# Call Pause #>
        fncPause

    }

    <# Exchange IRM configuration actions #>
    If ($Private:intMenuSelection -Eq "G") {

        <# Logging #>
        fncLogging -strLogFunction "fncShowMenu" -strLogDescription "COLLECT EXCHANGE IRM CONFIGURATION" -strLogValue "Selected"

        <# Clear console #>
        Clear-Host

        <# Call CollectExchangeIRMConfiguration #>
        fncCollectExchangeIRMConfiguration

        <# Call Pause #>
        fncPause

    }

    <# Compress logs actions #>
    If ($Private:intMenuSelection -Eq "Z") {

        <# Logging #>
        fncLogging -strLogFunction "fncShowMenu" -strLogDescription "COMPRESS LOGS" -strLogValue "Selected"

        <# Clear console #>
        Clear-Host

        <# Call CompressLogs #>
        fncCompressLogs

        <# Call Pause #>
        fncPause

    }

    <# Actions for required module updates #>
    If ($Private:intMenuSelection -Eq "Q") {

        <# Logging #>
        fncLogging -strLogFunction "fncShowMenu" -strLogDescription "UPDATE REQUIRED MODULES" -strLogValue "Selected"

        <# Clear console #>
        Clear-Host

        <# Call update function #>
        fncUpdateModules

        <# Call Pause #>
        fncPause

        <# Return to menu #>
        fncShowMenu

        Return

    }

    <# Exit menu actions #>
    If ($Private:intMenuSelection -Eq "X") {

        <# Logging #>
        fncLogging -strLogFunction "fncShowMenu" -strLogDescription "EXIT" -strLogValue "Selected"

        <# Clear variable #>
        $Global:bolComingFromMenu = $false

        <# Release variable (updates active) #>
        $Global:bolSkipRequiredUpdates = $false

        <# Set back window title #>
        $Global:host.UI.RawUI.WindowTitle = $Global:strDefaultWindowTitle

        <# Exit #>
        Break

    }
    Else {

        <# Clear console #>
        Clear-Host

        <# Call ShowMenu #>
        fncShowMenu

    }

}

<# Initialize module #>
fncInitialize

<# Detect enabled logging #>
fncValidateForActivatedLogging

<# Export functions #>
Export-ModuleMember -Function ComplianceUtility -Alias "CompUtil"

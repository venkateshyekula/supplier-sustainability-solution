#Requires -Version 5.1

<#
.SYNOPSIS
Installs the Supplier Sustainability SPFx app and registers
the Supplier Navigation Application Customizer only on the
configured Pre-Prod and Production SharePoint sites.

.REQUIREMENTS
Windows PowerShell 5.1
PnP.PowerShell 1.12.0
The SPFx package must already be uploaded and published in
the Tenant App Catalog.

.NOTES
Application Customizer Component ID:
f961f110-75b0-45dc-90ec-2accd2523d22
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$PreProdSiteUrl =
        "https://liquidtelecommunications.sharepoint.com/sites/ltsadev/sc",

    [Parameter(Mandatory = $true)]
    [string]$ProdSiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$PnPClientId,

    [Parameter(Mandatory = $true)]
    [string]$SolutionId,

    [Parameter(Mandatory = $false)]
    [string]$ComponentId =
        "f961f110-75b0-45dc-90ec-2accd2523d22"
)

Set-StrictMode -Version 2.0

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

function Write-Step {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Host ""
    Write-Host "============================================================" `
        -ForegroundColor DarkGray

    Write-Host $Message `
        -ForegroundColor Cyan

    Write-Host "============================================================" `
        -ForegroundColor DarkGray
}

function Write-Success {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Host "[SUCCESS] $Message" `
        -ForegroundColor Green
}

function Write-WarningMessage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Host "[WARNING] $Message" `
        -ForegroundColor Yellow
}

function Normalize-SiteUrl {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl
    )

    return $SiteUrl.Trim().TrimEnd("/")
}

function Get-SitePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl
    )

    $siteUri = New-Object System.Uri($SiteUrl)

    $sitePath = $siteUri.AbsolutePath.TrimEnd("/")

    if (:IsNullOrWhiteSpace($sitePath)) {
        return ""
    }

    return $sitePath
}

function Get-ApplicationCustomizerProperties {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SitePath,

        [Parameter(Mandatory = $true)]
        [string]$AllowedSiteUrls
    )

    $properties = [ordered]@{
        organizationName =
            "Cassava Technologies"

        confidentialityText =
            "All information is confidential and must be used in accordance with our privacy and security policies."

        homeUrl =
            "$SitePath/SitePages/Home.aspx"

        completedQuestionnairesUrl =
            "$SitePath/SitePages/Completed-Questionnaires.aspx"

        supportingDocumentsUrl =
            "$SitePath/SitePages/Supporting-Documents.aspx"

        dashboardUrl =
            "$SitePath/SitePages/Dashboard.aspx"

        helpUrl =
            "$SitePath/SitePages/Help.aspx"

        footerText =
            "Supplier ESG questionnaire and sustainability portal"

        privacyUrl =
            "$SitePath/SitePages/Privacy.aspx"

        accessibilityUrl =
            "$SitePath/SitePages/Accessibility.aspx"

        allowedSiteUrls =
            $AllowedSiteUrls
    }

    return $properties |
        ConvertTo-Json `
            -Compress `
            -Depth 10
}

function Test-RequiredPage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PageServerRelativeUrl
    )

    try {
        $page =
            Get-PnPFile `
                -Url $PageServerRelativeUrl `
                -AsFileObject `
                -ErrorAction Stop

        Write-Success `
            "Page found: $($page.ServerRelativeUrl)"
    }
    catch {
        Write-WarningMessage `
            "Page was not found or is inaccessible: $PageServerRelativeUrl"
    }
}

function Install-SolutionOnSite {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppSolutionId
    )

    Write-Host `
        "Checking whether the SPFx app is available..." `
        -ForegroundColor Gray

    $availableApp =
        Get-PnPApp `
            -Identity $AppSolutionId `
            -Scope Tenant `
            -ErrorAction Stop

    if ($null -eq $availableApp) {
        throw (
            "The SPFx app with Solution ID " +
            "$AppSolutionId was not found in the Tenant App Catalog."
        )
    }

    $installedApps =
        Get-PnPApp `
            -Scope Tenant `
            -ErrorAction Stop

    $installedApp =
        $installedApps |
        Where-Object {
            $_.Id -eq [Guid]$AppSolutionId -and
            -not :IsNullOrWhiteSpace(
                $_.InstalledVersion
            )
        } |
        Select-Object -First 1

    if ($null -eq $installedApp) {
        Write-Host `
            "Installing the app on the current site..." `
            -ForegroundColor Yellow

        Install-PnPApp `
            -Identity $AppSolutionId `
            -Scope Tenant `
            -Wait `
            -ErrorAction Stop

        Write-Success `
            "The SPFx app was installed on the site."
    }
    else {
        Write-Host `
            "The SPFx app is already installed. Checking for an update..." `
            -ForegroundColor Yellow

        try {
            Update-PnPApp `
                -Identity $AppSolutionId `
                -Scope Tenant `
                -ErrorAction Stop

            Write-Success `
                "The installed SPFx app was updated."
        }
        catch {
            Write-WarningMessage `
                "No app update was required, or the installed app is already current."
        }
    }
}

function Remove-ExistingRegistration {
    param(
        [Parameter(Mandatory = $true)]
        [Guid]$ClientSideComponentId
    )

    Write-Host `
        "Checking existing Application Customizer registrations..." `
        -ForegroundColor Gray

    $existingCustomizers =
        Get-PnPApplicationCustomizer `
            -Scope Web `
            -ErrorAction Stop

    $matchingCustomizers =
        $existingCustomizers |
        Where-Object {
            $_.ClientSideComponentId -eq
                $ClientSideComponentId
        }

    foreach ($customizer in $matchingCustomizers) {
        Write-Host `
            "Removing existing registration: $($customizer.Title)" `
            -ForegroundColor Yellow

        Remove-PnPApplicationCustomizer `
            -Identity $customizer.Id `
            -Scope Web `
            -Force `
            -ErrorAction Stop

        Write-Success `
            "Existing Application Customizer registration removed."
    }
}

function Register-ApplicationCustomizer {
    param(
        [Parameter(Mandatory = $true)]
        [Guid]$ClientSideComponentId,

        [Parameter(Mandatory = $true)]
        [string]$PropertiesJson
    )

    Write-Host `
        "Application Customizer properties:" `
        -ForegroundColor Yellow

    Write-Host $PropertiesJson

    Add-PnPApplicationCustomizer `
        -Title "Supplier Navigation" `
        -Name "SupplierNavigation" `
        -Description "Supplier Sustainability navigation and footer" `
        -ClientSideComponentId $ClientSideComponentId `
        -ClientSideComponentProperties $PropertiesJson `
        -Scope Web `
        -Sequence 10 `
        -ErrorAction Stop

    Write-Success `
        "The Supplier Navigation Application Customizer was registered."
}

function Verify-ApplicationCustomizer {
    param(
        [Parameter(Mandatory = $true)]
        [Guid]$ClientSideComponentId
    )

    $registration =
        Get-PnPApplicationCustomizer `
            -Scope Web `
            -ErrorAction Stop |
        Where-Object {
            $_.ClientSideComponentId -eq
                $ClientSideComponentId
        } |
        Select-Object -First 1

    if ($null -eq $registration) {
        throw (
            "The Application Customizer registration " +
            "could not be verified."
        )
    }

    Write-Success `
        "Application Customizer registration verified."

    Write-Host "Title: $($registration.Title)"
    Write-Host "Custom Action ID: $($registration.Id)"
    Write-Host "Component ID: $($registration.ClientSideComponentId)"
    Write-Host "Scope: Web"
}

function Deploy-ToSite {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,

        [Parameter(Mandatory = $true)]
        [string]$AllowedSiteUrls,

        [Parameter(Mandatory = $true)]
        [string]$EntraClientId,

        [Parameter(Mandatory = $true)]
        [string]$AppSolutionId,

        [Parameter(Mandatory = $true)]
        [Guid]$ClientSideComponentId
    )

    $normalizedSiteUrl =
        Normalize-SiteUrl `
            -SiteUrl $SiteUrl

    $sitePath =
        Get-SitePath `
            -SiteUrl $normalizedSiteUrl

    Write-Step `
        "Connecting to $normalizedSiteUrl"

    Connect-PnPOnline `
        -Url $normalizedSiteUrl `
        -ClientId $EntraClientId `
        -UseWebLogin `
        -ErrorAction Stop

    Write-Success `
        "Connected to $normalizedSiteUrl"

    Write-Step `
        "Installing or updating the SPFx application"

    Install-SolutionOnSite `
        -AppSolutionId $AppSolutionId

    Write-Step `
        "Checking the configured SharePoint pages"

    Test-RequiredPage `
        -PageServerRelativeUrl (
            "$sitePath/SitePages/Home.aspx"
        )

    Test-RequiredPage `
        -PageServerRelativeUrl (
            "$sitePath/SitePages/Completed-Questionnaires.aspx"
        )

    Test-RequiredPage `
        -PageServerRelativeUrl (
            "$sitePath/SitePages/Supporting-Documents.aspx"
        )

    Test-RequiredPage `
        -PageServerRelativeUrl (
            "$sitePath/SitePages/Dashboard.aspx"
        )

    Test-RequiredPage `
        -PageServerRelativeUrl (
            "$sitePath/SitePages/Help.aspx"
        )

    Write-Step `
        "Removing previous Supplier Navigation registration"

    Remove-ExistingRegistration `
        -ClientSideComponentId $ClientSideComponentId

    $propertiesJson =
        Get-ApplicationCustomizerProperties `
            -SitePath $sitePath `
            -AllowedSiteUrls $AllowedSiteUrls

    Write-Step `
        "Registering the Supplier Navigation Application Customizer"

    Register-ApplicationCustomizer `
        -ClientSideComponentId $ClientSideComponentId `
        -PropertiesJson $propertiesJson

    Write-Step `
        "Verifying the Application Customizer registration"

    Verify-ApplicationCustomizer `
        -ClientSideComponentId $ClientSideComponentId

    Disconnect-PnPOnline

    Write-Success `
        "Deployment completed for $normalizedSiteUrl"
}

try {
    Write-Host ""
    Write-Host `
        "Supplier Sustainability Deployment - PowerShell 5.1" `
        -ForegroundColor White `
        -BackgroundColor DarkBlue

    Write-Step `
        "Checking Windows PowerShell and PnP module"

    if ($PSVersionTable.PSVersion.Major -ne 5) {
        Write-WarningMessage `
            "This script was prepared for Windows PowerShell 5.1."
    }

    $pnpModule =
        Get-Module `
            -ListAvailable `
            -Name PnP.PowerShell |
        Where-Object {
            $_.Version -eq
                [Version]"1.12.0"
        } |
        Select-Object -First 1

    if ($null -eq $pnpModule) {
        throw @"
PnP.PowerShell 1.12.0 is not installed.

Run:

Install-Module PnP.PowerShell -RequiredVersion 1.12.0 -Scope CurrentUser -Force -AllowClobber
"@
    }

    Import-Module `
        PnP.PowerShell `
        -RequiredVersion 1.12.0 `
        -Force `
        -ErrorAction Stop

    Write-Success `
        "PnP.PowerShell 1.12.0 loaded."

    $normalizedPreProdUrl =
        Normalize-SiteUrl `
            -SiteUrl $PreProdSiteUrl

    $normalizedProdUrl =
        Normalize-SiteUrl `
            -SiteUrl $ProdSiteUrl

    $allowedSiteUrls =
        "$normalizedPreProdUrl;$normalizedProdUrl"

    Write-Host ""
    Write-Host "Allowed sites:" `
        -ForegroundColor Cyan

    Write-Host $allowedSiteUrls

    $applicationCustomizerGuid =
        [Guid]$ComponentId

    Write-Step `
        "Deploying to Pre-Prod"

    Deploy-ToSite `
        -SiteUrl $normalizedPreProdUrl `
        -AllowedSiteUrls $allowedSiteUrls `
        -EntraClientId $PnPClientId `
        -AppSolutionId $SolutionId `
        -ClientSideComponentId $applicationCustomizerGuid

    Write-Step `
        "Deploying to Production"

    Deploy-ToSite `
        -SiteUrl $normalizedProdUrl `
        -AllowedSiteUrls $allowedSiteUrls `
        -EntraClientId $PnPClientId `
        -AppSolutionId $SolutionId `
        -ClientSideComponentId $applicationCustomizerGuid

    Write-Step `
        "Deployment completed successfully"

    Write-Host "Pre-Prod: $normalizedPreProdUrl"
    Write-Host "Production: $normalizedProdUrl"

    Write-Host ""
    Write-Host `
        "The header and footer were registered only on these two sites." `
        -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "Deployment failed." `
        -ForegroundColor White `
        -BackgroundColor DarkRed

    Write-Host $_.Exception.Message `
        -ForegroundColor Red

    Write-Host ""

    exit 1
}
finally {
    Disconnect-PnPOnline `
        -ErrorAction SilentlyContinue
}
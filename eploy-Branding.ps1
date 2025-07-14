# To run this script:
# 1. Modify the configuration section below to match your environment
# 2. Create a CSV file with a column named "URL" containing SharePoint site URLs
# 3. Make sure you have PnP PowerShell module installed (script will prompt to install if needed)
# 4. Configure branding options:
#    - Set ChangeLogoImage to control whether to update the site logo
#    - Set ApplyThemeColors to control whether to apply theme colors
#    - Set ProcessSubsites to control whether to process subsites in addition to site collections
#    - Choose a color theme by setting the ColorTheme parameter if applying colors
# 5. Run the script: .\Deploy-Branding.ps1
#
# Note: This script is completely self-contained and doesn't require any additional configuration files

<#
.SYNOPSIS
    Updates SharePoint Online site branding (logo and theme) for multiple sites.

.DESCRIPTION
    This script reads a list of SharePoint Online site URLs from a CSV file,
    connects to each site based on its geo-location, and updates the site's logo
    and theme. Includes logging functionality to track progress and errors.

.EXAMPLE
    Simply modify the configuration section at the top of the script and run:
    .\Deploy-Branding.ps1

.NOTES
    Author: Bill Hadden | Mike Lee
    Date: 7/14/2025
    Requirements: 
    - PnP PowerShell module
    - SharePoint Online Management Shell
    - Appropriate permissions to modify site branding
#>

#region Configuration
# ===== CONFIGURATION SECTION - MODIFY THESE VALUES =====
$Config = @{
    # Path to the CSV file containing SharePoint site URLs (must have a URL column)
    CsvPath          = "C:\temp\sitelist-m365x61250205.csv"
    
    # Path to the logo image file that will be uploaded and set for all sites
    LogoPath         = "C:\temp\branding\logo.png"
    
    # Whether to change the site logo image
    # Set to $false if you only want to update theme colors without changing the logo
    ChangeLogoImage  = $true
    
    # Whether to apply the selected theme colors to the sites
    # Set to $false if you only want to update the logo without changing site colors
    ApplyThemeColors = $true
    
    # Whether to process subsites in addition to site collections
    # Set to $true to apply branding to all subsites of the site collections in the CSV
    ProcessSubsites  = $true
    
    # Theme selection - Choose from: 
    # Teal, Red, Orange, Green, Blue, 
    # Purple, Gray, Periwinkle, DarkYellow, DarkBlue, Custom
    
    # Available themes are defined in the $ColorThemes array below
    ColorTheme       = "Periwinkle"
    
    # Theme configuration - set to $null to skip theme application
    Theme            = @{
        Name      = "CompanyTheme"
        # The Colors property will be populated with the selected theme from $ColorThemes
        # No need to define colors here - they will be set based on ColorTheme selection
        Colors    = @{}
        
        # Set to $true to overwrite an existing theme with the same name
        Overwrite = $true
    }
    
    # Entra App Registration details
    TenantName       = "m365x61250205"
    AppId            = "5baa1427-1e90-4501-831d-a8e67465f0d9"
    Thumbprint       = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9"
    TenantId         = "85612ccb-4c28-4a34-88df-a538cc139a51"
}

# Set up log file path in the temp directory
$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$Config.LogPath = "$env:TEMP\SPO_Branding_${date}_logfile.log"

# Define color theme dictionaries for easy selection
$ColorThemes = @{
    # Custom theme - Modify these colors to match your brand
    # You can use the Microsoft Theme Designer tool to create a custom theme: https://aka.ms/themedesigner
    "Custom"     = @{
        "themePrimary"         = "#107c10";  # Primary brand color
        "themeDarkAlt"         = "#0c6b0c";  # Slightly darker than primary
        "themeDark"            = "#085708";  # Darker than primary
        "themeDarker"          = "#064306";  # Darkest version of primary
        "themeSecondary"       = "#218721";  # Slightly different shade of primary
        "themeTertiary"        = "#559e55";  # Medium shade of primary
        "themeLight"           = "#8fc28f";  # Light shade of primary
        "themeLighter"         = "#c6e7c6";  # Lighter shade of primary
        "themeLighterAlt"      = "#e5f4e5";  # Lightest shade of primary
        
        "neutralPrimary"       = "#333333";  # Main text color
        "neutralDark"          = "#212121";  # Darker text
        "neutralSecondary"     = "#666666";  # Secondary text
        "neutralTertiary"      = "#a6a6a6";  # Less important text
        "neutralTertiaryAlt"   = "#c8c8c8";  
        "neutralQuaternary"    = "#d0d0d0";
        "neutralQuaternaryAlt" = "#dadada";
        "neutralLight"         = "#e5e5e5";  # Light backgrounds
        "neutralLighter"       = "#f0f0f0";  # Lighter backgrounds
        "neutralLighterAlt"    = "#f8f8f8";  # Lightest backgrounds
        "neutralSecondaryAlt"  = "#767676";
        "neutralPrimaryAlt"    = "#3c3c3c";
        
        "black"                = "#000000";  # True black
        "white"                = "#ffffff";  # True white
        "primaryBackground"    = "#ffffff";  # Page background color
        "primaryText"          = "#333333";  # Primary text color
        "accent"               = "#107c10";  # Accent color for highlights
    };
    # Teal theme
    "Teal"       = @{
        "themeDarker"          = "#014446";
        "themeDark"            = "#025c5f";
        "themeDarkAlt"         = "#026d70";
        "themePrimary"         = "#03787c";
        "themeSecondary"       = "#13898d";
        "themeTertiary"        = "#49aeb1";
        "themeLight"           = "#98d6d8";
        "themeLighter"         = "#c5e9ea";
        "themeLighterAlt"      = "#f0f9fa";
        "black"                = "#000000";
        "neutralDark"          = "#212121";
        "neutralPrimary"       = "#333333";
        "neutralPrimaryAlt"    = "#3c3c3c";
        "neutralSecondary"     = "#666666";
        "neutralTertiary"      = "#a6a6a6";
        "neutralTertiaryAlt"   = "#c8c8c8";
        "neutralLight"         = "#eaeaea";
        "neutralLighter"       = "#f4f4f4";
        "neutralLighterAlt"    = "#f8f8f8";
        "white"                = "#ffffff";
        "neutralQuaternaryAlt" = "#dadada";
        "neutralQuaternary"    = "#d0d0d0";
        "neutralSecondaryAlt"  = "#767676";
        "primaryBackground"    = "#ffffff";
        "primaryText"          = "#333333";
        "accent"               = "#4f6bed";
    };
    # Red theme
    "Red"        = @{
        "themeDarker"          = "#751b1e";
        "themeDark"            = "#952226";
        "themeDarkAlt"         = "#c02b30";
        "themePrimary"         = "#d13438";
        "themeSecondary"       = "#d6494d";
        "themeTertiary"        = "#ecaaac";
        "themeLight"           = "#f6d6d8";
        "themeLighter"         = "#faebeb";
        "themeLighterAlt"      = "#fdf5f5";
        "black"                = "#000000";
        "neutralDark"          = "#212121";
        "neutralPrimary"       = "#333333";
        "neutralPrimaryAlt"    = "#3c3c3c";
        "neutralSecondary"     = "#666666";
        "neutralTertiary"      = "#a6a6a6";
        "neutralTertiaryAlt"   = "#c8c8c8";
        "neutralLight"         = "#eaeaea";
        "neutralLighter"       = "#f4f4f4";
        "neutralLighterAlt"    = "#f8f8f8";
        "white"                = "#ffffff";
        "neutralQuaternaryAlt" = "#dadada";
        "neutralQuaternary"    = "#d0d0d0";
        "neutralSecondaryAlt"  = "#767676";
        "primaryBackground"    = "#ffffff";
        "primaryText"          = "#333333";
        "accent"               = "#ca5010";
    };
    # Orange theme
    "Orange"     = @{
        "themeDarker"          = "#6f2d09";
        "themeDark"            = "#8d390b";
        "themeDarkAlt"         = "#b5490f";
        "themePrimary"         = "#ca5010";
        "themeSecondary"       = "#e55c12";
        "themeTertiary"        = "#f6b28d";
        "themeLight"           = "#fbdac9";
        "themeLighter"         = "#fdede4";
        "themeLighterAlt"      = "#fef6f1";
        "black"                = "#000000";
        "neutralDark"          = "#212121";
        "neutralPrimary"       = "#333333";
        "neutralPrimaryAlt"    = "#3c3c3c";
        "neutralSecondary"     = "#666666";
        "neutralTertiary"      = "#a6a6a6";
        "neutralTertiaryAlt"   = "#c8c8c8";
        "neutralLight"         = "#eaeaea";
        "neutralLighter"       = "#f4f4f4";
        "neutralLighterAlt"    = "#f8f8f8";
        "white"                = "#ffffff";
        "neutralQuaternaryAlt" = "#dadada";
        "neutralQuaternary"    = "#d0d0d0";
        "neutralSecondaryAlt"  = "#767676";
        "primaryBackground"    = "#ffffff";
        "primaryText"          = "#333333";
        "accent"               = "#986f0b";
    };
    # Green theme
    "Green"      = @{
        "themePrimary"         = "#498205";
        "themeLighterAlt"      = "#f6faf0";
        "themeLighter"         = "#dbebc7";
        "themeLight"           = "#bdda9b";
        "themeTertiary"        = "#85b44c";
        "themeSecondary"       = "#5a9117";
        "themeDarkAlt"         = "#427505";
        "themeDark"            = "#386304";
        "themeDarker"          = "#294903";
        "neutralLighterAlt"    = "#faf9f8";
        "neutralLighter"       = "#f3f2f1";
        "neutralLight"         = "#edebe9";
        "neutralQuaternaryAlt" = "#e1dfdd";
        "neutralQuaternary"    = "#d2d0ce";
        "neutralTertiaryAlt"   = "#c8c6c4";
        "neutralTertiary"      = "#a19f9d";
        "neutralSecondaryAlt"  = "#8a8886";
        "neutralSecondary"     = "#605e5c";
        "neutralPrimary"       = "#323130";
        "neutralPrimaryAlt"    = "#3b3a39";
        "neutralDark"          = "#201f1e";
        "black"                = "#000000";
        "white"                = "#ffffff";
        "primaryBackground"    = "#ffffff";
        "primaryText"          = "#333333";
        "accent"               = "#03787c";
    };
    # Blue theme
    "Blue"       = @{
        "themePrimary"         = "#0078d7";
        "themeLighterAlt"      = "#eff6fc";
        "themeLighter"         = "#deecf9";
        "themeLight"           = "#c7e0f4";
        "themeTertiary"        = "#71afe5";
        "themeSecondary"       = "#2b88d8";
        "themeDarkAlt"         = "#106ebe";
        "themeDark"            = "#005a9e";
        "themeDarker"          = "#004578";
        "neutralLighterAlt"    = "#f8f8f8";
        "neutralLighter"       = "#f4f4f4";
        "neutralLight"         = "#eaeaea";
        "neutralQuaternaryAlt" = "#dadada";
        "neutralQuaternary"    = "#d0d0d0";
        "neutralTertiaryAlt"   = "#c8c8c8";
        "neutralTertiary"      = "#a6a6a6";
        "neutralSecondaryAlt"  = "#767676";
        "neutralSecondary"     = "#666666";
        "neutralPrimary"       = "#333333";
        "neutralPrimaryAlt"    = "#3c3c3c";
        "neutralDark"          = "#212121";
        "black"                = "#000000";
        "white"                = "#ffffff";
        "primaryBackground"    = "#ffffff";
        "primaryText"          = "#333333";
        "accent"               = "#8764b8";
    };
    # Purple theme
    "Purple"     = @{
        "themePrimary"         = "#6b69d6";
        "themeLighterAlt"      = "#f8f7fd";
        "themeLighter"         = "#f0f0fb";
        "themeLight"           = "#e1e1f7";
        "themeTertiary"        = "#c1c0ee";
        "themeSecondary"       = "#7a78da";
        "themeDarkAlt"         = "#5250cf";
        "themeDark"            = "#3230b0";
        "themeDarker"          = "#27268a";
        "neutralLighterAlt"    = "#f8f8f8";
        "neutralLighter"       = "#f4f4f4";
        "neutralLight"         = "#eaeaea";
        "neutralQuaternaryAlt" = "#dadada";
        "neutralQuaternary"    = "#d0d0d0";
        "neutralTertiaryAlt"   = "#c8c8c8";
        "neutralTertiary"      = "#a6a6a6";
        "neutralSecondaryAlt"  = "#767676";
        "neutralSecondary"     = "#666666";
        "neutralPrimary"       = "#333333";
        "neutralPrimaryAlt"    = "#3c3c3c";
        "neutralDark"          = "#212121";
        "black"                = "#000000";
        "white"                = "#ffffff";
        "primaryBackground"    = "#ffffff";
        "primaryText"          = "#333333";
        "accent"               = "#038387";
    };
    # Gray theme
    "Gray"       = @{
        "themePrimary"         = "#5d5a58";
        "themeLighterAlt"      = "#f7f7f7";
        "themeLighter"         = "#efeeee";
        "themeLight"           = "#dfdedd";
        "themeTertiary"        = "#bbb9b8";
        "themeSecondary"       = "#6d6a67";
        "themeDarkAlt"         = "#53504e";
        "themeDark"            = "#403e3d";
        "themeDarker"          = "#323130";
        "neutralLighterAlt"    = "#f8f8f8";
        "neutralLighter"       = "#f4f4f4";
        "neutralLight"         = "#eaeaea";
        "neutralQuaternaryAlt" = "#dadada";
        "neutralQuaternary"    = "#d0d0d0";
        "neutralTertiaryAlt"   = "#c8c8c8";
        "neutralTertiary"      = "#a6a6a6";
        "neutralSecondaryAlt"  = "#767676";
        "neutralSecondary"     = "#666666";
        "neutralPrimary"       = "#333333";
        "neutralPrimaryAlt"    = "#3c3c3c";
        "neutralDark"          = "#212121";
        "black"                = "#000000";
        "white"                = "#ffffff";
        "primaryBackground"    = "#ffffff";
        "primaryText"          = "#333333";
        "accent"               = "#0078d4";
    };
    # Periwinkle theme
    "Periwinkle" = @{
        "themeDarker"          = "#383966";
        "themeDark"            = "#3D3E78";
        "themeDarkAlt"         = "#444791";
        "themePrimary"         = "#5B5FC7";
        "themeSecondary"       = "#7579EB";
        "themeTertiary"        = "#7F85F5";
        "themeLight"           = "#AAB1FA";
        "themeLighter"         = "#B6BCFA";
        "themeLighterAlt"      = "#C5CBFA";
        "black"                = "#000000";
        "neutralDark"          = "#201f1e";
        "neutralPrimary"       = "#323130";
        "neutralPrimaryAlt"    = "#3b3a39";
        "neutralSecondary"     = "#605e5c";
        "neutralTertiary"      = "#a19f9d";
        "neutralTertiaryAlt"   = "#c8c6c4";
        "neutralLight"         = "#edebe9";
        "neutralLighter"       = "#f3f2f1";
        "neutralLighterAlt"    = "#faf9f8";
        "white"                = "#ffffff";
        "neutralQuaternaryAlt" = "#dadada";
        "neutralQuaternary"    = "#d0d0d0";
        "neutralSecondaryAlt"  = "#767676";
        "primaryBackground"    = "#ffffff";
        "primaryText"          = "#333333";
        "accent"               = "#5B5FC7";
    };
    # Dark Yellow theme (dark theme)
    "DarkYellow" = @{
        "themePrimary"         = "#fce100";
        "themeLighterAlt"      = "#0d0b00";
        "themeLighter"         = "#191700";
        "themeLight"           = "#322d00";
        "themeTertiary"        = "#6a5f00";
        "themeSecondary"       = "#e3cc00";
        "themeDarkAlt"         = "#ffe817";
        "themeDark"            = "#ffed4b";
        "themeDarker"          = "#fff171";
        "neutralLighterAlt"    = "#282828";
        "neutralLighter"       = "#313131";
        "neutralLight"         = "#3f3f3f";
        "neutralQuaternaryAlt" = "#484848";
        "neutralQuaternary"    = "#4f4f4f";
        "neutralTertiaryAlt"   = "#6d6d6d";
        "neutralTertiary"      = "#c8c8c8";
        "neutralSecondaryAlt"  = "#d0d0d0";
        "neutralSecondary"     = "#dadada";
        "neutralPrimaryAlt"    = "#eaeaea";
        "neutralPrimary"       = "#ffffff";
        "neutralDark"          = "#f4f4f4";
        "black"                = "#f8f8f8";
        "white"                = "#1f1f1f";
        "primaryBackground"    = "#1f1f1f";
        "primaryText"          = "#ffffff";
        "error"                = "#ff5f5f";
        "accent"               = "#ffc83d";
    };
    # Dark Blue theme (dark theme)
    "DarkBlue"   = @{
        "themePrimary"         = "#00bcf2";
        "themeLighterAlt"      = "#00090c";
        "themeLighter"         = "#001318";
        "themeLight"           = "#002630";
        "themeTertiary"        = "#005066";
        "themeSecondary"       = "#00abda";
        "themeDarkAlt"         = "#0ecbff";
        "themeDark"            = "#44d6ff";
        "themeDarker"          = "#6cdfff";
        "neutralLighterAlt"    = "#2e3340";
        "neutralLighter"       = "#353a49";
        "neutralLight"         = "#404759";
        "neutralQuaternaryAlt" = "#474e62";
        "neutralQuaternary"    = "#4c546a";
        "neutralTertiaryAlt"   = "#646e8a";
        "neutralTertiary"      = "#c8c8c8";
        "neutralSecondaryAlt"  = "#d0d0d0";
        "neutralSecondary"     = "#dadada";
        "neutralPrimaryAlt"    = "#eaeaea";
        "neutralPrimary"       = "#ffffff";
        "neutralDark"          = "#f4f4f4";
        "black"                = "#f8f8f8";
        "white"                = "#262a35";
        "primaryBackground"    = "#262a35";
        "primaryText"          = "#ffffff";
        "error"                = "#ff5f5f";
        "accent"               = "#3a96dd";
    };
}
#endregion

# Function to write log entries
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Output to console with color coding
    switch ($Level) {
        "INFO" { Write-Host $logEntry -ForegroundColor Green }
        "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
        "ERROR" { Write-Host $logEntry -ForegroundColor Red }
    }
    
    # Write to log file
    Add-Content -Path $Config.LogPath -Value $logEntry
}

# Function to determine SharePoint geo-location and connect
function Connect-ToSharePointGeo {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl
    )
    
    # Extract the domain from the URL to determine geo-location
    $domain = ([System.Uri]$SiteUrl).Host
    
    Write-Log "Determining geo-location for domain: $domain"
    
    # Define geo-specific connection logic
    # Modify these patterns according to your multi-geo setup
    if ($domain -like "*.nam.*" -or $domain -like "*-nam.*") {
        Write-Log "Detected North America geo-location"
    }
    elseif ($domain -like "*.eur.*" -or $domain -like "*-eur.*") {
        Write-Log "Detected Europe geo-location"
    }
    elseif ($domain -like "*.apac.*" -or $domain -like "*-apac.*") {
        Write-Log "Detected Asia-Pacific geo-location"
    }
    else {
        # Default to global tenant
        Write-Log "Using default geo-location"
    }
    
    try {
        # Connect to SharePoint Online using PnP PowerShell with certificate thumbprint
        Write-Log "Connecting to SharePoint site: $SiteUrl"
        
        # Connect using App-Only authentication with certificate thumbprint
        Write-Log "Connecting with app registration using certificate thumbprint"
        Connect-PnPOnline -Url $SiteUrl -ClientId $Config.AppId -Thumbprint $Config.Thumbprint -Tenant $Config.TenantId
        
        Write-Log "Successfully connected to SharePoint site: $SiteUrl"
        return $true
    }
    catch {
        Write-Log "Failed to connect to $SiteUrl. Error: $_" -Level "ERROR"
        return $false
    }
}

# Function to update site branding
function Update-SiteBranding {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl
    )
    
    try {
        # Check if we need to process anything
        if (-not $Config.ChangeLogoImage -and -not $Config.ApplyThemeColors) {
            Write-Log "Both ChangeLogoImage and ApplyThemeColors are set to false. No branding changes will be made." -Level "WARNING"
            return $true
        }
        
        # Process logo if enabled
        if ($Config.ChangeLogoImage) {
            # Check if logo file exists
            if (-not (Test-Path -Path $Config.LogoPath)) {
                Write-Log "Logo file not found: $($Config.LogoPath)" -Level "ERROR"
                return $false
            }

            # Upload logo to site assets library
            Write-Log "Uploading logo from $($Config.LogoPath) to site assets"
            $assetLibrary = "SiteAssets"
            
            # Ensure SiteAssets library exists
            $web = Get-PnPWeb
            $assetLibraryExists = Get-PnPList -Identity $assetLibrary -ErrorAction SilentlyContinue
            
            if (-not $assetLibraryExists) {
                Write-Log "Creating SiteAssets library" -Level "INFO"
                New-PnPList -Title $assetLibrary -Template DocumentLibrary -OnQuickLaunch
            }
            
            # Upload the logo file
            $logoFile = Add-PnPFile -Path $Config.LogoPath -Folder $assetLibrary -ErrorAction Stop
            
            # Get the URL of the uploaded file
            $logoUrl = $logoFile.ServerRelativeUrl
            Write-Log "Logo uploaded successfully to $logoUrl"
            
            # Update site logo
            Write-Log "Setting site logo to $logoUrl"
            
            # Try using newer cmdlet format first, then fall back to older version if needed
            try {
                # For modern PnP.PowerShell module (newer versions)
                Set-PnPWeb -SiteLogoUrl $logoUrl
                Write-Log "Successfully set site logo using Set-PnPWeb cmdlet"
            }
            catch {
                try {
                    # For older PnP PowerShell module versions
                    Write-Log "Falling back to older cmdlet format..." -Level "INFO"
                    Set-PnPSite -LogoUrl $logoUrl
                    Write-Log "Successfully set site logo using Set-PnPSite cmdlet"
                }
                catch {
                    Write-Log "Unable to set site logo. Error: $_" -Level "ERROR"
                    return $false
                }
            }
        }
        else {
            Write-Log "ChangeLogoImage is set to false, skipping logo update" -Level "INFO"
        }
        
        # Apply theme if specified in config
        if ($null -ne $Config.Theme) {
            # Check if theme colors should be applied
            if (-not $Config.ApplyThemeColors) {
                Write-Log "ApplyThemeColors is set to false, skipping theme color application" -Level "INFO"
                return $true
            }

            # First, update colors based on selected theme before creating the theme
            if (-not [string]::IsNullOrEmpty($Config.ColorTheme)) {
                $selectedTheme = $Config.ColorTheme
                Write-Log "Using color theme: $selectedTheme"
                
                # Apply the selected theme colors
                if ($ColorThemes.ContainsKey($selectedTheme)) {
                    Write-Log "Applying $selectedTheme theme colors"
                    $Config.Theme.Colors = $ColorThemes[$selectedTheme]
                }
                else {
                    Write-Log "Unknown theme '$selectedTheme'. Valid options are: Custom, Teal, Red, Orange, Green, Blue, Purple, Gray, Periwinkle, DarkYellow, or DarkBlue. Using Green theme as fallback." -Level "WARNING"
                    $Config.Theme.Colors = $ColorThemes["Green"]
                }
            }
        
            # Then, connect to admin site to manage themes
            try {
                $adminUrl = "https://$($Config.TenantName)-admin.sharepoint.com"
                Write-Log "Connecting to admin site $adminUrl to manage themes..."
                
                # Store current connection information instead of the connection object
                Write-Log "Storing current site URL for reconnection later..."
                $originalSiteUrl = $SiteUrl
                
                # Connect to admin site
                try {
                    Connect-PnPOnline -Url $adminUrl -ClientId $Config.AppId -Thumbprint $Config.Thumbprint -Tenant $Config.TenantId
                    Write-Log "Successfully connected to admin site"
                }
                catch {
                    Write-Log "Failed to connect to admin site: $_" -Level "ERROR"
                    return $false
                }
                
                # Check if theme exists and handle theme management using direct SharePoint REST API calls
                # This avoids PnP parameter name inconsistencies
                try {
                    $themeName = $Config.Theme.Name
                    Write-Log "Managing theme '$themeName' using REST API approach..."
                    
                    # Create theme JSON definition
                    $themeJson = @{
                        "name"       = $themeName
                        "palette"    = $Config.Theme.Colors
                        "isInverted" = $false
                    } | ConvertTo-Json -Depth 10
                    
                    # Output the theme JSON for verification
                    Write-Log "Theme definition JSON to be applied:"
                    Write-Log $themeJson
                    
                    # Get the client context and SharePoint client object models
                    $context = Get-PnPContext
                    $web = $context.Web
                    $context.Load($web)
                    Invoke-PnPQuery
                    
                    # Create a tenant admin client context
                    $tenant = New-Object Microsoft.Online.SharePoint.TenantAdministration.Tenant($context)
                    
                    Write-Log "Removing any existing theme with the same name..."
                    try {
                        # Attempt to remove the theme if it exists (we don't check first, just try to remove and handle errors)
                        $tenant.DeleteTenantTheme($themeName)
                        $context.ExecuteQuery()
                        Write-Log "Existing theme removed successfully"
                    }
                    catch {
                        # Theme might not exist, which is fine
                        Write-Log "No existing theme found or unable to remove: $_" -Level "INFO"
                    }
                    
                    Write-Log "Creating new theme..."
                    try {
                        # Add the new theme
                        $tenant.AddTenantTheme($themeName, $themeJson)
                        $context.ExecuteQuery()
                        Write-Log "Theme created successfully"
                    }
                    catch {
                        Write-Log "Failed to create theme using client object model: $_" -Level "ERROR"
                        throw
                    }
                
                    # Reconnect to original site
                    try {
                        Write-Log "Reconnecting to original site ($originalSiteUrl)..."
                        Connect-PnPOnline -Url $originalSiteUrl -ClientId $Config.AppId -Thumbprint $Config.Thumbprint -Tenant $Config.TenantId
                        Write-Log "Successfully reconnected to original site"
                    }
                    catch {
                        Write-Log "Failed to reconnect to original site: $_" -Level "ERROR"
                        return $false
                    }
                
                    # Apply theme to current site using direct REST API call
                    try {
                        Write-Log "Applying theme '$themeName' to site using REST API..."
                    
                        # Get the web URL for the absolute path
                        $web = Get-PnPWeb
                        $webUrl = $web.Url
                        
                        # Use a simpler approach with Set-PnPWebTheme
                        try {
                            Write-Log "Trying Set-PnPWebTheme cmdlet..."
                            Set-PnPWebTheme -Theme $themeName
                            Write-Log "Successfully applied theme using Set-PnPWebTheme"
                        }
                        catch {
                            Write-Log "Set-PnPWebTheme failed, trying alternative approach: $_" -Level "WARNING"
                            
                            # Alternative approach using direct REST API
                            try {
                                # Format the REST endpoint correctly
                                $restBody = "{'name':'$themeName'}"
                                $restEndpoint = "$webUrl/_api/web/ApplyTheme"
                                
                                # Execute the REST call with proper parameters
                                Invoke-PnPSPRestMethod -Method Post -Url $restEndpoint -Content $restBody -ContentType "application/json;odata=verbose" -Headers @{"X-RequestDigest" = (Get-PnPRequestDigest) }
                                Write-Log "Successfully applied theme using REST API"
                            }
                            catch {
                                Write-Log "REST API approach failed: $_" -Level "ERROR"
                                # Continue execution even if theme application fails
                            }
                        }
                    }
                    catch {
                        Write-Log "Failed to apply theme: $_" -Level "ERROR"
                        # Continue execution
                    }
                }
                catch {
                    Write-Log "Error managing or applying theme: $_" -Level "ERROR"
                    # Don't return false here, as we want to continue even if theme application fails
                }
            }
            catch {
                Write-Log "Failed to connect to admin site or manage themes: $_" -Level "ERROR"
                # Continue execution even if theme management fails
            }
        }
        else {
            Write-Log "Theme configuration is null, skipping theme application"
        }
        
        Write-Log "Successfully updated branding for $SiteUrl"
        return $true
    }
    catch {
        Write-Log "Failed to update branding for $SiteUrl. Error: $_" -Level "ERROR"
        return $false
    }
}

# Function to get all subsites of a site collection recursively
function Get-AllSubsites {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl
    )
    
    try {
        Write-Log "Getting subsites for $SiteUrl" -Level "INFO"
        
        # Create a list to store all subsites
        $allSubsites = New-Object System.Collections.Generic.List[string]
        
        # Connect to the site collection
        if (Connect-ToSharePointGeo -SiteUrl $SiteUrl) {
            # Get all immediate subsites
            $subsites = Get-PnPSubWeb -Recurse -IncludeRootWeb:$false -ErrorAction SilentlyContinue
            
            if ($subsites -and $subsites.Count -gt 0) {
                Write-Log "Found $($subsites.Count) subsites under $SiteUrl" -Level "INFO"
                
                # Add each subsite URL to the list
                foreach ($subsite in $subsites) {
                    $allSubsites.Add($subsite.Url)
                }
            }
            else {
                Write-Log "No subsites found under $SiteUrl" -Level "INFO"
            }
            
            # Disconnect from the site collection
            Disconnect-PnPOnline
            Write-Log "Disconnected from $SiteUrl"
        }
        
        return $allSubsites
    }
    catch {
        Write-Log "Error getting subsites for $SiteUrl. Error: $_" -Level "ERROR"
        return @()
    }
}

# Main script execution
try {
    # Set execution policy for current process if needed
    if ((Get-ExecutionPolicy -Scope Process) -eq "Restricted") {
        Write-Host "Setting execution policy to Bypass for current process..."
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
    }

    # Check if PnP PowerShell module is installed
    if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
        Write-Host "PnP.PowerShell module is not installed. Would you like to install it now? (Y/N)" -ForegroundColor Yellow
        $response = Read-Host
        if ($response -eq "Y" -or $response -eq "y") {
            Write-Host "Installing PnP.PowerShell module..." -ForegroundColor Green
            Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
            Import-Module PnP.PowerShell
        }
        else {
            Write-Log "PnP.PowerShell module is required for this script. Please install it using: Install-Module -Name PnP.PowerShell" -Level "ERROR"
            exit 1
        }
    }
    
    # Initialize log file
    $logHeader = "=== SharePoint Branding Update Script Started at $(Get-Date) ==="
    Set-Content -Path $Config.LogPath -Value $logHeader
    Write-Log "Script started with configuration: CSV=$($Config.CsvPath), Logo=$($Config.LogoPath), ChangeLogoImage=$($Config.ChangeLogoImage), ApplyThemeColors=$($Config.ApplyThemeColors), ProcessSubsites=$($Config.ProcessSubsites), Theme=$($Config.Theme.Name)"
    
    # Check if CSV file exists
    if (-not (Test-Path -Path $Config.CsvPath)) {
        Write-Log "CSV file not found: $($Config.CsvPath)" -Level "ERROR"
        Write-Log "Please create a CSV file with a column named 'URL' containing SharePoint site URLs" -Level "ERROR"
        exit 1
    }
    
    # Import CSV file
    Write-Log "Importing site URLs from $($Config.CsvPath)"
    $sites = Import-Csv -Path $Config.CsvPath
    
    if (-not $sites -or $sites.Count -eq 0) {
        Write-Log "No sites found in the CSV file or the file format is incorrect" -Level "WARNING"
        exit 1
    }
    
    # Verify Entra App Registration details
    if ([string]::IsNullOrEmpty($Config.AppId) -or $Config.AppId -eq "12345678-1234-1234-1234-1234567890ab") {
        Write-Log "Please update the AppId in the configuration section with your actual app registration ID" -Level "WARNING"
    }
    
    if ([string]::IsNullOrEmpty($Config.Thumbprint) -or $Config.Thumbprint -eq "1234567890ABCDEF1234567890ABCDEF12345678") {
        Write-Log "Please update the Thumbprint in the configuration section with your actual certificate thumbprint" -Level "WARNING"
    }
    
    if ([string]::IsNullOrEmpty($Config.TenantId) -or $Config.TenantId -eq "12345678-1234-1234-1234-1234567890ab") {
        Write-Log "Please update the TenantId in the configuration section with your actual tenant ID" -Level "WARNING"
    }
    
    Write-Log "Using app registration authentication with AppID: $($Config.AppId)"
    Write-Log "Found $($sites.Count) sites to process"
    
    # Process each site
    $successCount = 0
    $failureCount = 0
    $totalSiteCount = 0
    $processedSiteCollections = 0
    $processedSubsites = 0
    
    foreach ($site in $sites) {
        # Ensure the CSV has a URL column
        if (-not $site.URL) {
            Write-Log "CSV file must contain a 'URL' column. Please check your CSV format." -Level "ERROR"
            exit 1
        }
        
        $siteUrl = $site.URL.Trim()
        $totalSiteCount++
        Write-Log "Processing site collection: $siteUrl"
        
        # Create a list to hold all sites to process (site collection + subsites if enabled)
        $sitesToProcess = New-Object System.Collections.Generic.List[string]
        $sitesToProcess.Add($siteUrl)
        
        # Get subsites if enabled
        if ($Config.ProcessSubsites) {
            Write-Log "Subsite processing is enabled. Getting all subsites for $siteUrl..." -Level "INFO"
            $subsites = Get-AllSubsites -SiteUrl $siteUrl
            
            if ($subsites -and $subsites.Count -gt 0) {
                Write-Log "Adding $($subsites.Count) subsites to processing queue" -Level "INFO"
                foreach ($subsite in $subsites) {
                    $sitesToProcess.Add($subsite)
                    $totalSiteCount++
                }
            }
        }
        
        # Process the site collection and all its subsites
        foreach ($currentSite in $sitesToProcess) {
            Write-Log "Processing site: $currentSite" -Level "INFO"
            
            # Determine if this is a site collection or subsite for logging
            $isSiteCollection = ($currentSite -eq $siteUrl)
            
            # Connect to the site based on geo-location
            if (Connect-ToSharePointGeo -SiteUrl $currentSite) {
                # Update site branding
                if (Update-SiteBranding -SiteUrl $currentSite) {
                    $successCount++
                    if ($isSiteCollection) {
                        $processedSiteCollections++
                    }
                    else {
                        $processedSubsites++
                    }
                }
                else {
                    $failureCount++
                }
                
                # Disconnect from the site
                Disconnect-PnPOnline
                Write-Log "Disconnected from $currentSite"
            }
            else {
                $failureCount++
            }
        }
    }
    
    # Summary
    Write-Log "Script completed. Processed $totalSiteCount total sites ($processedSiteCollections site collections and $processedSubsites subsites)."
    Write-Log "Successfully processed $successCount sites. Failed to process $failureCount sites."
    Write-Log "=== SharePoint Branding Update Script Completed at $(Get-Date) ==="
}
catch {
    Write-Log "An unexpected error occurred: $_" -Level "ERROR"
}
finally {
    # Ensure we're disconnected from any SharePoint sites
    try {
        Disconnect-PnPOnline
    }
    catch {}
}

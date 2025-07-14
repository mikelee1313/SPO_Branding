# SharePoint Branding Deployment Script

This PowerShell script automates the deployment of consistent branding (logo and theme) across multiple SharePoint Online sites, including multi-geo environments. The script offers flexible options for controlling which branding elements to update and whether to include subsites.

## Features

- Deploys consistent logo and theme to multiple SharePoint sites
- Selectively apply logo updates, theme colors, or both
- Choose from 11 predefined SharePoint color themes
- Process site collections and their subsites
- Supports multi-geo SharePoint environments
- Uses secure authentication via Entra ID App Registration with certificate
- Detailed logging of all operations
- Completely self-contained - no additional configuration files needed

## Prerequisites

1. PowerShell 5.1 or later
2. PnP.PowerShell module (script will prompt to install if not present)
3. Entra ID App Registration with appropriate permissions
4. Certificate for authentication (referenced by thumbprint)
5. CSV file with site URLs (see sample-sites.csv for format)

## Setup Instructions

1. **Configure Entra ID App Registration:**
   - Register a new app in Entra ID
   - Add API permissions for SharePoint (Sites.FullControl.All)
   - Create and upload a certificate
   - Note the App ID, Tenant ID, and certificate thumbprint

2. **Update Configuration:**
   - Modify the `$Config` hashtable at the top of the script
   - Set your CSV file path
   - Set the path to your logo image file
   - Configure branding options (ChangeLogoImage, ApplyThemeColors, ProcessSubsites)
   - Choose a color theme from the predefined options
   - Enter your Entra ID App details (TenantName, AppId, Thumbprint, TenantId)

3. **Prepare Site List:**
   - Create a CSV file with a column header named "URL"
   - Add SharePoint site URLs, one per line

## Usage

Simply run the script:

```powershell
.\Deploy-Branding.ps1
```

The script is completely self-contained and will:
- Verify prerequisites
- Check execution policy and adjust if needed
- Connect to each site using secure app-only authentication
- Process site collections and subsites (if enabled)
- Apply the specified logo and/or theme based on configuration
- Generate a detailed log file

## Configuration

The script uses a configuration section at the top for easy customization:

```powershell
$Config = @{
    # Path to the CSV file containing SharePoint site URLs (must have a URL column)
    CsvPath          = "C:\temp\sitelist.csv"
    
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
    # Teal, Red, Orange, Green, Blue, Purple, Gray, Periwinkle, DarkYellow, DarkBlue, Custom
    ColorTheme       = "Blue"
    
    # Theme configuration
    Theme            = @{
        Name      = "CompanyTheme"
        # Colors will be populated from the selected ColorTheme
        Colors    = @{}
        Overwrite = $true
    }
    
    # Entra App Registration details
    TenantName       = "contoso"
    AppId            = "12345678-1234-1234-1234-1234567890ab"
    Thumbprint       = "1234567890ABCDEF1234567890ABCDEF12345678"
    TenantId         = "12345678-1234-1234-1234-1234567890ab"
}
```

## Branding Options

The script provides flexible options for controlling which branding elements to update:

### Logo Options
- **ChangeLogoImage**: When set to `$true`, the script will upload and set the logo specified in `LogoPath`
- If set to `$false`, the script will skip logo updates entirely

### Theme Options
- **ApplyThemeColors**: When set to `$true`, the script will apply the color theme specified in `ColorTheme`
- If set to `$false`, the script will skip theme color updates entirely

### Site Processing Options
- **ProcessSubsites**: When set to `$true`, the script will process both site collections and all their subsites
- If set to `$false`, the script will only process the site collections listed in the CSV file

This flexibility allows you to:
1. Update only logos across sites
2. Update only theme colors across sites
3. Update both logos and theme colors
4. Apply branding to top-level site collections only or include all subsites

## Available Color Themes

The script includes 11 predefined SharePoint color themes:

1. **Custom** - A customizable theme you can modify to match your exact brand colors
2. **Teal** - A modern teal-based theme
3. **Red** - A vibrant red-based theme
4. **Orange** - A warm orange-based theme
5. **Green** - A fresh green-based theme
6. **Blue** - A professional blue-based theme
7. **Purple** - A rich purple-based theme
8. **Gray** - A neutral gray-based theme
9. **Periwinkle** - A soft periwinkle-based theme
10. **DarkYellow** - A dark theme with yellow accents
11. **DarkBlue** - A dark theme with blue accents

To select a theme, set the `ColorTheme` parameter in the configuration to one of these names. If you need custom colors, select "Custom" and modify the color definitions in the `$ColorThemes` hashtable.

## Sample CSV Format

```
URL
https://contoso.sharepoint.com/sites/site1
https://contoso.sharepoint.com/sites/site2
https://contoso-eur.sharepoint.com/sites/europe1
```

When `ProcessSubsites` is enabled, the script will automatically discover and process all subsites under each URL in this file.

## Authentication

The script uses certificate thumbprint authentication with an Entra ID App Registration:

### Entra ID App Registration Setup

1. Create an Entra ID App Registration in the Azure Portal
2. Generate a certificate and upload it to the app registration
3. Note the certificate thumbprint
4. Grant the app the necessary SharePoint permissions:
   - SharePoint: `Sites.FullControl.All`
5. Configure the app details in the script configuration section

This approach eliminates the need for user credentials or interactive login, providing a secure and automated solution for branding deployment.

## Logging

The script generates detailed logs that capture all actions and any errors encountered during execution. The log file is created in the system's temp directory by default, with a timestamp in the filename for easy identification. The logs include:

- Full configuration details including all branding options
- Connection status for each site
- Logo upload and application status
- Theme creation and application status
- Subsite discovery and processing details
- Summary of successful and failed site updates
- Separate counts for site collections and subsites processed

## Error Handling

The script includes robust error handling to ensure it continues processing sites even if some encounter errors. A summary of successful and failed site updates is provided at the end of execution.

## Troubleshooting

- Check the log file for detailed error messages
- Verify Entra ID App has proper permissions (Sites.FullControl.All)
- Ensure your certificate is valid and installed correctly
- Confirm the logo file exists at the specified path
- Make sure your selected color theme is properly defined
- If subsites aren't being processed, verify the ProcessSubsites setting is $true
- Check that the Entra ID App has permissions for both site collections and subsites

## License

This script is provided as-is with no warranties. Use at your own risk.

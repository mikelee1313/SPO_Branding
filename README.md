# SharePoint Branding Deployment Script

This PowerShell script automates the deployment of consistent branding (logo and theme) across multiple SharePoint Online sites, including multi-geo environments. The script offers flexible options for controlling which branding elements to update and whether to include subsites.

## Features

- Deploys consistent logo and theme to multiple SharePoint sites
- Selectively apply logo updates, theme colors, or both
- Choose from 11 predefined SharePoint color themes
- Process site collections and their subsites
- Supports multi-geo SharePoint environments
- Secure certificate-based authentication
- Smart throttling protection with exponential backoff
- Detailed logging of all operations
- Completely self-contained - no additional configuration files needed

## Prerequisites

1. PowerShell 5.1 or later
2. PnP.PowerShell module (script will prompt to install if not present)
3. Entra ID App Registration with appropriate permissions
4. Certificate for certificate-based authentication
5. CSV file with site URLs (see sample-sites.csv for format)

## Setup Instructions

1. **Configure Entra ID App Registration:**
   - Register a new app in Entra ID
   - Add API permissions for SharePoint (Sites.FullControl.All)
   - Either:
     - Create and upload a certificate for certificate-based authentication
   - Note the App ID, Tenant ID, and certificate thumbprint

2. **Update Configuration:**
   - Modify the `$Config` hashtable at the top of the script
   - Set your CSV file path
   - Set the path to your logo image file
   - Configure branding options (ChangeLogoImage, ApplyThemeColors, ProcessSubsites)
   - Choose a color theme from the predefined options
   - Enter your Entra ID App details (TenantName, AppId, TenantId)
   - Enter your certificate thumbprint
   - Configure throttling settings (MaxRetries, RetryInitialWait, RetryFactor)

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
- Connect to each site using secure app-only authentication with your chosen method
- Implement smart throttling protection with exponential backoff
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
    
    # Whether to apply an existing theme instead of creating a new one
    # When set to $true, the script will look for the theme specified in Theme.Name
    # and apply it without creating a new theme
    ApplyExistingTheme = $false
    
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
    TenantId         = "12345678-1234-1234-1234-1234567890ab"
    
    # Certificate Thumbprint for authentication
    Thumbprint       = "1234567890ABCDEF1234567890ABCDEF12345678"
    
    # Throttling settings
    MaxRetries       = 5           # Maximum number of retry attempts for throttled requests
    RetryInitialWait = 2           # Initial wait time in seconds before first retry
    RetryFactor      = 2           # Multiplicative factor for subsequent retry waits (exponential backoff)
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

### Theme Application Options
- **ApplyExistingTheme**: When set to `$true`, the script will look for a theme with the name specified in `Theme.Name` 
  and apply it without creating a new theme
- When set to `$false` (default), the script will create a new theme with the colors specified by `ColorTheme`
- This is useful when you already have a published theme in your tenant that you want to apply to multiple sites

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

The script supports two authentication methods via Entra ID App Registration:

1. **Certificate-based Authentication**
   - Set `AuthType = "Certificate"` in the configuration
   - Provide the certificate thumbprint in the `Thumbprint` parameter
   - The certificate must be installed in the certificate store of the machine running the script
   - More secure method, recommended for production environments

2. **Client Secret Authentication**
   - Set `AuthType = "ClientSecret"` in the configuration
   - Provide the client secret in the `ClientSecret` parameter
   - Simpler to set up but less secure than certificate-based authentication

### Entra ID App Registration Setup

1. Create an Entra ID App Registration in the Azure Portal
2. Grant the app the necessary SharePoint permissions:
   - SharePoint: `Sites.FullControl.All`
3. Choose your authentication method:
   - For certificate: Generate a certificate and upload it to the app registration, note the thumbprint
   - For client secret: Create a new client secret and save the value securely
4. Configure the app details and your chosen authentication method in the script configuration section

This approach eliminates the need for user credentials or interactive login, providing a secure and automated solution for branding deployment.

## Throttling Protection

The script implements smart throttling protection with exponential backoff to handle SharePoint API throttling gracefully:

- **MaxRetries**: Maximum number of retry attempts for throttled requests (default: 5)
- **RetryInitialWait**: Initial wait time in seconds before first retry (default: 2)
- **RetryFactor**: Multiplicative factor for subsequent retry waits (default: 2)

With these settings, if a request is throttled:
1. First retry will wait 2 seconds
2. Second retry will wait 4 seconds (2 × 2)
3. Third retry will wait 8 seconds (4 × 2)
4. Fourth retry will wait 16 seconds (8 × 2)
5. Fifth retry will wait 32 seconds (16 × 2)

This exponential backoff strategy helps the script gracefully handle throttling situations when deploying branding to many sites or in environments with high API traffic.

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
- For certificate authentication:
  - Ensure your certificate is valid and installed correctly on the machine running the script
  - Verify the thumbprint is correct in the configuration
- For client secret authentication:
  - Ensure the client secret is valid and has not expired
  - Check that the client secret value is correctly entered in the configuration
- Confirm the logo file exists at the specified path
- Make sure your selected color theme is properly defined
- If subsites aren't being processed, verify the ProcessSubsites setting is $true
- Check that the Entra ID App has permissions for both site collections and subsites
- If experiencing throttling issues:
  - Check logs for throttling patterns
  - Consider increasing RetryInitialWait or MaxRetries in high-traffic environments
  - Try running the script during off-peak hours

## License

This script is provided as-is with no warranties. Use at your own risk.

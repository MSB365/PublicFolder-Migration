
# Exchange Public Folder Migration Script

A comprehensive PowerShell script for migrating Exchange Public Folders from on-premises Exchange Server to Exchange Online with interactive discovery, confirmation, and detailed HTML reporting.

## 🚀 Overview

This business-critical script provides a complete solution for Exchange Public Folder migration, featuring:

- **Automated Discovery**: Identifies all Public Folders on Exchange on-premises servers
- **Interactive Interface**: User-friendly prompts and confirmations
- **Progress Tracking**: Real-time progress indicators and status updates
- **Comprehensive Reporting**: Detailed HTML reports with migration results
- **Error Handling**: Robust error handling and logging throughout the process
- **Exchange Online Integration**: Seamless connection to Exchange Online

## 📋 Prerequisites

### Required Software
- **Exchange Management Shell** (installed on Exchange Server or management workstation)
- **Exchange Online PowerShell V2 Module**
- **PowerShell 5.1** or later
- **.NET Framework 4.7.2** or later

### Required Permissions
- **Exchange Organization Management** permissions on on-premises Exchange
- **Exchange Online Administrator** or **Global Administrator** in Microsoft 365
- **Network connectivity** to Exchange Online

### Installation Commands
```powershell
# Install Exchange Online Management module
Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber

# Verify Exchange Management Shell is available
Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue
```

## ️ Installation

1. **Download the script** to your Exchange management workstation
2. **Place the script** in a directory accessible from Exchange Management Shell
3. **Ensure execution policy** allows script execution:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

```




## Usage

### Basic Usage

```powershell
# Run from Exchange Management Shell
.\Exchange-PublicFolder-Migration.ps1
```

### Step-by-Step Process

1. **Launch Exchange Management Shell** as Administrator
2. **Navigate** to the script directory
3. **Execute the script**:

```powershell
.\Exchange-PublicFolder-Migration.ps1

```


4. **Follow the interactive prompts**:

1. Review discovered Public Folders
2. Confirm migration parameters
3. Authenticate to Exchange Online
4. Monitor migration progress
5. Save the HTML report





## Features

### Public Folder Discovery

- Automatically discovers all Public Folders
- Retrieves folder statistics (item count, size, permissions)
- Displays hierarchical folder structure
- Identifies mail-enabled folders


### Interactive Dashboard

- Real-time progress indicators
- Colored console output for different message types
- Summary statistics and folder hierarchy display
- User confirmation prompts with detailed warnings


### Migration Process

- Batch creation with timestamp naming
- Individual folder migration tracking
- Success/failure status for each folder
- Comprehensive error logging


### HTML Reporting

- Professional, responsive HTML reports
- Migration statistics and summaries
- Detailed folder inventory tables
- Complete migration logs
- Mobile-friendly design


### ️ Error Handling

- Comprehensive try-catch blocks
- Detailed error logging with timestamps
- Graceful failure handling
- Connection validation and retry logic


## ️ Configuration

### Script Variables

```powershell
$Script:LogLevel = "INFO"           # Logging level (INFO, WARNING, ERROR, SUCCESS)
$Script:BatchPrefix = "PF-Migration" # Migration batch naming prefix
$Script:ScriptVersion = "1.0.0"    # Script version for reporting
```

### Customization Options

- **Batch Naming**: Modify `$Script:BatchPrefix` for custom batch names
- **Logging Level**: Adjust `$Script:LogLevel` for different verbosity levels
- **Report Styling**: Customize HTML report CSS in the `Generate-HTMLReport` function


## Report Features

The generated HTML report includes:

### Summary Dashboard

- Total folders discovered
- Successful migrations count
- Failed migrations count
- Total migration duration


### Detailed Tables

- **Migration Results**: Status, items migrated, timestamps, error messages
- **Folder Inventory**: Names, paths, item counts, sizes, mail-enabled status
- **Migration Log**: Complete chronological log with color-coded entries


### Visual Elements

- Responsive design for desktop and mobile
- Color-coded status indicators
- Professional styling with gradients and shadows
- Hover effects and interactive elements


## Troubleshooting

### Common Issues

#### Exchange Management Shell Connection

```powershell
# Verify Exchange Management Shell
Get-ExchangeServer | Select-Object Name, ServerRole
```

#### Exchange Online Connection

```powershell
# Test Exchange Online connectivity
Get-OrganizationConfig | Select-Object DisplayName
```

#### Module Installation Issues

```powershell
# Install with elevated privileges
Install-Module -Name ExchangeOnlineManagement -Force -Scope AllUsers
```

### Error Messages

| Error | Solution
|-----|-----
| "Exchange Management Shell connection failed" | Ensure script is run from Exchange Management Shell
| "ExchangeOnlineManagement module not found" | Install the module using `Install-Module`
| "Failed to connect to Exchange Online" | Verify credentials and network connectivity
| "No Public Folders found" | Check permissions and Public Folder existence


## Examples

### Example Output

```plaintext
================================================================================
Exchange Public Folder Migration Script v1.0.0
================================================================================

[2024-01-15 10:30:15] [INFO] Testing Exchange Management Shell connection...
[2024-01-15 10:30:16] [SUCCESS] Successfully connected to Exchange Server: EX01
[2024-01-15 10:30:16] [INFO] Starting Public Folder discovery...
[2024-01-15 10:30:18] [SUCCESS] Found 25 Public Folders. Gathering statistics...

================================================================================
PUBLIC FOLDER MIGRATION SUMMARY
================================================================================

SUMMARY STATISTICS:
  Total Folders Found: 25
  Total Items: 1,247
  Mail-Enabled Folders: 3
  Discovery Time: 1/15/2024 10:30:20 AM
```

### Sample Migration Batch

```plaintext
Migration Batch: PF-Migration-20240115-103020
- Company Announcements (Success: 45 items)
- HR Documents (Success: 123 items)
- Project Files (Success: 89 items)
- Sales Reports (Failed: Connection timeout)
```

## Security Considerations

- **Credentials**: Script uses interactive authentication for Exchange Online
- **Permissions**: Requires high-level Exchange permissions
- **Network**: Ensure secure network connection for migration
- **Logging**: Review logs for sensitive information before sharing


## Support

### Before Contacting Support

1. **Check Prerequisites**: Ensure all requirements are met
2. **Review Logs**: Check the generated HTML report for detailed error information
3. **Test Connectivity**: Verify Exchange and Exchange Online connections
4. **Check Permissions**: Ensure proper administrative permissions


### Support Information

- **Script Version**: 1.0.0
- **Author**: Exchange Migration Team
- **PowerShell Version**: 5.1+
- **Exchange Compatibility**: Exchange 2013, 2016, 2019


## License

This script is provided as-is for Exchange Public Folder migration purposes. Please test thoroughly in a non-production environment before using in production.



---

**⚠️ Important**: Always test migration scripts in a non-production environment first. Ensure you have proper backups before performing any migration operations.


# Exchange Public Folder Migration Script

A comprehensive PowerShell script for migrating Exchange on-premises Public Folders to Exchange Online with interactive user interface, progress tracking, and detailed HTML reporting.

## Features

- **Automated Discovery**: Identifies all available Public Folders on Exchange on-premises
- **Interactive Interface**: User-friendly prompts in Exchange Management Shell
- **Progress Tracking**: Real-time status updates during migration process
- **Exchange Online Integration**: Seamless connection and migration to Office 365
- **Professional Reporting**: Beautiful HTML reports with download functionality
- **Business-Critical Quality**: Comprehensive error handling and logging

## Prerequisites

### Required Software
- **Exchange Management Shell** (Exchange Server 2010/2013/2016/2019)
- **PowerShell 5.1** or later
- **ExchangeOnlineManagement Module** (auto-installed if missing)
- **.NET Framework 4.7.2** or later

### Required Permissions
- **Exchange Organization Management** role on-premises
- **Exchange Administrator** role in Exchange Online
- **Local Administrator** rights on the machine running the script

### Network Requirements
- Connectivity to Exchange on-premises server
- Internet access for Exchange Online connection
- Firewall ports 443 (HTTPS) and 80 (HTTP) open

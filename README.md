# AD Export Tool

A PowerShell script providing a graphical user interface (GUI) for exporting Active Directory (AD) user data to CSV or Excel formats.

## Prerequisites

- PowerShell 5.1 or higher
- Active Directory PowerShell module
- Windows Server: `Add-WindowsFeature RSAT-AD-PowerShell`
- Windows 10/11: `Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"`
- Excel installed for COM Object if exporting to Excel

## Configuration

Create a `config.json` file in the same directory as the script with the following structure:

```json
{
    "DC": [
        "domain-controller1.domain.com",
        "domain-controller2.domain.com"
    ]
}
```

## Features

- Export AD user data to CSV or Excel formats
- Customizable property selection for export
- Multiple domain controller support
- Flexible export options and formatting
- Batch search functionality for multiple users

## Usage

1. Configure domain controllers in config.json
2. Run the script:

```ps
.\ADTool.ps1
```
3. For batch search, create a text file with one username per line:

```txt
jsmith\
jane.doe\
user1@domain.com\
user2@domain.com
```

Both SAMAccountName (jsmith) and UserPrincipalName (user@domain.com) formats are supported.

## Notes

- The script creates a log file (ADExport.log) for troubleshooting
- Failed batch exports are logged separately
- Supports both single and batch user exports
- Export formats: Excel (.xlsx) and CSV (.csv)
# Azure File Share to Blob Storage Archival Script

A comprehensive PowerShell script for archiving files from Azure File Share to Azure Blob Storage based on age criteria. Designed for Azure Automation Account with Managed Identity authentication.

## 🚀 Features

- **Smart Age-Based Archival**: Archive files older than specified years using NTFS LastWriteTime
- **Server-Side Copy**: Efficient data transfer without local download
- **Comprehensive Verification**: Verify blob copies before any deletion
- **Stub File Creation**: Mark archived files with stub files to prevent re-processing
- **Batch Processing**: Handle large datasets efficiently
- **Single File Testing**: Debug and test individual files
- **Folder Path Selection**: Target specific folders for archival operations
- **Blob Tier Optimization**: Choose storage tiers (Hot/Cool/Cold/Archive) for cost optimization
- **Azure Automation Ready**: Optimized for Azure Automation Account
- **Dual Runtime Support**: Works with both PowerShell 5.1 and 7.2

## 📋 Prerequisites

- Azure Storage Account with File Share and Blob Container
- Azure Automation Account with Managed Identity
- Required PowerShell modules:
  - `Az.Storage`
  - `Az.Accounts` 
  - `Az.Resources`

## 🔧 Required Permissions

The Managed Identity needs these roles:
- **Storage Account Contributor** (for blob operations)
- **Storage File Data SMB Share Reader** (for file enumeration)
- **Storage File Data SMB Share Contributor** (for file deletion and stub creation)

## 🚀 Quick Start

### Basic Usage
```powershell
.\FileShareToBlob.ps1 -StorageAccountName "your-storage-account" -ResourceGroupName "YourResourceGroup" -FileShareName "your-fileshare" -BlobContainerName "your-container"
```

### Preview Mode (WhatIf)
```powershell
.\FileShareToBlob.ps1 -StorageAccountName "your-storage-account" -ResourceGroupName "YourResourceGroup" -FileShareName "your-fileshare" -BlobContainerName "your-container" -WhatIfOnly $true
```

### Full Archival with Deletion
```powershell
.\FileShareToBlob.ps1 -StorageAccountName "your-storage-account" -ResourceGroupName "YourResourceGroup" -FileShareName "your-fileshare" -BlobContainerName "your-container" -DeleteAfterVerify $true -CreateStubFiles $true
```

## 📖 Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `StorageAccountName` | string | Required | Azure Storage Account name |
| `ResourceGroupName` | string | Required | Azure Resource Group name |
| `FileShareName` | string | Required | Source Azure File Share name |
| `BlobContainerName` | string | Required | Destination Blob Container name |
| `ArchiveOlderThanYears` | int | 2 | Archive files older than N years (1-50) |
| `WhatIfOnly` | bool | false | Preview mode - no actual copying |
| `DeleteAfterVerify` | bool | false | Delete source files after verification |
| `CreateStubFiles` | bool | true | Create stub files to mark archived files |
| `StubFileSuffix` | string | ".archived" | Suffix for stub files |
| `MaxQueuePerBatch` | int | 5000 | Max files per batch (1-10000) |
| `VerifyTimeoutSec` | int | 1800 | Verification timeout in seconds |
| `VerifyPollIntervalSec` | int | 30 | Polling interval for verification |
| `VerifyBatchSize` | int | 50 | Files to verify per batch |
| `FolderPath` | string | "" | Target specific folder for archival |
| `BlobTier` | string | "Hot" | Storage tier: Hot/Cool/Cold/Archive |

## 🔍 Testing & Debugging

### Test Single File
```powershell
.\FileShareToBlob.ps1 -StorageAccountName "your-storage-account" -ResourceGroupName "YourResourceGroup" -FileShareName "your-fileshare" -BlobContainerName "your-container" -SingleFileCopyRelativePath "folder/file.txt"
```

### Debug File Processing
```powershell
.\FileShareToBlob.ps1 -StorageAccountName "your-storage-account" -ResourceGroupName "YourResourceGroup" -FileShareName "your-fileshare" -BlobContainerName "your-container" -ShowFirstNDecisions 10
```

### Test File Timestamp
```powershell
.\FileShareToBlob.ps1 -StorageAccountName "your-storage-account" -ResourceGroupName "YourResourceGroup" -FileShareName "your-fileshare" -BlobContainerName "your-container" -SingleFileTestRelativePath "folder/file.txt"
```

## 🏗️ Azure Automation Setup

### 1. Create Automation Account
```powershell
# Create Automation Account
New-AzAutomationAccount -ResourceGroupName "YourResourceGroup" -Name "FileArchiveAutomation" -Location "East US"
```

### 2. Enable Managed Identity
```powershell
# Enable system-assigned managed identity
Set-AzAutomationAccount -ResourceGroupName "YourResourceGroup" -Name "FileArchiveAutomation" -AssignSystemIdentity
```

### 3. Assign Permissions
```powershell
# Get the managed identity principal ID
$automationAccount = Get-AzAutomationAccount -ResourceGroupName "YourResourceGroup" -Name "FileArchiveAutomation"
$principalId = $automationAccount.Identity.PrincipalId

# Assign Storage Account Contributor role
New-AzRoleAssignment -ObjectId $principalId -RoleDefinitionName "Storage Account Contributor" -Scope "/subscriptions/{subscription-id}/resourceGroups/{rg}/providers/Microsoft.Storage/storageAccounts/{storage-account}"

# Assign Storage File Data SMB Share Reader role
New-AzRoleAssignment -ObjectId $principalId -RoleDefinitionName "Storage File Data SMB Share Reader" -Scope "/subscriptions/{subscription-id}/resourceGroups/{rg}/providers/Microsoft.Storage/storageAccounts/{storage-account}"

# Assign Storage File Data SMB Share Contributor role
New-AzRoleAssignment -ObjectId $principalId -RoleDefinitionName "Storage File Data SMB Share Contributor" -Scope "/subscriptions/{subscription-id}/resourceGroups/{rg}/providers/Microsoft.Storage/storageAccounts/{storage-account}"
```

### 4. Import Required Modules
```powershell
# Import Az.Storage module
New-AzAutomationModule -ResourceGroupName "YourResourceGroup" -AutomationAccountName "FileArchiveAutomation" -Name "Az.Storage" -ContentLinkUri "https://www.powershellgallery.com/api/v2/package/Az.Storage/5.0.0"

# Import Az.Accounts module  
New-AzAutomationModule -ResourceGroupName "YourResourceGroup" -AutomationAccountName "FileArchiveAutomation" -Name "Az.Accounts" -ContentLinkUri "https://www.powershellgallery.com/api/v2/package/Az.Accounts/2.12.0"

# Import Az.Resources module
New-AzAutomationModule -ResourceGroupName "YourResourceGroup" -AutomationAccountName "FileArchiveAutomation" -Name "Az.Resources" -ContentLinkUri "https://www.powershellgallery.com/api/v2/package/Az.Resources/6.0.0"
```

### 5. Create Runbook
1. Go to Azure Portal → Automation Account → Runbooks
2. Click "Create a runbook"
3. Name: "FileShareToBlob"
4. Type: "PowerShell"
5. Runtime: "PowerShell 7.2" (recommended)
6. Copy the script content
7. Save and Publish

## 📊 Usage Examples

### Example 1: Archive Files Older Than 1 Year
```powershell
.\FileShareToBlob.ps1 `
  -StorageAccountName "your-storage-account" `
  -ResourceGroupName "YourResourceGroup" `
  -FileShareName "your-fileshare" `
  -BlobContainerName "your-container" `
  -ArchiveOlderThanYears 1 `
  -DeleteAfterVerify $true `
  -CreateStubFiles $true
```

### Example 2: Preview 5-Year Archive
```powershell
.\FileShareToBlob.ps1 `
  -StorageAccountName "your-storage-account" `
  -ResourceGroupName "YourResourceGroup" `
  -FileShareName "your-fileshare" `
  -BlobContainerName "your-container" `
  -ArchiveOlderThanYears 5 `
  -WhatIfOnly $true `
  -ShowFirstNDecisions 10
```

### Example 3: Custom Stub Files
```powershell
.\FileShareToBlob.ps1 `
  -StorageAccountName "your-storage-account" `
  -ResourceGroupName "YourResourceGroup" `
  -FileShareName "your-fileshare" `
  -BlobContainerName "your-container" `
  -StubFileSuffix ".moved" `
  -DeleteAfterVerify $true
```

### Example 4: Archive Specific Folder to Cool Tier
```powershell
.\FileShareToBlob.ps1 `
  -StorageAccountName "your-storage-account" `
  -ResourceGroupName "YourResourceGroup" `
  -FileShareName "your-fileshare" `
  -BlobContainerName "your-container" `
  -FolderPath "documents/archive" `
  -BlobTier "Cool" `
  -DeleteAfterVerify $true
```

### Example 5: Long-Term Archive to Archive Tier
```powershell
.\FileShareToBlob.ps1 `
  -StorageAccountName "your-storage-account" `
  -ResourceGroupName "YourResourceGroup" `
  -FileShareName "your-fileshare" `
  -BlobContainerName "your-container" `
  -ArchiveOlderThanYears 5 `
  -BlobTier "Archive" `
  -DeleteAfterVerify $true
```

### Example 6: Preview Folder-Specific Archive
```powershell
.\FileShareToBlob.ps1 `
  -StorageAccountName "your-storage-account" `
  -ResourceGroupName "YourResourceGroup" `
  -FileShareName "your-fileshare" `
  -BlobContainerName "your-container" `
  -FolderPath "reports/2023" `
  -BlobTier "Cold" `
  -WhatIfOnly $true
```

## 🔍 How It Works

1. **Authentication**: Uses Managed Identity to authenticate with Azure
2. **File Enumeration**: Recursively scans the file share for files
3. **Age Filtering**: Uses NTFS LastWriteTime to determine file age
4. **Server-Side Copy**: Copies files from File Share to Blob Storage
5. **Verification**: Confirms successful blob creation
6. **Cleanup**: Optionally deletes source files and creates stub files

## 💰 Blob Tier Cost Optimization

The script supports different Azure Blob Storage tiers for cost optimization:

| Tier | Use Case | Cost | Access Time | Minimum Retention |
|------|----------|------|-------------|-------------------|
| **Hot** | Frequently accessed files | Highest | Immediate | None |
| **Cool** | Infrequently accessed files | Lower | Immediate | 30 days |
| **Cold** | Rarely accessed files | Lower | Immediate | 30 days |
| **Archive** | Long-term storage | Lowest | 1-15 hours | 180 days |

### Cost Optimization Examples

**For Active Projects (Hot Tier)**
```powershell
-BlobTier "Hot"  # Default - for frequently accessed files
```

**For Completed Projects (Cool Tier)**
```powershell
-BlobTier "Cool"  # For infrequently accessed files
```

**For Long-Term Storage (Archive Tier)**
```powershell
-BlobTier "Archive"  # For rarely accessed files (5+ years old)
```

## 📁 Folder Path Selection

Target specific folders for archival operations instead of processing the entire file share:

### Basic Folder Targeting
```powershell
-FolderPath "documents/archive"  # Archive only files in documents/archive folder
```

### Nested Folder Examples
```powershell
-FolderPath "projects/2023"      # Archive files in projects/2023 folder
-FolderPath "reports/old"        # Archive files in reports/old folder
-FolderPath "temp"               # Archive files in temp folder
```

### Combined with Blob Tiers
```powershell
# Archive old documents to Cool tier
-FolderPath "documents/old" -BlobTier "Cool"

# Archive completed projects to Archive tier
-FolderPath "projects/completed" -BlobTier "Archive"
```

## 📝 Stub Files

When `CreateStubFiles = true` and `DeleteAfterVerify = true`, the script creates stub files with:
- **Content**: Archive timestamp and original file path
- **Suffix**: Configurable (default: `.archived`)
- **Purpose**: Track archived files and prevent re-processing

Example stub file content:
```
Archived on: 2024-01-15 14:30:25 UTC
Original file: documents/old-report.pdf
```

## 🚨 Important Notes

- **NTFS LastWriteTime**: Uses actual file modification time, not Azure service timestamps
- **Server-Side Copy**: No data downloads to local machine
- **Batch Processing**: Handles large file shares efficiently
- **Error Handling**: Comprehensive logging and error recovery
- **Idempotent**: Safe to run multiple times (stub files prevent re-processing)

## 🐛 Troubleshooting

### Common Issues

1. **"You cannot call a method on a null-valued expression"**
   - Ensure proper permissions are assigned
   - Check that file share and blob container exist

2. **No files found for archival**
   - Verify file ages meet the cutoff criteria
   - Check NTFS LastWriteTime vs service LastModified

3. **Verification timeouts**
   - Increase `VerifyTimeoutSec` parameter
   - Reduce `VerifyBatchSize` for large files

### Debug Mode
```powershell
# Enable detailed logging
.\FileShareToBlob.ps1 -StorageAccountName "your-storage-account" -ResourceGroupName "YourResourceGroup" -FileShareName "your-fileshare" -BlobContainerName "your-container" -ShowFirstNDecisions 20 -DebugTargetRelativePath "problematic-file.txt"
```

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📞 Support

For issues and questions:
- Create an issue in this repository
- Check the troubleshooting section above
- Review Azure Automation logs for detailed error information

## 🔄 Version History

- **v1.0**: Initial release with full archival functionality
- Comprehensive documentation and examples
- Azure Automation integration
- Stub file creation and verification

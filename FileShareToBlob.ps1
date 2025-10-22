<#
.SYNOPSIS
    Azure File Share to Blob Storage Archival Script with Verification and Stub File Creation

.DESCRIPTION
    This PowerShell script archives files from Azure File Share to Azure Blob Storage based on age criteria.
    It provides comprehensive functionality including file copying, verification, deletion, and stub file creation.
    Designed to run in Azure Automation Account with Managed Identity authentication.

.FEATURES
    - Archives files older than specified years (configurable cutoff)
    - Uses NTFS LastWriteTime for accurate file age determination
    - Server-side copy for efficient data transfer
    - Comprehensive verification before deletion
    - Optional stub file creation to mark archived files
    - Batch processing for large datasets
    - Single file testing and debugging capabilities
    - Support for both PowerShell 5.1 and 7.2 runtimes

.PARAMETER StorageAccountName
    Name of the Azure Storage Account containing the file share and blob container.
    Example: "mystorageaccount"

.PARAMETER ResourceGroupName
    Name of the Azure Resource Group containing the storage account.
    Example: "MyResourceGroup"

.PARAMETER FileShareName
    Name of the Azure File Share to archive files from.
    Example: "myshare"

.PARAMETER BlobContainerName
    Name of the Azure Blob Container to copy files to.
    Example: "archive"

.PARAMETER SubscriptionId
    Azure Subscription ID. If not provided, uses current context.
    Example: "your-subscription-id-here"

.PARAMETER MaxQueuePerBatch
    Maximum number of files to process in each batch (1-10000).
    Default: 5000

.PARAMETER WhatIfOnly
    Preview mode - shows what would be copied without actually copying.
    Default: false

.PARAMETER DeleteAfterVerify
    Delete source files after successful verification of blob copies.
    Default: false

.PARAMETER VerifyTimeoutSec
    Timeout in seconds for verification process (default: 1800 = 30 minutes).
    Example: 3600

.PARAMETER VerifyPollIntervalSec
    Interval in seconds between verification checks (default: 30).
    Example: 60

.PARAMETER VerifyBatchSize
    Number of files to verify in each batch (default: 50).
    Example: 100

.PARAMETER SingleFileTestRelativePath
    Test a single file's timestamp logic without copying.
    Example: "folder/file.txt"

.PARAMETER SingleFileCopyRelativePath
    Copy a single specific file for testing.
    Example: "folder/file.txt"

.PARAMETER DebugTargetRelativePath
    Debug a specific file's processing.
    Example: "folder/file.txt"

.PARAMETER ShowFirstNDecisions
    Show first N file processing decisions for debugging.
    Example: 10

.PARAMETER UseSasListing
    Use SAS-based listing instead of RBAC (recommended for Automation).
    Default: true

.PARAMETER CreateStubFiles
    Create stub files to mark archived files.
    Default: true

.PARAMETER StubFileSuffix
    Suffix for stub files (default: ".archived").
    Example: ".moved"

.PARAMETER ArchiveOlderThanYears
    Archive files older than this many years (1-50).
    Default: 2

.PARAMETER FolderPath
    Optional folder path to limit archival to a specific folder within the file share.
    If not specified, processes all folders. Use forward slashes for path separators.
    Example: "documents/archive" or "reports/2023"

.PARAMETER BlobTier
    Azure Blob Storage tier to use for archived files. Options: Hot, Cool, Archive, Cold.
    Default: Hot (for frequently accessed files). Use Cool/Cold/Archive for cost optimization.
    Example: "Cool" for infrequently accessed files, "Archive" for long-term storage

.EXAMPLE
    # Basic archival of files older than 2 years
    .\FileShareToBlob.ps1 -StorageAccountName "mystorage" -ResourceGroupName "MyRG" -FileShareName "myshare" -BlobContainerName "archive"

.EXAMPLE
    # Preview what would be archived (WhatIf mode)
    .\FileShareToBlob.ps1 -StorageAccountName "mystorage" -ResourceGroupName "MyRG" -FileShareName "myshare" -BlobContainerName "archive" -WhatIfOnly $true

.EXAMPLE
    # Archive files older than 1 year with deletion and stubs
    .\FileShareToBlob.ps1 -StorageAccountName "mystorage" -ResourceGroupName "MyRG" -FileShareName "myshare" -BlobContainerName "archive" -ArchiveOlderThanYears 1 -DeleteAfterVerify $true -CreateStubFiles $true

.EXAMPLE
    # Test a single file
    .\FileShareToBlob.ps1 -StorageAccountName "mystorage" -ResourceGroupName "MyRG" -FileShareName "myshare" -BlobContainerName "archive" -SingleFileCopyRelativePath "folder/file.txt" -DeleteAfterVerify $true

.EXAMPLE
    # Debug first 5 file decisions
    .\FileShareToBlob.ps1 -StorageAccountName "mystorage" -ResourceGroupName "MyRG" -FileShareName "myshare" -BlobContainerName "archive" -ShowFirstNDecisions 5

.EXAMPLE
    # Archive files older than 5 years with custom stub suffix
    .\FileShareToBlob.ps1 -StorageAccountName "mystorage" -ResourceGroupName "MyRG" -FileShareName "myshare" -BlobContainerName "archive" -ArchiveOlderThanYears 5 -StubFileSuffix ".moved"

.EXAMPLE
    # Archive files from a specific folder only
    .\FileShareToBlob.ps1 -StorageAccountName "mystorage" -ResourceGroupName "MyRG" -FileShareName "myshare" -BlobContainerName "archive" -FolderPath "documents/archive"

.EXAMPLE
    # Preview archival for a specific folder (WhatIf mode)
    .\FileShareToBlob.ps1 -StorageAccountName "mystorage" -ResourceGroupName "MyRG" -FileShareName "myshare" -BlobContainerName "archive" -FolderPath "reports/2023" -WhatIfOnly $true

.EXAMPLE
    # Archive files to Cool tier for cost optimization
    .\FileShareToBlob.ps1 -StorageAccountName "mystorage" -ResourceGroupName "MyRG" -FileShareName "myshare" -BlobContainerName "archive" -BlobTier "Cool"

.EXAMPLE
    # Archive files to Archive tier for long-term storage
    .\FileShareToBlob.ps1 -StorageAccountName "mystorage" -ResourceGroupName "MyRG" -FileShareName "myshare" -BlobContainerName "archive" -BlobTier "Archive" -ArchiveOlderThanYears 5

.EXAMPLE
    # Archive specific folder to Cold tier with deletion
    .\FileShareToBlob.ps1 -StorageAccountName "mystorage" -ResourceGroupName "MyRG" -FileShareName "myshare" -BlobContainerName "archive" -FolderPath "old-documents" -BlobTier "Cold" -DeleteAfterVerify $true

.NOTES
    - Requires Az.Storage, Az.Accounts, and Az.Resources modules
    - Designed for Azure Automation Account with Managed Identity
    - Uses NTFS LastWriteTime for accurate file age determination
    - Supports both PowerShell 5.1 and 7.2 runtimes
    - Includes comprehensive error handling and logging
    - Creates stub files to track archived files and prevent re-processing
    - Blob tier selection helps optimize storage costs (Hot/Cool/Cold/Archive)
    - Archive tier has retrieval delays and costs - use for long-term storage only

.AUTHOR
    Azure File Share Archival Script

.VERSION
    1.0
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$StorageAccountName = "",

    [Parameter(Mandatory=$false)]
    [string]$ResourceGroupName = "",

    [Parameter(Mandatory=$false)]
    [string]$FileShareName = "",

    [Parameter(Mandatory=$false)]
    [string]$BlobContainerName = "",

    [Parameter(Mandatory=$false)]
    [string]$SubscriptionId = $null,

    [Parameter(Mandatory=$false)]
    [ValidateRange(1,10000)]
    [int]$MaxQueuePerBatch = 5000,

    # Only preview actions; no copy/delete
    #[switch]$WhatIfOnly,
    [bool]$WhatIfOnly = $true,

    # Verification + deletion controls
        [bool]$DeleteAfterVerify = $false,
    [int]$VerifyTimeoutSec = 86400,          # 24h default timeout for long copies
    [int]$VerifyPollIntervalSec = 10,        # Poll interval for copy status
    [int]$VerifyBatchSize = 500,             # How many blobs to poll per cycle

    # Optional: debug a specific source relative path (e.g., 'folder/file.txt' or 'fdrive.txt')
    [string]$DebugTargetRelativePath = '',

    # Optional: test a single file (relative path within share) and print timestamps then exit
    [string]$SingleFileTestRelativePath = $null,

    # Optional: copy a single file (relative path) honoring the NTFS cutoff
    [string]$SingleFileCopyRelativePath = $null,

    # Optional: log the first N decisions during scan (0 to disable)
    [int]$ShowFirstNDecisions = 5
    ,
    # Use SAS-based REST listing instead of RBAC listing (recommended in Automation)
    [bool]$UseSasListing = $true,
    
    # Create stub files to track copied files and skip them on subsequent runs
    [bool]$CreateStubFiles = $true,
    [string]$StubFileSuffix = ".archived",

    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 50)]
    [int]$ArchiveOlderThanYears = 2,

    [Parameter(Mandatory=$false)]
    [string]$FolderPath = "",

    [Parameter(Mandatory=$false)]
    [ValidateSet("Hot", "Cool", "Archive", "Cold")]
    [string]$BlobTier = "Hot"
)

Import-Module Az.Accounts -ErrorAction Stop
Import-Module Az.Storage  -ErrorAction Stop
Import-Module Az.Resources -ErrorAction Stop

function Connect-Cloud {
    try {
        Connect-AzAccount -Identity | Out-Null
        if ($SubscriptionId) { Set-AzContext -SubscriptionId $SubscriptionId | Out-Null }
    } catch {
        Write-Warning "Managed Identity sign-in failed. Falling back to device login (for local testing)."
        Connect-AzAccount -UseDeviceAuthentication | Out-Null
        if ($SubscriptionId) { Set-AzContext -SubscriptionId $SubscriptionId | Out-Null }
    }
}

function Test-StubFileExists {
    param(
        [string]$RelativePath,
        [object]$Context,
        [string]$ShareName,
        [string]$StubSuffix
    )
    $stubPath = "$RelativePath$StubSuffix"
    try {
        $stubFile = Get-AzStorageFile -ShareName $ShareName -Path $stubPath -Context $Context -ErrorAction SilentlyContinue
        return ($stubFile -ne $null)
    } catch {
        return $false
    }
}

function New-StubFile {
    param(
        [string]$RelativePath,
        [object]$Context,
        [string]$ShareName,
        [string]$StubSuffix
    )
    $stubPath = "$RelativePath$StubSuffix"
    try {
        $stubContent = "Archived on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC')`nOriginal file: $RelativePath"
        $stubFile = New-AzStorageFile -ShareName $ShareName -Path $stubPath -Context $Context -ErrorAction Stop
        $stubFile.UploadText($stubContent) | Out-Null
        Write-Output ("[STUB] Created stub file: {0}" -f $stubPath)
        return $true
    } catch {
        Write-Output ("[STUB-ERROR] Failed to create stub for '{0}': {1}" -f $RelativePath, $_.Exception.Message)
        return $false
    }
}

# Returns [DateTime] in UTC if available, otherwise $null
function Get-FileLastWriteTimeUtc {
    param(
        [Parameter(Mandatory=$true)][string]$FileUrlWithSas
    )
    # Attempt 1: Use Azure.Storage.Files.Shares SDK if available
    try {
        $loaded = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'Azure.Storage.Files.Shares' }
        if (-not $loaded) {
            try { [void][System.Reflection.Assembly]::Load('Azure.Storage.Files.Shares') } catch { }
        }
        if ([Type]::GetType('Azure.Storage.Files.Shares.ShareFileClient, Azure.Storage.Files.Shares')) {
            $sf = [Azure.Storage.Files.Shares.ShareFileClient]::new([System.Uri]$FileUrlWithSas)
            $props = $sf.GetProperties()
            $fw = $props.Value.SmbProperties.FileLastWrittenOn
            if ($fw) { return ([DateTimeOffset]$fw).UtcDateTime }
        }
    } catch { }

    # Attempt 2: REST HEAD with recent storage version
    $headers = @{ 'x-ms-version' = '2023-11-03' }
    try {
        $resp = Invoke-WebRequest -Method Head -Uri $FileUrlWithSas -Headers $headers -UseBasicParsing -TimeoutSec 60 -ErrorAction Stop
        $val = $resp.Headers['x-ms-file-last-write-time']
        if (-not $val) { return $null }
        if ($val -eq 'preserve' -or $val -eq 'now') { return $null }
        try { return ([DateTimeOffset]::Parse($val)).UtcDateTime } catch { return $null }
    } catch { return $null }
}

function Get-StorageContexts {
    $acct = Get-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $StorageAccountName -ErrorAction Stop
    $key = (Get-AzStorageAccountKey -ResourceGroupName $ResourceGroupName -Name $StorageAccountName)[0].Value
    $ctx = New-AzStorageContext -StorageAccountName $StorageAccountName -StorageAccountKey $key
    return [pscustomobject]@{
        StorageAccount = $acct
        Context        = $ctx
    }
}

function Ensure-Container {
    param(
        [object]$Context,
        [string]$ContainerName
    )
    $container = Get-AzStorageContainer -Name $ContainerName -Context $Context -ErrorAction SilentlyContinue
    if (-not $container) {
        New-AzStorageContainer -Name $ContainerName -Context $Context | Out-Null
    }
}

function New-FileShareSas {
    param(
        [object]$Context,
        [string]$ShareName,
        [datetime]$ExpiryUtc
    )
    $sas = New-AzStorageShareSASToken -Context $Context -Name $ShareName -Permission rl -ExpiryTime $ExpiryUtc -FullUri:$false
    if ($sas -and $sas[0] -ne '?') { $sas = "?$sas" }
    return $sas
}

function Get-AllFilesRecursive {
    param(
        [object]$Context,
        [string]$ShareName,
        [string]$ShareSas,
        [bool]$UseSasListing,
        [string]$AccountName,
        [string]$FolderPath = ""
    )
    function List-Path {
        param([string]$p)
        if ($UseSasListing) {
            $dir = if ([string]::IsNullOrEmpty($p)) { '' } else { "/$p" }
            # Try SDK first
            try {
                if (-not ([Type]::GetType('Azure.Storage.Files.Shares.ShareDirectoryClient, Azure.Storage.Files.Shares'))) {
                    try { [void][System.Reflection.Assembly]::Load('Azure.Storage.Files.Shares') } catch {}
                }
                if ([Type]::GetType('Azure.Storage.Files.Shares.ShareDirectoryClient, Azure.Storage.Files.Shares')) {
                    if ($p -eq '') { Write-Output "[ENUM] Using SAS-based SDK listing" }
                    $dirUrl = "https://$AccountName.file.core.windows.net/$ShareName$dir$ShareSas"
                    $sd = [Azure.Storage.Files.Shares.ShareDirectoryClient]::new([System.Uri]$dirUrl)
                    $entries = $sd.GetFilesAndDirectories()
                    $items = @()
                    foreach ($e in $entries) {
                        $obj = New-Object PSObject -Property @{ Name = $e.Name; IsDirectory = [bool]$e.IsDirectory }
                        $items += $obj
                    }
                    return $items
                }
            } catch {
                Write-Output ("[ENUM] SDK list failed for path '{0}': {1}" -f $p, $_.Exception.Message)
            }
            # Fallback to REST
            if ($p -eq '') { Write-Output "[ENUM] Using SAS-based REST listing" }
            $base = "https://$AccountName.file.core.windows.net/$ShareName"
            $marker = $null
            $items = @()
            do {
                $sasTail = if ([string]::IsNullOrEmpty($ShareSas)) { '' } elseif ($ShareSas.StartsWith('?')) { $ShareSas.Substring(1) } else { $ShareSas }
               # $sasTail = if ($ShareSas.StartsWith('?')) { $ShareSas.Substring(1) } else { $ShareSas }
                $uri = "$base$dir?restype=directory&comp=list&$sasTail" + $(if ($marker) { "&marker=$([uri]::EscapeDataString($marker))" } else { '' })
                try {
                    $resp = Invoke-WebRequest -Method Get -Uri $uri -Headers @{ 'x-ms-version' = '2023-11-03' } -UseBasicParsing -TimeoutSec 60 -ErrorAction Stop
                    [xml]$xml = $resp.Content
                    foreach ($d in $xml.EnumerationResults.Entries.Directory) { $items += (New-Object PSObject -Property @{ Name = $d.Name; IsDirectory = $true }) }
                    foreach ($f in $xml.EnumerationResults.Entries.File) { $items += (New-Object PSObject -Property @{ Name = $f.Name; IsDirectory = $false }) }
                    $marker = $xml.EnumerationResults.NextMarker; if ([string]::IsNullOrEmpty($marker)) { $marker = $null }
                } catch {
                    Write-Output ("[ENUM] REST list failed for path '{0}': {1}" -f $p, $_.Exception.Message)
                    break
                }
            } while ($marker)
            return $items
        } else {
            $res = @()
            try {
                if ([string]::IsNullOrEmpty($p)) { $res = Get-AzStorageFile -ShareName $ShareName -Context $Context -ErrorAction Stop }
                else { $res = Get-AzStorageFile -ShareName $ShareName -Path $p -Context $Context -ErrorAction Stop }
            } catch {}
            return $res
        }
    }

    try {
    $stack = New-Object System.Collections.Stack
    # Start from specified folder path or root if not specified
    $startPath = if ([string]::IsNullOrEmpty($FolderPath)) { '' } else { $FolderPath.Trim('/') }
    $stack.Push($startPath)
    while ($stack.Count -gt 0) {
        $currentPath = [string]$stack.Pop()
            try {
                $listing = List-Path -p $currentPath
                if ($currentPath -eq '') {
                    $preview = ($listing | Select-Object -First 5 | ForEach-Object { $_.Name }) -join ', '
                    if ($preview) { Write-Output ("[ENUM] Root entries: {0}" -f $preview) } else { Write-Output "[ENUM] Root listing returned 0 entries" }
                }
        foreach ($item in $listing) {
                    $isDir = $false
                    if ($UseSasListing) { $isDir = $item.IsDirectory } else { $isDir = ($item -is [Microsoft.Azure.Storage.File.CloudFileDirectory]) }
                    if ($isDir) {
                $dirName = $item.Name
                $next = if ([string]::IsNullOrEmpty($currentPath)) { $dirName } else { "$currentPath/$dirName" }
                $stack.Push($next)
                    } else {
                        if ($UseSasListing) {
                            # Yield synthetic file object with Uri-like properties for downstream
                            $full = if ([string]::IsNullOrEmpty($currentPath)) { $item.Name } else { "$currentPath/$($item.Name)" }
                            $obj = New-Object PSObject -Property @{ Name = $item.Name; RelativePath = $full; Uri = [Uri]::new("https://$AccountName.file.core.windows.net/$ShareName/$full") }
                            $obj
                        } else {
                $item
                        }
                    }
                }
            } catch {
                Write-Output ("[ENUM-ERROR] Path '{0}': {1}" -f $currentPath, $_.Exception.Message)
            }
        }
    } catch {
        Write-Output ("[ENUM-ERROR] Stack traversal failed: {0}" -f $_.Exception.Message)
    }
}

function Queue-BlobCopiesForOldFiles {
    param(
        [object]$Context,
        [string]$ShareName,
        [string]$ContainerName,
        [datetime]$OlderThanUtc,
        [string]$ShareSas,
        [int]$MaxQueuePerBatch = 5000,
        [bool]$WhatIfOnly = $false,
        [bool]$UseSasListing = $true,
        [string]$FolderPath = ""
    )
    Write-Output ("[DEBUG] Queue-BlobCopiesForOldFiles called with WhatIfOnly={0}" -f $WhatIfOnly)
    $queued = 0
    $batchQueued = 0
    $copyJobs = New-Object System.Collections.Generic.List[object]
    $scanned = 0
    $noHeaderCount = 0
    $fallbackUsedCount = 0
    $skippedNewer = 0
    $sampleLogged = 0
    $decisionLogged = 0

    Write-Output ("[DEBUG] Inside Queue-BlobCopiesForOldFiles, calling Get-AllFilesRecursive...")
    try { $files = Get-AllFilesRecursive -Context $Context -ShareName $ShareName -ShareSas $ShareSas -UseSasListing:$UseSasListing -AccountName $StorageAccountName -FolderPath $FolderPath }
    catch { Write-Output ("[ENUM] Listing failed: {0}" -f $_.Exception.Message); return [pscustomobject]@{ QueuedCount = 0; Jobs = @() } }
    
    # Debug first few objects to see their structure
    $debugCount = 0
    foreach ($f in $files) {
        if ($debugCount -lt 3) {
            Write-Output ("[DEBUG] Object {0}: Type={1}, HasRelativePath={2}, RelativePath='{3}', HasUri={4}" -f $debugCount, $f.GetType().FullName, ($f.PSObject.Properties.Name -contains 'RelativePath'), $f.RelativePath, ($f.PSObject.Properties.Name -contains 'Uri'))
        }
        $debugCount++
        if ($debugCount -ge 3) { break }
    }
    # Materialize the enumeration so we can count and preview
    $fileList = @()
    foreach ($f in $files) { $fileList += ,$f }
    $previewNames = ($fileList | Select-Object -First 5 | ForEach-Object {
        if ($_ -is [Microsoft.Azure.Storage.File.CloudFile]) { $_.Name }
        elseif ($_ -and $_.PSObject.Properties['RelativePath']) { $_.RelativePath }
        else { $_.Name }
    }) -join ', '
    Write-Output ("[ENUM] Total items enumerated: {0}{1}" -f $fileList.Count, $(if ($previewNames) { "; First: $previewNames" } else { '' }))
    if ($fileList.Count -eq 0) { return [pscustomobject]@{ QueuedCount = 0; Jobs = @() } }

    foreach ($file in $fileList) {
        $scanned++
        # Build URLs and paths once
        if ($file -ne $null -and $file.PSObject.Properties.Name -contains 'RelativePath' -and $file.RelativePath) {
            $relativePath = $file.RelativePath
            $destBlobName = $relativePath
            $srcUrl = "https://$StorageAccountName.file.core.windows.net/$ShareName/$relativePath$ShareSas"
            $uri = [Uri]::new("https://$StorageAccountName.file.core.windows.net/$ShareName/$relativePath")
            $relativeSegments = $relativePath.Split('/')
        } else {
            if ($file.Uri) {
            $uri = $file.Uri
            $segments = $uri.AbsolutePath.TrimStart('/').Split('/')
            $relativeSegments = $segments[1..($segments.Length - 1)]
            $destBlobName = ($relativeSegments -join '/')
            $srcUrl = "$uri$ShareSas"
            } else {
                Write-Output ("[SKIP] File object has no Uri property: {0}" -f $file.Name)
                continue
            }
        }

        # Obtain NTFS/Samba LastWriteTime via HEAD (x-ms-file-last-write-time)
        $lastWriteUtc = Get-FileLastWriteTimeUtc -FileUrlWithSas $srcUrl
        if (-not $lastWriteUtc) {
            $noHeaderCount++
            # Fallback to service LastModified only for SDK CloudFile objects
            if ($file -is [Microsoft.WindowsAzure.Commands.Common.Storage.ResourceModel.AzureStorageFile]) {
                try { $file.FetchAttributes() } catch {}
                if ($file.Properties -and $file.Properties.LastModified) {
                    $lastWriteUtc = $file.Properties.LastModified.UtcDateTime; $fallbackUsedCount++
                }
            }
        }
        if (-not $lastWriteUtc) { continue }

        # Debug for a specific relative path
        if ($DebugTargetRelativePath) {
            $relJoin = ($relativeSegments -join '/')
            if ($relJoin -ieq $DebugTargetRelativePath) {
                $svcLastMod = $null
                try { if (-not $file.Properties.LastModified) { $file.FetchAttributes() } } catch {}
                if ($file.Properties.LastModified) { $svcLastMod = $file.Properties.LastModified.UtcDateTime }
                Write-Output ("[DEBUG] Target={0} | NTFS LastWrite={1:o} | Service LastModified={2:o} | Cutoff={3:o}" -f $relJoin, $lastWriteUtc, $svcLastMod, $OlderThanUtc)
            }
        }

        if ($lastWriteUtc -le $OlderThanUtc) {
            if ($ShowFirstNDecisions -gt 0 -and $decisionLogged -lt $ShowFirstNDecisions) {
                Write-Output ("[SCAN] COPY: {0} | LastWriteUtc={1:o} <= Cutoff={2:o}" -f ($relativeSegments -join '/'), $lastWriteUtc, $OlderThanUtc)
                $decisionLogged++
            }
            # Build source relative path for deletion (Azure Files API)
            $sourceRelativePath = ($relativeSegments -join '/')
            if ($WhatIfOnly) {
                Write-Output ("[WHATIF] Would copy: share '{0}' file '{1}' -> {2}/{3}" -f $ShareName, ($relativeSegments -join '/'), $ContainerName, $destBlobName)
            } else {
                try {
                    # Try to set blob tier during copy if possible
                    $copyJob = Start-AzStorageBlobCopy -SrcShareName $ShareName -SrcFilePath ($relativeSegments -join '/') -DestContainer $ContainerName -DestBlob $destBlobName -Context $Context -Force
                    $jobObj = [pscustomobject]@{ 
                        Container = $ContainerName; 
                        BlobName = $destBlobName; 
                        CopyId = $copyJob.CopyId;
                        SourceShare = $ShareName;
                        SourceRelativePath = ($relativeSegments -join '/');
                        BlobTier = $BlobTier
                    }
                    $copyJobs.Add($jobObj)
                    Write-Output ("[DEBUG] Created job object: Container={0}, BlobName={1}, Properties={2}" -f $jobObj.Container, $jobObj.BlobName, ($jobObj | Get-Member -MemberType Property | Select-Object -ExpandProperty Name | Sort-Object))
                    Write-Output ("Queued copy: share '{0}' file '{1}' -> '{2}/{3}'" -f $ShareName, ($relativeSegments -join '/'), $ContainerName, $destBlobName)
                } catch {
                    Write-Output ("Failed to queue copy for '{0}': {1}" -f ($relativeSegments -join '/'), $_.Exception.Message)
                }
            }

            $queued++
            $batchQueued++
            if ($batchQueued -ge $MaxQueuePerBatch) {
                Write-Verbose "Throttling: queued $batchQueued operations, pausing 5 seconds..."
                Start-Sleep -Seconds 5
                $batchQueued = 0
            }
        } else {
            $skippedNewer++
            if ($ShowFirstNDecisions -gt 0 -and $decisionLogged -lt $ShowFirstNDecisions) {
                Write-Output ("[SCAN] SKIP: {0} | LastWriteUtc={1:o} > Cutoff={2:o}" -f ($relativeSegments -join '/'), $lastWriteUtc, $OlderThanUtc)
                $decisionLogged++
            } elseif ($sampleLogged -lt 5) {
                Write-Verbose ("Skip newer: {0} | LastWriteUtc={1:o} > Cutoff={2:o}" -f $uri.AbsoluteUri, $lastWriteUtc, $OlderThanUtc)
                $sampleLogged++
            }
        }
    }

    Write-Output ("Scanned files: {0}; With NTFS header missing: {1}; Fallback used: {2}; Skipped newer: {3}" -f $scanned, $noHeaderCount, $fallbackUsedCount, $skippedNewer)

    Write-Output ("[DEBUG] Queue-BlobCopiesForOldFiles returning: QueuedCount={0}, JobsCount={1}" -f $queued, $copyJobs.Count)
    if ($copyJobs.Count -gt 0) {
        Write-Output ("[DEBUG] First job in return: Container={0}, BlobName={1}" -f $copyJobs[0].Container, $copyJobs[0].BlobName)
    }
    return [pscustomobject]@{
        QueuedCount = $queued
        Jobs        = $copyJobs
    }
}

function Wait-Verify-And-Delete {
    param(
        [object]$Context,
        [Parameter(Mandatory=$true)]
        [object]$Jobs,                       # accept array, list or any enumerable
        [int]$TimeoutSec,
        [int]$PollIntervalSec,
        [int]$BatchSize,
        [bool]$DeleteAfterVerify = $false,
        [bool]$WhatIfOnly = $false,
        [bool]$CreateStubFiles = $false,
        [string]$StubFileSuffix = ".archived"
    )

    Write-Output "[VERIFY-FUNC] Function entry point reached!"
    # Normalize Jobs into a System.Collections.Generic.List[object]
    $pending = [System.Collections.Generic.List[object]]::new()
    if ($null -eq $Jobs) {
        Write-Output "[VERIFY-FUNC] Jobs is null or empty."
    } else {
        # If it's already a generic list, add elements; if it's an array/enumerable, iterate and add
        try {
            foreach ($j in $Jobs) { $pending.Add($j) }
        } catch {
            # Single object (not enumerable) — just add it
            $pending.Add($Jobs)
        }
    }

    Write-Output ("[VERIFY-FUNC] Called with JobsCount={0}, DeleteAfterVerify={1}, WhatIfOnly={2}" -f $pending.Count, $DeleteAfterVerify, $WhatIfOnly)

    if ($pending.Count -eq 0) {
        Write-Output "No jobs to verify."
        return [pscustomobject]@{
            Verified = 0
            Deleted  = 0
            Failed   = 0
            TimedOut = 0
        }
    }

    $deadline = (Get-Date).ToUniversalTime().AddSeconds($TimeoutSec)
    $verified = 0
    $deleted  = 0
    $failed   = 0
    $timedOut = 0

    while ($pending.Count -gt 0) {
        $now = (Get-Date).ToUniversalTime()
        if ($now -ge $deadline) {
            $timedOut += $pending.Count
            Write-Warning ("[VERIFY] Deadline reached, marking {0} remaining jobs as timed out." -f $pending.Count)
            $pending.Clear()
            break
        }

        # Work in batches
        $countThisBatch = [Math]::Min($BatchSize, $pending.Count)
        $batch = for ($i = 0; $i -lt $countThisBatch; $i++) { $pending[$i] }

        $toRemove = New-Object System.Collections.Generic.List[object]

        foreach ($job in $batch) {
            try {
                # Defensive: ensure job has Container and BlobName
                $container = $null; $blobName = $null
                if ($job -and $job.PSObject.Properties['Container']) { $container = $job.Container }
                if ($job -and $job.PSObject.Properties['BlobName'])  { $blobName  = $job.BlobName  }

                Write-Output ("[VERIFY] Checking blob existence for {0}/{1}" -f ($container), ($blobName))
                $blob = Get-AzStorageBlob -Container $container -Blob $blobName -Context $Context -ErrorAction SilentlyContinue
                if ($blob) {
                        $verified++
                    Write-Output ("[VERIFY] Verified copy: {0}/{1}" -f $container, $blobName)
                    
                    # Set blob tier if specified and different from current tier
                    if ($job -and $job.PSObject.Properties['BlobTier'] -and $job.BlobTier) {
                        try {
                            $currentTier = $blob.BlobTier
                            if ($currentTier -ne $job.BlobTier) {
                                # Use Azure Storage SDK approach (most reliable in Automation Account)
                                try {
                                    # Load Azure.Storage.Blobs SDK if available
                                    if (-not ([Type]::GetType('Azure.Storage.Blobs.BlobClient, Azure.Storage.Blobs'))) {
                                        try { [void][System.Reflection.Assembly]::Load('Azure.Storage.Blobs') } catch { }
                                    }
                                    
                                    if ([Type]::GetType('Azure.Storage.Blobs.BlobClient, Azure.Storage.Blobs')) {
                                        # Use Azure Storage SDK with storage account key
                                        $storageKey = $Context.StorageAccount.Credentials.ExportBase64EncodedKey()
                                        $credential = [Azure.Storage.StorageSharedKeyCredential]::new($Context.StorageAccountName, $storageKey)
                                        $blobUri = "https://$($Context.StorageAccountName).blob.core.windows.net/$container/$blobName"
                                        $blobClient = [Azure.Storage.Blobs.BlobClient]::new([System.Uri]$blobUri, $credential)
                                        $blobClient.SetAccessTier($job.BlobTier)
                                        Write-Output ("[BLOB-TIER] Set tier for {0}/{1} to {2} using Azure Storage SDK" -f $container, $blobName, $job.BlobTier)
                                    } else {
                                        Write-Warning ("[BLOB-TIER] Azure Storage SDK not available, skipping tier setting for {0}/{1}" -f $container, $blobName)
                                    }
                                } catch {
                                    Write-Warning ("[BLOB-TIER] Failed to set tier for {0}/{1}: {2}" -f $container, $blobName, $_.Exception.Message)
                                }
                            } else {
                                Write-Output ("[BLOB-TIER] {0}/{1} already has tier {2}" -f $container, $blobName, $job.BlobTier)
                            }
                        } catch {
                            Write-Warning ("[BLOB-TIER] Failed to set tier for {0}/{1}: {2}" -f $container, $blobName, $_.Exception.Message)
                        }
                    }

                        if ($DeleteAfterVerify) {
                        Write-Output ("[DELETE] DeleteAfterVerify is true; WhatIfOnly={0}" -f $WhatIfOnly)

                        # Log job's source info (defensive)
                        $srcShare = $null; $srcRel = $null
                        if ($job -and $job.PSObject.Properties['SourceShare']) { $srcShare = $job.SourceShare }
                        if ($job -and $job.PSObject.Properties['SourceRelativePath']) { $srcRel = $job.SourceRelativePath }

                        Write-Output ("[DELETE] Source info: Share='{0}' | Path='{1}'" -f $srcShare, $srcRel)

                            if ($WhatIfOnly) {
                            Write-Output ("[WHATIF] Would delete source: {0}/{1}" -f $srcShare, $srcRel)
                            } else {
                            try {
                                if (-not $srcShare -or -not $srcRel) {
                                    throw [System.Exception] "Missing SourceShare or SourceRelativePath for job; cannot delete."
                                }
                                # Ensure we remove only the exact relative path (no leading slash)
                                $cleanPath = $srcRel.TrimStart('/')

                                Remove-AzStorageFile -ShareName $srcShare -Path $cleanPath -Context $Context -ErrorAction Stop
                                Write-Output ("[DELETE] Successfully deleted source: {0}/{1}" -f $srcShare, $cleanPath)
                                
                                # Create stub file if requested
                                if ($CreateStubFiles) {
                                    $stubPath = $cleanPath + $StubFileSuffix
                                    try {
                                        $stubContent = "This file has been archived to Azure Blob Storage. Original file: $cleanPath"
                                        # Create a temporary file with the content
                                        $tempFile = [System.IO.Path]::GetTempFileName()
                                        $stubContent | Out-File -FilePath $tempFile -Encoding UTF8
                                        
                                        # Upload the temporary file to Azure File Share
                                        Set-AzStorageFileContent -ShareName $srcShare -Path $stubPath -Source $tempFile -Context $Context -ErrorAction Stop
                                        
                                        # Clean up temporary file
                                        Remove-Item $tempFile -Force
                                        
                                        Write-Output ("[STUB] Created stub file: {0}/{1}" -f $srcShare, $stubPath)
                                    } catch {
                                        Write-Warning ("[STUB] Failed to create stub file '{0}/{1}': {2}" -f $srcShare, $stubPath, $_.Exception.Message)
                                        # Clean up temporary file if it exists
                                        if (Test-Path $tempFile) { Remove-Item $tempFile -Force }
                                    }
                                }
                                
                                $deleted++
                            } catch {
                                $failed++
                                Write-Warning ("[DELETE] Failed to delete source '{0}/{1}': {2}" -f $srcShare, $srcRel, $_.Exception.Message)
                            }
                        }
                    } else {
                        Write-Output ("[DELETE] DeleteAfterVerify is false, skipping deletion")
                    }

                    $toRemove.Add($job)
                } else {
                        $failed++
                    Write-Output ("[VERIFY] Blob not found: {0}/{1}" -f $container, $blobName)
                    $toRemove.Add($job)
                }
            } catch {
                $failed++
                Write-Warning ("[VERIFY] Error checking copy state for job: {0}" -f $_.Exception.Message)
                # defensive: add to remove list to avoid infinite loop for bad job
                $toRemove.Add($job)
            }
        }

        # Remove processed items from pending (by instance equality)
        foreach ($done in $toRemove) { [void]$pending.Remove($done) }

        if ($pending.Count -gt 0) {
            Start-Sleep -Seconds $PollIntervalSec
        }
    }

    return [pscustomobject]@{
        Verified = $verified
        Deleted  = $deleted
        Failed   = $failed
        TimedOut = $timedOut
    }
}

# Test function recognition
Write-Output ("[DEBUG] Function definition complete. Testing function recognition...")
$testResult = Get-Command Wait-Verify-And-Delete -ErrorAction SilentlyContinue
Write-Output ("[DEBUG] Function recognition test: {0}" -f ($testResult -ne $null))

# Simple test function
function Test-SimpleFunction {
    param([string]$Message)
    Write-Output ("[TEST-FUNC] Simple function called with: {0}" -f $Message)
    return "Success"
}

# Test the simple function
Write-Output ("[DEBUG] Testing simple function...")
$testResult = Test-SimpleFunction -Message "Hello World"
Write-Output ("[DEBUG] Simple function result: {0}" -f $testResult)

# Test Wait-Verify-And-Delete with empty jobs
Write-Output ("[DEBUG] Testing Wait-Verify-And-Delete with empty jobs...")
$emptyJobs = [System.Collections.Generic.List[object]]::new()
$testResult = Wait-Verify-And-Delete -Context $null -Jobs $emptyJobs -TimeoutSec 10 -PollIntervalSec 1 -BatchSize 10 -DeleteAfterVerify $false -WhatIfOnly $false
Write-Output ("[DEBUG] Wait-Verify-And-Delete test result: {0}" -f $testResult)

# Test Queue-BlobCopiesForOldFiles function recognition
Write-Output ("[DEBUG] Testing Queue-BlobCopiesForOldFiles function recognition...")
$testResult = Get-Command Queue-BlobCopiesForOldFiles -ErrorAction SilentlyContinue
Write-Output ("[DEBUG] Queue-BlobCopiesForOldFiles function recognition: {0}" -f ($testResult -ne $null))

# Test Queue-BlobCopiesForOldFiles function call
Write-Output ("[DEBUG] Testing Queue-BlobCopiesForOldFiles function call...")
try {
    $testResult = Queue-BlobCopiesForOldFiles -Context $null -ShareName "test" -ContainerName "test" -OlderThanUtc (Get-Date) -ShareSas "test" -WhatIfOnly $true
    Write-Output ("[DEBUG] Queue-BlobCopiesForOldFiles test call result: {0}" -f $testResult)
} catch {
    Write-Output ("[DEBUG] Queue-BlobCopiesForOldFiles test call failed: {0}" -f $_.Exception.Message)
}

# ----------------------------
# Main
# ----------------------------
$ErrorActionPreference = 'Stop'
$start = Get-Date
Connect-Cloud

Write-Output "Starting archival copy from Azure Files to Blob..."
Write-Output "Subscription: $((Get-AzContext).Subscription.Id)"
Write-Output "Storage Account: $StorageAccountName | Resource Group: $ResourceGroupName"
Write-Output "File Share: $FileShareName | Destination Container: $BlobContainerName"
Write-Output "WhatIfOnly: $WhatIfOnly | DeleteAfterVerify: $DeleteAfterVerify"
Write-Output "Stub Files: $CreateStubFiles | Suffix: $StubFileSuffix"
Write-Output ("Params: SingleFileTestRelativePath='{0}' | SingleFileCopyRelativePath='{1}' | DebugTargetRelativePath='{2}' | ArchiveOlderThanYears={3} | FolderPath='{4}' | BlobTier='{5}'" -f $SingleFileTestRelativePath, $SingleFileCopyRelativePath, $DebugTargetRelativePath, $ArchiveOlderThanYears, $FolderPath, $BlobTier)

$ctxInfo = Get-StorageContexts
$ctx = $ctxInfo.Context

Ensure-Container -Context $ctx -ContainerName $BlobContainerName

$olderThanUtc = (Get-Date).ToUniversalTime().AddYears(-$ArchiveOlderThanYears)
$shareSas = New-FileShareSas -Context $ctx -ShareName $FileShareName -ExpiryUtc ((Get-Date).ToUniversalTime().AddHours(8))

# Single-file copy path (if provided)
if ($SingleFileCopyRelativePath) {
    $normalized = $SingleFileCopyRelativePath.Trim().Trim('"')
    $root = "https://$StorageAccountName.file.core.windows.net/$FileShareName/"
    $fileUrl = "$root$normalized$shareSas"
    $ntfs = Get-FileLastWriteTimeUtc -FileUrlWithSas $fileUrl
    if (-not $ntfs) {
        Write-Output ("[SINGLE-COPY] NTFS timestamp unavailable for {0}; skipping." -f $normalized)
        return
    }
    Write-Output ("[SINGLE-COPY] Path={0} | NTFS LastWrite={1:o} | Cutoff={2:o}" -f $normalized, $ntfs, $olderThanUtc)
    if ($ntfs -gt $olderThanUtc) {
        Write-Output "[SINGLE-COPY] File newer than cutoff; skipping."
        return
    }
    if ($WhatIfOnly) {
        Write-Output ("[WHATIF] Would copy: share '{0}' file '{1}' -> {2}/{3}" -f $FileShareName, $normalized, $BlobContainerName, $normalized)
        return
    }
    try {
        $copyJob = Start-AzStorageBlobCopy -SrcShareName $FileShareName -SrcFilePath $normalized -DestContainer $BlobContainerName -DestBlob $normalized -Context $ctx -Force
        Write-Output ("[SINGLE-COPY] Queued copy -> {0}/{1}" -f $BlobContainerName, $normalized)
        
        # Create a job object for verification
        $singleJob = [pscustomobject]@{
            Container = $BlobContainerName
            BlobName = $normalized
            CopyId = $copyJob.CopyId
            SourceShare = $FileShareName
            SourceRelativePath = $normalized
            BlobTier = $BlobTier
        }
        
        # Wait for copy to complete and then verify/delete
        Write-Output ("[SINGLE-COPY] Waiting for copy to complete...")
        $timeout = 300 # 5 minutes timeout
        $startTime = Get-Date
        $completed = $false
        
        while (-not $completed -and ((Get-Date) - $startTime).TotalSeconds -lt $timeout) {
            Start-Sleep -Seconds 5
            try {
                $blob = Get-AzStorageBlob -Container $BlobContainerName -Blob $normalized -Context $ctx -ErrorAction SilentlyContinue
                if ($blob) {
                    $completed = $true
                    Write-Output ("[SINGLE-COPY] Copy completed, verifying...")
                    
                    # Verify the copy
                    Write-Output ("[VERIFY] Checking blob existence for {0}/{1}" -f $BlobContainerName, $normalized)
                    if ($blob) {
                        Write-Output ("[VERIFY] Verified copy: {0}/{1}" -f $BlobContainerName, $normalized)
                        
                        # Set blob tier if specified
                        if ($BlobTier -and $BlobTier -ne "Hot") {
                            try {
                                $currentTier = $blob.BlobTier
                                if ($currentTier -ne $BlobTier) {
                                    # Use Azure Storage SDK approach (most reliable in Automation Account)
                                    try {
                                        # Load Azure.Storage.Blobs SDK if available
                                        if (-not ([Type]::GetType('Azure.Storage.Blobs.BlobClient, Azure.Storage.Blobs'))) {
                                            try { [void][System.Reflection.Assembly]::Load('Azure.Storage.Blobs') } catch { }
                                        }
                                        
                                        if ([Type]::GetType('Azure.Storage.Blobs.BlobClient, Azure.Storage.Blobs')) {
                                            # Use Azure Storage SDK with storage account key
                                            $storageKey = $ctx.StorageAccount.Credentials.ExportBase64EncodedKey()
                                            $credential = [Azure.Storage.StorageSharedKeyCredential]::new($ctx.StorageAccountName, $storageKey)
                                            $blobUri = "https://$($ctx.StorageAccountName).blob.core.windows.net/$BlobContainerName/$normalized"
                                            $blobClient = [Azure.Storage.Blobs.BlobClient]::new([System.Uri]$blobUri, $credential)
                                            $blobClient.SetAccessTier($BlobTier)
                                            Write-Output ("[BLOB-TIER] Set tier for {0}/{1} to {2} using Azure Storage SDK" -f $BlobContainerName, $normalized, $BlobTier)
                                        } else {
                                            Write-Warning ("[BLOB-TIER] Azure Storage SDK not available, skipping tier setting for {0}/{1}" -f $BlobContainerName, $normalized)
                                        }
                                    } catch {
                                        Write-Warning ("[BLOB-TIER] Failed to set tier for {0}/{1}: {2}" -f $BlobContainerName, $normalized, $_.Exception.Message)
                                    }
                                } else {
                                    Write-Output ("[BLOB-TIER] {0}/{1} already has tier {2}" -f $BlobContainerName, $normalized, $BlobTier)
                                }
                            } catch {
                                Write-Warning ("[BLOB-TIER] Failed to set tier for {0}/{1}: {2}" -f $BlobContainerName, $normalized, $_.Exception.Message)
                            }
                        }
                        
                        if ($DeleteAfterVerify) {
                            Write-Output ("[DELETE] DeleteAfterVerify is true; WhatIfOnly={0}" -f $WhatIfOnly)
                            Write-Output ("[DELETE] Source info: Share='{0}' | Path='{1}'" -f $FileShareName, $normalized)
                            
                            if ($WhatIfOnly) {
                                Write-Output ("[WHATIF] Would delete source: {0}/{1}" -f $FileShareName, $normalized)
                            } else {
                                try {
                                    $cleanPath = $normalized.TrimStart('/')
                                    Remove-AzStorageFile -ShareName $FileShareName -Path $cleanPath -Context $ctx -ErrorAction Stop
                                    Write-Output ("[DELETE] Successfully deleted source: {0}/{1}" -f $FileShareName, $cleanPath)
                                    
                                    # Create stub file if requested
                                    if ($CreateStubFiles) {
                                        $stubPath = $cleanPath + $StubFileSuffix
                                        try {
                                            $stubContent = "This file has been archived to Azure Blob Storage. Original file: $cleanPath"
                                            $tempFile = [System.IO.Path]::GetTempFileName()
                                            $stubContent | Out-File -FilePath $tempFile -Encoding UTF8
                                            Set-AzStorageFileContent -ShareName $FileShareName -Path $stubPath -Source $tempFile -Context $ctx -ErrorAction Stop
                                            Remove-Item $tempFile -Force
                                            Write-Output ("[STUB] Created stub file: {0}/{1}" -f $FileShareName, $stubPath)
                                        } catch {
                                            Write-Warning ("[STUB] Failed to create stub file '{0}/{1}': {2}" -f $FileShareName, $stubPath, $_.Exception.Message)
                                            if (Test-Path $tempFile) { Remove-Item $tempFile -Force }
                                        }
                                    }
                                } catch {
                                    Write-Warning ("[DELETE] Failed to delete source '{0}/{1}': {2}" -f $FileShareName, $normalized, $_.Exception.Message)
                                }
                            }
                        }
                    }
                }
            } catch {
                Write-Output ("[SINGLE-COPY] Still waiting for copy to complete...")
            }
        }
        
        if (-not $completed) {
            Write-Warning ("[SINGLE-COPY] Copy verification timed out after {0} seconds" -f $timeout)
        }
        
    } catch {
        Write-Output ("[SINGLE-COPY] Failed to queue copy: {0}" -f $_.Exception.Message)
    }
    return
}

# Single-file test path (if provided)
if ($SingleFileTestRelativePath) {
    $root = "https://$StorageAccountName.file.core.windows.net/$FileShareName/"
    $fileUrl = "$root$SingleFileTestRelativePath$shareSas"
    $ntfs = Get-FileLastWriteTimeUtc -FileUrlWithSas $fileUrl
    $svc = $null
    try {
        # Try to get service LastModified via REST HEAD
        $headers = @{ 'x-ms-version' = '2023-11-03' }
        $resp = Invoke-WebRequest -Method Head -Uri $fileUrl -Headers $headers -UseBasicParsing -TimeoutSec 60 -ErrorAction Stop
        $lm = $resp.Headers['Last-Modified']
        if ($lm) { $svc = ([DateTimeOffset]::Parse($lm)).UtcDateTime }
    } catch {}
    Write-Output ("[SINGLE-TEST] Path={0} | NTFS LastWrite={1:o} | Service LastModified={2:o} | Cutoff={3:o}" -f $SingleFileTestRelativePath, $ntfs, $svc, $olderThanUtc)
    if ($ntfs -and ($ntfs -le $olderThanUtc)) {
        Write-Output "[SINGLE-TEST] This file qualifies for copy based on NTFS timestamp."
    } elseif ($svc -and ($svc -le $olderThanUtc)) {
        Write-Output "[SINGLE-TEST] This file qualifies for copy based on service LastModified fallback."
    } else {
        Write-Output "[SINGLE-TEST] This file does NOT qualify based on available timestamps."
    }
    Write-Output "[SINGLE-TEST] Exiting early by request."
    return
}

Write-Output ("Copying files last modified on/before: {0} (UTC) (older than {1} years)" -f $olderThanUtc.ToString("u"), $ArchiveOlderThanYears)
if (-not [string]::IsNullOrEmpty($FolderPath)) {
    Write-Output ("Processing folder: {0}" -f $FolderPath)
} else {
    Write-Output "Processing all folders in the file share"
}
Write-Output ("[DEBUG] About to call Get-AllFilesRecursive with: ShareName='{0}', UseSasListing={1}, AccountName='{2}', FolderPath='{3}'" -f $FileShareName, $UseSasListing, $StorageAccountName, $FolderPath)

# Test enumeration directly first
Write-Output "[DEBUG] Testing enumeration directly..."
try {
    $testFiles = Get-AllFilesRecursive -Context $ctx -ShareName $FileShareName -ShareSas $shareSas -UseSasListing:$UseSasListing -AccountName $StorageAccountName -FolderPath $FolderPath
    $testCount = 0
    foreach ($f in $testFiles) { $testCount++ }
    Write-Output ("[DEBUG] Direct enumeration returned {0} items" -f $testCount)
} catch {
    Write-Output ("[ENUM-DIRECT-ERROR] {0}" -f $_.Exception.Message)
    return
}

# Use the existing Queue-BlobCopiesForOldFiles function which handles verification properly
Write-Output "[DEBUG] Using Queue-BlobCopiesForOldFiles function..."
try {
    Write-Output ("[DEBUG] Calling Queue-BlobCopiesForOldFiles with WhatIfOnly={0}" -f $WhatIfOnly)
    $queueResult = Queue-BlobCopiesForOldFiles -Context $ctx -ShareName $FileShareName -ContainerName $BlobContainerName -OlderThanUtc $olderThanUtc -ShareSas $shareSas -MaxQueuePerBatch $MaxQueuePerBatch -WhatIfOnly $WhatIfOnly -UseSasListing $UseSasListing -FolderPath $FolderPath
    Write-Output ("[DEBUG] Queue-BlobCopiesForOldFiles returned: QueuedCount={0}, JobsCount={1}" -f $queueResult.QueuedCount, $queueResult.Jobs.Count)
Write-Output ("Queued {0} copy operation(s)." -f $queueResult.QueuedCount)
} catch {
    Write-Output ("[QUEUE-ERROR] Failed to queue copies: {0}" -f $_.Exception.Message)
    Write-Output ("[QUEUE-ERROR] Stack trace: {0}" -f $_.ScriptStackTrace)
    $queueResult = [pscustomobject]@{ QueuedCount = 0; Jobs = @() }
}

if (-not $WhatIfOnly -and $queueResult.QueuedCount -gt 0) {
    Write-Output "Verifying server-side copies..."
    Write-Output ("[DEBUG] Jobs to verify: {0}" -f $queueResult.Jobs.Count)
    if ($queueResult.Jobs.Count -gt 0) {
        Write-Output ("[DEBUG] First job: Container={0}, BlobName={1}" -f $queueResult.Jobs[0].Container, $queueResult.Jobs[0].BlobName)
        Write-Output ("[DEBUG] First job type: {0}" -f $queueResult.Jobs[0].GetType().FullName)
        Write-Output ("[DEBUG] First job properties: {0}" -f ($queueResult.Jobs[0] | Get-Member -MemberType Property | Select-Object -ExpandProperty Name | Sort-Object))
    }
    Write-Output ("[DEBUG] Calling Wait-Verify-And-Delete with DeleteAfterVerify={0}, WhatIfOnly={1}" -f $DeleteAfterVerify, $WhatIfOnly)
    Write-Output ("[DEBUG] Function exists: {0}" -f (Get-Command Wait-Verify-And-Delete -ErrorAction SilentlyContinue -ne $null))
    Write-Output ("[DEBUG] Jobs parameter type: {0}, Count: {1}" -f $queueResult.Jobs.GetType().FullName, $queueResult.Jobs.Count)
    try {
        $verifyStats = Wait-Verify-And-Delete -Context $ctx -Jobs $queueResult.Jobs -TimeoutSec $VerifyTimeoutSec -PollIntervalSec $VerifyPollIntervalSec -BatchSize $VerifyBatchSize -DeleteAfterVerify $DeleteAfterVerify -WhatIfOnly $WhatIfOnly -CreateStubFiles $CreateStubFiles -StubFileSuffix $StubFileSuffix
        Write-Output ("[DEBUG] Function call completed successfully")
    } catch {
        Write-Output ("[DEBUG] Function call failed: {0}" -f $_.Exception.Message)
        $verifyStats = [pscustomobject]@{ Verified = 0; Deleted = 0; Failed = 0; TimedOut = 0 }
    }

    Write-Output ("Verification complete. Verified: {0}, Deleted: {1}, Failed: {2}, TimedOut: {3}" -f `
        $verifyStats.Verified, $verifyStats.Deleted, $verifyStats.Failed, $verifyStats.TimedOut)
} elseif ($WhatIfOnly) {
    Write-Output "[WHATIF] Skipping verification and deletion."
}

$elapsed = (Get-Date) - $start
Write-Output ("Done. Elapsed: {0}" -f $elapsed.ToString())
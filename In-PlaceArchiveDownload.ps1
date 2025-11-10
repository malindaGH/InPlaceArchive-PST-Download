# ===================Script========================
# Get The Email Address of the Mailbox
$addressOrSite = Read-Host "Enter The Email Address of The Mailbox "

# Validate email input
if ([string]::IsNullOrWhiteSpace($addressOrSite) -or $addressOrSite -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
    Write-Host "`n[ERROR] Invalid email address entered. Exiting script..." -ForegroundColor Red
    return
}

# Define global Variable to use in entire Script.
$global:folderQueries = @()

# ===========================================
# Authenticate with Exchange Online and Retrieve Folder IDs
try {
    Write-Host "`nConnecting to Exchange Online PowerShell..." -ForegroundColor Yellow
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
} catch {
    Write-Host "`n[ERROR] Failed to connect to Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
    return
}

try {
    Write-Host "`nRetrieving mailbox folder statistics for $addressOrSite ..." -ForegroundColor Yellow
    $folderStatistics = Get-MailboxFolderStatistics -Identity $addressOrSite -Archive -ErrorAction Stop
} catch {
    Write-Host "`n[ERROR] Failed to retrieve mailbox folder data. Verify the email and permissions." -ForegroundColor Red
    return
}

if (-not $folderStatistics) {
    Write-Host "`n[ERROR] No folders found for the specified mailbox. Exiting script..." -ForegroundColor Red
    return
}

foreach ($folderStatistic in $folderStatistics) {
    $folderId = $folderStatistic.FolderId
    $folderPath = $folderStatistic.FolderPath
    $encoding = [System.Text.Encoding]::GetEncoding("us-ascii")
    $nibbler = $encoding.GetBytes("0123456789ABCDEF")
    $folderIdBytes = [Convert]::FromBase64String($folderId)
    $indexIdBytes = New-Object byte[] 48
    $indexIdIdx = 0
    $folderIdBytes | Select-Object -Skip 23 -First 24 | ForEach-Object {
        $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -shr 4]
        $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -band 0xF]
    }
    $folderQuery = "folderid:$($encoding.GetString($indexIdBytes))"
    $folderStat = New-Object PSObject
    Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderPath -Value $folderPath
    Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderQuery -Value $folderQuery
    $global:folderQueries += $folderStat
}

Write-Host "`n-----Exchange Archive Mailbox Folders That Is Going To Search-----" -ForegroundColor Cyan
$global:folderQueries | Format-Table

# ===========================================
# Build ContentMatchQuery string from all folder IDs
$folderQueryString = ($global:folderQueries | ForEach-Object { $_.FolderQuery }) -join " OR "

if (-not $folderQueryString) {
    Write-Host "`n[ERROR] Failed to build folder ID query string. Exiting..." -ForegroundColor Red
    return
}

# ===========================================
# Create Compliance Search Command
$searchName = Read-Host "Enter an Unique Search Name For the Compliance Search "
if ([string]::IsNullOrWhiteSpace($searchName)) {
    Write-Host "`n[ERROR] Search name cannot be empty. Exiting..." -ForegroundColor Red
    return
}

$complianceSearchCommand = "New-ComplianceSearch -Name `"$searchName`" -ExchangeLocation `"$addressOrSite`" -ContentMatchQuery `"$folderQueryString`""
Write-Host "`nGenerated Compliance Search command (not executed):`n" -ForegroundColor Yellow
#Write-Host $complianceSearchCommand -ForegroundColor Cyan

# ===========================================
# Connect to Purview (IPP) session
try {
    if (-not (Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.ComputerName -like "*compliance*" })) {
        Write-Host "`nConnecting to Microsoft Purview (Compliance Center) PowerShell..." -ForegroundColor Yellow
        Connect-IPPSSession -WarningAction SilentlyContinue -ErrorAction Stop
    } else {
        Write-Host "Already connected to IPP session." -ForegroundColor Green
    }
} catch {
    Write-Host "`n[ERROR] Failed to connect to Microsoft Purview PowerShell: $($_.Exception.Message)" -ForegroundColor Red
    return
}


# ===========================================
# Ask to create and optionally start the Compliance Search
$runChoice = Read-Host "`nDo you want to execute the Compliance Search command now? (Y/N)"
if ($runChoice -match '^[Yy]$') {

    try {
        # Run the compliance search creation command with -ErrorAction Stop
        Invoke-Expression "$complianceSearchCommand -ErrorAction Stop | Out-Null"
        Write-Host "`nCompliance Search command executed successfully." -ForegroundColor Green

        # Nested prompt to start the search
        $startChoice = Read-Host "`nDo you want to start the Compliance Search now? (Y/N)"
        if ($startChoice -match '^[Yy]$') {
            try {
                Start-ComplianceSearch -Identity $searchName -ErrorAction Stop
                Write-Host "`nCompliance Search started successfully." -ForegroundColor Green
                Get-ComplianceSearch -Identity $searchName | Format-Table
            } catch {
                Write-Host "`n[ERROR] Failed to start Compliance Search: $($_.Exception.Message)" -ForegroundColor Red
                return
            }
        } else {
            Write-Host "`nCompliance Search not started. You can run manually using:`n" -ForegroundColor Yellow
            Write-Host "Start-ComplianceSearch -Identity $searchName" -ForegroundColor Cyan
            return
        }

    } catch {
        Write-Host "`n[ERROR] Failed to create Compliance Search: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

} else {
    Write-Host "`nCompliance Search not created. You can run manually using:`n" -ForegroundColor Yellow
    Write-Host $complianceSearchCommand -ForegroundColor Cyan
    return

}

<#
.SYNOPSIS
    Exports Exchange mailboxes and their email aliases to CSV (simple format - one row per mailbox)

.DESCRIPTION
    This script reads all user mailboxes from Exchange and exports information
    including all email aliases to a CSV file. All aliases are combined in a single column
    for easier overview and reporting.
    
    Features:
    - Filter by Active Directory OU
    - Filter by primary SMTP domain
    - Comprehensive mailbox information (size, item count, last logon)
    - Error handling and progress indicators
    - UTF-8 encoding support

.PARAMETER OutputPath
    Path for the output CSV file. Default: MailboxAliases_Simple_YYYYMMDD_HHMMSS.csv

.PARAMETER FilterByOU
    Filter mailboxes by Active Directory Organizational Unit (e.g., "OU=Users,DC=company,DC=com")

.PARAMETER FilterByDomain
    Filter mailboxes by primary SMTP domain (e.g., "company.com")

.PARAMETER IncludeStatistics
    Include mailbox statistics (size, item count, last logon). May slow down execution.

.PARAMETER AliasSeparator
    Character to separate multiple aliases. Default: "; "

.EXAMPLE
    .\Export-MailboxAliases-Simple.ps1
    Exports all mailboxes with basic information

.EXAMPLE
    .\Export-MailboxAliases-Simple.ps1 -FilterByDomain "company.com" -IncludeStatistics
    Exports only mailboxes with primary domain "company.com" including statistics

.EXAMPLE
    .\Export-MailboxAliases-Simple.ps1 -FilterByOU "OU=Sales,OU=Users,DC=company,DC=com" -AliasSeparator " | "
    Exports only mailboxes from the Sales OU with custom alias separator

.NOTES
    Requires Exchange Management Shell or Exchange Online PowerShell module
    Author: PowerShell Exchange Export Script
    Version: 1.0
#>

param(
    [string]$OutputPath = "",
    [string]$FilterByOU = "",
    [string]$FilterByDomain = "",
    [switch]$IncludeStatistics = $false,
    [string]$AliasSeparator = "; "
)

# Generate default output filename if not provided
if ([string]::IsNullOrEmpty($OutputPath)) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $OutputPath = "MailboxAliases_Simple_$timestamp.csv"
}

Write-Host "Exchange Mailbox and Aliases Export (Simple Format)" -ForegroundColor Green
Write-Host "===============================================" -ForegroundColor Green
Write-Host "Output file: $OutputPath" -ForegroundColor Yellow

# Initialize results array
$results = @()

try {
    # Build filter parameters for Get-Mailbox
    $mailboxParams = @{
        RecipientTypeDetails = 'UserMailbox'
        ResultSize = 'Unlimited'
    }

    # Add OU filter if specified
    if (![string]::IsNullOrEmpty($FilterByOU)) {
        $mailboxParams.OrganizationalUnit = $FilterByOU
        Write-Host "Filtering by OU: $FilterByOU" -ForegroundColor Yellow
    }

    # Get mailboxes
    Write-Host "Retrieving mailboxes..." -ForegroundColor Cyan
    $mailboxes = Get-Mailbox @mailboxParams

    # Apply domain filter if specified
    if (![string]::IsNullOrEmpty($FilterByDomain)) {
        Write-Host "Filtering by primary domain: $FilterByDomain" -ForegroundColor Yellow
        $mailboxes = $mailboxes | Where-Object { $_.PrimarySmtpAddress -like "*@$FilterByDomain" }
    }

    $totalMailboxes = $mailboxes.Count
    Write-Host "Found $totalMailboxes mailboxes to process" -ForegroundColor Green

    if ($totalMailboxes -eq 0) {
        Write-Warning "No mailboxes found matching the specified criteria."
        return
    }

    $counter = 0
    foreach ($mailbox in $mailboxes) {
        $counter++
        $percentComplete = [math]::Round(($counter / $totalMailboxes) * 100, 2)
        
        Write-Progress -Activity "Processing Mailboxes" -Status "Processing $($mailbox.DisplayName) ($counter of $totalMailboxes)" -PercentComplete $percentComplete

        try {
            # Get mailbox statistics if requested
            $mailboxStats = $null
            if ($IncludeStatistics) {
                try {
                    $mailboxStats = Get-MailboxStatistics -Identity $mailbox.Identity -ErrorAction SilentlyContinue
                } catch {
                    Write-Warning "Could not retrieve statistics for $($mailbox.DisplayName): $($_.Exception.Message)"
                }
            }

            # Get all email addresses for this mailbox
            $emailAddresses = $mailbox.EmailAddresses | Where-Object { $_.PrefixString -eq "smtp" -or $_.PrefixString -eq "SMTP" }
            
            # Separate primary and aliases
            $primaryEmail = $mailbox.PrimarySmtpAddress.ToString()
            $aliases = $emailAddresses | Where-Object { $_.SmtpAddress -ne $primaryEmail } | ForEach-Object { $_.SmtpAddress }

            # Combine all aliases into a single string
            $allAliases = if ($aliases.Count -gt 0) { 
                ($aliases | Sort-Object) -join $AliasSeparator 
            } else { 
                "" 
            }

            # Create mailbox info object
            $mailboxInfo = [PSCustomObject]@{
                DisplayName = $mailbox.DisplayName
                UserPrincipalName = $mailbox.UserPrincipalName
                SamAccountName = $mailbox.SamAccountName
                PrimarySmtpAddress = $primaryEmail
                AllAliases = $allAliases
                AliasCount = $aliases.Count
                MailboxDatabase = $mailbox.Database
                OrganizationalUnit = $mailbox.OrganizationalUnit
                MailboxSizeMB = if ($mailboxStats -and $mailboxStats.TotalItemSize) { 
                    [math]::Round(($mailboxStats.TotalItemSize.Value.ToBytes() / 1MB), 2) 
                } else { "N/A" }
                ItemCount = if ($mailboxStats) { $mailboxStats.ItemCount } else { "N/A" }
                LastLogonTime = if ($mailboxStats) { $mailboxStats.LastLogonTime } else { "N/A" }
                WhenCreated = $mailbox.WhenCreated
                HiddenFromAddressListsEnabled = $mailbox.HiddenFromAddressListsEnabled
            }

            $results += $mailboxInfo

        } catch {
            Write-Warning "Error processing mailbox $($mailbox.DisplayName): $($_.Exception.Message)"
            
            # Add error row
            $errorInfo = [PSCustomObject]@{
                DisplayName = $mailbox.DisplayName
                UserPrincipalName = $mailbox.UserPrincipalName
                SamAccountName = $mailbox.SamAccountName
                PrimarySmtpAddress = "ERROR"
                AllAliases = "Error: $($_.Exception.Message)"
                AliasCount = 0
                MailboxDatabase = $mailbox.Database
                OrganizationalUnit = $mailbox.OrganizationalUnit
                MailboxSizeMB = "N/A"
                ItemCount = "N/A"
                LastLogonTime = "N/A"
                WhenCreated = $mailbox.WhenCreated
                HiddenFromAddressListsEnabled = $mailbox.HiddenFromAddressListsEnabled
            }
            $results += $errorInfo
        }
    }

    Write-Progress -Activity "Processing Mailboxes" -Completed

    # Export to CSV
    Write-Host "Exporting $($results.Count) records to CSV..." -ForegroundColor Cyan
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

    Write-Host "Export completed successfully!" -ForegroundColor Green
    Write-Host "File saved: $OutputPath" -ForegroundColor Yellow
    Write-Host "Total mailboxes: $($results.Count)" -ForegroundColor Yellow

} catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    Write-Error $_.ScriptStackTrace
}

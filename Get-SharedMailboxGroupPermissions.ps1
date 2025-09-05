<#
.SYNOPSIS
    Exports shared mailbox permissions for groups to CSV file.

.DESCRIPTION
    This script connects to Exchange Online, retrieves all shared mailboxes,
    and exports information about which groups have full access permissions
    to each shared mailbox into a CSV file.

.PARAMETER OutputPath
    Specifies the path for the output CSV file. Default is current directory.

.EXAMPLE
    .\Get-SharedMailboxGroupPermissions.ps1
    
.EXAMPLE
    .\Get-SharedMailboxGroupPermissions.ps1 -OutputPath "C:\Reports\SharedMailboxReport.csv"
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "SharedMailboxGroupPermissions_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Import Exchange Online module
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Write-Host "Exchange Online module imported successfully." -ForegroundColor Green
}
catch {
    Write-Error "Failed to import Exchange Online module. Please install it using: Install-Module -Name ExchangeOnlineManagement"
    exit 1
}

# Connect to Exchange Online if not already connected
try {
    $session = Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened"}
    if (-not $session) {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
        Connect-ExchangeOnline -ShowProgress $true
        Write-Host "Connected to Exchange Online successfully." -ForegroundColor Green
    } else {
        Write-Host "Already connected to Exchange Online." -ForegroundColor Green
    }
}
catch {
    Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
    exit 1
}

# Initialize results array
$results = @()
$errorCount = 0

try {
    # Get all shared mailboxes
    Write-Host "Retrieving shared mailboxes..." -ForegroundColor Yellow
    $sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
    
    if ($sharedMailboxes.Count -eq 0) {
        Write-Warning "No shared mailboxes found in the organization."
        return
    }
    
    Write-Host "Found $($sharedMailboxes.Count) shared mailboxes. Processing permissions..." -ForegroundColor Green
    
    # Process each shared mailbox
    $counter = 0
    foreach ($mailbox in $sharedMailboxes) {
        $counter++
        $percentComplete = [math]::Round(($counter / $sharedMailboxes.Count) * 100, 2)
        Write-Progress -Activity "Processing Shared Mailboxes" -Status "Processing $($mailbox.DisplayName) ($counter of $($sharedMailboxes.Count))" -PercentComplete $percentComplete
        
        try {
            # Get mailbox permissions
            $permissions = Get-MailboxPermission -Identity $mailbox.Identity | 
                Where-Object { 
                    $_.AccessRights -contains "FullAccess" -and 
                    $_.IsInherited -eq $false -and 
                    $_.User -ne "NT AUTHORITY\SELF" -and
                    $_.User -notlike "S-1-5-*"
                }
            
            # Check if any permissions exist
            if ($permissions) {
                foreach ($permission in $permissions) {
                    # Verify if the trustee is a group
                    try {
                        $recipient = Get-Recipient -Identity $permission.User -ErrorAction SilentlyContinue
                        if ($recipient -and ($recipient.RecipientType -like "*Group*" -or $recipient.RecipientTypeDetails -like "*Group*")) {
                            $results += [PSCustomObject]@{
                                'Shared Mailbox Name' = $mailbox.DisplayName
                                'Shared Mailbox Email' = $mailbox.PrimarySmtpAddress
                                'Shared Mailbox Alias' = $mailbox.Alias
                                'Group with Full Access' = $permission.User
                                'Permission Type' = ($permission.AccessRights -join ", ")
                                'Is Inherited' = $permission.IsInherited
                                'Date Checked' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                            }
                        }
                    }
                    catch {
                        # If we can't resolve the recipient, still include it but note it
                        $results += [PSCustomObject]@{
                            'Shared Mailbox Name' = $mailbox.DisplayName
                            'Shared Mailbox Email' = $mailbox.PrimarySmtpAddress
                            'Shared Mailbox Alias' = $mailbox.Alias
                            'Group with Full Access' = "$($permission.User) (Could not verify as group)"
                            'Permission Type' = ($permission.AccessRights -join ", ")
                            'Is Inherited' = $permission.IsInherited
                            'Date Checked' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        }
                    }
                }
            } else {
                # No group permissions found for this mailbox
                $results += [PSCustomObject]@{
                    'Shared Mailbox Name' = $mailbox.DisplayName
                    'Shared Mailbox Email' = $mailbox.PrimarySmtpAddress
                    'Shared Mailbox Alias' = $mailbox.Alias
                    'Group with Full Access' = "No groups with full access"
                    'Permission Type' = "N/A"
                    'Is Inherited' = "N/A"
                    'Date Checked' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
            }
        }
        catch {
            $errorCount++
            Write-Warning "Error processing mailbox '$($mailbox.DisplayName)': $($_.Exception.Message)"
            
            # Add error entry to results
            $results += [PSCustomObject]@{
                'Shared Mailbox Name' = $mailbox.DisplayName
                'Shared Mailbox Email' = $mailbox.PrimarySmtpAddress
                'Shared Mailbox Alias' = $mailbox.Alias
                'Group with Full Access' = "ERROR: $($_.Exception.Message)"
                'Permission Type' = "ERROR"
                'Is Inherited' = "ERROR"
                'Date Checked' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
    }
    
    Write-Progress -Activity "Processing Shared Mailboxes" -Completed
    
    # Export results to CSV
    if ($results.Count -gt 0) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "`nReport exported successfully to: $OutputPath" -ForegroundColor Green
        
        # Display summary
        $totalMailboxes = ($results | Where-Object { $_.'Group with Full Access' -ne "ERROR: *" }).Count
        $mailboxesWithGroupAccess = ($results | Where-Object { $_.'Group with Full Access' -ne "No groups with full access" -and $_.'Group with Full Access' -notlike "ERROR:*" }).Count
        $mailboxesWithoutGroupAccess = ($results | Where-Object { $_.'Group with Full Access' -eq "No groups with full access" }).Count
        
        Write-Host "`n=== SUMMARY ===" -ForegroundColor Cyan
        Write-Host "Total shared mailboxes processed: $totalMailboxes" -ForegroundColor White
        Write-Host "Mailboxes with group full access: $mailboxesWithGroupAccess" -ForegroundColor Green
        Write-Host "Mailboxes without group full access: $mailboxesWithoutGroupAccess" -ForegroundColor Yellow
        if ($errorCount -gt 0) {
            Write-Host "Mailboxes with errors: $errorCount" -ForegroundColor Red
        }
        Write-Host "Report location: $OutputPath" -ForegroundColor White
    } else {
        Write-Warning "No data to export."
    }
}
catch {
    Write-Error "An error occurred during script execution: $($_.Exception.Message)"
    exit 1
}
finally {
    Write-Host "`nScript execution completed." -ForegroundColor Green
}

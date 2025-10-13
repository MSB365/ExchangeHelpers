<#
.SYNOPSIS
    Exports all mailbox permissions (Send As, Send On Behalf, Full Access) from on-premises Exchange Server.

.DESCRIPTION
    This script reads all permissions assigned to each mailbox and exports them to a CSV file.
    It also generates an HTML report for easy viewing.

.PARAMETER OutputPath
    The directory where the CSV and HTML files will be saved. Default is the script's directory.

.EXAMPLE
    .\Export-MailboxPermissions.ps1
    .\Export-MailboxPermissions.ps1 -OutputPath "C:\ExchangeReports"
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = $PSScriptRoot
)

# Ensure output directory exists
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# Generate timestamp for file names
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvFile = Join-Path $OutputPath "MailboxPermissions_$timestamp.csv"
$htmlFile = Join-Path $OutputPath "MailboxPermissions_$timestamp.html"

Write-Host "Starting mailbox permissions export..." -ForegroundColor Green
Write-Host "Output directory: $OutputPath" -ForegroundColor Cyan

# Initialize results array
$results = @()

try {
    # Get all mailboxes
    Write-Host "Retrieving all mailboxes..." -ForegroundColor Yellow
    $mailboxes = Get-Mailbox -ResultSize Unlimited | Sort-Object DisplayName
    $totalMailboxes = $mailboxes.Count
    $currentMailbox = 0

    Write-Host "Found $totalMailboxes mailboxes. Processing..." -ForegroundColor Cyan

    foreach ($mailbox in $mailboxes) {
        $currentMailbox++
        $percentComplete = [math]::Round(($currentMailbox / $totalMailboxes) * 100, 2)
        Write-Progress -Activity "Processing Mailboxes" -Status "Processing $($mailbox.DisplayName) ($currentMailbox of $totalMailboxes)" -PercentComplete $percentComplete

        # Get Full Access permissions
        try {
            $fullAccessPerms = Get-MailboxPermission -Identity $mailbox.Identity | 
                Where-Object { $_.User -notlike "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false -and $_.User -notlike "S-1-5-*" }
            
            foreach ($perm in $fullAccessPerms) {
                $results += [PSCustomObject]@{
                    MailboxIdentity = $mailbox.Identity
                    MailboxDisplayName = $mailbox.DisplayName
                    MailboxPrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                    MailboxType = $mailbox.RecipientTypeDetails
                    PermissionType = "FullAccess"
                    TrustedUser = $perm.User
                    TrustedUserDisplay = $perm.User
                    AccessRights = ($perm.AccessRights -join ";")
                    Deny = $perm.Deny
                    InheritanceType = $perm.InheritanceType
                }
            }
        }
        catch {
            Write-Warning "Error getting Full Access permissions for $($mailbox.DisplayName): $_"
        }

        # Get Send As permissions
        try {
            $sendAsPerms = Get-ADPermission -Identity $mailbox.DistinguishedName | 
                Where-Object { $_.ExtendedRights -like "Send-As" -and $_.User -notlike "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false -and $_.User -notlike "S-1-5-*" }
            
            foreach ($perm in $sendAsPerms) {
                $results += [PSCustomObject]@{
                    MailboxIdentity = $mailbox.Identity
                    MailboxDisplayName = $mailbox.DisplayName
                    MailboxPrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                    MailboxType = $mailbox.RecipientTypeDetails
                    PermissionType = "SendAs"
                    TrustedUser = $perm.User
                    TrustedUserDisplay = $perm.User
                    AccessRights = "Send-As"
                    Deny = $perm.Deny
                    InheritanceType = $perm.InheritanceType
                }
            }
        }
        catch {
            Write-Warning "Error getting Send As permissions for $($mailbox.DisplayName): $_"
        }

        # Get Send On Behalf permissions
        try {
            if ($mailbox.GrantSendOnBehalfTo -and $mailbox.GrantSendOnBehalfTo.Count -gt 0) {
                foreach ($user in $mailbox.GrantSendOnBehalfTo) {
                    $results += [PSCustomObject]@{
                        MailboxIdentity = $mailbox.Identity
                        MailboxDisplayName = $mailbox.DisplayName
                        MailboxPrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                        MailboxType = $mailbox.RecipientTypeDetails
                        PermissionType = "SendOnBehalf"
                        TrustedUser = $user
                        TrustedUserDisplay = $user
                        AccessRights = "Send-On-Behalf"
                        Deny = $false
                        InheritanceType = "N/A"
                    }
                }
            }
        }
        catch {
            Write-Warning "Error getting Send On Behalf permissions for $($mailbox.DisplayName): $_"
        }
    }

    Write-Progress -Activity "Processing Mailboxes" -Completed

    # Export to CSV
    Write-Host "`nExporting to CSV..." -ForegroundColor Yellow
    $results | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
    Write-Host "CSV exported to: $csvFile" -ForegroundColor Green

    # Generate HTML Report
    Write-Host "Generating HTML report..." -ForegroundColor Yellow
    
    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Exchange Mailbox Permissions Report</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }
        .header {
            background-color: #0078d4;
            color: white;
            padding: 20px;
            border-radius: 5px;
            margin-bottom: 20px;
        }
        .summary {
            background-color: white;
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .summary-item {
            display: inline-block;
            margin-right: 30px;
            padding: 10px;
        }
        .summary-label {
            font-weight: bold;
            color: #666;
        }
        .summary-value {
            font-size: 24px;
            color: #0078d4;
            font-weight: bold;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            background-color: white;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        th {
            background-color: #0078d4;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: 600;
            position: sticky;
            top: 0;
        }
        td {
            padding: 10px;
            border-bottom: 1px solid #ddd;
        }
        tr:hover {
            background-color: #f0f0f0;
        }
        .permission-type {
            padding: 4px 8px;
            border-radius: 3px;
            font-weight: bold;
            font-size: 12px;
        }
        .full-access {
            background-color: #ff6b6b;
            color: white;
        }
        .send-as {
            background-color: #ffa500;
            color: white;
        }
        .send-on-behalf {
            background-color: #4ecdc4;
            color: white;
        }
        .footer {
            margin-top: 20px;
            text-align: center;
            color: #666;
            font-size: 12px;
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>Exchange Mailbox Permissions Report</h1>
        <p>Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
    </div>
    
    <div class="summary">
        <div class="summary-item">
            <div class="summary-label">Total Permissions</div>
            <div class="summary-value">$($results.Count)</div>
        </div>
        <div class="summary-item">
            <div class="summary-label">Full Access</div>
            <div class="summary-value">$(($results | Where-Object {$_.PermissionType -eq "FullAccess"}).Count)</div>
        </div>
        <div class="summary-item">
            <div class="summary-label">Send As</div>
            <div class="summary-value">$(($results | Where-Object {$_.PermissionType -eq "SendAs"}).Count)</div>
        </div>
        <div class="summary-item">
            <div class="summary-label">Send On Behalf</div>
            <div class="summary-value">$(($results | Where-Object {$_.PermissionType -eq "SendOnBehalf"}).Count)</div>
        </div>
        <div class="summary-item">
            <div class="summary-label">Mailboxes Processed</div>
            <div class="summary-value">$totalMailboxes</div>
        </div>
    </div>
    
    <table>
        <thead>
            <tr>
                <th>Mailbox Display Name</th>
                <th>Mailbox Email</th>
                <th>Mailbox Type</th>
                <th>Permission Type</th>
                <th>Trusted User</th>
                <th>Access Rights</th>
                <th>Deny</th>
            </tr>
        </thead>
        <tbody>
"@

    foreach ($result in $results) {
        $permClass = switch ($result.PermissionType) {
            "FullAccess" { "full-access" }
            "SendAs" { "send-as" }
            "SendOnBehalf" { "send-on-behalf" }
        }
        
        $htmlContent += @"
            <tr>
                <td>$($result.MailboxDisplayName)</td>
                <td>$($result.MailboxPrimarySmtpAddress)</td>
                <td>$($result.MailboxType)</td>
                <td><span class="permission-type $permClass">$($result.PermissionType)</span></td>
                <td>$($result.TrustedUserDisplay)</td>
                <td>$($result.AccessRights)</td>
                <td>$($result.Deny)</td>
            </tr>
"@
    }

    $htmlContent += @"
        </tbody>
    </table>
    
    <div class="footer">
        <p>Exchange Mailbox Permissions Export | CSV File: $csvFile</p>
    </div>
</body>
</html>
"@

    $htmlContent | Out-File -FilePath $htmlFile -Encoding UTF8
    Write-Host "HTML report generated: $htmlFile" -ForegroundColor Green

    # Summary
    Write-Host "`n========== EXPORT SUMMARY ==========" -ForegroundColor Cyan
    Write-Host "Total permissions exported: $($results.Count)" -ForegroundColor White
    Write-Host "  - Full Access: $(($results | Where-Object {$_.PermissionType -eq 'FullAccess'}).Count)" -ForegroundColor White
    Write-Host "  - Send As: $(($results | Where-Object {$_.PermissionType -eq 'SendAs'}).Count)" -ForegroundColor White
    Write-Host "  - Send On Behalf: $(($results | Where-Object {$_.PermissionType -eq 'SendOnBehalf'}).Count)" -ForegroundColor White
    Write-Host "Mailboxes processed: $totalMailboxes" -ForegroundColor White
    Write-Host "====================================" -ForegroundColor Cyan

}
catch {
    Write-Error "An error occurred during export: $_"
    Write-Error $_.Exception.Message
}

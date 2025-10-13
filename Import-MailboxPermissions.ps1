<#
.SYNOPSIS
    Imports mailbox permissions from a CSV file exported by Export-MailboxPermissions.ps1

.DESCRIPTION
    This script reads a CSV file containing mailbox permissions and applies them to mailboxes
    on the target Exchange server. It generates an HTML report showing success and failures.

.PARAMETER CsvPath
    The path to the CSV file containing the permissions to import.

.PARAMETER OutputPath
    The directory where the HTML report will be saved. Default is the script's directory.

.PARAMETER WhatIf
    Shows what would happen if the script runs without actually making changes.

.EXAMPLE
    .\Import-MailboxPermissions.ps1 -CsvPath "C:\Export\MailboxPermissions_20250113_120000.csv"
    .\Import-MailboxPermissions.ps1 -CsvPath "C:\Export\MailboxPermissions_20250113_120000.csv" -WhatIf
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = $PSScriptRoot,
    
    [Parameter(Mandatory=$false)]
    [switch]$WhatIf
)

# Validate CSV file exists
if (-not (Test-Path $CsvPath)) {
    Write-Error "CSV file not found: $CsvPath"
    exit 1
}

# Ensure output directory exists
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# Generate timestamp for report file
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$htmlFile = Join-Path $OutputPath "MailboxPermissionsImport_$timestamp.html"

Write-Host "Starting mailbox permissions import..." -ForegroundColor Green
Write-Host "Source CSV: $CsvPath" -ForegroundColor Cyan
Write-Host "Output directory: $OutputPath" -ForegroundColor Cyan

if ($WhatIf) {
    Write-Host "WHATIF MODE: No changes will be made" -ForegroundColor Yellow
}

# Initialize results tracking
$importResults = @()
$successCount = 0
$failureCount = 0
$skippedCount = 0

try {
    # Import CSV
    Write-Host "`nReading CSV file..." -ForegroundColor Yellow
    $permissions = Import-Csv -Path $CsvPath
    $totalPermissions = $permissions.Count
    $currentPermission = 0

    Write-Host "Found $totalPermissions permissions to import. Processing..." -ForegroundColor Cyan

    foreach ($perm in $permissions) {
        $currentPermission++
        $percentComplete = [math]::Round(($currentPermission / $totalPermissions) * 100, 2)
        Write-Progress -Activity "Importing Permissions" -Status "Processing $($perm.MailboxDisplayName) - $($perm.PermissionType) ($currentPermission of $totalPermissions)" -PercentComplete $percentComplete

        $status = "Unknown"
        $errorMessage = ""
        $action = ""

        try {
            # Verify mailbox exists
            $targetMailbox = Get-Mailbox -Identity $perm.MailboxPrimarySmtpAddress -ErrorAction SilentlyContinue
            
            if (-not $targetMailbox) {
                $status = "Skipped"
                $errorMessage = "Target mailbox not found"
                $skippedCount++
            }
            else {
                # Verify trusted user exists
                $trustedUser = $null
                try {
                    $trustedUser = Get-Recipient -Identity $perm.TrustedUser -ErrorAction SilentlyContinue
                }
                catch {
                    $status = "Skipped"
                    $errorMessage = "Trusted user not found"
                    $skippedCount++
                }

                if ($trustedUser) {
                    # Apply permission based on type
                    switch ($perm.PermissionType) {
                        "FullAccess" {
                            $action = "Add Full Access permission for $($perm.TrustedUser) on $($perm.MailboxDisplayName)"
                            
                            if ($WhatIf) {
                                Write-Host "WHATIF: $action" -ForegroundColor Yellow
                                $status = "WhatIf"
                            }
                            else {
                                Add-MailboxPermission -Identity $targetMailbox.Identity -User $perm.TrustedUser -AccessRights FullAccess -InheritanceType All -ErrorAction Stop | Out-Null
                                $status = "Success"
                                $successCount++
                                Write-Host "✓ $action" -ForegroundColor Green
                            }
                        }
                        
                        "SendAs" {
                            $action = "Add Send As permission for $($perm.TrustedUser) on $($perm.MailboxDisplayName)"
                            
                            if ($WhatIf) {
                                Write-Host "WHATIF: $action" -ForegroundColor Yellow
                                $status = "WhatIf"
                            }
                            else {
                                Add-ADPermission -Identity $targetMailbox.DistinguishedName -User $perm.TrustedUser -ExtendedRights "Send-As" -ErrorAction Stop | Out-Null
                                $status = "Success"
                                $successCount++
                                Write-Host "✓ $action" -ForegroundColor Green
                            }
                        }
                        
                        "SendOnBehalf" {
                            $action = "Add Send On Behalf permission for $($perm.TrustedUser) on $($perm.MailboxDisplayName)"
                            
                            if ($WhatIf) {
                                Write-Host "WHATIF: $action" -ForegroundColor Yellow
                                $status = "WhatIf"
                            }
                            else {
                                # Get current GrantSendOnBehalfTo
                                $currentDelegates = @($targetMailbox.GrantSendOnBehalfTo)
                                
                                # Add new delegate if not already present
                                if ($currentDelegates -notcontains $perm.TrustedUser) {
                                    $currentDelegates += $perm.TrustedUser
                                    Set-Mailbox -Identity $targetMailbox.Identity -GrantSendOnBehalfTo $currentDelegates -ErrorAction Stop
                                    $status = "Success"
                                    $successCount++
                                    Write-Host "✓ $action" -ForegroundColor Green
                                }
                                else {
                                    $status = "Skipped"
                                    $errorMessage = "Permission already exists"
                                    $skippedCount++
                                }
                            }
                        }
                        
                        default {
                            $status = "Failed"
                            $errorMessage = "Unknown permission type: $($perm.PermissionType)"
                            $failureCount++
                        }
                    }
                }
            }
        }
        catch {
            $status = "Failed"
            $errorMessage = $_.Exception.Message
            $failureCount++
            Write-Host "✗ Failed: $action - $errorMessage" -ForegroundColor Red
        }

        # Record result
        $importResults += [PSCustomObject]@{
            MailboxDisplayName = $perm.MailboxDisplayName
            MailboxEmail = $perm.MailboxPrimarySmtpAddress
            MailboxType = $perm.MailboxType
            PermissionType = $perm.PermissionType
            TrustedUser = $perm.TrustedUser
            AccessRights = $perm.AccessRights
            Status = $status
            ErrorMessage = $errorMessage
            Action = $action
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
    }

    Write-Progress -Activity "Importing Permissions" -Completed

    # Generate HTML Report
    Write-Host "`nGenerating HTML report..." -ForegroundColor Yellow
    
    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Exchange Mailbox Permissions Import Report</title>
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
            font-weight: bold;
        }
        .success-value { color: #28a745; }
        .failure-value { color: #dc3545; }
        .skipped-value { color: #ffc107; }
        .total-value { color: #0078d4; }
        
        table {
            width: 100%;
            border-collapse: collapse;
            background-color: white;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin-bottom: 20px;
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
        .status-badge {
            padding: 4px 8px;
            border-radius: 3px;
            font-weight: bold;
            font-size: 12px;
            text-align: center;
            display: inline-block;
            min-width: 70px;
        }
        .status-success {
            background-color: #28a745;
            color: white;
        }
        .status-failed {
            background-color: #dc3545;
            color: white;
        }
        .status-skipped {
            background-color: #ffc107;
            color: black;
        }
        .status-whatif {
            background-color: #17a2b8;
            color: white;
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
        .error-message {
            color: #dc3545;
            font-size: 12px;
            font-style: italic;
        }
        .filter-buttons {
            margin-bottom: 15px;
            background-color: white;
            padding: 15px;
            border-radius: 5px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .filter-btn {
            padding: 8px 16px;
            margin-right: 10px;
            border: none;
            border-radius: 3px;
            cursor: pointer;
            font-weight: bold;
            transition: all 0.3s;
        }
        .filter-btn:hover {
            opacity: 0.8;
        }
        .filter-all { background-color: #0078d4; color: white; }
        .filter-success { background-color: #28a745; color: white; }
        .filter-failed { background-color: #dc3545; color: white; }
        .filter-skipped { background-color: #ffc107; color: black; }
        .footer {
            margin-top: 20px;
            text-align: center;
            color: #666;
            font-size: 12px;
        }
    </style>
    <script>
        function filterTable(status) {
            var rows = document.querySelectorAll('tbody tr');
            rows.forEach(function(row) {
                if (status === 'all') {
                    row.style.display = '';
                } else {
                    var statusCell = row.cells[4].textContent.trim();
                    if (statusCell.toLowerCase() === status.toLowerCase()) {
                        row.style.display = '';
                    } else {
                        row.style.display = 'none';
                    }
                }
            });
        }
    </script>
</head>
<body>
    <div class="header">
        <h1>Exchange Mailbox Permissions Import Report</h1>
        <p>Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
        <p>Source CSV: $CsvPath</p>
        $(if ($WhatIf) { "<p style='background-color: #ffc107; color: black; padding: 10px; border-radius: 3px; display: inline-block;'><strong>WHATIF MODE - No changes were made</strong></p>" })
    </div>
    
    <div class="summary">
        <div class="summary-item">
            <div class="summary-label">Total Processed</div>
            <div class="summary-value total-value">$totalPermissions</div>
        </div>
        <div class="summary-item">
            <div class="summary-label">Successful</div>
            <div class="summary-value success-value">$successCount</div>
        </div>
        <div class="summary-item">
            <div class="summary-label">Failed</div>
            <div class="summary-value failure-value">$failureCount</div>
        </div>
        <div class="summary-item">
            <div class="summary-label">Skipped</div>
            <div class="summary-value skipped-value">$skippedCount</div>
        </div>
    </div>
    
    <div class="filter-buttons">
        <strong>Filter by Status:</strong>
        <button class="filter-btn filter-all" onclick="filterTable('all')">Show All</button>
        <button class="filter-btn filter-success" onclick="filterTable('success')">Success Only</button>
        <button class="filter-btn filter-failed" onclick="filterTable('failed')">Failed Only</button>
        <button class="filter-btn filter-skipped" onclick="filterTable('skipped')">Skipped Only</button>
    </div>
    
    <table>
        <thead>
            <tr>
                <th>Mailbox Display Name</th>
                <th>Mailbox Email</th>
                <th>Permission Type</th>
                <th>Trusted User</th>
                <th>Status</th>
                <th>Details</th>
                <th>Timestamp</th>
            </tr>
        </thead>
        <tbody>
"@

    foreach ($result in $importResults) {
        $statusClass = switch ($result.Status) {
            "Success" { "status-success" }
            "Failed" { "status-failed" }
            "Skipped" { "status-skipped" }
            "WhatIf" { "status-whatif" }
            default { "status-skipped" }
        }
        
        $permClass = switch ($result.PermissionType) {
            "FullAccess" { "full-access" }
            "SendAs" { "send-as" }
            "SendOnBehalf" { "send-on-behalf" }
        }
        
        $details = if ($result.ErrorMessage) {
            "<span class='error-message'>$($result.ErrorMessage)</span>"
        } else {
            $result.Action
        }
        
        $htmlContent += @"
            <tr>
                <td>$($result.MailboxDisplayName)</td>
                <td>$($result.MailboxEmail)</td>
                <td><span class="permission-type $permClass">$($result.PermissionType)</span></td>
                <td>$($result.TrustedUser)</td>
                <td><span class="status-badge $statusClass">$($result.Status)</span></td>
                <td>$details</td>
                <td>$($result.Timestamp)</td>
            </tr>
"@
    }

    $htmlContent += @"
        </tbody>
    </table>
    
    <div class="footer">
        <p>Exchange Mailbox Permissions Import | Report generated from: $CsvPath</p>
    </div>
</body>
</html>
"@

    $htmlContent | Out-File -FilePath $htmlFile -Encoding UTF8
    Write-Host "HTML report generated: $htmlFile" -ForegroundColor Green

    # Summary
    Write-Host "`n========== IMPORT SUMMARY ==========" -ForegroundColor Cyan
    Write-Host "Total permissions processed: $totalPermissions" -ForegroundColor White
    Write-Host "  ✓ Successful: $successCount" -ForegroundColor Green
    Write-Host "  ✗ Failed: $failureCount" -ForegroundColor Red
    Write-Host "  ⊘ Skipped: $skippedCount" -ForegroundColor Yellow
    Write-Host "====================================" -ForegroundColor Cyan
    
    if ($WhatIf) {
        Write-Host "`nWHATIF MODE was enabled - no actual changes were made" -ForegroundColor Yellow
    }

}
catch {
    Write-Error "An error occurred during import: $_"
    Write-Error $_.Exception.Message
}

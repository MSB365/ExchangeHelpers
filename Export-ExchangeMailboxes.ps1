<#
.SYNOPSIS
    Exports all Exchange mailboxes to PST files with HTML report
.DESCRIPTION
    This script exports User, Shared, Resource, and Room mailboxes to individual PST files
    named by their primary SMTP address and generates an HTML report of the process.
.PARAMETER ExportPath
    The UNC path where PST files will be exported (must be accessible by Exchange server)
.EXAMPLE
    .\Export-ExchangeMailboxes.ps1 -ExportPath "\\SERVER\ExchangeExports"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$ExportPath
)

# Ensure Exchange Management Shell is loaded
if (!(Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
    Write-Host "Loading Exchange Management Shell..." -ForegroundColor Yellow
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
}

# Verify export path exists
if (!(Test-Path $ExportPath)) {
    Write-Host "Creating export directory: $ExportPath" -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
}

# Initialize results array
$exportResults = @()
$startTime = Get-Date

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Exchange Mailbox Export Script" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Get all mailboxes (User, Shared, Resource, Room)
Write-Host "Retrieving all mailboxes..." -ForegroundColor Yellow
$allMailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {
    $_.RecipientTypeDetails -in @('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox')
}

Write-Host "Found $($allMailboxes.Count) mailboxes to export`n" -ForegroundColor Green

$counter = 0
foreach ($mailbox in $allMailboxes) {
    $counter++
    $primarySMTP = $mailbox.PrimarySmtpAddress.ToString()
    $displayName = $mailbox.DisplayName
    $mailboxType = $mailbox.RecipientTypeDetails
    
    # Sanitize filename (remove invalid characters)
    $fileName = $primarySMTP -replace '[\\/:*?"<>|]', '_'
    $pstPath = Join-Path $ExportPath "$fileName.pst"
    
    Write-Host "[$counter/$($allMailboxes.Count)] Processing: $displayName ($primarySMTP)" -ForegroundColor Cyan
    Write-Host "  Type: $mailboxType" -ForegroundColor Gray
    
    $exportStatus = "Success"
    $errorMessage = ""
    $exportStartTime = Get-Date
    
    try {
        # Create mailbox export request
        $exportRequest = New-MailboxExportRequest -Mailbox $mailbox.Identity -FilePath $pstPath -ErrorAction Stop
        
        Write-Host "  Export request created: $($exportRequest.Name)" -ForegroundColor Green
        Write-Host "  Waiting for export to complete..." -ForegroundColor Yellow
        
        # Wait for export to complete (check every 10 seconds)
        do {
            Start-Sleep -Seconds 10
            $status = Get-MailboxExportRequest -Identity $exportRequest.Identity
            Write-Host "  Status: $($status.Status) - $($status.PercentComplete)% complete" -ForegroundColor Gray
        } while ($status.Status -notin @('Completed', 'Failed', 'CompletedWithWarning'))
        
        if ($status.Status -eq 'Failed') {
            $exportStatus = "Failed"
            $errorMessage = $status.Message
            Write-Host "  Export FAILED: $errorMessage" -ForegroundColor Red
        } elseif ($status.Status -eq 'CompletedWithWarning') {
            $exportStatus = "Completed with Warnings"
            $errorMessage = $status.Message
            Write-Host "  Export completed with warnings" -ForegroundColor Yellow
        } else {
            Write-Host "  Export completed successfully!" -ForegroundColor Green
        }
        
        # Get mailbox statistics
        $mailboxStats = Get-MailboxStatistics -Identity $mailbox.Identity
        $itemCount = $mailboxStats.ItemCount
        $totalSize = $mailboxStats.TotalItemSize.Value.ToMB()
        
        # Remove the export request to clean up
        Remove-MailboxExportRequest -Identity $exportRequest.Identity -Confirm:$false
        
    } catch {
        $exportStatus = "Failed"
        $errorMessage = $_.Exception.Message
        Write-Host "  ERROR: $errorMessage" -ForegroundColor Red
        $itemCount = "N/A"
        $totalSize = "N/A"
    }
    
    $exportEndTime = Get-Date
    $duration = ($exportEndTime - $exportStartTime).ToString("hh\:mm\:ss")
    
    # Add to results
    $exportResults += [PSCustomObject]@{
        DisplayName = $displayName
        PrimarySMTP = $primarySMTP
        MailboxType = $mailboxType
        Status = $exportStatus
        ItemCount = $itemCount
        SizeMB = $totalSize
        PSTFile = "$fileName.pst"
        Duration = $duration
        ErrorMessage = $errorMessage
        ExportTime = $exportEndTime.ToString("yyyy-MM-dd HH:mm:ss")
    }
    
    Write-Host ""
}

$endTime = Get-Date
$totalDuration = ($endTime - $startTime).ToString("hh\:mm\:ss")

# Generate HTML Report
Write-Host "Generating HTML report..." -ForegroundColor Yellow

$successCount = ($exportResults | Where-Object { $_.Status -eq "Success" }).Count
$failedCount = ($exportResults | Where-Object { $_.Status -eq "Failed" }).Count
$warningCount = ($exportResults | Where-Object { $_.Status -eq "Completed with Warnings" }).Count

$htmlReport = @"
<!DOCTYPE html>
<html>
<head>
    <title>Exchange Mailbox Export Report</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        h1 {
            color: #0078d4;
            border-bottom: 3px solid #0078d4;
            padding-bottom: 10px;
        }
        h2 {
            color: #333;
            margin-top: 30px;
        }
        .summary {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin: 20px 0;
        }
        .summary-box {
            padding: 20px;
            border-radius: 6px;
            text-align: center;
        }
        .summary-box h3 {
            margin: 0;
            font-size: 2em;
            font-weight: bold;
        }
        .summary-box p {
            margin: 5px 0 0 0;
            color: #666;
        }
        .success { background-color: #dff6dd; color: #0f5132; }
        .failed { background-color: #f8d7da; color: #842029; }
        .warning { background-color: #fff3cd; color: #664d03; }
        .total { background-color: #cfe2ff; color: #084298; }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        th {
            background-color: #0078d4;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: 600;
        }
        td {
            padding: 10px;
            border-bottom: 1px solid #ddd;
        }
        tr:hover {
            background-color: #f5f5f5;
        }
        .status-success { color: #0f5132; font-weight: bold; }
        .status-failed { color: #842029; font-weight: bold; }
        .status-warning { color: #664d03; font-weight: bold; }
        .info {
            background-color: #e7f3ff;
            padding: 15px;
            border-left: 4px solid #0078d4;
            margin: 20px 0;
        }
        .footer {
            margin-top: 30px;
            padding-top: 20px;
            border-top: 1px solid #ddd;
            color: #666;
            font-size: 0.9em;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Exchange Mailbox Export Report</h1>
        
        <div class="info">
            <strong>Export Path:</strong> $ExportPath<br>
            <strong>Start Time:</strong> $($startTime.ToString("yyyy-MM-dd HH:mm:ss"))<br>
            <strong>End Time:</strong> $($endTime.ToString("yyyy-MM-dd HH:mm:ss"))<br>
            <strong>Total Duration:</strong> $totalDuration
        </div>
        
        <h2>Summary</h2>
        <div class="summary">
            <div class="summary-box total">
                <h3>$($exportResults.Count)</h3>
                <p>Total Mailboxes</p>
            </div>
            <div class="summary-box success">
                <h3>$successCount</h3>
                <p>Successful</p>
            </div>
            <div class="summary-box warning">
                <h3>$warningCount</h3>
                <p>Warnings</p>
            </div>
            <div class="summary-box failed">
                <h3>$failedCount</h3>
                <p>Failed</p>
            </div>
        </div>
        
        <h2>Export Details</h2>
        <table>
            <thead>
                <tr>
                    <th>Display Name</th>
                    <th>Primary SMTP</th>
                    <th>Type</th>
                    <th>Status</th>
                    <th>Items</th>
                    <th>Size (MB)</th>
                    <th>PST File</th>
                    <th>Duration</th>
                    <th>Export Time</th>
                    <th>Error Message</th>
                </tr>
            </thead>
            <tbody>
"@

foreach ($result in $exportResults) {
    $statusClass = switch ($result.Status) {
        "Success" { "status-success" }
        "Failed" { "status-failed" }
        default { "status-warning" }
    }
    
    $htmlReport += @"
                <tr>
                    <td>$($result.DisplayName)</td>
                    <td>$($result.PrimarySMTP)</td>
                    <td>$($result.MailboxType)</td>
                    <td class="$statusClass">$($result.Status)</td>
                    <td>$($result.ItemCount)</td>
                    <td>$($result.SizeMB)</td>
                    <td>$($result.PSTFile)</td>
                    <td>$($result.Duration)</td>
                    <td>$($result.ExportTime)</td>
                    <td>$($result.ErrorMessage)</td>
                </tr>
"@
}

$htmlReport += @"
            </tbody>
        </table>
        
        <div class="footer">
            Report generated on $($endTime.ToString("yyyy-MM-dd HH:mm:ss")) by Exchange Mailbox Export Script
        </div>
    </div>
</body>
</html>
"@

# Save HTML report
$reportPath = Join-Path $ExportPath "MailboxExportReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$htmlReport | Out-File -FilePath $reportPath -Encoding UTF8

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Export Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Total Mailboxes: $($exportResults.Count)" -ForegroundColor White
Write-Host "Successful: $successCount" -ForegroundColor Green
Write-Host "Warnings: $warningCount" -ForegroundColor Yellow
Write-Host "Failed: $failedCount" -ForegroundColor Red
Write-Host "Total Duration: $totalDuration" -ForegroundColor White
Write-Host "`nHTML Report saved to: $reportPath" -ForegroundColor Cyan
Write-Host "PST files saved to: $ExportPath`n" -ForegroundColor Cyan
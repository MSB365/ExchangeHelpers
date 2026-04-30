<#
.SYNOPSIS
    Creates an HTML report about Exchange Online mailboxes.

.DESCRIPTION
    This script connects to Exchange Online and creates a detailed HTML report
    with information about user mailboxes, shared mailboxes and the actually
    used storage.

.NOTES
    Prerequisite: ExchangeOnlineManagement module must be installed.
    Installation: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
	
	==============================================================================
	Created on:         2026/04/30
	Created by:         Drago Petrovic
	Organization:       MSB365.blog
	Filename:           Get-ExchangeOnlineReport.ps1
	Current version:    V1.0     

	Find us on:
	* Website:         https://www.msb365.net
	* Technet:         https://social.technet.microsoft.com/Profile/MSB365
	* LinkedIn:        https://www.linkedin.com/in/drago-petrovic/
	* MVP Profile:     https://mvp.microsoft.com/de-de/PublicProfile/5003446
==============================================================================

.EXAMPLE
    .\Get-ExchangeOnlineReport.ps1
#>

#region Configuration
$ReportPath = "$PSScriptRoot\ExchangeOnline_Report_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').html"
$CompanyName = "Your Company" # Adjust if desired
#endregion

#region Functions
function Convert-SizeToGB {
    param([string]$SizeString)
    
    if ([string]::IsNullOrWhiteSpace($SizeString)) { return 0 }
    
    # Format: "1.234 GB (1,324,234,234 bytes)" or "23.45 MB (24,567,890 bytes)"
    if ($SizeString -match '\(([0-9,]+) bytes\)') {
        $bytes = [double]($matches[1] -replace ',', '')
        return [math]::Round($bytes / 1GB, 2)
    }
    return 0
}

function Write-Status {
    param(
        [string]$Message,
        [string]$Type = "Info"
    )
    $colors = @{
        "Info"    = "Cyan"
        "Success" = "Green"
        "Warning" = "Yellow"
        "Error"   = "Red"
    }
    $prefix = @{
        "Info"    = "[INFO]"
        "Success" = "[OK]  "
        "Warning" = "[WARN]"
        "Error"   = "[ERR] "
    }
    Write-Host "$($prefix[$Type]) $Message" -ForegroundColor $colors[$Type]
}
#endregion

#region Check module and establish connection
Write-Status "Checking ExchangeOnlineManagement module..." -Type Info

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Status "Module not found. Installing ExchangeOnlineManagement..." -Type Warning
    try {
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
        Write-Status "Module successfully installed." -Type Success
    } catch {
        Write-Status "Error installing module: $_" -Type Error
        exit 1
    }
}

Import-Module ExchangeOnlineManagement -ErrorAction Stop

Write-Status "Connecting to Exchange Online..." -Type Info
try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Status "Connection to Exchange Online established." -Type Success
} catch {
    Write-Status "Connection failed: $_" -Type Error
    exit 1
}
#endregion

#region Collect data
Write-Status "Collecting mailbox information..." -Type Info

try {
    # Retrieve all mailboxes
    $AllMailboxes = Get-EXOMailbox -ResultSize Unlimited -PropertySets All -ErrorAction Stop
    
    $UserMailboxes   = $AllMailboxes | Where-Object { $_.RecipientTypeDetails -eq 'UserMailbox' }
    $SharedMailboxes = $AllMailboxes | Where-Object { $_.RecipientTypeDetails -eq 'SharedMailbox' }
    $RoomMailboxes   = $AllMailboxes | Where-Object { $_.RecipientTypeDetails -eq 'RoomMailbox' }
    $EquipMailboxes  = $AllMailboxes | Where-Object { $_.RecipientTypeDetails -eq 'EquipmentMailbox' }
    
    Write-Status "Found: $($UserMailboxes.Count) user mailboxes, $($SharedMailboxes.Count) shared mailboxes" -Type Success
    Write-Status "Retrieving statistics for $($AllMailboxes.Count) mailboxes (this may take a while)..." -Type Info

    $MailboxDetails = @()
    $counter = 0
    $total = $AllMailboxes.Count
    
    foreach ($mbx in $AllMailboxes) {
        $counter++
        Write-Progress -Activity "Loading mailbox statistics" `
                       -Status "$counter of $total - $($mbx.DisplayName)" `
                       -PercentComplete (($counter / $total) * 100)
        
        try {
            $stats = Get-EXOMailboxStatistics -Identity $mbx.UserPrincipalName -ErrorAction Stop
            $usedGB = Convert-SizeToGB -SizeString $stats.TotalItemSize.ToString()
            
            # Read quota
            $quotaGB = 0
            if ($mbx.ProhibitSendQuota -and $mbx.ProhibitSendQuota -ne 'Unlimited') {
                $quotaGB = Convert-SizeToGB -SizeString $mbx.ProhibitSendQuota.ToString()
            }
            
            $MailboxDetails += [PSCustomObject]@{
                DisplayName       = $mbx.DisplayName
                UserPrincipalName = $mbx.UserPrincipalName
                Type              = $mbx.RecipientTypeDetails
                UsedGB            = $usedGB
                QuotaGB           = $quotaGB
                ItemCount         = $stats.ItemCount
                LastLogonTime     = $stats.LastLogonTime
            }
        } catch {
            Write-Status "Error processing mailbox $($mbx.DisplayName): $_" -Type Warning
        }
    }
    Write-Progress -Activity "Loading mailbox statistics" -Completed
    
} catch {
    Write-Status "Error retrieving mailbox data: $_" -Type Error
    Disconnect-ExchangeOnline -Confirm:$false
    exit 1
}
#endregion

#region Calculations
$TotalUsedGB    = [math]::Round(($MailboxDetails | Measure-Object -Property UsedGB -Sum).Sum, 2)
$TotalQuotaGB   = [math]::Round(($MailboxDetails | Measure-Object -Property QuotaGB -Sum).Sum, 2)
$UserUsedGB     = [math]::Round((($MailboxDetails | Where-Object Type -eq 'UserMailbox') | Measure-Object -Property UsedGB -Sum).Sum, 2)
$SharedUsedGB   = [math]::Round((($MailboxDetails | Where-Object Type -eq 'SharedMailbox') | Measure-Object -Property UsedGB -Sum).Sum, 2)
$AvgMailboxSize = if ($MailboxDetails.Count -gt 0) { [math]::Round($TotalUsedGB / $MailboxDetails.Count, 2) } else { 0 }
$UsagePercent   = if ($TotalQuotaGB -gt 0) { [math]::Round(($TotalUsedGB / $TotalQuotaGB) * 100, 1) } else { 0 }

# Top 10 largest mailboxes
$TopMailboxes = $MailboxDetails | Sort-Object UsedGB -Descending | Select-Object -First 10

# Tenant info
try {
    $TenantInfo = Get-OrganizationConfig -ErrorAction Stop
    $TenantName = $TenantInfo.DisplayName
    $TenantDomain = ($TenantInfo.Identity)
} catch {
    $TenantName = "N/A"
    $TenantDomain = "N/A"
}
#endregion

#region Create HTML report
Write-Status "Creating HTML report..." -Type Info

$ReportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Generate top mailbox table rows
$TopMailboxRows = ""
foreach ($mbx in $TopMailboxes) {
    $percent = if ($mbx.QuotaGB -gt 0) { [math]::Round(($mbx.UsedGB / $mbx.QuotaGB) * 100, 1) } else { 0 }
    $barColor = if ($percent -gt 90) { "#e74c3c" } elseif ($percent -gt 75) { "#f39c12" } else { "#27ae60" }
    
    $TopMailboxRows += @"
        <tr>
            <td><strong>$($mbx.DisplayName)</strong><br><span class="upn">$($mbx.UserPrincipalName)</span></td>
            <td><span class="badge badge-$($mbx.Type)">$($mbx.Type)</span></td>
            <td class="number">$($mbx.UsedGB) GB</td>
            <td class="number">$($mbx.QuotaGB) GB</td>
            <td>
                <div class="progress-bar">
                    <div class="progress-fill" style="width: $percent%; background: $barColor;"></div>
                    <span class="progress-text">$percent%</span>
                </div>
            </td>
            <td class="number">$($mbx.ItemCount)</td>
        </tr>
"@
}

# Distribution table for all mailbox types
$DistributionRows = @"
        <tr>
            <td><span class="badge badge-UserMailbox">User Mailbox</span></td>
            <td class="number">$($UserMailboxes.Count)</td>
            <td class="number">$UserUsedGB GB</td>
        </tr>
        <tr>
            <td><span class="badge badge-SharedMailbox">Shared Mailbox</span></td>
            <td class="number">$($SharedMailboxes.Count)</td>
            <td class="number">$SharedUsedGB GB</td>
        </tr>
"@

if ($RoomMailboxes.Count -gt 0) {
    $RoomUsedGB = [math]::Round((($MailboxDetails | Where-Object Type -eq 'RoomMailbox') | Measure-Object -Property UsedGB -Sum).Sum, 2)
    $DistributionRows += @"
        <tr>
            <td><span class="badge badge-RoomMailbox">Room Mailbox</span></td>
            <td class="number">$($RoomMailboxes.Count)</td>
            <td class="number">$RoomUsedGB GB</td>
        </tr>
"@
}

if ($EquipMailboxes.Count -gt 0) {
    $EquipUsedGB = [math]::Round((($MailboxDetails | Where-Object Type -eq 'EquipmentMailbox') | Measure-Object -Property UsedGB -Sum).Sum, 2)
    $DistributionRows += @"
        <tr>
            <td><span class="badge badge-EquipmentMailbox">Equipment Mailbox</span></td>
            <td class="number">$($EquipMailboxes.Count)</td>
            <td class="number">$EquipUsedGB GB</td>
        </tr>
"@
}

$Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Exchange Online Report - $ReportDate</title>
<style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        min-height: 100vh;
        padding: 20px;
        color: #333;
    }
    .container {
        max-width: 1400px;
        margin: 0 auto;
        background: #fff;
        border-radius: 12px;
        box-shadow: 0 20px 60px rgba(0,0,0,0.3);
        overflow: hidden;
    }
    .header {
        background: linear-gradient(135deg, #0078d4 0%, #005a9e 100%);
        color: #fff;
        padding: 40px;
        position: relative;
    }
    .header::after {
        content: '';
        position: absolute;
        bottom: 0; left: 0; right: 0;
        height: 4px;
        background: linear-gradient(90deg, #00bcf2, #0078d4, #005a9e);
    }
    .header h1 {
        font-size: 32px;
        font-weight: 600;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        gap: 15px;
    }
    .header .icon {
        font-size: 40px;
    }
    .header .subtitle {
        font-size: 16px;
        opacity: 0.9;
    }
    .meta-info {
        background: #f8f9fa;
        padding: 20px 40px;
        border-bottom: 1px solid #e0e0e0;
        display: flex;
        flex-wrap: wrap;
        gap: 30px;
    }
    .meta-info div {
        font-size: 14px;
        color: #555;
    }
    .meta-info strong {
        color: #0078d4;
        margin-right: 5px;
    }
    .content { padding: 40px; }
    .section { margin-bottom: 40px; }
    .section h2 {
        color: #0078d4;
        font-size: 22px;
        margin-bottom: 20px;
        padding-bottom: 10px;
        border-bottom: 2px solid #e0e0e0;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .stats-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
        gap: 20px;
        margin-bottom: 30px;
    }
    .stat-card {
        background: #fff;
        border: 1px solid #e0e0e0;
        border-radius: 10px;
        padding: 25px;
        position: relative;
        overflow: hidden;
        transition: transform 0.2s, box-shadow 0.2s;
    }
    .stat-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 10px 25px rgba(0,0,0,0.1);
    }
    .stat-card::before {
        content: '';
        position: absolute;
        top: 0; left: 0;
        width: 4px; height: 100%;
        background: #0078d4;
    }
    .stat-card.users::before    { background: #0078d4; }
    .stat-card.shared::before   { background: #00bcf2; }
    .stat-card.storage::before  { background: #107c10; }
    .stat-card.quota::before    { background: #ff8c00; }
    .stat-card.average::before  { background: #5c2d91; }
    .stat-card.usage::before    { background: #d83b01; }
    .stat-label {
        font-size: 13px;
        color: #666;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 10px;
        font-weight: 600;
    }
    .stat-value {
        font-size: 36px;
        font-weight: 700;
        color: #1a1a1a;
        line-height: 1;
        margin-bottom: 5px;
    }
    .stat-unit {
        font-size: 14px;
        color: #888;
        font-weight: normal;
    }
    .stat-sub {
        font-size: 13px;
        color: #666;
        margin-top: 8px;
    }
    table {
        width: 100%;
        border-collapse: collapse;
        background: #fff;
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    thead { background: #f3f2f1; }
    th {
        padding: 14px 16px;
        text-align: left;
        font-size: 13px;
        font-weight: 600;
        color: #323130;
        text-transform: uppercase;
        letter-spacing: 0.3px;
        border-bottom: 2px solid #e0e0e0;
    }
    td {
        padding: 14px 16px;
        border-bottom: 1px solid #f0f0f0;
        font-size: 14px;
    }
    tbody tr:hover { background: #f8f9fa; }
    tbody tr:last-child td { border-bottom: none; }
    .number { text-align: right; font-variant-numeric: tabular-nums; font-weight: 500; }
    .upn { font-size: 12px; color: #888; }
    .badge {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.3px;
    }
    .badge-UserMailbox      { background: #deecf9; color: #0078d4; }
    .badge-SharedMailbox    { background: #cff4fc; color: #055160; }
    .badge-RoomMailbox      { background: #d4edda; color: #155724; }
    .badge-EquipmentMailbox { background: #fff3cd; color: #856404; }
    .progress-bar {
        position: relative;
        background: #f0f0f0;
        border-radius: 10px;
        height: 22px;
        overflow: hidden;
        min-width: 120px;
    }
    .progress-fill {
        height: 100%;
        border-radius: 10px;
        transition: width 0.3s;
    }
    .progress-text {
        position: absolute;
        top: 50%; left: 50%;
        transform: translate(-50%, -50%);
        font-size: 12px;
        font-weight: 600;
        color: #1a1a1a;
        text-shadow: 0 0 3px rgba(255,255,255,0.8);
    }
    .global-progress {
        background: #f8f9fa;
        border-radius: 10px;
        padding: 20px;
        margin-top: 20px;
    }
    .global-progress-bar {
        background: #e0e0e0;
        height: 30px;
        border-radius: 15px;
        overflow: hidden;
        margin-top: 10px;
        position: relative;
    }
    .global-progress-fill {
        height: 100%;
        background: linear-gradient(90deg, #107c10, #00bcf2);
        border-radius: 15px;
        display: flex;
        align-items: center;
        justify-content: center;
        color: #fff;
        font-weight: 600;
        font-size: 14px;
        transition: width 0.5s;
    }
    .footer {
        background: #f8f9fa;
        padding: 20px 40px;
        text-align: center;
        font-size: 12px;
        color: #888;
        border-top: 1px solid #e0e0e0;
    }
    .footer strong { color: #0078d4; }
</style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1><span class="icon">&#9993;</span> Exchange Online Report</h1>
        <div class="subtitle">Overview of mailboxes and storage usage</div>
    </div>

    <div class="meta-info">
        <div><strong>Generated:</strong> $ReportDate</div>
        <div><strong>Tenant:</strong> $TenantName</div>
        <div><strong>Organization:</strong> $TenantDomain</div>
        <div><strong>Total Mailboxes:</strong> $($AllMailboxes.Count)</div>
    </div>

    <div class="content">

        <div class="section">
            <h2>&#128202; Summary</h2>
            <div class="stats-grid">
                <div class="stat-card users">
                    <div class="stat-label">User Mailboxes</div>
                    <div class="stat-value">$($UserMailboxes.Count)</div>
                    <div class="stat-sub">active user mailboxes</div>
                </div>
                <div class="stat-card shared">
                    <div class="stat-label">Shared Mailboxes</div>
                    <div class="stat-value">$($SharedMailboxes.Count)</div>
                    <div class="stat-sub">shared mailboxes</div>
                </div>
                <div class="stat-card storage">
                    <div class="stat-label">Used Storage</div>
                    <div class="stat-value">$TotalUsedGB <span class="stat-unit">GB</span></div>
                    <div class="stat-sub">actually consumed storage</div>
                </div>
                <div class="stat-card quota">
                    <div class="stat-label">Allocated Quota</div>
                    <div class="stat-value">$TotalQuotaGB <span class="stat-unit">GB</span></div>
                    <div class="stat-sub">total available storage</div>
                </div>
                <div class="stat-card average">
                    <div class="stat-label">Average per Mailbox</div>
                    <div class="stat-value">$AvgMailboxSize <span class="stat-unit">GB</span></div>
                    <div class="stat-sub">across all mailboxes</div>
                </div>
                <div class="stat-card usage">
                    <div class="stat-label">Overall Usage</div>
                    <div class="stat-value">$UsagePercent <span class="stat-unit">%</span></div>
                    <div class="stat-sub">of allocated quota</div>
                </div>
            </div>

            <div class="global-progress">
                <strong>Total Storage Usage: $TotalUsedGB GB of $TotalQuotaGB GB</strong>
                <div class="global-progress-bar">
                    <div class="global-progress-fill" style="width: $UsagePercent%;">$UsagePercent%</div>
                </div>
            </div>
        </div>

        <div class="section">
            <h2>&#128193; Distribution by Mailbox Type</h2>
            <table>
                <thead>
                    <tr>
                        <th>Mailbox Type</th>
                        <th class="number">Count</th>
                        <th class="number">Used Storage</th>
                    </tr>
                </thead>
                <tbody>
$DistributionRows
                </tbody>
            </table>
        </div>

        <div class="section">
            <h2>&#127942; Top 10 Largest Mailboxes</h2>
            <table>
                <thead>
                    <tr>
                        <th>Mailbox</th>
                        <th>Type</th>
                        <th class="number">Used</th>
                        <th class="number">Quota</th>
                        <th>Usage</th>
                        <th class="number">Items</th>
                    </tr>
                </thead>
                <tbody>
$TopMailboxRows
                </tbody>
            </table>
        </div>

    </div>

    <div class="footer">
        Report automatically generated with <strong>PowerShell &amp; Exchange Online Management</strong> &middot; $ReportDate
    </div>
</div>
</body>
</html>
"@

try {
    $Html | Out-File -FilePath $ReportPath -Encoding UTF8 -ErrorAction Stop
    Write-Status "Report successfully saved: $ReportPath" -Type Success
    
    # Open report in browser
    Start-Process $ReportPath
} catch {
    Write-Status "Error saving the report: $_" -Type Error
}
#endregion

#region Cleanup
Write-Status "Disconnecting from Exchange Online..." -Type Info
Disconnect-ExchangeOnline -Confirm:$false
Write-Status "Done!" -Type Success
#endregion

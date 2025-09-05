<#
.SYNOPSIS
    Adds additional alias addresses to Exchange Online mailboxes from a CSV file.

.DESCRIPTION
    This script reads a CSV file containing mailbox identities and their corresponding alias addresses,
    then adds these aliases to the specified mailboxes in Exchange Online.
    
    CSV Format Required:
    - Email: The primary email address or identity of the mailbox
    - SMTP: The alias email address to be added

.NOTES
    Author: PowerShell Script for Exchange Online
    Requires: ExchangeOnlineManagement module
    Version: 1.0
#>

# Import required modules
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Write-Host "✓ ExchangeOnlineManagement module loaded successfully" -ForegroundColor Green
}
catch {
    Write-Host "✗ Failed to load ExchangeOnlineManagement module. Please install it using:" -ForegroundColor Red
    Write-Host "Install-Module -Name ExchangeOnlineManagement" -ForegroundColor Yellow
    exit 1
}

# Add Windows Forms for file dialog
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to show file picker dialog
function Get-CsvFile {
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Title = "Select CSV file with mailbox aliases"
    $fileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $fileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    
    $result = $fileDialog.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $fileDialog.FileName
    }
    else {
        Write-Host "No file selected. Exiting script." -ForegroundColor Yellow
        exit 0
    }
}

# Function to validate CSV structure
function Test-CsvStructure {
    param([string]$FilePath)
    
    try {
        $csvData = Import-Csv -Path $FilePath -ErrorAction Stop
        
        if (-not $csvData) {
            throw "CSV file is empty"
        }
        
        $requiredColumns = @("Email", "SMTP")
        $csvColumns = $csvData[0].PSObject.Properties.Name
        
        foreach ($column in $requiredColumns) {
            if ($column -notin $csvColumns) {
                throw "Required column '$column' not found in CSV. Available columns: $($csvColumns -join ', ')"
            }
        }
        
        return $csvData
    }
    catch {
        Write-Host "✗ CSV validation failed: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

# Function to connect to Exchange Online
function Connect-ExchangeOnlineWithRetry {
    $maxAttempts = 3
    $attempt = 1
    
    while ($attempt -le $maxAttempts) {
        try {
            Write-Host "Attempting to connect to Exchange Online (Attempt $attempt/$maxAttempts)..." -ForegroundColor Cyan
            Connect-ExchangeOnline -ShowProgress $true -ErrorAction Stop
            Write-Host "✓ Successfully connected to Exchange Online" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Host "✗ Connection attempt $attempt failed: $($_.Exception.Message)" -ForegroundColor Red
            $attempt++
            if ($attempt -le $maxAttempts) {
                Start-Sleep -Seconds 5
            }
        }
    }
    
    Write-Host "✗ Failed to connect to Exchange Online after $maxAttempts attempts" -ForegroundColor Red
    return $false
}

# Function to add alias to mailbox
function Add-MailboxAlias {
    param(
        [string]$MailboxIdentity,
        [string]$AliasAddress
    )
    
    try {
        # Validate email format
        if ($AliasAddress -notmatch '^[^\s@]+@[^\s@]+\.[^\s@]+$') {
            throw "Invalid email format: $AliasAddress"
        }
        
        # Get current mailbox
        $mailbox = Get-Mailbox -Identity $MailboxIdentity -ErrorAction Stop
        
        # Check if alias already exists
        if ($mailbox.EmailAddresses -contains "smtp:$AliasAddress") {
            return @{
                Success = $false
                Message = "Alias already exists"
                Mailbox = $mailbox.DisplayName
                PrimaryEmail = $mailbox.PrimarySmtpAddress
                Alias = $AliasAddress
            }
        }
        
        # Add the alias
        Set-Mailbox -Identity $MailboxIdentity -EmailAddresses @{Add="smtp:$AliasAddress"} -ErrorAction Stop
        
        return @{
            Success = $true
            Message = "Alias added successfully"
            Mailbox = $mailbox.DisplayName
            PrimaryEmail = $mailbox.PrimarySmtpAddress
            Alias = $AliasAddress
        }
    }
    catch {
        return @{
            Success = $false
            Message = $_.Exception.Message
            Mailbox = $MailboxIdentity
            PrimaryEmail = "Unknown"
            Alias = $AliasAddress
        }
    }
}

# Main script execution
Write-Host "=== Exchange Online Alias Management Script ===" -ForegroundColor Cyan
Write-Host "This script will add alias addresses to mailboxes from a CSV file." -ForegroundColor White
Write-Host ""

# Get CSV file from user
Write-Host "Please select your CSV file..." -ForegroundColor Yellow
$csvFilePath = Get-CsvFile
Write-Host "Selected file: $csvFilePath" -ForegroundColor Green
Write-Host ""

# Validate CSV structure
Write-Host "Validating CSV file structure..." -ForegroundColor Yellow
$csvData = Test-CsvStructure -FilePath $csvFilePath
Write-Host "✓ CSV file validation passed. Found $($csvData.Count) records to process." -ForegroundColor Green
Write-Host ""

# Connect to Exchange Online
if (-not (Connect-ExchangeOnlineWithRetry)) {
    exit 1
}
Write-Host ""

# Initialize counters and results
$totalRecords = $csvData.Count
$successCount = 0
$failureCount = 0
$skippedCount = 0
$results = @()

Write-Host "Processing $totalRecords records..." -ForegroundColor Cyan
Write-Host "=" * 80

# Process each record
$counter = 1
foreach ($record in $csvData) {
    $email = $record.Email.Trim()
    $smtp = $record.SMTP.Trim()
    
    Write-Host "[$counter/$totalRecords] Processing: $email -> $smtp" -ForegroundColor White
    
    if ([string]::IsNullOrWhiteSpace($email) -or [string]::IsNullOrWhiteSpace($smtp)) {
        Write-Host "  ⚠ Skipping - Empty email or SMTP value" -ForegroundColor Yellow
        $skippedCount++
        $results += @{
            Success = $false
            Message = "Empty email or SMTP value"
            Mailbox = $email
            PrimaryEmail = "N/A"
            Alias = $smtp
        }
    }
    else {
        $result = Add-MailboxAlias -MailboxIdentity $email -AliasAddress $smtp
        $results += $result
        
        if ($result.Success) {
            Write-Host "  ✓ Success: Added alias to $($result.Mailbox)" -ForegroundColor Green
            $successCount++
        }
        else {
            if ($result.Message -eq "Alias already exists") {
                Write-Host "  ⚠ Skipped: $($result.Message)" -ForegroundColor Yellow
                $skippedCount++
            }
            else {
                Write-Host "  ✗ Failed: $($result.Message)" -ForegroundColor Red
                $failureCount++
            }
        }
    }
    
    $counter++
    Start-Sleep -Milliseconds 500  # Small delay to avoid throttling
}

# Display summary
Write-Host ""
Write-Host "=" * 80
Write-Host "OPERATION SUMMARY" -ForegroundColor Cyan
Write-Host "=" * 80
Write-Host "Total Records Processed: $totalRecords" -ForegroundColor White
Write-Host "Successfully Added: $successCount" -ForegroundColor Green
Write-Host "Already Existed (Skipped): $skippedCount" -ForegroundColor Yellow
Write-Host "Failed: $failureCount" -ForegroundColor Red
Write-Host ""

# Display detailed results
if ($results.Count -gt 0) {
    Write-Host "DETAILED RESULTS:" -ForegroundColor Cyan
    Write-Host "-" * 80
    
    $results | ForEach-Object {
        $status = if ($_.Success) { "✓ SUCCESS" } elseif ($_.Message -eq "Alias already exists") { "⚠ SKIPPED" } else { "✗ FAILED" }
        $color = if ($_.Success) { "Green" } elseif ($_.Message -eq "Alias already exists") { "Yellow" } else { "Red" }
        
        Write-Host "$status | Mailbox: $($_.Mailbox) | Alias: $($_.Alias)" -ForegroundColor $color
        if (-not $_.Success -and $_.Message -ne "Alias already exists") {
            Write-Host "         Reason: $($_.Message)" -ForegroundColor Red
        }
    }
}

# Export results to CSV
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$resultsCsvPath = Join-Path -Path (Split-Path $csvFilePath -Parent) -ChildPath "AliasResults_$timestamp.csv"

try {
    $results | Select-Object @{Name="Status";Expression={if($_.Success){"Success"}elseif($_.Message -eq "Alias already exists"){"Skipped"}else{"Failed"}}}, 
                            Mailbox, PrimaryEmail, Alias, Message | 
    Export-Csv -Path $resultsCsvPath -NoTypeInformation -ErrorAction Stop
    
    Write-Host ""
    Write-Host "Results exported to: $resultsCsvPath" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "⚠ Warning: Could not export results to CSV: $($_.Exception.Message)" -ForegroundColor Yellow
}

# Disconnect from Exchange Online
Write-Host ""
Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Yellow
try {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
    Write-Host "✓ Disconnected successfully" -ForegroundColor Green
}
catch {
    Write-Host "⚠ Warning: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Script execution completed!" -ForegroundColor Cyan
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

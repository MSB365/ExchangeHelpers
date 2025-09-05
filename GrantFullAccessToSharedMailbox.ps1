# PowerShell script to grant Full Access permissions to mail-enabled security groups on shared mailboxes
# Reads from CSV file with MailboxAlias and PermissionHolder columns

# Load Windows Forms for file dialogs
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to show CSV file selection dialog
function Select-CsvFile {
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "Select CSV File with Mailbox Permissions"
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $openFileDialog.FileName
    } else {
        Write-Host "No CSV file selected. Exiting script." -ForegroundColor Red
        exit 1
    }
}

# Function to show log file save location dialog
function Select-LogFileLocation {
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Title = "Choose Location to Save Log File"
    $saveFileDialog.Filter = "Log Files (*.log)|*.log|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    $saveFileDialog.DefaultExt = "log"
    $saveFileDialog.FileName = "MailboxPermissions_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    $saveFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $saveFileDialog.FileName
    } else {
        Write-Host "No log file location selected. Exiting script." -ForegroundColor Red
        exit 1
    }
}

# Function to write to log file and console
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [string]$LogPath
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Write to console with color coding
    switch ($Level) {
        "ERROR" { Write-Host $logEntry -ForegroundColor Red }
        "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
        default { Write-Host $logEntry -ForegroundColor White }
    }
    
    # Write to log file
    Add-Content -Path $LogPath -Value $logEntry
}



# Main script execution
try {
    # Show file selection dialogs
    Write-Host "Starting Shared Mailbox Permission Grant Script" -ForegroundColor Cyan
    Write-Host "Please select the CSV file and log file location..." -ForegroundColor Yellow
    
    $CsvPath = Select-CsvFile
    $LogPath = Select-LogFileLocation
    
    Write-Log -Message "Script started" -LogPath $LogPath
    Write-Log -Message "CSV File: $CsvPath" -LogPath $LogPath
    Write-Log -Message "Log File: $LogPath" -LogPath $LogPath
    
    # Check if Exchange Online module is available
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Log -Message "ExchangeOnlineManagement module not found. Please install it first: Install-Module -Name ExchangeOnlineManagement" -Level "ERROR" -LogPath $LogPath
        exit 1
    }
    
    # Check if connected to Exchange Online
    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-Log -Message "Connected to Exchange Online" -Level "SUCCESS" -LogPath $LogPath
    }
    catch {
        Write-Log -Message "Not connected to Exchange Online. Please run Connect-ExchangeOnline first." -Level "ERROR" -LogPath $LogPath
        exit 1
    }
    
    # Validate CSV file exists
    if (-not (Test-Path $CsvPath)) {
        Write-Log -Message "CSV file not found: $CsvPath" -Level "ERROR" -LogPath $LogPath
        exit 1
    }
    
    # Import and validate CSV
    try {
        $csvData = Import-Csv -Path $CsvPath
        Write-Log -Message "Successfully imported CSV with $($csvData.Count) rows" -LogPath $LogPath
    }
    catch {
        Write-Log -Message "Failed to import CSV: $($_.Exception.Message)" -Level "ERROR" -LogPath $LogPath
        exit 1
    }
    
    # Validate required columns exist
    $requiredColumns = @('MailboxAlias', 'PermissionHolder')
    $csvColumns = $csvData[0].PSObject.Properties.Name
    
    foreach ($column in $requiredColumns) {
        if ($column -notin $csvColumns) {
            Write-Log -Message "Required column '$column' not found in CSV. Available columns: $($csvColumns -join ', ')" -Level "ERROR" -LogPath $LogPath
            exit 1
        }
    }
    
    Write-Log -Message "CSV validation successful. Required columns found." -Level "SUCCESS" -LogPath $LogPath
    
    # Initialize counters
    $successCount = 0
    $errorCount = 0
    $skippedCount = 0
    
    # Process each row in the CSV
    foreach ($row in $csvData) {
        $mailboxAlias = $row.MailboxAlias.Trim()
        $permissionHolder = $row.PermissionHolder.Trim()
        
        # Skip empty rows
        if ([string]::IsNullOrWhiteSpace($mailboxAlias) -or [string]::IsNullOrWhiteSpace($permissionHolder)) {
            Write-Log -Message "Skipping row with empty MailboxAlias or PermissionHolder" -Level "WARNING" -LogPath $LogPath
            $skippedCount++
            continue
        }
        
        Write-Log -Message "Processing: Granting '$permissionHolder' Full Access to '$mailboxAlias'" -LogPath $LogPath
        
        try {
            # Check if mailbox exists
            $mailbox = Get-Mailbox -Identity $mailboxAlias -ErrorAction Stop
            
            # Check if permission already exists
            $existingPermissions = Get-MailboxPermission -Identity $mailboxAlias -User $permissionHolder -ErrorAction SilentlyContinue
            
            if ($existingPermissions | Where-Object { $_.AccessRights -contains "FullAccess" }) {
                Write-Log -Message "Full Access permission already exists for '$permissionHolder' on '$mailboxAlias'" -Level "WARNING" -LogPath $LogPath
                $skippedCount++
                continue
            }
            
            # Grant Full Access permission
            Add-MailboxPermission -Identity $mailboxAlias -User $permissionHolder -AccessRights FullAccess -InheritanceType All -Confirm:$false
            
            Write-Log -Message "Successfully granted Full Access to '$permissionHolder' on '$mailboxAlias'" -Level "SUCCESS" -LogPath $LogPath
            $successCount++
        }
        catch {
            Write-Log -Message "Failed to grant permission for '$permissionHolder' on '$mailboxAlias': $($_.Exception.Message)" -Level "ERROR" -LogPath $LogPath
            $errorCount++
        }
    }
    
    # Summary
    Write-Log -Message "=== SCRIPT COMPLETED ===" -LogPath $LogPath
    Write-Log -Message "Total processed: $($csvData.Count)" -LogPath $LogPath
    Write-Log -Message "Successful: $successCount" -Level "SUCCESS" -LogPath $LogPath
    Write-Log -Message "Errors: $errorCount" -Level $(if ($errorCount -gt 0) { "ERROR" } else { "INFO" }) -LogPath $LogPath
    Write-Log -Message "Skipped: $skippedCount" -Level "WARNING" -LogPath $LogPath
    Write-Log -Message "Log file saved to: $LogPath" -LogPath $LogPath
    
}
catch {
    Write-Log -Message "Script failed with error: $($_.Exception.Message)" -Level "ERROR" -LogPath $LogPath
    exit 1
}

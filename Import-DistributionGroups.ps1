# Import-DistributionGroups.ps1
# Script to import Distribution Groups and Members from CSV export

# Input CSV file path - Update this to match your export file
$InputFile = "C:\Temp\DistributionGroups_Export_20251013_150015.csv"

# Log file
$LogFile = "C:\Temp\DistributionGroups_Import_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

# Function to write log
function Write-Log {
    param($Message, $Color = "White")
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$TimeStamp] $Message"
    Write-Host $LogMessage -ForegroundColor $Color
    Add-Content -Path $LogFile -Value $LogMessage
}

Write-Log "Starting Distribution Group Import Process" "Green"
Write-Log "Reading CSV file: $InputFile" "Cyan"

# Check if file exists
if (-not (Test-Path $InputFile)) {
    Write-Log "ERROR: CSV file not found at $InputFile" "Red"
    exit
}

# Import CSV
$ImportData = Import-Csv -Path $InputFile

Write-Log "Found $($ImportData.Count) records in CSV" "Yellow"

# Get unique distribution groups
$UniqueGroups = $ImportData | Select-Object 'Group Name', 'Group Display Name', 'Group Primary SMTP', 'Group Type' -Unique

Write-Log "Found $($UniqueGroups.Count) unique Distribution Groups to create" "Yellow"

# Track statistics
$GroupsCreated = 0
$GroupsSkipped = 0
$MembersAdded = 0
$MembersFailed = 0

# Create Distribution Groups
foreach ($Group in $UniqueGroups) {
    $GroupName = $Group.'Group Name'
    $GroupDisplayName = $Group.'Group Display Name'
    $GroupPrimarySMTP = $Group.'Group Primary SMTP'
    
    Write-Log "`nProcessing Group: $GroupName" "Cyan"
    
    # Check if group already exists
    $ExistingGroup = Get-DistributionGroup -Identity $GroupName -ErrorAction SilentlyContinue
    
    if ($ExistingGroup) {
        Write-Log "Group '$GroupName' already exists. Skipping creation." "Yellow"
        $GroupsSkipped++
    } else {
        try {
            # Create the distribution group
            New-DistributionGroup -Name $GroupName `
                                  -DisplayName $GroupDisplayName `
                                  -PrimarySmtpAddress $GroupPrimarySMTP `
                                  -ErrorAction Stop
            
            Write-Log "Successfully created group: $GroupName" "Green"
            $GroupsCreated++
            
            # Small delay to ensure group is fully created
            Start-Sleep -Seconds 2
        }
        catch {
            Write-Log "ERROR creating group '$GroupName': $($_.Exception.Message)" "Red"
            continue
        }
    }
    
    # Add members to the group
    $GroupMembers = $ImportData | Where-Object { $_.'Group Name' -eq $GroupName }
    
    foreach ($Member in $GroupMembers) {
        $MemberSMTP = $Member.'Member Primary SMTP'
        $MemberName = $Member.'Member Name'
        
        # Skip if no members
        if ($MemberSMTP -eq "No Members" -or [string]::IsNullOrWhiteSpace($MemberSMTP)) {
            continue
        }
        
        try {
            # Check if member exists in Exchange
            $Recipient = Get-Recipient -Identity $MemberSMTP -ErrorAction SilentlyContinue
            
            if ($Recipient) {
                # Check if member is already in the group
                $ExistingMember = Get-DistributionGroupMember -Identity $GroupName -ErrorAction SilentlyContinue | 
                                  Where-Object { $_.PrimarySmtpAddress -eq $MemberSMTP }
                
                if ($ExistingMember) {
                    Write-Log "  Member '$MemberName' already exists in group. Skipping." "Gray"
                } else {
                    # Add member to distribution group
                    Add-DistributionGroupMember -Identity $GroupName -Member $MemberSMTP -ErrorAction Stop
                    Write-Log "  Added member: $MemberName ($MemberSMTP)" "Green"
                    $MembersAdded++
                }
            } else {
                Write-Log "  WARNING: Recipient '$MemberName' ($MemberSMTP) not found in Exchange. Skipping." "Yellow"
                $MembersFailed++
            }
        }
        catch {
            Write-Log "  ERROR adding member '$MemberName' ($MemberSMTP): $($_.Exception.Message)" "Red"
            $MembersFailed++
        }
    }
}

# Summary
Write-Log "`n========================================" "Cyan"
Write-Log "Import Process Completed" "Green"
Write-Log "========================================" "Cyan"
Write-Log "Groups Created: $GroupsCreated" "Yellow"
Write-Log "Groups Skipped (already exist): $GroupsSkipped" "Yellow"
Write-Log "Members Added: $MembersAdded" "Yellow"
Write-Log "Members Failed/Not Found: $MembersFailed" "Yellow"
Write-Log "Log file saved to: $LogFile" "Cyan"
<#
.SYNOPSIS
    Creates mail-enabled security groups in Exchange Online and adds members from CSV file.

.DESCRIPTION
    This script reads a CSV file containing GroupName and UserPrincipalName columns,
    creates mail-enabled security groups, and adds the specified members to each group.

.PARAMETER CsvPath
    Path to the CSV file containing GroupName and UserPrincipalName columns.

.PARAMETER Domain
    The domain to use for group email addresses (e.g., "contoso.com").

.EXAMPLE
    .\Create-MailEnabledSecurityGroups.ps1 -CsvPath "C:\Groups.csv" -Domain "contoso.com"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,
    
    [Parameter(Mandatory=$true)]
    [string]$Domain
)

# Function to write colored output
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

# Function to connect to Exchange Online
function Connect-ExchangeOnlineIfNeeded {
    try {
        # Check if already connected
        $session = Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened"}
        
        if (-not $session) {
            Write-ColorOutput "Connecting to Exchange Online..." "Yellow"
            Connect-ExchangeOnline -ShowProgress $true
            Write-ColorOutput "Successfully connected to Exchange Online!" "Green"
        } else {
            Write-ColorOutput "Already connected to Exchange Online." "Green"
        }
    }
    catch {
        Write-ColorOutput "Failed to connect to Exchange Online: $($_.Exception.Message)" "Red"
        exit 1
    }
}

# Function to validate CSV file
function Test-CsvFile {
    param([string]$Path)
    
    if (-not (Test-Path $Path)) {
        Write-ColorOutput "CSV file not found: $Path" "Red"
        return $false
    }
    
    try {
        $csvData = Import-Csv $Path
        $requiredColumns = @("GroupName", "UserPrincipalName")
        $csvColumns = $csvData[0].PSObject.Properties.Name
        
        foreach ($column in $requiredColumns) {
            if ($column -notin $csvColumns) {
                Write-ColorOutput "Required column '$column' not found in CSV file." "Red"
                return $false
            }
        }
        
        Write-ColorOutput "CSV file validation successful." "Green"
        return $true
    }
    catch {
        Write-ColorOutput "Error reading CSV file: $($_.Exception.Message)" "Red"
        return $false
    }
}

# Function to create mail-enabled security group
function New-MailEnabledSecurityGroup {
    param(
        [string]$GroupName,
        [string]$Domain
    )
    
    try {
        # Create alias from group name (remove spaces and special characters)
        $alias = $GroupName -replace '[^a-zA-Z0-9]', ''
        $primarySmtpAddress = "$alias@$Domain"
        
        # Check if group already exists
        $existingGroup = Get-DistributionGroup -Identity $GroupName -ErrorAction SilentlyContinue
        
        if ($existingGroup) {
            Write-ColorOutput "Group '$GroupName' already exists. Skipping creation." "Yellow"
            return $existingGroup
        }
        
        # Create the mail-enabled security group
        $group = New-DistributionGroup -Name $GroupName -Type "Security" -Alias $alias -PrimarySmtpAddress $primarySmtpAddress
        
        Write-ColorOutput "Successfully created group: $GroupName ($primarySmtpAddress)" "Green"
        return $group
    }
    catch {
        Write-ColorOutput "Failed to create group '$GroupName': $($_.Exception.Message)" "Red"
        return $null
    }
}

# Function to add member to group
function Add-GroupMember {
    param(
        [string]$GroupName,
        [string]$UserPrincipalName
    )
    
    try {
        # Check if user exists
        $user = Get-Recipient -Identity $UserPrincipalName -ErrorAction SilentlyContinue
        
        if (-not $user) {
            Write-ColorOutput "User '$UserPrincipalName' not found. Skipping." "Yellow"
            return $false
        }
        
        # Check if user is already a member
        $existingMember = Get-DistributionGroupMember -Identity $GroupName -ResultSize Unlimited | Where-Object {$_.PrimarySmtpAddress -eq $UserPrincipalName}
        
        if ($existingMember) {
            Write-ColorOutput "User '$UserPrincipalName' is already a member of '$GroupName'." "Yellow"
            return $true
        }
        
        # Add user to group
        Add-DistributionGroupMember -Identity $GroupName -Member $UserPrincipalName
        Write-ColorOutput "Added '$UserPrincipalName' to group '$GroupName'" "Green"
        return $true
    }
    catch {
        Write-ColorOutput "Failed to add '$UserPrincipalName' to group '$GroupName': $($_.Exception.Message)" "Red"
        return $false
    }
}

# Main script execution
Write-ColorOutput "=== Exchange Online Mail-Enabled Security Groups Creation Script ===" "Cyan"
Write-ColorOutput "CSV Path: $CsvPath" "White"
Write-ColorOutput "Domain: $Domain" "White"
Write-ColorOutput ""

# Validate CSV file
if (-not (Test-CsvFile -Path $CsvPath)) {
    exit 1
}

# Connect to Exchange Online
Connect-ExchangeOnlineIfNeeded

# Import CSV data
Write-ColorOutput "Reading CSV data..." "Yellow"
$csvData = Import-Csv $CsvPath

# Get unique group names
$uniqueGroups = $csvData | Select-Object -Property GroupName -Unique
Write-ColorOutput "Found $($uniqueGroups.Count) unique groups to create." "White"

# Statistics tracking
$stats = @{
    GroupsCreated = 0
    GroupsSkipped = 0
    MembersAdded = 0
    MembersFailed = 0
}

# Create groups first
Write-ColorOutput "`nCreating mail-enabled security groups..." "Cyan"
foreach ($groupInfo in $uniqueGroups) {
    $group = New-MailEnabledSecurityGroup -GroupName $groupInfo.GroupName -Domain $Domain
    if ($group) {
        $stats.GroupsCreated++
    } else {
        $stats.GroupsSkipped++
    }
}

# Add members to groups
Write-ColorOutput "`nAdding members to groups..." "Cyan"
foreach ($row in $csvData) {
    if ([string]::IsNullOrWhiteSpace($row.GroupName) -or [string]::IsNullOrWhiteSpace($row.UserPrincipalName)) {
        Write-ColorOutput "Skipping row with empty GroupName or UserPrincipalName" "Yellow"
        continue
    }
    
    $success = Add-GroupMember -GroupName $row.GroupName -UserPrincipalName $row.UserPrincipalName
    if ($success) {
        $stats.MembersAdded++
    } else {
        $stats.MembersFailed++
    }
}

# Display final statistics
Write-ColorOutput "`n=== EXECUTION SUMMARY ===" "Cyan"
Write-ColorOutput "Groups Created: $($stats.GroupsCreated)" "Green"
Write-ColorOutput "Groups Skipped: $($stats.GroupsSkipped)" "Yellow"
Write-ColorOutput "Members Added: $($stats.MembersAdded)" "Green"
Write-ColorOutput "Members Failed: $($stats.MembersFailed)" "Red"
Write-ColorOutput "`nScript execution completed!" "Cyan"

# Optionally disconnect from Exchange Online
$disconnect = Read-Host "`nDisconnect from Exchange Online? (y/N)"
if ($disconnect -eq 'y' -or $disconnect -eq 'Y') {
    Disconnect-ExchangeOnline -Confirm:$false
    Write-ColorOutput "Disconnected from Exchange Online." "Green"
}

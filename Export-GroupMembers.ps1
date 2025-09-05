<#
.SYNOPSIS
    Exports all members from EntraID security groups listed in a CSV file.

.DESCRIPTION
    This script reads a CSV file containing EntraID security group names,
    retrieves all members from those groups, and exports the results to a new CSV file.

.PARAMETER InputCsvPath
    Path to the input CSV file containing group names (must have a 'GroupName' column)

.PARAMETER OutputCsvPath
    Path for the output CSV file that will contain group members

.EXAMPLE
    .\Export-GroupMembers.ps1 -InputCsvPath "C:\Groups.csv" -OutputCsvPath "C:\GroupMembers.csv"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$InputCsvPath,
    
    [Parameter(Mandatory = $true)]
    [string]$OutputCsvPath
)

# Check if Microsoft Graph PowerShell module is installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Groups)) {
    Write-Error "Microsoft Graph PowerShell module is not installed. Please install it using: Install-Module Microsoft.Graph"
    exit 1
}

# Import required modules
Import-Module Microsoft.Graph.Groups
Import-Module Microsoft.Graph.Users

try {
    # Connect to Microsoft Graph with required scopes
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Connect-MgGraph -Scopes "Group.Read.All", "User.Read.All" -NoWelcome
    
    # Verify connection
    $context = Get-MgContext
    if (-not $context) {
        throw "Failed to connect to Microsoft Graph"
    }
    Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
    
    # Read the input CSV file
    Write-Host "Reading input CSV file: $InputCsvPath" -ForegroundColor Yellow
    if (-not (Test-Path $InputCsvPath)) {
        throw "Input CSV file not found: $InputCsvPath"
    }
    
    $groups = Import-Csv $InputCsvPath
    
    # Validate CSV structure
    if (-not ($groups | Get-Member -Name "GroupName" -MemberType NoteProperty)) {
        throw "CSV file must contain a 'GroupName' column"
    }
    
    Write-Host "Found $($groups.Count) groups in CSV file" -ForegroundColor Green
    
    # Initialize array to store all members
    $allMembers = @()
    $processedGroups = 0
    $totalGroups = $groups.Count
    
    # Process each group
    foreach ($groupRow in $groups) {
        $groupName = $groupRow.GroupName.Trim()
        $processedGroups++
        
        Write-Host "[$processedGroups/$totalGroups] Processing group: $groupName" -ForegroundColor Cyan
        
        try {
            # Find the group by display name
            $group = Get-MgGroup -Filter "displayName eq '$groupName'" -ErrorAction Stop
            
            if (-not $group) {
                Write-Warning "Group not found: $groupName"
                continue
            }
            
            if ($group.Count -gt 1) {
                Write-Warning "Multiple groups found with name '$groupName'. Using the first one."
                $group = $group[0]
            }
            
            # Get group members
            $members = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop
            
            Write-Host "  Found $($members.Count) members" -ForegroundColor Gray
            
            # Process each member
            foreach ($member in $members) {
                try {
                    # Get user details to retrieve UPN
                    $user = Get-MgUser -UserId $member.Id -Property "UserPrincipalName,DisplayName,Mail" -ErrorAction Stop
                    
                    $memberInfo = [PSCustomObject]@{
                        GroupName = $groupName
                        GroupId = $group.Id
                        UserPrincipalName = $user.UserPrincipalName
                        DisplayName = $user.DisplayName
                        Email = $user.Mail
                        UserId = $user.Id
                    }
                    
                    $allMembers += $memberInfo
                }
                catch {
                    Write-Warning "  Failed to get details for member ID: $($member.Id) - $($_.Exception.Message)"
                }
            }
        }
        catch {
            Write-Error "Failed to process group '$groupName': $($_.Exception.Message)"
        }
    }
    
    # Export results to CSV
    if ($allMembers.Count -gt 0) {
        Write-Host "Exporting $($allMembers.Count) members to: $OutputCsvPath" -ForegroundColor Yellow
        $allMembers | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8
        Write-Host "Export completed successfully!" -ForegroundColor Green
        
        # Display summary
        Write-Host "`nSummary:" -ForegroundColor Cyan
        Write-Host "  Total groups processed: $processedGroups" -ForegroundColor White
        Write-Host "  Total members exported: $($allMembers.Count)" -ForegroundColor White
        Write-Host "  Output file: $OutputCsvPath" -ForegroundColor White
        
        # Show unique groups with member counts
        $groupSummary = $allMembers | Group-Object GroupName | Sort-Object Name
        Write-Host "`nMembers per group:" -ForegroundColor Cyan
        foreach ($group in $groupSummary) {
            Write-Host "  $($group.Name): $($group.Count) members" -ForegroundColor White
        }
    }
    else {
        Write-Warning "No members found in any of the specified groups"
    }
}
catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    exit 1
}
finally {
    # Disconnect from Microsoft Graph
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Host "Disconnected from Microsoft Graph" -ForegroundColor Gray
    }
    catch {
        # Ignore disconnect errors
    }
}

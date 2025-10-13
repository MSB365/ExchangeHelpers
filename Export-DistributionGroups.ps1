# Export-DistributionGroups.ps1
# Script to export all Distribution Groups with Member Names and Primary SMTP Addresses

# Output file path
$OutputFile = "C:\Temp\DistributionGroups_Export_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Create array to store results
$Results = @()

Write-Host "Connecting to Exchange Management Shell..." -ForegroundColor Green

# Get all Distribution Groups
Write-Host "Retrieving all Distribution Groups..." -ForegroundColor Green
$DistributionGroups = Get-DistributionGroup -ResultSize Unlimited | Sort-Object Name

Write-Host "Found $($DistributionGroups.Count) Distribution Groups" -ForegroundColor Yellow

# Loop through each Distribution Group
foreach ($Group in $DistributionGroups) {
    Write-Host "Processing: $($Group.Name)" -ForegroundColor Cyan
    
    # Get members of the distribution group
    $Members = Get-DistributionGroupMember -Identity $Group.Identity -ResultSize Unlimited
    
    if ($Members) {
        foreach ($Member in $Members) {
            # Create custom object for each member
            $Result = [PSCustomObject]@{
                'Group Name'              = $Group.Name
                'Group Display Name'      = $Group.DisplayName
                'Group Primary SMTP'      = $Group.PrimarySmtpAddress
                'Group Type'              = $Group.GroupType
                'Member Name'             = $Member.Name
                'Member Display Name'     = $Member.DisplayName
                'Member Primary SMTP'     = $Member.PrimarySmtpAddress
                'Member Type'             = $Member.RecipientType
                'Member Alias'            = $Member.Alias
            }
            $Results += $Result
        }
    } else {
        # Add group even if it has no members
        $Result = [PSCustomObject]@{
            'Group Name'              = $Group.Name
            'Group Display Name'      = $Group.DisplayName
            'Group Primary SMTP'      = $Group.PrimarySmtpAddress
            'Group Type'              = $Group.GroupType
            'Member Name'             = "No Members"
            'Member Display Name'     = "No Members"
            'Member Primary SMTP'     = "No Members"
            'Member Type'             = "N/A"
            'Member Alias'            = "N/A"
        }
        $Results += $Result
    }
}

# Export to CSV
Write-Host "`nExporting results to CSV..." -ForegroundColor Green
$Results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8

Write-Host "Export completed successfully!" -ForegroundColor Green
Write-Host "Total Groups: $($DistributionGroups.Count)" -ForegroundColor Yellow
Write-Host "Total Records: $($Results.Count)" -ForegroundColor Yellow
Write-Host "File saved to: $OutputFile" -ForegroundColor Yellow
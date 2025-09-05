# ExchangeHelper

## Overview
The **ExchangeHelper** repository contains a collection of PowerShell scripts designed to assist Exchange Engineers with migration tasks and management of Exchange Online environments. These scripts automate common tasks such as creating mail-enabled security groups, managing mailbox aliases, exporting group memberships, and handling shared mailbox permissions.

## Scripts

Below is a list of the scripts included in this repository, along with their descriptions and functionality:

### 1. Create-MailEnabledSecurityGroups.ps1
**Description**: Creates mail-enabled security groups in Exchange Online and adds members from a CSV file.  
**Functionality**: This script reads a CSV file containing `GroupName` and `UserPrincipalName` columns, creates mail-enabled security groups, and adds the specified members to each group.

### 2. EXOMailbox-addAlias.ps1
**Description**: Adds additional alias addresses to Exchange Online mailboxes from a CSV file.  
**Functionality**: This script reads a CSV file containing mailbox identities and their corresponding alias addresses, then adds these aliases to the specified mailboxes in Exchange Online.

### 3. Export-GroupMembers.ps1
**Description**: Exports all members from EntraID security groups listed in a CSV file.  
**Functionality**: This script reads a CSV file containing EntraID security group names, retrieves all members from those groups, and exports the results to a new CSV file.

### 4. Export-MailboxAliases-Simple.ps1
**Description**: Exports Exchange mailboxes and their email aliases to a CSV file in a simple format (one row per mailbox).  
**Functionality**: This script reads all user mailboxes from Exchange and exports information, including all email aliases, to a CSV file. All aliases are combined in a single column for easier overview and reporting.

### 5. Export-MailboxAliases.ps1
**Description**: Exports Exchange mailboxes and their email aliases to a CSV file in a detailed format (one row per alias).  
**Functionality**: This script reads all user mailboxes from Exchange and exports detailed information, including all email aliases, to a CSV file. Each alias gets its own row for detailed analysis.

### 6. Get-SharedMailboxGroupPermissions.ps1
**Description**: Exports shared mailbox permissions for groups to a CSV file.  
**Functionality**: This script connects to Exchange Online, retrieves all shared mailboxes, and exports information about which groups have full access permissions to each shared mailbox into a CSV file.

### 7. GrantFullAccessToSharedMailbox.ps1
**Description**: Grants Full Access permissions to mail-enabled security groups on shared mailboxes.  
**Functionality**: This script reads from a CSV file with `MailboxAlias` and `PermissionHolder` columns and grants Full Access permissions to the specified mail-enabled security groups for the listed shared mailboxes.

## Usage
1. **Prerequisites**:
   - Ensure you have the necessary permissions to manage Exchange Online and EntraID.
   - Install the required PowerShell modules (e.g., `ExchangeOnlineManagement`).
   - Prepare CSV files as required by each script, ensuring they contain the correct column headers.

2. **Running the Scripts**:
   - Clone or download this repository.
   - Open PowerShell and navigate to the repository folder.
   - Run each script with appropriate parameters (refer to individual script documentation or comments for details).

3. **CSV File Formats**:
   - Each script expects specific CSV file formats. Ensure the CSV files are correctly formatted as described in the script descriptions.

## License
This repository is licensed under the [MIT License](LICENSE).

### More about PowerShell and Graph Scripts here on my GitHub Profile or at my Blog: https://www.msb365.blog

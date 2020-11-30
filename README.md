# Powershell Scripts
PowerShell Scripts to manipulate AD &amp; Exchange Objects

## Create_PublicFolder_Redacted.ps1

This script creates a Public Folder, which is often used in a corporate environment as an alternative to a Shared Mailbox, and is intended to be run by staff who work on a Service Desk
- Before running the script, the staff member would need to update the input file, with the relevant name & email address for the new Public Folder
- The script itself is split into 2 parts to allow time for replication between AD & Exchange, and the staff member is prompted to confirm if they want to run each part of the script.
- The  first part reads the input file, create the PF, mail-enables it and creates the relevant Auth & SendAs security groups. 
- The second part mail-enables the security groups and adds them to the PF.
- Then staff member then needs to manually add end-users to the Public Folder using AD


## Export_ADGroups_ManagedBy_Redacted.ps1

This script exports a list of AD groups that are ‘managed by’ a given user 


## FindChange_ADGroupsManagedByUser_Redacted.ps1

This script find what AD security groups a given manager owns, and changes the ownership to another new manager – only one new manager can be specified

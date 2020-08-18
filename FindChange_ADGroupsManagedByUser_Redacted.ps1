<# 	

.SYNOPSIS
Script to find what AD security groups a given manager owns, and change the ownership to another new manager – only one new manager can be specified

If you want to change ownership to multiple new managers, see the pair of scripts Export_ManagerOf.ps1 & Import_NewManager.ps1 
http://sharepoint/_URL_REDACTED_/Scripts.aspx

.DESCRIPTION
This script outputs to text file in the current directory, for audit purposes only.
The file contains a list of AD groups managed by the old manager. 
The script then updates the ManagedBy field of each AD Group, to the new manager.
However if the old manager doesnt manage any AD groups, then nothing will be outputted to the file, and nothing updated in AD. 

.EXAMPLE
./FindChange_ADGroupsManagedByUser.ps1
Input the username of the old manager, so the groups they managed can be saved to a file
bloggsj
Input the username of the new manager that you want to take over the AD groups from the old manager
smithj

.INPUTS
$old_manager = name of AD User (Manager), to find out what AD Groups they manage/own. 
$new_manager = name of new AD User (Manager) that you want to transfer AD groups (ownership) to
FYI: this script will not find any Direct Reports of the AD user/manager - to change AD Groups & Direct Reports, use the scripts Export_ManagerOf.ps1 & Import_NewManager.ps1 

.OUTPUTS
File named "ChangeManager_from_$old_manager_to_$new_manager_$timestamp.txt" in current directory eg ChangeManager_from_bloggsj_to_smithj_2017-07-04-1044

.NOTES
Script created by - Patricia Hayden on 6 April 2017
Last Updated by - Patricia on 4 July 2017 (version 2 - added synopsis/inputs/output etc; addecd 'wait for input' option if script is run from console & tested; renamed script)
Contact - patricia.hayden.0@gmail.com
See Script KB webpage - http://sharepoint/_URL_REDACTED_/Scripts.aspx

#>

Import-Module ActiveDirectory

#find out the username of the old manager
$old_manager = Read-Host -Prompt "Input the username of the Old Manager, so the groups they managed can be saved to a file"

#find out the username of the new manager
$new_manager = Read-Host -Prompt "Input the username of the New Manager that you want to take over as manager of the AD groups that $old_manager currently owns."

#create file to store details of AD groups that old manager owns
$timestamp = (Get-Date -Format yyyy-MM-dd-hhmm)
$Filename =  "ChangeManager_from_" + $old_manager + "_to_" + $new_manager + "_" + $timestamp + ".txt"

# Prompt user to confirm they want to change manager of groups in AD - use this 2 line format to have a coloured confirmation prompt
Write-Host "`nCAUTION! You are about to change all the AD Groups that $old_manager owns/manages, to $new_manager - Enter y to confirm, or anything else to exit" -ForegroundColor Cyan
$confirmation = Read-Host 
    
if ($confirmation -eq 'y') {

    try {
        #list all the AD groups for the old manager
        $list_of_groups = Get-ADGroup -LDAPFilter "(ManagedBy=$((Get-ADuser -Identity $old_manager).distinguishedname))" 

        ForEach ($Group_Member in $list_of_groups) {
    	   #find the name of the group
    	   $temp_id = Get-ADGroup -Identity $Group_Member
    	   $temp_samname = Get-ADGroup -Identity $Group_Member -Properties SamAccountName | Select SamAccountName
    	
    	   #find the manager of the group	
    	   $temp_mgr = Get-ADGroup -Identity $Group_Member -Properties ManagedBy | Select-Object ManagedBy	

    	   #write the AD group name & OLD manager username to a text file
    	   "$temp_samname;  $temp_mgr" | Out-File -Append $Filename
    	
    	   #update the manager to the NEW manager username provided
    	   Set-ADGroup -Identity $temp_id -ManagedBy $new_manager
        }
        
        write-host “`nScript finished: For a list of security groups that were managed by $old_manager, before they were assigned to $new_manager, see file:  $Filename"
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] { # in case the user cant be found
        write-host “Exception Type: $($_.Exception.GetType().FullName).
        `rException Message: $($_.Exception.Message).
        `rAn Error occured - Perhaps the user $old_manager or $new_manager doesnt exist or their username was mis-spelt” -ForegroundColor Cyan
    }
    catch {
        write-host "Some other error occured. Perhaps the user doesn't manage any AD Groups. The output file may have been created but is likely empty"  -ForegroundColor Cyan
        write-host “Exception Type: $($_.Exception.GetType().FullName)
            `rException Message: $($_.Exception.Message)" -ForegroundColor Cyan
    }

} # end of confirmation prompt
else {
    write-host “Script cancelled: Nothing has been changed in AD."
}

# If running in the console, wait for Enter key press before closing the window, so any error messages can be seen.
if ($Host.Name -eq "ConsoleHost") {
    Read-Host -Prompt "Press Enter to quit"
}
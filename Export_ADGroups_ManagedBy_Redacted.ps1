<# 	

.SYNOPSIS
Script to export a list of AD Groups that are "managed by" a given user

.DESCRIPTION
This script outputs to text file called "ADGroups_ManagedBy_$username $timestamp.txt" in the current directory. 
The text file must then be opened in Excel and the data organized using 'Text to Columns' and delimiters Tab, Semicolon, Comma, (Other) Equals sign. 
Then use Text to Columns again, with the } delimiter against the column for AD Group Name
This script assumes the user is not deleted ie the user has NOT had all their group memberships removed from their AD account

.EXAMPLE
./Export_ADGroups_ManagedBy.ps1
Input the username of the user whose AD group membership and managers you want to export, eg bloggsjoe

.INPUTS
$Username = name of the AD user, to find all the AD Groups that are managed by them. FYI this will not find their Direct Reports.

.OUTPUTS
File named "ADGroups_ManagedBy_$username $timestamp.txt" in current directory eg ADGroups_ManagedBy_bloggsjoe 2017-06-12-1124

.NOTES
Script created by Davin Stirling, on 12 June 2017. 
Last Updated by Davin on 12 June 2017 (version 1). Took about 1hr, including testing & comments & basic exception handling
Last Updated 30 June (version 2) - added 'wait for input' option if script is run from console
Contact - davin.stirling@gmail.com


#>


Import-Module ActiveDirectory

#prompt for the username required
$Username = Read-Host -Prompt "Input the username of the required user, to export a list of AD Groups that are Managed By that user"

# create file to output results to
$timestamp = (Get-Date -Format yyyy-MM-dd-hhmm)
$Filename = "ADGroups_ManagedBy_" + $Username  + "_" + $timestamp + ".txt"

try {
    #list all the AD groups that this user manages
    $list_of_groups = Get-ADGroup -LDAPFilter "(ManagedBy=$((Get-ADuser -Identity $Username).distinguishedname))"

    #iterate through list, to output the properties SamAccountName & ManagedBy to file
    ForEach ($Group_Member in $list_of_groups) {
	   #find the name of the group
	   $temp_name = Get-ADGroup -Identity $Group_Member -Properties SamAccountName | Select-Object  SamAccountName
	
	   #find the manager of the group	
	   $temp_mgr = Get-ADGroup -Identity $Group_Member -Properties ManagedBy | Select-Object ManagedBy	

	   #write the name & manager of the AD group to a text file
	   "$temp_name;  $temp_mgr" | Out-File -Append $Filename
    }
    write-host “`nScript finished: For a list of security groups that are managed by this user, see file:  $Filename"

}
catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] { # in case the user cant be found
    write-host “Exception Type: $($_.Exception.GetType().FullName).
    `rException Message: $($_.Exception.Message).
    `rAn Error occured - Perhaps the user $Username doesnt exist or their username was mis-spelt” -ForegroundColor Cyan
}
catch {
    write-host "Some other error occured. Perhaps the user doesn't manage any AD Groups. Output file has not been created"  -ForegroundColor Cyan
    write-host “Exception Type: $($_.Exception.GetType().FullName)
        `rException Message: $($_.Exception.Message)" -ForegroundColor Cyan
}

# If running in the console, wait for Enter key press before closing the window, so any error messages can be seen.
if ($Host.Name -eq "ConsoleHost") {
    Read-Host -Prompt "Press Enter to quit"
}

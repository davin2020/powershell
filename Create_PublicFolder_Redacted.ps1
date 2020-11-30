<# 
	Powershell Script to Create a Public Folder
	Created by - Davin Stirling on 28 June 2017, based on Create_PF_part1.ps1 and Create_PF_part2.ps1
	Last Updated - 30 June 2017 (version 3 - combined previous scripts into 1, added some error checking & user confirmation prompts; updated exit-prompt & imported Exchange module properly)
	
    Input Parameters are read from a file called PublicFolder_Input.txt
    
    See ReadMe file in same directory, or Knowledge Base website for instructions - 
    http://sharepoint/_URL_REDACTED_/Script%20-%20Create%20Public%20Folder.aspx
    
	Caveat - this script contains minimal error checking!
    Reminder: 	FYI: backtick ` is used to escape a variable name, so its outputted to the console
#>

Import-Module ActiveDirectory
# the proper way to import Exchange modules
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1 #The script is dot sourced here
Connect-ExchangeServer -auto

# setup console buffer size, window size & title, code taken from - 
# https://blogs.technet.microsoft.com/heyscriptingguy/2006/12/04/how-can-i-expand-the-width-of-the-windows-powershell-console/
$pshost = get-host
$pswindow = $pshost.ui.rawui

$newsize = $pswindow.buffersize
$newsize.height = 3000
$newsize.width = 200
$pswindow.buffersize = $newsize

$newsize = $pswindow.windowsize
$newsize.height = 50
$newsize.width = 150
$pswindow.windowsize = $newsize

$pswindow.windowtitle = “Running Script - Create_PublicFolder.ps1, written by Davin Stirling”

<#
---------------------
Try and Read in paramaeters from input file 
Assumes file exists, as try/catch for Exception System.Management.Automation.ItemNotFoundException doenst work as expected 
---------------------
#>
try {
    $separator = ":"
    $InputText = Get-Content PublicFolder_Input.txt 
		
    [string]$PF_Alias = $InputText[0].split($separator)[1]	
    [string]$PF_Name = $InputText[1].split($separator)[1]
    [string]$PF_Email = $InputText[2].split($separator)[1]
    [string]$PF_Location = $InputText[3].split($separator)[1]
    [string]$PF_Description = $InputText[4].split($separator)[1] + " at location " + $PF_Location
    [string]$PF_DL_OU = $InputText[5].split($separator)[1]
    [string]$PF_Manager = $InputText[6].split($separator)[1]

    # create additional variables based on parameters inputted
    [string]$PF_Name_Location = $PF_Location +"\"+ $PF_Name
    [string]$PF_AuthGroup = "Auth-" + $PF_Alias
    [string]$PF_SendAsGroup = "SendAs-" + $PF_Alias
    [string]$PF_AuthGroupLocation = “OU=" + $PF_DL_OU + ",OU=Distribution Lists,OU=WWW,DC=XXX,DC=YYY,DC=ZZZ" # original location redacted, need to specify relevant location according to AD OU structure

    # output variables to console
    Write-Host "`nA new Public Folder will be created based on the input file PublicFolder_Input.txt, which contains the below data:" -ForegroundColor Cyan
    
    Write-Host "Alias: 			`$PF_Alias, 	 	value $PF_Alias"
    Write-Host "Name: 			`$PF_Name, 	 	value $PF_Name"
    Write-Host "Email: 			`$PF_Email, 	 	value $PF_Email"
    Write-Host "Location: 		`$PF_Location, 	 	value $PF_Location" 
    Write-Host "Description: 		`$PF_Description, 	value $PF_Description"
    Write-Host "Distribution List OU: 	`$PF_DL_OU, 	 	value $PF_DL_OU "
    Write-Host "Manager: 		`$PF_Manager, 	 	value $PF_Manager"

    Write-Host "Full Name & Location: 	`$PF_Name_Location, 	value $PF_Name_Location"
    Write-Host "Auth group: 		`$PF_AuthGroup, 		value $PF_AuthGroup"
    Write-Host "SendAs group: 		`$PF_SendAsGroup, 	value $PF_SendAsGroup"
    Write-Host "OU for Secuirty Groups: `$PF_AuthGroupLocation, 	value $PF_AuthGroupLocation"

    <#
    ==========================================================================================================================
    Check if user wants to run Part 1 of script - to create new Public Folder, mail-enable it and create Auth & SendAs groups 
    ==========================================================================================================================
    #>
    Write-Host "`nStep 1 - Are you sure you want to create a new Public Folder based on the above input? Enter Y to continue, or anything else to skip this step:" -ForegroundColor Cyan 
    $confirmation_part1 = Read-Host

    if ($confirmation_part1 -eq 'y') {   
        Write-Host "1A. Creating public folder..."
        New-PublicFolder -Name $PF_Name -Path $PF_Location

        Write-Host "1B. Mail enabling the public folder..."
        Enable-MailPublicFolder -Identity $PF_Name_Location

        Write-Host "1C. Setting alias and user-specified email addresses..." #need $PF_Alias
        Set-MailPublicFolder -Identity $PF_Name -Alias $PF_Alias  -EmailAddresses smtp:$PF_Email  

        Read-Host -Prompt "Press Enter to continue"
        Write-Host "1D. Showing current email addresses for new mail-enabled PF... `$PF_Name_Location value $PF_Name_Location"
        Get-MailPublicFolder -Identity $PF_Name_Location   	
            
        #Read-Host -Prompt "Press Enter to continue"
        Write-Host "1E. Creating distribution security Auth AD group & setting manager..."
        New-AdGroup –name $PF_AuthGroup –groupscope Universal –path $PF_AuthGroupLocation
        Set-ADGroup -Identity $PF_AuthGroup –Description $PF_Description –ManagedBy $PF_Manager

        Write-Host "1F. Creating distribution security SendAs AD group & setting manager..."
        New-AdGroup –name $PF_SendAsGroup –groupscope Universal –path $PF_AuthGroupLocation
        Set-ADGroup -Identity $PF_SendAsGroup –Description $PF_Description –ManagedBy $PF_Manager

        # remind user to wait 20mins
        $currenttime = Get-date
        Write-Host "Please wait at least 20 mins from now for AD/Exchange replication: $currenttime"
        Start-Sleep -s 10
        Write-Host "While you wait, please manually add some users to the AD security groups $PF_AuthGroup and $PF_SendAsGroup. `nPlease also check the email addresses created on the Public Folder and select the appropriate default"
    } # end of part 1 confirmation   
      
    else {
    	Write-Host "Skipping Part 1 - creating new Public Folder."
    }

    <#
    ================================================================================================================
    Check if user wants to run Part 2 of script - to mail-enable security groups and add them to the Public Folder 
    ================================================================================================================
    #>
    Write-Host "`nStep 2 - Are you sure you want mail-enable the security groups and add them to the Public Folder? Enter Y to continue, or anything else to skip this step:" -ForegroundColor Cyan 
    $confirmation_part2 = Read-Host

    if ($confirmation_part2 -eq 'y') { 
        Write-Host "2A. Mail enabling Auth security group..."
        Enable-DistributionGroup -Identity $PF_AuthGroup

        Write-Host "2B. Adding Auth group to PF..."
        Add-PublicFolderClientPermission -Identity $PF_Name_Location –User $PF_AuthGroup –AccessRights PublishingEditor

        #Read-Host -Prompt "Press Enter to continue"
        Write-Host "2C. Adding SendAs group to PF..."
        Add-ADPermission –Identity $PF_Name –User $PF_SendAsGroup –ExtendedRights ‘Send-as'

        Write-Host "2D. Showing PF permissions for PF `$PF_Name_Location $PF_Name_Location and security group `$PF_AuthGroup $PF_AuthGroup..."
        Get-PublicFolderClientPermission $PF_Name_Location  –User $PF_AuthGroup  
        
        # trying to show permissions doesnt always work...not sure why...needs more work
        Write-Host "2E. Showing PF permissions for PF `$PF_Name $PF_Name and security group `$PF_SendAsGroup $PF_SendAsGroup..."
        Get-ADPermission –Identity "$PF_Name" –User $PF_SendAsGroup  
        # need backtick ` to escape double-quotes around  PF Name, in case it contains a space

        Write-Host "`nScript has finished and created new Public Folder, if there are no errors above. Please send a test email to the public folder and add any members to the relevant security groups $PF_AuthGroup & $PF_SendAsGroup." -ForegroundColor Cyan 
    } # end of part 2 confirmation
        
    else {
    	Write-Host "Skipping Part 2 - adding security groups to new Public Folder."
    }

} # end of trying to read input file -------------
catch {
    Write-Host "`nSome other error occured -"
    write-host “Exception Type: $($_.Exception.GetType().FullName)
        `rException Message: $($_.Exception.Message)" -ForegroundColor Cyan
}

# If running in the console, wait for Enter key press before closing the window, so any error messages can be seen.
if ($Host.Name -eq "ConsoleHost") {
    Read-Host -Prompt "Press Enter to quit"
}

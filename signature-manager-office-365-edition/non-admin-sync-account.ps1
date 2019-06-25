#
<#
.SYNOPSIS
    Applies the permissions required to update signatures as the sync account in Signature Manager Office 365 Edition
.DESCRIPTION
    This script is designed to be run as part of the steps of the following Knowledge Base article.
    https://support.exclaimer.com/hc/en-gb/articles/360004374711-Using-a-dedicated-non-admin-user-for-aggregation-and-updating-signatures-with-Signature-Manager-Office-365-Edition

    Once complete, this will provide you with an account that can have its licence removed and still be used as a sync admin for Exclaimer

    Please refer to the REQUIREMENTS for the information needed to run this script correctly.
.NOTES
    Email: support@exclaimer.com
	Date: 3rd December 2018
.PRODUCTS
	Signature Manager Office 365 Edition
	Signature Manager Outlook Edition
.REQUIREMENTS
    - Global Administrator Credentials
	- An account to use as the admin
	- For Signature Manager Outlook Edition, you will need to use OWA Update from Server
.HISTORY
    1.0 - Original script completed with a new management role called MailboxMessageConfiguration
    1.1 - Changed script on advise of customer to create management role ExclaimerMailboxMessageConfiguration to improve clarity
#>

$cred = get-credential -message "Enter the Office 365 Admin email address and password"

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Cred -Authentication Basic -AllowRedirection
Import-PSSession $Session

$user = Read-Host "Enter the username you wish to use to as your Office 365 Administrator account"

if ((get-user $user) -eq $null)
{

	Write-Host "Can't find that user" -Fore "Red" -Back "Black"

}
else
{

	New-ManagementRole -Name "ExclaimerMailboxMessageConfiguration" -Parent "User Options"

	foreach ($role in Get-ManagementRoleEntry "ExclaimerMailboxMessageConfiguration\*")
	{
		if ($role.name -ne "Set-MailboxMessageConfiguration")
		{
			Remove-ManagementRoleEntry $("ExclaimerMailboxMessageConfiguration\"+$role.name) -Confirm:$false
		}
	}

	New-ManagementRoleAssignment -Role "ApplicationImpersonation" -User $user

	New-ManagementRoleAssignment -Role "ExclaimerMailboxMessageConfiguration" -User $user

	Add-RoleGroupMember -Identity "View-Only Organization Management" -Member $user
}
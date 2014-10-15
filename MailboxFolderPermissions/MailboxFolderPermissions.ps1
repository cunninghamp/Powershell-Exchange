<#
.SYNOPSIS
MailboxFolderPermissions.ps1

.DESCRIPTION 
A proof of concept script for adding mailbox folder
permissions to all folders in a mailbox.

.OUTPUTS
Console output for progress.

.PARAMETER Mailbox
The mailbox that the folder permissions will be added to.

.PARAMETER User
The user you are granting mailbox folder permissions to.

.PARAMETER Access
The permissions to grant for each folder.

.EXAMPLE
.\MailboxFolderPermissions.ps1 -Mailbox alex.heyne -User alan.reid -Access Reviewer

This will grant Alan Reid "Reviewer" access to all folders in Alex Heyne's mailbox.

.LINK
http://exchangeserverpro.com/grant-read-access-exchange-mailbox/

.NOTES
Written by: Paul Cunningham

For more Exchange Server tips, tricks and news
check out Exchange Server Pro.

* Website:	http://exchangeserverpro.com
* Twitter:	http://twitter.com/exchservpro

Find me on:

* My Blog:	http://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	http://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp

Change Log:
V1.00, 15/10/2014 - Initial version
#>

#requires -version 2

[CmdletBinding()]
param (
	[Parameter( Mandatory=$true)]
	[string]$Mailbox,
    
	[Parameter( Mandatory=$true)]
	[string]$User,
    
  	[Parameter( Mandatory=$true)]
	[string]$Access
)


#...................................
# Variables
#...................................

$exclusions = @("/Sync Issues",
                "/Sync Issues/Conflicts",
                "/Sync Issues/Local Failures",
                "/Sync Issues/Server Failures",
                "/Recoverable Items",
                "/Deletions",
                "/Purges",
                "/Versions"
                )


#...................................
# Initialize
#...................................

#Add Exchange 2010 snapin if not already loaded in the PowerShell session
if (!(Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
{
	try
	{
		Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
	}
	catch
	{
		#Snapin was not loaded
		Write-Warning $_.Exception.Message
		EXIT
	}
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	Connect-ExchangeServer -auto -AllowClobber
}


#Set scope to include entire forest
if (!(Get-ADServerSettings).ViewEntireForest)
{
	Set-ADServerSettings -ViewEntireForest $true -WarningAction SilentlyContinue
}


#...................................
# Script
#...................................

$mailboxfolders = @(Get-MailboxFolderStatistics $Mailbox | Where {!($exclusions -icontains $_.FolderPath)} | Select FolderPath)

foreach ($mailboxfolder in $mailboxfolders)
{
    $folder = $mailboxfolder.FolderPath.Replace("/","\")
    if ($folder -match "Top of Information Store")
    {
       $folder = $folder.Replace(“\Top of Information Store”,”\”)
    }
    $identity = "$($mailbox):$folder"
    Write-Host "Adding $user to $identity with $access permissions"
    Add-MailboxFolderPermission -Identity $identity -User $user -AccessRights $Access -ErrorAction SilentlyContinue
}


#...................................
# End
#...................................
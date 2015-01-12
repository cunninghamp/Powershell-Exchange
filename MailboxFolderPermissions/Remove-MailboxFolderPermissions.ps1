<#
.SYNOPSIS
Remove-MailboxFolderPermissions.ps1

.DESCRIPTION 
A proof of concept script for removing mailbox folder
permissions to all folders in a mailbox.

.OUTPUTS
Console output for progress.

.PARAMETER Mailbox
The mailbox that the folder permissions will be removed from.

.PARAMETER User
The user you are removing mailbox folder permissions for.

.EXAMPLE
.\Remove-MailboxFolderPermissions.ps1 -Mailbox alex.heyne -User alan.reid

This will remove Alan Reid's permissions to all folders in Alex Heyne's mailbox.

.LINK
http://exchangeserverpro.com/powershell-script-remove-permissions-exchange-mailbox/

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
V1.00, 12/01/2014 - Initial version
#>

#requires -version 2

[CmdletBinding()]
param (
	[Parameter( Mandatory=$true)]
	[string]$Mailbox,
    
	[Parameter( Mandatory=$true)]
	[string]$User
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
    Write-Host "Checking $identity for permissions for user $user"
    if (Get-MailboxFolderPermission -Identity $identity -User $user -ErrorAction SilentlyContinue)
    {
        try
        {
            Remove-MailboxFolderPermission -Identity $identity -User $User -Confirm:$false -ErrorAction STOP
            Write-Host -ForegroundColor Green "Removed!"
        }
        catch
        {
            Write-Warning $_.Exception.Message
        }
    }
}


#...................................
# End
#...................................
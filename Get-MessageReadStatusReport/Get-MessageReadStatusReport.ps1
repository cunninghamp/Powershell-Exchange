<#
.SYNOPSIS
Get-MessageReadStatusReport.ps1

.PARAMETER InputFileName


.EXAMPLE
.\New-LabUsers.ps1
Uses the RandomNameList.txt file to generate the list of user accounts in an OU
called "Company" in Active Directory.

.EXAMPLE
.\New-LabUsers.ps1 -InputFileName .\MyNames.txt -PasswordLength 8 -OU TestLab
Uses the MyNames.txt file to generate the list of user accounts in an OU
called "TestLab" in Active Directory.

.NOTES
Script written by: Paul Cunningham

Find me on:

* My Blog:	http://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	http://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp

For more Exchange Server tips, tricks and news
check out Exchange Server Pro.

* Website:	http://exchangeserverpro.com
* Twitter:	http://twitter.com/exchservpro

Change Log
V1.00, 9/06/2016 - Initial version
#>

[CmdletBinding()]
param (
	
	[Parameter( Mandatory=$true)]
	[string]$Mailbox,

    [Parameter( Mandatory=$true)]
    [string]$MessageId

	)

$output = @()

#First check if read tracking is enabled, and exit if it is not
if (!(Get-OrganizationConfig).ReadTrackingEnabled) {
    throw "Read tracking is not enabled for your Exchange organization."
}

#Get the message ID
$msg = Search-MessageTrackingReport -Identity $Mailbox -BypassDelegateChecking -MessageId $MessageId

#Should be exactly 1 result for the script to work
if ($msg.count -ne 1) {
    throw "$($msg).count messages found with that ID. Something is wrong."
}

#Get the message tracking report
$report = Get-MessageTrackingReport -Identity $msg.MessageTrackingReportId -BypassDelegateChecking

#Extract the recipient tracking events from the message tracking report
$recipienttrackingevents = @($report | Select -ExpandProperty RecipientTrackingEvents)

#Generate a list of recipients to loop through
$recipients = $recipienttrackingevents | select recipientaddress

#Loop through recipients list retrieving the read status (if present) for each recipient
foreach ($recipient in $recipients) {

    $events = Get-MessageTrackingReport -Identity $msg.MessageTrackingReportId -BypassDelegateChecking `
    -RecipientPathFilter $recipient.RecipientAddress -ReportTemplate RecipientPath
        
    $outputline = $events.RecipientTrackingEvents[-1] | Select RecipientAddress,Status,EventDescription

    $output += $outputline
}

#Here's the report
$output
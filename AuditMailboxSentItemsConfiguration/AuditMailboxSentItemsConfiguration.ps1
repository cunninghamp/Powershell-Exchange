<#
.SYNOPSIS
AuditMailboxSentItemsConfiguration.ps1

.DESCRIPTION 
Description

.OUTPUTS
Results are output to a CSV file.

.PARAMETER all
Generates a report for all mailboxes in the organization.

.PARAMETER server
Generates a report for all mailboxes on the specified server.

.PARAMETER database
Generates a report for all mailboxes on the specified database.

.EXAMPLE
.\AuditMailboxSentItemsConfiguration.ps1 -All -Verbose

.LINK
http://exchangeserverpro.com/powershell-script-audit-mailbox-sent-items-configurations

.NOTES
Written By: Paul Cunningham
Website:	http://exchangeserverpro.com
Twitter:	http://twitter.com/exchservpro

Change Log
V1.00, 10/04/2014, Initial version
#>

[CmdletBinding()]
param(
	[Parameter(ParameterSetName='database')] [string]$database,
	[Parameter(ParameterSetName='server')] [string]$server,
	[Parameter(ParameterSetName='all')] [switch]$all
)


#Set recipient scope
Set-ADServerSettings -ViewEntireForest $true


#Generate report file name with random strings for uniqueness
#Thanks to @proxb and @chrisbrownie for the help with random string generatio

$timestamp = Get-Date -UFormat %Y%m%d-%H%M
$random = -join(48..57+65..90+97..122 | ForEach-Object {[char]$_} | Get-Random -Count 6)
$reportfile = "MailboxSentItemsConfigReport-$timestamp-$random.csv"

$report = @()

Write-Verbose "Fetching mailbox list"

if ($all) { $mailboxes = @(Get-Mailbox -resultsize Unlimited) }
if ($server) { $mailboxes = @(Get-Mailbox -server $server -resultsize Unlimited) }
if ($database) { $mailboxes = @(Get-Mailbox -database $database -resultsize unlimited) }


#Loop through the mailbox list
foreach ($mailbox in $mailboxes)
{
    Write-Host "Processing $mailbox"

    $sentitemsconfig = Get-MailboxSentItemsConfiguration $mailbox
    $sendas = $sentitemsconfig.SendAsItemsCopiedTo
    $sendonbehalf = $sentitemsconfig.SendOnBehalfOfItemsCopiedTo

    if (($sendas -ne "Sender") -or ($sendonbehalf -ne "Sender"))
    {
        $reportObj = New-Object PSObject
        $reportObj | Add-Member NoteProperty -Name "Mailbox" -Value $mailbox
        $reportObj | Add-Member NoteProperty -Name "SendAsItemsCopiedTo" -Value $sendas
        $reportObj | Add-Member NoteProperty -Name "SendonBehalfItemsCopiedTo" -Value $sendonbehalf

        Write-Verbose "- Send as: $sendas"
        Write-Verbose "- Send on Behalf: $sendonbehalf"
        Write-Verbose "Adding $mailbox to report"

        $report += $reportObj
    }
}

#Output results if any
if ($report)
{
    Write-Host "Results output to $reportfile."
    $report | Export-Csv $reportfile -NoTypeInformation -Encoding UTF8
}
else
{
    Write-Host "No mailboxes with non-default sent items configuration were found."
}
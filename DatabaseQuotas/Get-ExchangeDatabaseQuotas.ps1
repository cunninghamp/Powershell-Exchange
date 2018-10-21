<#
.SYNOPSIS
Get-ExchangeDatabaseQuotas.ps1 - Exchange Database Storage Quota Report Script

.DESCRIPTION 
Generates a report of the storage quota configurations for Exchange Server databases

.OUTPUTS
Outputs to CSV files

.EXAMPLE
.\StorageQuotas.ps1
Reports storage quota configuration for all Exchange mailbox 
and public folder databases and outputs to CSV files.

.LINK
http://exchangeserverpro.com/powershell-script-audit-exchange-server-database-storage-quotas/

.NOTES
Written by: Paul Cunningham

Find me on:

* My Blog:	https://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	https://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp

Change Log
V1.00, 11/03/2014 - Initial Version
V1.01, 26/08/2015 - Minor update
V1.02, 03/11/2016 - Changed values from GB to MB for accuracy
#>

#requires -version 2


#...................................
# Variables
#...................................

$mbxreport = @()
$pfreport = @()

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$mbxreportfile = "$myDir\ExchangeDatabaseQuotas-MailboxDB.csv"
$pfreportfile = "$myDir\ExchangeDatabaseQuotas-PublicFolderDB.csv"

$mbxquotas = @("IssueWarningQuota",
            "ProhibitSendQuota",
            "ProhibitSendReceiveQuota",
            "RecoverableItemsQuota",
            "RecoverableItemsWarningQuota"
            )

$pfquotas = @("ProhibitPostQuota",
            "IssueWarningQuota"
            )


#...................................
# Script
#...................................

$mbxdatabases = @(Get-MailboxDatabase | select MasterServerOrAvailabilityGroup,Name,IssueWarningQuota,ProhibitSendQuota,ProhibitSendReceiveQuota,RecoverableItemsQuota,RecoverableItemsWarningQuota)
$pfdatabases = @(Get-PublicFolderDatabase | select Server,Name,IssueWarningQuota,ProhibitPostQuota)


if ($mbxdatabases)
{
    foreach ($database in $mbxdatabases)
    {
        Write-Host "Processing $($database.Name)"
        $mbxreportObj = New-Object PSObject
	    $mbxreportObj | Add-Member NoteProperty -Name "DAG/Server" -Value $database.MasterServerOrAvailabilityGroup
	    $mbxreportObj | Add-Member NoteProperty -Name "Database" -Value $database.Name
 
        foreach ($quota in $mbxquotas)
        {
            if (($database."$quota").IsUnlimited -eq $true)
            {
                $mbxreportObj | Add-Member NoteProperty -Name "$quota (MB)" -Value "Unlimited"
            }
            else
            {
                $mbxreportObj | Add-Member NoteProperty -Name "$quota (MB)" -Value $($database."$quota").Value.ToMB()
            }
        }
    
        $mbxreport += $mbxreportObj
    }

    $mbxreport | Export-CSV -NoTypeInformation -Path $mbxreportfile -Encoding UTF8
    Write-Host "Mailbox database storage quota report saved as $mbxreportfile"
}
else
{
    Write-Host "No mailbox databases found."
}


if ($pfdatabases)
{
    foreach ($database in $pfdatabases)
    {
        Write-Host "Processing $($database.Name)"
        $pfreportObj = New-Object PSObject
	    $pfreportObj | Add-Member NoteProperty -Name "Database" -Value $database.Name
 
        foreach ($quota in $pfquotas)
        {
            if (($database."$quota").IsUnlimited -eq $true)
            {
                $pfreportObj | Add-Member NoteProperty -Name "$quota (MB)" -Value "Unlimited"
            }
            else
            {
                $pfreportObj | Add-Member NoteProperty -Name "$quota (MB)" -Value $($database."$quota").Value.ToMB()
            }
        }
    
        $pfreport += $pfreportObj
    }

    $pfreport | Export-CSV -NoTypeInformation -Path $pfreportfile -Encoding UTF8
    Write-Host "Public folder database storage quota report saved as $pfreportfile"
}
else
{
    Write-Host "No public folder databases found."
}

Write-Host "Finished."

#...................................
# Finished
#...................................

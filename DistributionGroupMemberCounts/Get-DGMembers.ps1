<#
.SYNOPSIS
Get-DGMembers.ps1 - Get the members email address of every distribution group

.DESCRIPTION 
This PowerShell script returns the members email addres of every distribution group
in the Exchange organization.

.OUTPUTS
Results are output to a CSV file.

.EXAMPLE
.\Get-DGMembers.ps1
Created the report of distribution group member.

#>

#requires -version 2

[CmdletBinding()]
param ()


#...................................
# Variables
#...................................

$now = Get-Date											#Used for timestamps
$date = $now.ToShortDateString()						#Short date format for email message subject

$report = @()

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path


#...................................
# Script
#...................................

#Add Exchange 2010 snapin if not already loaded in the PowerShell session
if (Test-Path $env:ExchangeInstallPath\bin\RemoteExchange.ps1)
{
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	Connect-ExchangeServer -auto -AllowClobber
}
else
{
    Write-Warning "Exchange Server management tools are not installed on this computer."
    EXIT
}

#Set scope to entire forest
Set-ADServerSettings -ViewEntireForest:$true

#Get distribution groups
$distgroups = @(Get-DistributionGroup -ResultSize Unlimited)

#Process each distribution group
foreach ($dg in $distgroups)
{
	$list = @(Get-DistributionGroupMember -Identity $dg.Name)
	
	$members = ($list | Select-Object -ExpandProperty PrimarySmtpAddress) -join ";"
	$count = ($list | Select-Object -ExpandProperty PrimarySmtpAddress).Count

    $reportObj = New-Object PSObject
    $reportObj | Add-Member NoteProperty -Name "Group Name" -Value $dg.Name
    $reportObj | Add-Member NoteProperty -Name "DN" -Value $dg.distinguishedName
    $reportObj | Add-Member NoteProperty -Name "Manager" -Value $dg.managedby.Name
	$reportObj | Add-Member NoteProperty -Name "MembersCount" -Value $count
    $reportObj | Add-Member NoteProperty -Name "MembersList" -Value $members

    $report += $reportObj

}

$report | Export-CSV -Path $myDir\DistributionGroupMembers.csv -NoTypeInformation -Encoding UTF8


#...................................
# Finished
#...................................

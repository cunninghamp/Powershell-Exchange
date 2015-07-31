<#
.SYNOPSIS
MailFlowHeatMap.ps1 - Mail flow latency heat map generation script.

.DESCRIPTION 
Generates a HTML report of the mail flow latency
between mailbox servers in the organization.

.OUTPUTS
Report is output to a HTML file.

.EXAMPLE
PS C:\> .\MailFlowHeatMap.ps1
Generates the mail flow heat map file.

.LINK
http://exchangeserverpro.com/create-exchange-mail-flow-latency-heat-map-powershell

.NOTES
Written By: Paul Cunningham

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
V1.0, 22/06/2012 - Initial version
#>

#...................................
# Variables
#...................................

$report = @()
$mailboxservers = @()

[int]$hot = 30
[int]$warm = 20
[int]$cool = 10
[int]$cold = 10
$filename = "mailflowheatmap.html"


#...................................
# Script
#...................................

#Add dependencies
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue

#Get list of mailbox servers
$tempservers = @(Get-ExchangeServer | Where {$_.IsMailboxServer} | Sort Site,Name)

#Discard servers with no mailbox databases (eg Public Folder servers)
foreach ($server in $tempservers)
{
	if ($((Get-MailboxDatabase -server $server | where {$_.Recovery -ne $true}).Count) -gt 0)
	{
		$mailboxservers += $server
	}
}

#Test mail flow for each server against all other servers
foreach ($sourceserver in $mailboxservers)
{

	$reportObj = New-Object PSObject
	$reportObj | Add-Member NoteProperty -Name "Server" -Value $sourceserver.name
	
	foreach ($targetserver in $mailboxservers)
    {
        try {
            Write-Host "Testing $sourceserver to $targetserver"
    		$result = Test-Mailflow -Identity $sourceserver -TargetMailboxServer $targetserver -ErrorAction Stop
    	}
    	catch
    	{
    		$msg = $_.Exception.Message
            Write-Warning "Message latency from $sourceserver to $targetserver not tested due to an error."
			Write-Warning $msg
            $result = $null
    	}
    		
    	if ($result)
    	{
			[double]$latency = "{0:N2}" -f $($result.MessageLatencyTime.TotalSeconds)
    		$reportObj | Add-Member NoteProperty -Name "To $targetserver" -Value $latency
    	}
    	else
    	{
			$reportObj | Add-Member NoteProperty -Name "To $targetserver" -Value "Error"
    	}
    }
	$report = $report += $reportObj
}

$reportime = Get-Date

$headers = $report | Get-Member -MemberType NoteProperty | Select Name

$htmlhead="<html>
			<style>
			BODY{font-family: Arial; font-size: 8pt;}
			H1{font-size: 16px;}
			H2{font-size: 14px;}
			H3{font-size: 12px;}
			TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
			TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
			TD{border: 1px solid black; padding: 5px; text-align: center;}
			td.cold{background: #B3E6B3;}
			td.cool{background: #E2F09C;}
			td.warm{background: #FF6E6E;}
			td.hot{background: #E34949; color: #ffffff;}
			td.error{color:#FF0000;}
			</style>
			<body>
			<h1 align=""center"">Message Latency Heat Map</h1>
			<h3 align=""center"">Generated: $reportime</h3>
			<p>Message latency values are in seconds.</p>"

$htmltableheadstart = "<p>
					<table>
					<tr>"
foreach ($header in $headers)
{
	$name = $header.Name
	$htmltableheaders = $htmltableheaders + "<th>$name</th>"
}

$htmltableheadend = "</tr>"

$htmltableheader = $htmltableheadstart + $htmltableheaders + $htmltableheadend

foreach ($reportline in $report)
{
	$htmltablerow = "<tr>"
	foreach ($header in $headers)
	{
		$name = $header.Name
		$value = $reportline.$name
		if ($name -eq "Server")
		{
			$htmltablerow = $htmltablerow += "<td>$value</td>"
		}
		elseif ($value -eq "Error")
		{			
			$htmltablerow = $htmltablerow += "<td class=""error"">$value</td>"
		}
		elseif ($value -gt $hot)
		{			
			$htmltablerow = $htmltablerow += "<td class=""hot"">$value</td>"
		}
		elseif ($value -gt $warm)
		{			
			$htmltablerow = $htmltablerow += "<td class=""warm"">$value</td>"
		}
		elseif ($value -gt $cool)
		{			
			$htmltablerow = $htmltablerow += "<td class=""cool"">$value</td>"
		}
		elseif ($value -le $cold)
		{			
			$htmltablerow = $htmltablerow += "<td class=""cold"">$value</td>"
		}
		else
		{			
			$htmltablerow = $htmltablerow += "<td>$value</td>"
		}
	}
	$htmltablerow = $htmltablerow + "</tr>"
	$htmltable = $htmltable + $htmltablerow
}

$htmltable = $htmltableheader + $htmltable

$htmltail = "</body>
			</html>"
			
$htmlreport = $htmlhead + $htmltable + $htmltail

$htmlreport | Out-File $filename

Write-Output "Done."


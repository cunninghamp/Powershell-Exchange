<#
.SYNOPSIS
Test-TransportServerBackPressure.ps1 - Script to check Transport Servers for Back Pressure.

.DESCRIPTION 
Checks the event logs of Hub Transport servers for "back pressure" events.

.OUTPUTS
Results are output to the PowerShell window.

.PARAMETER server
Perform a check of a single server

.EXAMPLE
.\Test-TransportServerBackPressure.ps1
Checks all Hub Transport servers in the organization and outputs the results to the shell window.

.EXAMPLE
.\Test-TransportServerBackPressure.ps1 -server HO-EX2010-MB1
Checks the server HO-EX2010-MB1 and outputs the results to the shell window.

.LINK
http://exchangeserverpro.com/powershell-script-check-hub-transport-servers-for-back-pressure-events

.NOTES
Written by: Paul Cunningham

Find me on:

* My Blog:	https://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	https://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp
#>

#requires -version 2

param (
	[Parameter( Mandatory=$false)]
	[string]$server
)

#Add Exchange 2010 snapin if not already loaded
if (!(Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue
}

#...................................
# Script
#...................................


#Check if a single server was specified
if ($server)
{
	#Run for single specified server
	try
	{
		[array]$servers = @(Get-ExchangeServer $server -ErrorAction Stop)
	}
	catch
	{
		#Couldn't find Exchange server of that name
		Write-Warning $_.Exception.Message
		
		#Exit because single server was specified and couldn't be found in the organization
		EXIT
	}
}
else
{
	#Get list of Hub Transport servers in the organization
	[array]$servers = @(Get-ExchangeServer | Where-Object {$_.IsHubTransportServer})
}

#Check each server
foreach($server in $servers)
{
	$events = @(Invoke-Command –Computername $server –ScriptBlock { Get-EventLog -LogName Application | Where-Object {$_.Source -eq "MSExchangeTransport" -and $_.Category -eq "ResourceManager"} })
	$count = $events.count

	if ($count -lt 1)
	{
		Write-Host "$server has no back pressure events found."
	}
	else
	{
		$lastevent = $events | Select-Object -First 1

		$now = Get-Date
		$timewritten = $lastevent.TimeWritten
		$ago = "{0:N0}" -f ($now - $timewritten).TotalHours 
		
		switch ($lastevent.EventID)
		{
			"15006" { $BPstate = "Critical (Diskspace)" }
			"15007" { $BPstate = "Critical (Memory)" }
			default { $BPstate = $lastevent.ReplacementStrings[1] }
		}

		Write-Host "$server is $BPstate as of $ago hours ago"

	}
}


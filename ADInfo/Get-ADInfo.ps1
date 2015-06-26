<#
.SYNOPSIS
Get-ADInfo.ps1 - PowerShell script to collect Active Directory information

.DESCRIPTION 
This PowerShell Script collects some basic information about an Active Directory
environment that is useful for verifying the pre-requisites for an Exchange
Server deployment or upgrade.

.OUTPUTS
Results are output to the console and to a HTML file.

.EXAMPLE
.\Get-ADInfo.ps1
Runs the script and generates the output.

.NOTES
Written by: Paul Cunningham

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
V1.00, 26/06/2015 - Initial version
#>

#requires -version 2

$htmlreport = $null
$htmlbody = $null
$htmlfile = "ADInfo.html"
$spacer = "<br />"

Import-Module ActiveDirectory

#---------------------------------------------------------------------
# Collect AD Forest information and convert to HTML fragment
#---------------------------------------------------------------------

$forest = Get-ADForest

$htmlbody += "<h3>Forest Details</h3>"

$forestinfo = New-Object PSObject

Write-Host -ForeGroundColor Yellow "*** Forest: $($forest.RootDomain) ***"
Write-Host ""
$forestinfo  | Add-Member NoteProperty -Name "Forest" -Value $($forest.RootDomain)

Write-Host "Forest Mode: $($forest.ForestMode)"
$forestinfo  | Add-Member NoteProperty -Name "Forest Mode" -Value $($forest.ForestMode)

Write-Host "Schema Master: $($forest.SchemaMaster)"
$forestinfo  | Add-Member NoteProperty -Name "Schema Master" -Value $($forest.SchemaMaster)

Write-Host "Domain Naming Master: $($forest.DomainNamingMaster)"
$forestinfo  | Add-Member NoteProperty -Name "Domain Naming Master" -Value $($forest.DomainNamingMaster)

Write-Host "UPN Suffixes: $($forest.UPNSuffixes)"
$forestinfo  | Add-Member NoteProperty -Name "UPN Suffixes" -Value $($forest.UPNSuffixes)

$htmlbody += $forestinfo | ConvertTo-Html -Fragment
$htmlbody += $spacer

#---------------------------------------------------------------------
# Collect AD Domain information and convert to HTML fragment
#---------------------------------------------------------------------

$htmlbody += "<h3>Domain Details</h3>"

$domains = @($forest | Select -ExpandProperty:Domains)

Foreach ($domain in $domains)
{
    Write-Host ""
    Write-Host -ForeGroundColor Yellow "*** Domain: $domain ***"
    Write-Host ""

    $domaindetails = Get-ADDomain $domain

    $domaininfo = New-Object PSObject
    
    $domaininfo | Add-Member NoteProperty -Name "Name" -Value $domaindetails.Name
    
    Write-Host "NetBIOS Name: $($domaindetails.NetBIOSName)"
    $domaininfo | Add-Member NoteProperty -Name "NetBIOS Name" -Value $domaindetails.NetBIOSName

    Write-Host "Domain Mode: $($domaindetails.DomainMode)"
    $domaininfo | Add-Member NoteProperty -Name "Mode" -Value $($domaindetails.DomainMode)
    
    Write-Host "PDC Emulator: $($domaindetails.PDCEmulator)"
    $domaininfo | Add-Member NoteProperty -Name "PDC Emulator" -Value $($domaindetails.PDCEmulator)
    
    Write-Host "Infrastructure Master: $($domaindetails.InfrastructureMaster)"
    $domaininfo | Add-Member NoteProperty -Name "Infrastructure Master" -Value $($domaindetails.InfrastructureMaster)   
    
    Write-Host "RID Master: $($domaindetails.RIDMaster)"
    $domaininfo | Add-Member NoteProperty -Name "RID Master" -Value $($domaindetails.RIDMaster)

    $htmlbody += $domaininfo | ConvertTo-Html -Fragment
    $htmlbody += $spacer

}

$htmlbody += "<h3>Global Catalog Servers by Site/OS</h3>"

$domaincontrollers = @(Get-ADDomainController -Filter {IsGlobalCatalog -eq $true})

$gcs = @($domaincontrollers | Group-Object -Property:Site,OperatingSystem | Select @{Expression="Name";Label="Site, OS"},Count)

Write-Host ""
Write-Host -ForeGroundColor Yellow "*** Global Catalogs by Site/OS ***"

$gcs = $gcs | Sort "Site, OS"
$gcs | ft -auto

$htmlbody += $gcs | ConvertTo-Html -Fragment


Write-Verbose "Producing HTML report"
    
$reportime = Get-Date

#Common HTML head and styles
$htmlhead="<html>
			<style>
			BODY{font-family: Arial; font-size: 8pt;}
			H1{font-size: 20px;}
			H2{font-size: 18px;}
			H3{font-size: 16px;}
			TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
			TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
			TD{border: 1px solid black; padding: 5px; }
			td.pass{background: #7FFF00;}
			td.warn{background: #FFE600;}
			td.fail{background: #FF0000; color: #ffffff;}
			td.info{background: #85D4FF;}
			</style>
			<body>
			<h1 align=""center"">Active Directory Information</h1>
			<h3 align=""center"">Generated: $reportime</h3>"

$htmltail = "</body>
		</html>"

$htmlreport = $htmlhead + $htmlbody + $htmltail

$htmlreport | Out-File $htmlfile -Encoding Utf8
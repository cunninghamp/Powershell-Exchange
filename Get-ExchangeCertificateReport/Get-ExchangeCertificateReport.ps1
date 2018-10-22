<#
.SYNOPSIS
CertificateReport.ps1 - Exchange Server 2010 SSL Certificate Report Script

.DESCRIPTION 
Generates a report of the SSL certificates installed on Exchange Server 2010 servers

.OUTPUTS
Outputs to a HTML file.

.EXAMPLE
.\CertificateReport.ps1
Reports SSL certificates for Exchange Server 2010 servers and outputs to a HTML file.

.LINK
http://exchangeserverpro.com/powershell-script-ssl-certificate-report

.NOTES
Written by: Paul Cunningham

Find me on:

* My Blog:	https://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	https://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp

Change Log
V1.00, 13/03/2014 - Initial Version
V1.01, 13/03/2014 - Minor bug fix

#>

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$reportfile = "$myDir\CertificateReport.html"

$htmlreport = @()

$exchangeservers = @(Get-ExchangeServer)

foreach ($server in $exchangeservers)
{
    $htmlsegment = @()
    
    $serverdetails = "Server: $($server.Name) ($($server.ServerRole))"
    Write-Host $serverdetails
    
    $certificates = @(Get-ExchangeCertificate -Server $server)

    $certtable = @()

    foreach ($cert in $certificates)
    {
        
        $iis = $null
        $smtp = $null
        $pop = $null
        $imap = $null
        $um = $null       
        
        $subject = ((($cert.Subject -split ",")[0]) -split "=")[1]
                
        if ($($cert.IsSelfSigned))
        {
            $selfsigned = "Yes"
        }
        else
        {
            $selfsigned = "No"
        }

        $issuer = ((($cert.Issuer -split ",")[0]) -split "=")[1]

        #$domains = @($cert | Select -ExpandProperty:CertificateDomains)
        $certdomains = @($cert | Select -ExpandProperty:CertificateDomains)
        if ($($certdomains.Count) -gt 1)
        {
            $domains = $null
            $domains = $certdomains -join ", "
        }
        else
        {
            $domains = $certdomains[0]
        }

        #$services = @($cert | Select -ExpandProperty:Services)
        $services = $cert.ServicesStringForm.ToCharArray()

        if ($services -icontains "W") {$iis = "Yes"}
        if ($services -icontains "S") {$smtp = "Yes"}
        if ($services -icontains "P") {$pop = "Yes"}
        if ($services -icontains "I") {$imap = "Yes"}
        if ($services -icontains "U") {$um = "Yes"}

        $certObj = New-Object PSObject
        $certObj | Add-Member NoteProperty -Name "Subject" -Value $subject
        $certObj | Add-Member NoteProperty -Name "Status" -Value $cert.Status
        $certObj | Add-Member NoteProperty -Name "Expires" -Value $cert.NotAfter.ToShortDateString()
        $certObj | Add-Member NoteProperty -Name "Self Signed" -Value $selfsigned
        $certObj | Add-Member NoteProperty -Name "Issuer" -Value $issuer
        $certObj | Add-Member NoteProperty -Name "SMTP" -Value $smtp
        $certObj | Add-Member NoteProperty -Name "IIS" -Value $iis
        $certObj | Add-Member NoteProperty -Name "POP" -Value $pop
        $certObj | Add-Member NoteProperty -Name "IMAP" -Value $imap
        $certObj | Add-Member NoteProperty -Name "UM" -Value $um
        $certObj | Add-Member NoteProperty -Name "Thumbprint" -Value $cert.Thumbprint
        $certObj | Add-Member NoteProperty -Name "Domains" -Value $domains
        
        $certtable += $certObj
    }

    $htmlcerttable = $certtable | ConvertTo-Html -Fragment

    $htmlserver = "<p>$serverdetails</p>" + $htmlcerttable

    $htmlreport += $htmlserver
}


$htmlhead="<html>
			<style>
			BODY{font-family: Arial; font-size: 10pt;}
			H1{font-size: 16px;}
			H2{font-size: 14px;}
			H3{font-size: 12px;}
			TABLE{border: 1px solid black; border-collapse: collapse; font-size: 10pt;}
			TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
			TD{border: 1px solid black; padding: 5px; }
			td.pass{background: #7FFF00;}
			td.warn{background: #FFE600;}
			td.fail{background: #FF0000; color: #ffffff;}
			td.info{background: #85D4FF;}
			</style>
			<body>
			<h3 align=""center"">Exchange Server 2010 Certificate Report</h3>"

$htmltail = "</body>
			</html>"

$htmlreport = $htmlhead + $htmlreport + $htmltail

$htmlreport | Out-File -Encoding utf8 -FilePath $reportfile



<#
These functions can be added to your PowerShell profile to allow you to
quickly and easily connect to Exchange on-premises.

For more info on how they are used see:
http://exchangeserverpro.com/powershell-function-to-connect-to-exchange-on-premises/
http://exchangeserverpro.com/create-powershell-profile


Written by: Paul Cunningham

Find me on:

* My Blog:	https://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	https://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp

#>

# This function will prompt for authentication and connect to Exchange on-premises
# Use -URL to specify a URL, or set the default URL to your most preferred server.
Function Connect-Exchange {

    param(
        [Parameter( Mandatory=$false)]
        [string]$URL="ex2016srv1.exchangeserverpro.net"
    )
    
    $Credentials = Get-Credential -Message "Enter your Exchange admin credentials"

    $ExOPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$URL/PowerShell/ -Authentication Kerberos -Credential $Credentials

    Import-PSSession $ExOPSession

}

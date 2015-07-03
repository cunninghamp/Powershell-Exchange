#Get-ADInfo.ps1 - PowerShell script to collect Active Directory information

This PowerShell Script collects some basic information about an Active Directory environment that is useful for verifying the pre-requisites for an Exchange
Server deployment or upgrade.

Results are output to the console and to a HTML file.

##Example

```
.\Get-ADInfo.ps1
```

Runs the script and generates the output.

```
[PS] C:\Scripts>.\Get-ADInfo.ps1
*** Forest: exchangeserverpro.net ***

Forest Mode: Windows2003Forest
Schema Master: S1DC1.exchangeserverpro.net
Domain Naming Master: S1DC1.exchangeserverpro.net
UPN Suffixes: esp.local

*** Domain: exchangeserverpro.net ***

NetBIOS Name: ESPNET
Domain Mode: Windows2003Domain
PDC Emulator: S1DC1.exchangeserverpro.net
Infrastructure Master: S1DC1.exchangeserverpro.net
RID Master: S1DC1.exchangeserverpro.net

*** Global Catalogs by Site/OS ***

Site, OS                                       Count
--------                                       -----
DataCenter1, Windows Server 2012 R2 Datacenter     1
```


##Credits
Written by: Paul Cunningham

Find me on:

* My Blog:	http://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	http://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp

For more Exchange Server tips, tricks and news check out Exchange Server Pro.

* Website:	http://exchangeserverpro.com
* Twitter:	http://twitter.com/exchservpro

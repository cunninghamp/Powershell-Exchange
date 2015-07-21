#Get-Exchange2013AVExclusions.ps1 - Generate list of exclusions for antivirus software.

This PowerShell script generates a list of file, folder, process file extension exclusions for configuring antivirus software that will be running on an Exchange 2013 server. The correct exclusions are recommended to prevent antivirus software from interfering with the operation of Exchange Server 2013.

This script is based on information published by Microsoft here:

https://technet.microsoft.com/en-us/library/bb332342(v=exchg.150).aspx

Use this script to generate the exclusion list based on a single server. You can then apply the same exclusions to all servers that have the same configuration. If your antivirus software has a policy-based administration console then that can make the configuration of multiple servers more efficient.

Run the script in the Exchange Management Shell locally on the server you wish to generate the exclusions list for.

Results are output to text files.

##Usage

```
.\Get-Exchange2013AVExclusions.ps1
```

##Credits

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

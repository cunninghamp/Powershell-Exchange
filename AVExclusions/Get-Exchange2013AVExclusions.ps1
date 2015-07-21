<#
.SYNOPSIS
Get-Exchange2013AVExclusions.ps1 - Generate list of exclusions for antivirus software.

.DESCRIPTION 
This PowerShell script generates a list of file, folder, process file extension exclusions
for configuring antivirus software that will be running on an Exchange 2013 server. The 
correct exclusions are recommended to prevent antivirus software from interfering with
the operation of Exchange Server 2013.

This script is based on information published by Microsoft here:
https://technet.microsoft.com/en-us/library/bb332342(v=exchg.150).aspx

Use this script to generate the exclusion list based on a single server. You can then
apply the same exclusions to all servers that have the same configuration. If your antivirus
software has a policy-based administration console then that can make the configuration
of multiple servers more efficient.

Run the script in the Exchange Management Shell locally on the server you wish to generate
the exclusions list for.

.OUTPUTS
Results are output to text files.

.EXAMPLE
.\Get-Exchange2013AVExclusions.ps1

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
V1.00, 21/7/2015 - Initial version
#>

#requires -version 2


#...................................
# Variables
#...................................
#
# Three separate text files are created by the script. Microsoft recommends
# configuring file/folder, process, and file types in case one method of 
# exclusion fails, or in case a path changes later when the server is
# reconfigured.
#

$server = $ENV:ComputerName

#This text file lists the file and folder paths to exclude from antivirus scanning.
$outputfile_paths = "av-exclusions-$($server)-paths.txt"

#This text file lists the processes to exclude from antivirus scanning.
$outputfile_procs = "av-exclusions-$($server)-procs.txt"

#This test file lists the file extensions to exclude from antivirus scanning.
$outputfile_extensions = "av-exclusions-$($server)-extensions.txt"


#...................................
# Script
#...................................

# Start the file/folder paths text file
"### Antivirus exclusion paths for $server ###" | Out-File $outputfile_paths
"" | Out-File $outputfile_paths -Append

### Mailbox Server ###

#Databases

$databases = @(Get-MailboxDatabase -Server $server | Sort Name | Select EdbFilePath,LogFolderPath)

$databases.EdbFilePath.PathName | Out-File $outputfile_paths -Append

$databases.LogFolderPath.PathName | Out-File $outputfile_paths -Append

"$($exinstall)Mailbox\MDBTEMP" | Out-File $outputfile_paths -Append


#Group Metrics

"$($exinstall)GroupMetrics" | Out-File $outputfile_paths -Append


#Log files

$serverlogs = Get-MailboxServer $server | Select DataPath,CalendarRepairLogPath,LogPathForManagedFolders,MigrationLogFilePath,`
                                                    TransportSyncLogFilePath,TransportSyncMailboxHealthLogFilePath

$names = @($serverlogs | Get-Member | Where {$_.membertype -eq "NoteProperty"})
foreach ($name in $names) {$serverlogs.($name.Name).PathName | Out-File $outputfile_paths -Append}


#OAB

"$($exinstall)ClientAccess\OAB" | Out-File $outputfile_paths -Append

#IIS System Files

"$($env:SystemRoot)\System32\Inetsrv" | Out-File $outputfile_paths -Append

#Cluster

"$($env:windir)\Cluster" | Out-File $outputfile_paths -Append

### Transport Services ###

$transportservice = Get-TransportService $server | Select ConnectivityLogPath,MessageTrackingLogPath,IrmLogPath,ActiveUserStatisticsLogPath,`
                                        ServerStatisticsLogPath,ReceiveProtocolLogPath,RoutingTableLogPath,SendProtocolLogPath,`
                                        QueueLogPath,WlmLogPath,AgentLogPath,FlowControlLogPath,ProcessingSchedulerLogPath,`
                                        ResourceLogPath,DnsLogPath,JournalLogPath,TransportMaintenanceLogPath,PipelineTracingPath,`
                                        PickupDirectoryPath,ReplayDirectoryPath,RootDropDirectoryPath

$names = @($transportservice | Get-Member | Where {$_.membertype -eq "NoteProperty"})

foreach ($name in $names) {$transportservice.($name.Name).PathName | Out-File $outputfile_paths -Append}

#Queue and DB files#

$xmlfile = "$($exinstall)Bin\EdgeTransport.exe.config"

if (!(Test-Path $xmlfile)) {Write-Host "File not found"; EXIT}

[xml]$edgetransportconfig = Get-Content $xmlfile

$ETConfigPaths = @()

foreach ($item in $edgetransportconfig.configuration.appSettings.add)
{
    if ($item.key -eq "QueueDatabasePath")
    {
        $QueueDatabasePath = $item.value
        if (!($ETConfigPaths -contains $QueueDatabasePath))
        {
            $ETConfigPaths += $QueueDatabasePath
        }
    }

    if ($item.key -eq "QueueDatabaseLoggingPath")
    {
        $QueueDatabaseLoggingPath = $item.value
        if (!($ETConfigPaths -contains $QueueDatabaseLoggingPath))
        {
            $ETConfigPaths += $QueueDatabaseLoggingPath
        }        
    }

    if ($item.key -eq "IPFilterDatabasePath")
    {
        $IPFilterDatabasePath = $item.value
        if (!($ETConfigPaths -contains $IPFilterDatabasePath))
        {
            $ETConfigPaths += $IPFilterDatabasePath
        }    
    }

    if ($item.key -eq "IPFilterDatabaseLoggingPath")
    {
        $IPFilterDatabaseLoggingPath = $item.value
        if (!($ETConfigPaths -contains $IPFilterDatabaseLoggingPath))
        {
            $ETConfigPaths += $IPFilterDatabaseLoggingPath
        }    
    }
}

$ETConfigPaths | Out-File $outputfile_paths -Append

#Content conversions

"$($env:windir)\temp" | Out-File $outputfile_paths -Append
"$($exinstall)Working\OleConverter" | Out-File $outputfile_paths -Append

#Malware and DLP scanning

"$($exinstall)FIP-FS" | Out-File $outputfile_paths -Append

#Mailbox Transport

$mailboxtransport = @(Get-MailboxTransportService $server | Select *logpath*)

$names = @($mailboxtransport | Get-Member | Where {$_.membertype -eq "NoteProperty"})

foreach ($name in $names) {$mailboxtransport.($name.Name).PathName | Out-File $outputfile_paths -Append}


### Unified Messaging ###

#Grammars

"$($exinstall)UnifiedMessaging\grammars" | Out-File $outputfile_paths -Append

#Voice prompts

"$($exinstall)UnifiedMessaging\Prompts" | Out-File $outputfile_paths -Append

#Voice mail

"$($exinstall)UnifiedMessaging\voicemail" | Out-File $outputfile_paths -Append

#Temp

"$($exinstall)UnifiedMessaging\temp" | Out-File $outputfile_paths -Append


### Setup ###

"$($env:SystemRoot)\Temp\ExchangeSetup" | Out-File $outputfile_paths -Append


### Client Access ###

#IIS

"$($env:SystemDrive)\inetpub\temp\IIS Temporary Compressed Files" | Out-File $outputfile_paths -Append

"$($env:SystemDrive)\inetpub\logs\logfiles\w3svc" | Out-File $outputfile_paths -Append

#POP and IMAP logging

"$($exinstall)Logging\POP3" | Out-File $outputfile_paths -Append

"$($exinstall)Logging\IMAP4" | Out-File $outputfile_paths -Append

#FE Transport

$fetransport = @(Get-FrontEndTransportService $server | Select *logpath*)

$names = @($fetransport | Get-Member | Where {$_.membertype -eq "NoteProperty"})

foreach ($name in $names) {$fetransport.($name.Name).PathName | Out-File $outputfile_paths -Append}


### Process Exclusions ###

#Start the process exclusions text file
"### Antivirus exclusion procs for $server ###" | Out-File $outputfile_procs
"" | Out-File $outputfile_procs -Append

"$($env:SystemRoot)\System32\Dsamain.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\EdgeTransport.exe" | Out-File $outputfile_procs -Append
"$($exinstall)FIP-FS\Bin\fms.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\Search\Ceres\HostController\hostcontrollerservice.exe" | Out-File $outputfile_procs -Append
"$($env:SystemRoot)\System32\inetsrv\inetinfo.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\Microsoft.Exchange.AntispamUpdateSvc.exe" | Out-File $outputfile_procs -Append
"$($exinstall)TransportRoles\agents\Hygiene\Microsoft.Exchange.ContentFilter.Wrapper.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\Microsoft.Exchange.Diagnostics.Service.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\Microsoft.Exchange.Directory.TopologyService.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\Microsoft.Exchange.EdgeCredentialSvc.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\Microsoft.Exchange.EdgeSyncSvc.exe" | Out-File $outputfile_procs -Append
"$($exinstall)FrontEnd\PopImap\Microsoft.Exchange.Imap4.exe" | Out-File $outputfile_procs -Append
"$($exinstall)ClientAccess\PopImap\Microsoft.Exchange.Imap4service.exe" | Out-File $outputfile_procs -Append
"$($exinstall)FrontEnd\PopImap\Microsoft.Exchange.Pop3.exe" | Out-File $outputfile_procs -Append
"$($exinstall)ClientAccess\PopImap\Microsoft.Exchange.Pop3service.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\Microsoft.Exchange.ProtectedServiceHost.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\Microsoft.Exchange.RPCClientAccess.Service.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\Microsoft.Exchange.Search.Service.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\Microsoft.Exchange.Servicehost.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\Microsoft.Exchange.Store.Service.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\Microsoft.Exchange.Store.Worker.exe" | Out-File $outputfile_procs -Append
"$($exinstall)FrontEnd\CallRouter\Microsoft.Exchange.UM.CallRouter.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\MSExchangeDagMgmt.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\MSExchangeDelivery.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\MSExchangeFrontendTransport.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\MSExchangeHMHost.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\MSExchangeHMWorker.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\MSExchangeMailboxAssistants.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\MSExchangeMailboxReplication.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\MSExchangeMigrationWorkflow.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\MSExchangeRepl.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\MSExchangeSubmission.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\MSExchangeTransport.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\MSExchangeTransportLogSearch.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\MSExchangeThrottling.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\Search\Ceres\Runtime\1.0\Noderunner.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\OleConverter.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\Search\Ceres\ParserServer\ParserServer.exe" | Out-File $outputfile_procs -Append
"$($env:SystemRoot)\System32\WindowsPowerShell\v1.0\Powershell.exe" | Out-File $outputfile_procs -Append
"$($exinstall)FIP-FS\Bin\ScanEngineTest.exe" | Out-File $outputfile_procs -Append
"$($exinstall)FIP-FS\Bin\ScanningProcess.exe" | Out-File $outputfile_procs -Append
"$($exinstall)ClientAccess\Owa\Bin\DocumentViewing\TranscodingService.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\UmService.exe" | Out-File $outputfile_procs -Append
"$($exinstall)Bin\UmWorkerProcess.exe" | Out-File $outputfile_procs -Append
"$($exinstall)FIP-FS\Bin\UpdateService.exe" | Out-File $outputfile_procs -Append
"$($env:SystemRoot)\System32\inetsrv\W3wp.exe" | Out-File $outputfile_procs -Append


### File Extension Exclusions ###

# Start the file type exclusions text file
"### Antivirus exclusion extensions for $server ###" | Out-File $outputfile_extensions
"" | Out-File $outputfile_extensions -Append

".config" | Out-File $outputfile_extensions -Append
".dia" | Out-File $outputfile_extensions -Append
".wsb" | Out-File $outputfile_extensions -Append
".chk" | Out-File $outputfile_extensions -Append
".edb" | Out-File $outputfile_extensions -Append
".jrs" | Out-File $outputfile_extensions -Append
".jsl" | Out-File $outputfile_extensions -Append
".log" | Out-File $outputfile_extensions -Append
".que" | Out-File $outputfile_extensions -Append
".lzx" | Out-File $outputfile_extensions -Append
".ci" | Out-File $outputfile_extensions -Append
".dir" | Out-File $outputfile_extensions -Append
".wid" | Out-File $outputfile_extensions -Append
".000" | Out-File $outputfile_extensions -Append
".001" | Out-File $outputfile_extensions -Append
".002" | Out-File $outputfile_extensions -Append
".cfg" | Out-File $outputfile_extensions -Append
".grxml" | Out-File $outputfile_extensions -Append
".dsc" | Out-File $outputfile_extensions -Append
".txt" | Out-File $outputfile_extensions -Append


Write-Host "Done."

#...................................
# Finished
#...................................
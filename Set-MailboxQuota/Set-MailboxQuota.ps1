<#
.SYNOPSIS
Set-MailboxQuota.ps1 - Configure mailbox quotas for Exchange mailboxes

.DESCRIPTION 
This PowerShell script allows you to quickly set mailbox quota values for
one or more Exchange mailboxes. You can use this script to set the mailboxes
to use the database quota defaults, increase by a specified percentage, or
decrease by a specified percentage.

.OUTPUTS
Results are output to console.

.PARAMETER Mailbox
The name of the mailbox you want to modify.

.PARAMETER UseDatabaseDefaults
This switch will set the specified mailboxes to use the default quotas
configured on the mailbox database on which they are hosted.

.PARAMETER IncreaseByPercentage
This parameter allows you to specify the percentage by which existing
quota values on a mailbox should be increased. If the mailbox is currently
using the default database quota levels the increase will be based off
those values.

If any of the three quota settings are set to "Unlimited" then this
parameter will not make any changes to the mailbox.

.PARAMETER DecreaseByPercentage
This parameter allows you to specify the percentage by which existing
quota values on a mailbox should be decreased. If the mailbox is currently
using the default database quota levels the increase will be based off
those values.

If any of the three quota settings are set to "Unlimited" then this
parameter will not make any changes to the mailbox.

.EXAMPLE
.\Set-MailboxQuota.ps1 -Mailbox Adam.Wally -UseDatabaseDefaults
The mailbox user Adam.Wally will be configured to use the quota
values configured on the mailbox database on which they are hosted.

.EXAMPLE
.\Set-MailboxQuota.ps1 -Mailbox Adam.Wally -IncreaseByPercentage 5
The mailbox user Adam.Wally will have their mailbox quota values
increased by 5 percent.

.EXAMPLE
Get-Mailbox -Database DB04 | .\Set-MailboxQuota.ps1 -UseDatabaseDefaults
All mailboxes hosted on database DB04 will be configured to use the
default database quota values.

.NOTES
Written by: Paul Cunningham

Find me on:

* My Blog:	https://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	https://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp

Change Log
V1.00, 23/10/2015 - Initial version
#>


[CmdletBinding()]
param (
	
	[Parameter( ParameterSetName='Increase' )]
	[int]$IncreaseByPercentage,

	[Parameter( ParameterSetName='Decrease' )]
	[int]$DecreaseByPercentage,

	[Parameter( ParameterSetName='UseDefaults' )]
	[switch]$UseDatabaseDefaults,

    [Parameter( ValueFromPipeline=$True, Mandatory=$true)]
    [string[]]$Mailbox

	)


#...................................
# Script
#...................................

Begin {

    #...................................
    # Functions
    #...................................

    Function Convert-QuotaStringToKB() {

        Param([string]$CurrentQuota)

        [string]$CurrentQuota = ($CurrentQuota.Split("("))[1]
        [string]$CurrentQuota = ($CurrentQuota.Split(" bytes)"))[0]
        $CurrentQuota = $CurrentQuota.Replace(",","")
        [int]$CurrentQuotaInKB = "{0:F0}" -f ($CurrentQuota/1024)

        return $CurrentQuotaInKB
    }

}

Process {

    foreach ($i in $mailbox)
    {

        $UnlimitedWarning = $null
        $UnlimitedSend = $null
        $UnlimitedSendReceive = $null

        #Verify mailbox exists
        try
        {
            $mbx = Get-Mailbox $i -ErrorAction STOP

            #Current mailbox quota details
            Write-Host "----------------------------------------" -ForegroundColor White
            Write-Host "Mailbox: $mbx" -ForegroundColor White
            Write-Host "----------------------------------------" -ForegroundColor White
            Write-Host ""
            Write-Host "Uses Database Defaults: $($mbx.UseDatabaseQuotaDefaults)"
            Write-Host ""

            if ($mbx.UseDatabaseQuotaDefaults)
            {
                $quotas = @(Get-MailboxDatabase $mbx.Database | Select IssueWarningQuota,ProhibitSendQuota,ProhibitSendReceiveQuota)

                Write-Host "Warning Quota: $($quotas.IssueWarningQuota)"
                Write-Host "Prohibit Send Quota: $($quotas.ProhibitSendQuota)"
                Write-Host "Prohibit Send/Receive Quota: $($quotas.ProhibitSendReceiveQuota)"
                Write-Host ""

                [string]$strIssueWarningQuota = $quotas.IssueWarningQuota.ToString()
                [string]$strProhibitSendQuota = $quotas.ProhibitSendQuota.ToString()
                [string]$strProhibitSendReceiveQuota = $quotas.ProhibitSendReceiveQuota.ToString()

                if ($strIssueWarningQuota -ne "Unlimited")
                {
                    $CurrentIssueWarningQuotaInKB = Convert-QuotaStringToKB $strIssueWarningQuota
                }
                else
                {
                    $UnlimitedWarning = $true
                }

                if ($strProhibitSendQuota -ne "Unlimited")
                {
                    $CurrentProhibitSendQuotaInKB = Convert-QuotaStringToKB $strProhibitSendQuota
                }
                else
                {
                    $UnlimitedSend = $true
                }

                if ($strProhibitSendReceiveQuota -ne "Unlimited")
                {
                    $CurrentProhibitSendReceiveQuotaInKB = Convert-QuotaStringToKB $strProhibitSendReceiveQuota
                }
                else
                {
                    $UnlimitedSendReceive = $true
                }

            }
            else
            {
                Write-Host "Warning Quota: $($mbx.IssueWarningQuota)"
                Write-Host "Prohibit Send Quota: $($mbx.ProhibitSendQuota)"
                Write-Host "Prohibit Send/Receive Quota: $($mbx.ProhibitSendReceiveQuota)"
                Write-Host ""

                if ($($mbx.IssueWarningQuota) -ne "Unlimited")
                {
                    $CurrentIssueWarningQuotaInKB = $mbx.IssueWarningQuota.Value.ToKB()
                }
                else
                {
                    $UnlimitedWarning = $true
                }
                
                if ($($mbx.ProbibitSendQuota) -ne "Unlimited")
                {
                    $CurrentProhibitSendQuotaInKB = $mbx.ProhibitSendQuota.Value.ToKB()
                }
                else
                {
                    $UnlimitedSend = $true
                }
                
                if ($($mbx.ProhibitSendReceiveQuota) -ne "Unlimited")
                {
                    $CurrentProhibitSendReceiveQuotaInKB = $mbx.ProhibitSendReceiveQuota.Value.ToKB()
                }
                else
                {
                    $UnlimitedSendReceive = $true
                }
            }


            #Process UseDatabaseDefaults scenario
            if ($UseDatabaseDefaults)
            {
                if ($mbx.UseDatabaseQuotaDefaults)
                {
                    Write-Host "$mbx already uses database quota defaults, no change required." -ForegroundColor White
                }
                else
                {
                    try
                    {
                        Set-Mailbox $mbx -UseDatabaseQuotaDefaults:$true -ErrorAction STOP
                        Write-Host "$mbx has been configured to use database quota defaults." -ForegroundColor White
                    }
                    catch
                    {
                        Write-Warning $_.Exception.Message
                    }
                }
            }

            
            #Process Increase Scenario
            if ($IncreaseByPercentage)
            {

                Write-Host "Calculating new quotas" -ForegroundColor White

                if ($UnlimitedWarning -eq $true)
                {
                    Write-Host "Warning quota is unlimited and mailbox will not be modified."
                }
                else
                {
                    Write-Host "Current warning quota: $CurrentIssueWarningQuotaInKB KB"
                    $NewIssueWarningQuotaInKB = $CurrentIssueWarningQuotaInKB + ($CurrentIssueWarningQuotaInKB * $IncreaseByPercentage)/100
                    Write-Host "New warning quota: $NewIssueWarningQuotaInKB KB"
                }

                if ($UnlimitedSend -eq $true)
                {
                    Write-Host "Send quota is unlimited and mailbox will not be modified."
                }
                else
                {
                    Write-Host "Current send quota: $CurrentProhibitSendQuotaInKB KB"
                    $NewProhibitSendQuotaInKB = $CurrentProhibitSendQuotaInKB + ($CurrentProhibitSendQuotaInKB * $IncreaseByPercentage)/100
                    Write-Host "New send quota: $NewProhibitSendQuotaInKB KB"
                }

                if ($UnlimitedSendReceive -eq $true)
                {
                    Write-Host "Send/Receive quota is unlimited and mailbox will not be modified."
                }
                else
                {
                    Write-Host "Current send/rec quota: $CurrentProhibitSendReceiveQuotaInKB KB"
                    $NewProhibitSendReceiveQuotaInKB = $CurrentProhibitSendReceiveQuotaInKB + ($CurrentProhibitSendReceiveQuotaInKB * $IncreaseByPercentage)/100
                    Write-Host "New send/rec quota: $NewProhibitSendReceiveQuotaInKB KB"
                }
            }


            #Process Decrease Scenario
            if ($DecreaseByPercentage)
            {
                Write-Host "Calculating new quotas in KB" -ForegroundColor White

                if ($UnlimitedWarning -eq $true)
                {
                    Write-Host "Warning quota is unlimited and mailbox will not be modified."
                }
                else
                {
                    Write-Host "Current warning quota: $CurrentIssueWarningQuotaInKB KB"
                    $NewIssueWarningQuotaInKB = $CurrentIssueWarningQuotaInKB - ($CurrentIssueWarningQuotaInKB * $DecreaseByPercentage)/100
                    Write-Host "New warning quota: $NewIssueWarningQuotaInKB KB"
                }

                if ($UnlimitedSend -eq $true)
                {
                    Write-Host "Send quota is unlimited and mailbox will not be modified."
                }
                else
                {
                    Write-Host "Current send quota: $CurrentProhibitSendQuotaInKB KB"
                    $NewProhibitSendQuotaInKB = $CurrentProhibitSendQuotaInKB - ($CurrentProhibitSendQuotaInKB * $DecreaseByPercentage)/100
                    Write-Host "New send quota: $NewProhibitSendQuotaInKB KB"
                }

                if ($UnlimitedSendReceive -eq $true)
                {
                    Write-Host "Send/Receive quota is unlimited and mailbox will not be modified."
                }
                else
                {
                    Write-Host "Current send/rec quota: $CurrentProhibitSendReceiveQuotaInKB KB"
                    $NewProhibitSendReceiveQuotaInKB = $CurrentProhibitSendReceiveQuotaInKB - ($CurrentProhibitSendReceiveQuotaInKB * $DecreaseByPercentage)/100
                    Write-Host "New send/rec quota: $NewProhibitSendReceiveQuotaInKB KB"
                }
            }


            #If either Increase or Decrease scenario, set the new quota values
            if (!($UseDatabaseDefaults))
            {
                if ($UnlimitedWarning -eq $true -or $UnlimitedSend -eq $true -or $UnlimitedSendReceive -eq $true)
                {
                    Write-Host "One or more quota values are set to unlimited, mailbox will not be modified." -ForegroundColor White
                }
                else
                {
                    try
                    {
                        Write-Host "Setting new quotas" -ForegroundColor White
                        if ($mbx.UseDatabaseQuotaDefaults)                
                        {
                            Set-Mailbox $mbx -UseDatabaseQuotaDefaults $false -IssueWarningQuota "$($NewIssueWarningQuotaInKB)KB" -ProhibitSendQuota "$($NewProhibitSendQuotaInKB)KB" -ProhibitSendReceiveQuota "$($NewProhibitSendReceiveQuotaInKB)KB" -ErrorAction STOP
                        }
                        else
                        {
                            Set-Mailbox $mbx -IssueWarningQuota "$($NewIssueWarningQuotaInKB)KB" -ProhibitSendQuota "$($NewProhibitSendQuotaInKB)KB" -ProhibitSendReceiveQuota "$($NewProhibitSendReceiveQuotaInKB)KB" -ErrorAction STOP
                        }
                        if ($IncreaseByPercentage)
                        {
                            Write-Host "Quotas increased by $IncreaseByPercentage percent"
                        }
                        if ($DecreaseByPercentage)
                        {
                            Write-Host "Quotas decreased by $DecreaseByPercentage percent"
                        }
                    }
                    catch
                    {
                        Write-Warning $_.Exception.Message
                    }
                }
            }
            Write-Host "----------------------------------------" -ForegroundColor White
            Write-Host ""
        }
        catch
        {
            #Mailbox was not found
            Write-Warning $_.Exception.Message
        }

    }

}

End {

}

#...................................
# Finished
#...................................

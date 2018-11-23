# **********************************************************************************
#
# Script Name: Update-ALs.ps1
# Version: 1.0
# Author: Davde M
# Date Created: 10/16/2012
# _______________________________________
#
# MODIFICATIONS:
# Date Modified: 7-25-13
# Modified By: Dave M
# Reason for modification: Upgrade Address Lists to Exchange 2010
# What was modified: Added Get-AddressList | Set-AddressList -ForceUpgrade
#
# Description: Updates all GALs and ALs.
#
# Usage:
# ./Update-ALs.ps1
#
# **********************************************************************************

<#
	.SYNOPSIS
		Updates all ALs and GALs

	.DESCRIPTION
		Gets all Address Lists and Global Address Lists and forces the system to update them.

	.EXAMPLE
		./Script_Name_Here.ps1

		Description
		===========
		Updates all ALs and GALs
		
	.NOTES
		Requires Janus and Exchange Module Loaded.

#>

# Functions and Filters


# Main Script
# The Begin section executes once regardless of how many objects are passed through the pipeline
begin
{
$error.clear()
if ($global:JanusPSModule -ne $true) {Import-Module Janus -erroraction stop}
Show-JStatus

if ($global:E2010SnapIn -ne $true)
	{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
	$global:E2010SnapIn=$?
	$global:ExchangeSnapIn=$global:E2010SnapIn
	}

# Initialization
$syslogobject=New-JSyslogger -dest_host "p-ucslog02.janus.cap"
[string]$LOG=""
[string]$Report="Address List Update Report:`n"
}

# The process section runs once for each object in the pipeline
process
{
$LOG="Get-AddressList `| Update-AddressList"
Get-AddressList | Update-AddressList
$success=$?
$LOG=$LOG + " - Successful: $success"
$syslogobject.send($log)

$report=$report + "`n" + $log + "`n"

$LOG="Get-AddressList | Set-AddressList -ForceUpgrade"
Get-AddressList | Set-AddressList -ForceUpgrade
$success=$?
$LOG=$LOG + " - Successful: $success"
$syslogobject.send($log)

$report=$report + "`n" + $log + "`n"

$LOG="Get-GlobalAddressList `| Update-GlobalAddressList"
Get-GlobalAddressList | Update-GlobalAddressList
$success=$?
$LOG=$LOG + " - Successful: $success"
$syslogobject.send($log)

$report=$report + "`n" + $log + "`n"

}

# The End section executes once regardless of how many objects are passed through the pipeline
end
{
$log=$error | out-string
$syslogobject.send($log)

$report=$report + "`n`nError Log:`n" + $log + "`n"

Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Address List Update Report" -Body $report

remove-variable success
remove-variable report
remove-variable log
remove-variable syslogobject
}

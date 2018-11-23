# **********************************************************************************
#
# Script Name: XM03-CSV-Mover.ps1
# Version: 1.0
# Author: Dave M
# Date Created: 8-9-12
# _______________________________________
#
# MODIFICATIONS:
# Date Modified: N/A
# Modified By: N/A
# Reason for modification: N/A
# What was modified: N/A
#
# Description: Moves users based on the contents of \\janus.cap\groups$\IT\ITSM and Exchange\ExchangeMigrationDailySchedule.csv
#
# Usage:
# ./XM03-CSV-Mover.ps1
#
# **********************************************************************************

<#
	.SYNOPSIS
		Automates User Moves

	.DESCRIPTION
		Moves users based on the contents of \\janus.cap\groups$\IT\ITSM and Exchange\ExchangeMigrationDailySchedule.csv

	.EXAMPLE
		./XM03-CSV-Mover.ps1
		
	.NOTES
		Requires Exchange 2020 SnapIn and Janus PS Module.

#>

# Functions and Filters


# Main Script
# The Begin section executes once regardless of how many objects are passed through the pipeline
begin
{
[DateTime]$date=Get-Date
[string]$datecode=get-date $date -DisplayHint Date -Format yyyyMMdd
Start-Transcript \\p-ucadm01.janus.cap\d$\Scripts\Logs\XM03-CSV-Mover-$datecode.LOG

$test=Show-KVSTemp
$global:JanusPSModule=$?
if ($global:JanusPSModule -ne $true) {Import-Module Janus -erroraction stop}

$test=Get-DatabaseAvailabilityGroup
$global:E2010SnapIn=$?

# Load Exchange 2010 snapin if it is not already
if ($global:E2010SnapIn -ne $true)
	{
	. 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
	$test=Get-DatabaseAvailabilityGroup
	$success=$?
	If ($success -ne $true)
		{
		write-error ("Unable to load Exchange 2010 Module, exiting...") -erroraction stop
		} else
		{
		$global:E2010SnapIn = $true
		}
	}

$cdatetime=$null
$Schedule=$null
$target=$null
$XM3_1=$null
$XM3_2=$null
$XM3_3=$null
$XM3_4=$null
$XM3_5=$null
$XM3_6=$null
$XM3_7=$null
$XM3_8=$null
$csv=$null
$SUsers=$null

$RTime=get-date -Format MMddyyHHmm
$sname=gc env:computername
do {
	$csv=Import-Csv "\\janus.cap\groups$\IT\ITSM and Exchange\ExchangeMigrationDailySchedule.csv"
	$success=$?
} while ($success -ne $true)

if ($csv.count -eq 0) {write-warning "No mailboxes in CSV file at \\janus.cap\groups`$\IT\ITSM and Exchange\ExchangeMigrationDailySchedule.csv, exiting..." -warningaction stop}

$SUsers=$CSV | where { [system.datetime]$_.Day_Time -gt (get-date).addminutes(-15) -and [system.datetime]$_.Day_Time -lt (get-date) }

}

# The process section runs once for each object in the pipeline
process
{
# This will gather the mailbox info for users destined to mobe to USXM03-DB1 and p[ut them into an array
$XM3_1=$SUsers | where { $_.Target_Store -like "USXM03-DB1" } | foreach { $_.Mailbox } | get-mailbox
$XM3_2=$SUsers | where { $_.Target_Store -like "USXM03-DB2" } | foreach { $_.Mailbox } | get-mailbox
$XM3_3=$SUsers | where { $_.Target_Store -like "USXM03-DB3" } | foreach { $_.Mailbox } | get-mailbox
$XM3_4=$SUsers | where { $_.Target_Store -like "USXM03-DB4" } | foreach { $_.Mailbox } | get-mailbox
$XM3_5=$SUsers | where { $_.Target_Store -like "USXM03-DB5" } | foreach { $_.Mailbox } | get-mailbox
$XM3_6=$SUsers | where { $_.Target_Store -like "USXM03-DB6" } | foreach { $_.Mailbox } | get-mailbox
$XM3_7=$SUsers | where { $_.Target_Store -like "USXM03-DB7" } | foreach { $_.Mailbox } | get-mailbox
$XM3_8=$SUsers | where { $_.Target_Store -like "USXM03-DB8" } | foreach { $_.Mailbox } | get-mailbox

# don't process if the array is null
if ($XM3_1 -ne "" -and $XM3_1 -ne $null -and $XM3_1 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM03-DB1"

# Logging
    $XM3_1 >> "D:\Scripts\Logs\USXM01-DB1.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM3_1 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM3_1 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited -DomainController "$(Get-JDCs -GC)"  -MRSServer $((Get-TransportServer | Where {$_.Name -like "*P-UCUS*"} | Get-Random).Name) -confirm:$false
    $movesuccess=$?

# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM3_1 just failed. User `"$XM3_1`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM3_1 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }

# don't process if the array is null
if ($XM3_2 -ne "" -and $XM3_2 -ne $null -and $XM3_2 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM03-DB2"

# Logging
    $XM3_2 >> "D:\Scripts\Logs\USXM01-DB1.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM3_2 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM3_2 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited -DomainController "$(Get-JDCs -GC)"  -MRSServer $((Get-TransportServer | Where {$_.Name -like "*P-UCUS*"} | Get-Random).Name) -confirm:$false
    $movesuccess=$?

# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM3_2 just failed. User `"$XM3_2`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM3_2 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }

# don't process if the array is null
if ($XM3_3 -ne "" -and $XM3_3 -ne $null -and $XM3_3 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM03-DB3"

# Logging
    $XM3_3 >> "D:\Scripts\Logs\USXM01-DB1.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM3_3 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM3_3 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited -DomainController "$(Get-JDCs -GC)"  -MRSServer $((Get-TransportServer | Where {$_.Name -like "*P-UCUS*"} | Get-Random).Name) -confirm:$false
    $movesuccess=$?

# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM3_3 just failed. User `"$XM3_3`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM3_3 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }

# don't process if the array is null
if ($XM3_4 -ne "" -and $XM3_4 -ne $null -and $XM3_4 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM03-DB4"

# Logging
    $XM3_4 >> "D:\Scripts\Logs\USXM01-DB1.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM3_4 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM3_4 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited -DomainController "$(Get-JDCs -GC)"  -MRSServer $((Get-TransportServer | Where {$_.Name -like "*P-UCUS*"} | Get-Random).Name) -confirm:$false
    $movesuccess=$?

# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM3_4 just failed. User `"$XM3_4`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM3_4 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }

# don't process if the array is null
if ($XM3_5 -ne "" -and $XM3_5 -ne $null -and $XM3_5 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM03-DB5"

# Logging
    $XM3_5 >> "D:\Scripts\Logs\USXM01-DB1.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM3_5 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM3_5 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited -DomainController "$(Get-JDCs -GC)"  -MRSServer $((Get-TransportServer | Where {$_.Name -like "*P-UCUS*"} | Get-Random).Name) -confirm:$false
    $movesuccess=$?

# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM3_5 just failed. User `"$XM3_5`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM3_5 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }

# don't process if the array is null
if ($XM3_6 -ne "" -and $XM3_6 -ne $null -and $XM3_6 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM03-DB6"

# Logging
    $XM3_6 >> "D:\Scripts\Logs\USXM01-DB1.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM3_6 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM3_6 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited -DomainController "$(Get-JDCs -GC)"  -MRSServer $((Get-TransportServer | Where {$_.Name -like "*P-UCUS*"} | Get-Random).Name) -confirm:$false
    $movesuccess=$?

# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM3_6 just failed. User `"$XM3_6`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM3_6 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }

# don't process if the array is null
if ($XM3_7 -ne "" -and $XM3_7 -ne $null -and $XM3_7 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM03-DB7"

# Logging
    $XM3_7 >> "D:\Scripts\Logs\USXM01-DB1.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM3_7 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM3_7 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited -DomainController "$(Get-JDCs -GC)"  -MRSServer $((Get-TransportServer | Where {$_.Name -like "*P-UCUS*"} | Get-Random).Name) -confirm:$false
    $movesuccess=$?

# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM3_7 just failed. User `"$XM3_7`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM3_7 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }

# don't process if the array is null
if ($XM3_8 -ne "" -and $XM3_8 -ne $null -and $XM3_8 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM03-DB8"

# Logging
    $XM3_8 >> "D:\Scripts\Logs\USXM01-DB1.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM3_8 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM3_8 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited -DomainController "$(Get-JDCs -GC)"  -MRSServer $((Get-TransportServer | Where {$_.Name -like "*P-UCUS*"} | Get-Random).Name) -confirm:$false
    $movesuccess=$?

# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM3_8 just failed. User `"$XM3_8`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM3_8 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }
}

# The End section executes once regardless of how many objects are passed through the pipeline
end
{

remove-variable cdatetime
remove-variable Schedule
remove-variable target
remove-variable XM3_1
remove-variable XM3_2
remove-variable XM3_3
remove-variable XM3_4
remove-variable XM3_5
remove-variable XM3_6
remove-variable XM3_7
remove-variable XM3_8
remove-variable csv
remove-variable SUsers
remove-variable RTime
remove-variable sname
remove-variable date
remove-variable datecode
remove-variable success
remove-variable test
remove-variable store

$error | out-string
stop-transcript
}

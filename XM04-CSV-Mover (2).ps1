# **********************************************************************************
#
# Script Name: XM04-CSV-Mover.ps1
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
# ./XM04-CSV-Mover.ps1
#
# **********************************************************************************

<#
	.SYNOPSIS
		Automates User Moves

	.DESCRIPTION
		Moves users based on the contents of \\janus.cap\groups$\IT\ITSM and Exchange\ExchangeMigrationDailySchedule.csv

	.EXAMPLE
		./XM04-CSV-Mover.ps1
		
	.NOTES
		Requires Exchange 2020 SnapIn and Janus PS Module.

#>

# Functions and Filters


# Main Script
# The Begin section executes once regardless of how many objects are passed through the pipeline
begin
{
$test=Show-KVSTemp
$global:JanusPSModule=$?
if ($global:JanusPSModule -ne $true) {Import-Module Janus -erroraction stop}

# Unload X2k7 snapins
Get-PSSnapin | where {$_.name -like "*exchange*"} | Remove-PSSnapin

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
$XM4_1=$null
$XM4_2=$null
$XM4_3=$null
$XM4_4=$null
$XM4_5=$null
$XM4_6=$null
$XM4_7=$null
$XM4_8=$null
$csv=$null
$SUsers=$null

$RTime=get-date -Format MMddyyHHmm
$sname=gc env:computername
do {
	$csv=Import-Csv "\\janus.cap\groups$\IT\ITSM and Exchange\ExchangeMigrationDailySchedule.csv"
	$success=$?
} while ($success -ne $true)
$SUsers=$CSV | where { [system.datetime]$_.Day_Time -gt (get-date).addminutes(-15) -and [system.datetime]$_.Day_Time -lt (get-date) }

}

# The process section runs once for each object in the pipeline
process
{
# This will gather the mailbox info for users destined to mobe to USXM04-DB1 and p[ut them into an array
$XM4_1=$SUsers | where { $_.Target_Store -like "USXM04-DB1" } | foreach { $_.Mailbox } | get-mailbox
$XM4_2=$SUsers | where { $_.Target_Store -like "USXM04-DB2" } | foreach { $_.Mailbox } | get-mailbox
$XM4_3=$SUsers | where { $_.Target_Store -like "USXM04-DB3" } | foreach { $_.Mailbox } | get-mailbox
$XM4_4=$SUsers | where { $_.Target_Store -like "USXM04-DB4" } | foreach { $_.Mailbox } | get-mailbox
$XM4_5=$SUsers | where { $_.Target_Store -like "USXM04-DB5" } | foreach { $_.Mailbox } | get-mailbox
$XM4_6=$SUsers | where { $_.Target_Store -like "USXM04-DB6" } | foreach { $_.Mailbox } | get-mailbox
$XM4_7=$SUsers | where { $_.Target_Store -like "USXM04-DB7" } | foreach { $_.Mailbox } | get-mailbox
$XM4_8=$SUsers | where { $_.Target_Store -like "USXM04-DB8" } | foreach { $_.Mailbox } | get-mailbox

# don't process if the array is null
if ($XM4_1 -ne "" -and $XM4_1 -ne $null -and $XM4_1 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM04-DB1"

# Logging
    $XM4_1 >> "D:\Scripts\Logs\USXM01-DB1.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM4_1 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM4_1 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited -DomainController p-jcdcd07.janus.cap -confirm:$false
    $movesuccess=$?

# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM4_1 just failed. User `"$XM4_1`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM4_1 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }

# don't process if the array is null
if ($XM4_2 -ne "" -and $XM4_2 -ne $null -and $XM4_2 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM04-DB2"

# Logging
    $XM4_2 >> "D:\Scripts\Logs\USXM01-DB1.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM4_2 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM4_2 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited -DomainController p-jcdcd07.janus.cap -confirm:$false
    $movesuccess=$?

# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM4_2 just failed. User `"$XM4_2`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM4_2 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }

# don't process if the array is null
if ($XM4_3 -ne "" -and $XM4_3 -ne $null -and $XM4_3 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM04-DB3"

# Logging
    $XM4_3 >> "D:\Scripts\Logs\USXM01-DB1.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM4_3 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM4_3 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited -DomainController p-jcdcd07.janus.cap -confirm:$false
    $movesuccess=$?

# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM4_3 just failed. User `"$XM4_3`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM4_3 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }

# don't process if the array is null
if ($XM4_4 -ne "" -and $XM4_4 -ne $null -and $XM4_4 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM04-DB4"

# Logging
    $XM4_4 >> "D:\Scripts\Logs\USXM01-DB1.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM4_4 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM4_4 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited -DomainController p-jcdcd07.janus.cap -confirm:$false
    $movesuccess=$?

# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM4_4 just failed. User `"$XM4_4`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM4_4 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }

# don't process if the array is null
if ($XM4_5 -ne "" -and $XM4_5 -ne $null -and $XM4_5 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM04-DB5"

# Logging
    $XM4_5 >> "D:\Scripts\Logs\USXM01-DB1.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM4_5 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM4_5 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited -DomainController p-jcdcd07.janus.cap -confirm:$false
    $movesuccess=$?

# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM4_5 just failed. User `"$XM4_5`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM4_5 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }

# don't process if the array is null
if ($XM4_6 -ne "" -and $XM4_6 -ne $null -and $XM4_6 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM04-DB6"

# Logging
    $XM4_6 >> "D:\Scripts\Logs\USXM01-DB1.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM4_6 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM4_6 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited -DomainController p-jcdcd07.janus.cap -confirm:$false
    $movesuccess=$?

# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM4_6 just failed. User `"$XM4_6`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM4_6 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }

# don't process if the array is null
if ($XM4_7 -ne "" -and $XM4_7 -ne $null -and $XM4_7 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM04-DB7"

# Logging
    $XM4_7 >> "D:\Scripts\Logs\USXM01-DB1.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM4_7 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM4_7 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited -DomainController p-jcdcd07.janus.cap -confirm:$false
    $movesuccess=$?

# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM4_7 just failed. User `"$XM4_7`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM4_7 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }

# don't process if the array is null
if ($XM4_8 -ne "" -and $XM4_8 -ne $null -and $XM4_8 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM04-DB8"

# Logging
    $XM4_8 >> "D:\Scripts\Logs\USXM01-DB1.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM4_8 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM4_8 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited -DomainController p-jcdcd07.janus.cap -confirm:$false
    $movesuccess=$?

# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM4_8 just failed. User `"$XM4_8`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM4_8 to $store is starting using PID $pid on $sname."
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
remove-variable XM4_1
remove-variable XM4_2
remove-variable XM4_3
remove-variable XM4_4
remove-variable XM4_5
remove-variable XM4_6
remove-variable XM4_7
remove-variable XM4_8
remove-variable csv
remove-variable SUsers
remove-variable RTime
remove-variable sname
}

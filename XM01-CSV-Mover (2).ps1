# **********************************************************************************
#
# Script Name: XM01-CSV-Mover.ps1
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
# ./XM01-CSV-Mover.ps1
#
# **********************************************************************************

<#
	.SYNOPSIS
		Automates User Moves

	.DESCRIPTION
		Moves users based on the contents of \\janus.cap\groups$\IT\ITSM and Exchange\ExchangeMigrationDailySchedule.csv

	.EXAMPLE
		./XM01-CSV-Mover.ps1
		
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
		} else # If ($success -ne $true)
		{
		$global:E2010SnapIn = $true
		} # else
	} # if ($global:E2010SnapIn -ne $true)

$cdatetime=$null
$Schedule=$null
$target=$null
$XM1_1=$null
$XM1_2=$null
$XM1_3=$null
$XM1_4=$null
$XM1_5=$null
$XM1_6=$null
$XM1_7=$null
$XM1_8=$null
$csv=$null
$SUsers=$null
[array]$XMRs=@()

$RTime=get-date -Format MMddyyHHmm
$sname=gc env:computername
do {
	$csv=Import-Csv "\\janus.cap\groups$\IT\ITSM and Exchange\ExchangeMigrationDailySchedule.csv"
	$success=$?
} while ($success -ne $true)
$SUsers=$CSV | where { [system.datetime]$_.Day_Time -gt (get-date).addminutes(-15) -and [system.datetime]$_.Day_Time -lt (get-date) }
$BADXMRs=Get-MoveRequest | where { ($_.Status -like "*warning*") -or($_.Status -like "*error*") -or($_.Status -like "*fail*") }
$XMRs=Get-MoveRequest | where { $_.Status -eq "Completed" }

} # begin

# The process section runs once for each object in the pipeline
process
{
# MoveRequest Cleanup
# Does not execute if there are no complete move requests
if ($XMRs -ne $null)
	{
# Do this to each move request
	Foreach ($XMR IN $XMRs)
		{
		$XMR | fl >> "D:\Scripts\Logs\$($XMR.name).log"
		Get-MoveRequestStatistics -id $($XMR.name) -IncludeReport | fl >> "D:\Scripts\Logs\$($XMR.name).log"
# Get their DN so we can look up the moved user
		$DN=$XMR.DistinguishedName
# Look up the moved user so we can get tehir AdminCount
		$ADE=get-jadentry -id $DN -exact -pso -properties AdminCount,SAMAccountNAme,displayname
# Get their AD ID so we can use it later
		$ADID=$ADE.SAMAccountName
# Get their AdminCount
		$Admin=$ADE.AdminCount
# AdminCount 0 is a normal user and their mailbox move will not affect their ActiveSync settings
		If ($admin -gt 0)
			{
# Check the Inheritanc on their security if they are/were an admin
			Set-JInheritance -ID $ADID
			"Was AD Inheritance set successfully: $?" >> "D:\Scripts\Logs\$($XMR.name).log"
			} else # If ($admin -gt 0)
			{
			"AD Inheritance was not modified." >> "D:\Scripts\Logs\$($XMR.name).log"
			} # else
# Remove the old move request
		Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move success" -Body "The move of mailbox $($ADE.displayname) completed successfully."
		$XMR | Remove-MoveRequest -confirm:$false
		if ($? -eq $false) {$error | out-string > d:\scripts\logs\Remove-moverequest.log}
		} # Foreach ($XMR IN $XMRs)
	} # if ($XMRs -ne $null)
# Bad Move Request Processing
if ($BADXMRs -ne $null)
	{
	foreach ($XMR IN $BADXMRs)
		{
	        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Critical Alert: Mailbox Move Failure" -Body "The move of mailbox $($XMR.DisplayName) just failed. Open D:\Scripts\Logs\$($XMR.name).log to see a report of issue."
		$XMR | fl >> "D:\Scripts\Logs\$($XMR.name).log"
		Get-MoveRequestStatistics -id $($XMR.name) -IncludeReport | fl >> "D:\Scripts\Logs\$($XMR.name).log"
		$XMR | Remove-MoveRequest -confirm:$false
		} # foreach ($XMR IN $BADXMRs)
	} # if ($BADXMRs -ne $null)
# This will gather the mailbox info for users destined to mobe to USXM01-DB1 and p[ut them into an array
$XM1_1=$SUsers | where { $_.Target_Store -like "USXM01-DB1" } | foreach { $_.Mailbox } | get-mailbox
$XM1_2=$SUsers | where { $_.Target_Store -like "USXM01-DB2" } | foreach { $_.Mailbox } | get-mailbox
$XM1_3=$SUsers | where { $_.Target_Store -like "USXM01-DB3" } | foreach { $_.Mailbox } | get-mailbox
$XM1_4=$SUsers | where { $_.Target_Store -like "USXM01-DB4" } | foreach { $_.Mailbox } | get-mailbox
$XM1_5=$SUsers | where { $_.Target_Store -like "USXM01-DB5" } | foreach { $_.Mailbox } | get-mailbox
$XM1_6=$SUsers | where { $_.Target_Store -like "USXM01-DB6" } | foreach { $_.Mailbox } | get-mailbox
$XM1_7=$SUsers | where { $_.Target_Store -like "USXM01-DB7" } | foreach { $_.Mailbox } | get-mailbox
$XM1_8=$SUsers | where { $_.Target_Store -like "USXM01-DB8" } | foreach { $_.Mailbox } | get-mailbox

# don't process if the array is null
if ($XM1_1 -ne "" -and $XM1_1 -ne $null -and $XM1_1 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM01-DB1"

# Logging
    $XM1_1 >> "D:\Scripts\Logs\USXM01-DB1.log"

# Let the oncall person know a move is happening

# Generate a move request for each element of the array
    $XM1_1 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited
    $movesuccess=$?
# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM1_1 just failed. User `"$XM1_1`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM1_1 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }

# don't process if the array is null
if ($XM1_2 -ne "" -and $XM1_2 -ne $null -and $XM1_2 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM01-DB2"

# Logging
    $XM1_2 >> "D:\Scripts\Logs\USXM01-DB2.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM1_2 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM1_2 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited
    $movesuccess=$?
# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM1_2 just failed. User `"$XM1_2`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM1_2 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }

# don't process if the array is null
if ($XM1_3 -ne "" -and $XM1_3 -ne $null -and $XM1_3 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM01-DB3"

# Logging
    $XM1_3 >> "D:\Scripts\Logs\USXM01-DB3.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM1_3 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM1_3 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited
    $movesuccess=$?
# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM1_3 just failed. User `"$XM1_3`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM1_3 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }

# don't process if the array is null
if ($XM1_4 -ne "" -and $XM1_4 -ne $null -and $XM1_4 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM01-DB4"

# Logging
    $XM1_4 >> "D:\Scripts\Logs\USXM01-DB4.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM1_4 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM1_4 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited
    $movesuccess=$?

# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM1_4 just failed. User `"$XM1_4`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM1_4 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }

# don't process if the array is null
if ($XM1_5 -ne "" -and $XM1_5 -ne $null -and $XM1_5 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM01-DB5"

# Logging
    $XM1_5 >> "D:\Scripts\Logs\USXM01-DB5.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM1_5 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM1_5 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited
    $movesuccess=$?

# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM1_5 just failed. User `"$XM1_5`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM1_5 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }

# don't process if the array is null
if ($XM1_6 -ne "" -and $XM1_6 -ne $null -and $XM1_6 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM01-DB6"

# Logging
    $XM1_6 >> "D:\Scripts\Logs\USXM01-DB6.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM1_6 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM1_6 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited
    $movesuccess=$?

# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM1_6 just failed. User `"$XM1_6`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM1_6 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }

# don't process if the array is null
if ($XM1_7 -ne "" -and $XM1_7 -ne $null -and $XM1_7 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM01-DB7"

# Logging
    $XM1_7 >> "D:\Scripts\Logs\USXM01-DB7.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM1_7 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM1_7 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited
    $movesuccess=$?

# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM1_7 just failed. User `"$XM1_7`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM1_7 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }

# don't process if the array is null
if ($XM1_8 -ne "" -and $XM1_8 -ne $null -and $XM1_8 -ne " ")
    {
$error.clear()
# Store is a text value, set it and forget it
    $store="USXM01-DB8"

# Logging
    $XM1_8 >> "D:\Scripts\Logs\USXM01-DB8.log"

# Let the oncall person know a move is happening
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM1_8 to $store is starting using PID $pid on $sname."

# Generate a move request for each element of the array
    $XM1_8 | New-MoveRequest -targetdatabase "$store" -AcceptLargeDataLoss -baditemlimit Unlimited
    $movesuccess=$?

# IF the move failed, we need to alert the on-call person
    if ($movesuccess -ne $true)
        {
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "High Alert: Mailbox Move Failure" -Body "The move of mailbox(es) $XM1_8 just failed. User `"$XM1_8`" | get-moverequest to see a report of issues."
	"ERROR - New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
        }
        else
            {
	    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move Initiation" -Body "The move of mailbox(es) $XM1_8 to $store is starting using PID $pid on $sname."
            "New-MoveRequest -targetdatabase `"$store`" -AcceptLargeDataLoss -baditemlimit Unlimited - Successful: $movesuccess" >> "D:\Scripts\Logs\USXM01-DB1.log"
            }
	$error | out-string >> "D:\Scripts\Logs\USXM01-DB1.log"
    }
} # process

# The End section executes once regardless of how many objects are passed through the pipeline
end
{

remove-variable cdatetime
remove-variable Schedule
remove-variable target
remove-variable XM1_1
remove-variable XM1_2
remove-variable XM1_3
remove-variable XM1_4
remove-variable XM1_5
remove-variable XM1_6
remove-variable XM1_7
remove-variable XM1_8
remove-variable csv
remove-variable SUsers
remove-variable RTime
remove-variable sname
remove-variable XMRs
} # end

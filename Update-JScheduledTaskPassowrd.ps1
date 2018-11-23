# **********************************************************************************
#
# Script Name: Update-JScheduledTaskPassword.ps1
# Version: 1.0
# Author: Dave M
# Date Created: 12/10/14
# _______________________________________
#
# MODIFICATIONS:
# Date Modified: N/A
# Modified By: N/A
# Reason for modification: N/A
# What was modified: N/A
#
# Description: Changes the password on all scheduled tasks that use the specified ID.
#
# Usage:
# ./Update-JScheduledTaskPassword.ps1 -ID bpp_exspaudit -server p-ucslog03 -password '12345678'
#
# **********************************************************************************

<#
	.SYNOPSIS
		Changes the password on all scheduled tasks that use the specified ID.

	.DESCRIPTION
		Changes the password on all scheduled tasks that use the specified ID.

	.PARAMETER  ID
		Specify the Run As ID that we want to change the password for.

	.PARAMETER  Server
		Specify the Server to check Scheduled Tasks on.
		NOTE: This is optional. If not specified it will run against every server with an AD description of "UC Server".

	.PARAMETER  Password
		Specify the Password to change the Scheduled Task to use.
		NOTE: This is optional. If not specified it use Get-JPassword to retrieve the current password for BPP_EXSPAudit.

	.EXAMPLE
		./Update-JScheduledTaskPassword.ps1 -ID bpp_exspaudit -server p-ucslog03 -password '12345'

		DESCRIPTION
		===========
		Inspects the Scheduled Tasks on p-ucslog03 and sets the password to "12345" on anty task that has a Run as ID of bpp_exspaudit.
		
	.EXAMPLE
		./Update-JScheduledTaskPassword.ps1 -ID bpp_exspaudit -password '12345'

		DESCRIPTION
		===========
		Inspects the Scheduled Tasks on every server with an AD Description of "UC Server" and sets the password to "12345" on anty task that has a Run as ID of bpp_exspaudit.
		
	.NOTES
		Requires the Janus module loaded.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the first parameter")]
		[String]
        [Alias("Identity")]
$ID,
        [Parameter(Position = 1, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the second parameter")]
		[String]
        [Alias("Computer")]
        [Alias("System")]
        [Alias("DNSHostNAme")]
        [Alias("SAMAccountName")]
        [Alias("ComputerName")]
        [Alias("ServerName")]
$server="*",
        [Parameter(Position = 2, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the third parameter")]
		[String]
$password="12345678"
)

# Functions and Filters


# Main Script
# The Begin section executes once regardless of how many objects are passed through the pipeline
begin
{
# Is the Janus Module loaded?
if ($global:JanusPSModule -ne $true) {Import-Module Janus -erroraction stop}

# Initialization
$tasks=@()
$CPUs=@()

# If serve is "*", then we need get every server with a description of "UC Server"
if ($server -eq "*")
	{
if ($d) {write-host "Looking up servers"}
	$CPUs=Get-JADEntry "UC Server" -description -pso -properties dnshostname,samaccountname,description
	} else
	{
if ($d) {write-host "Testing Specified Server $server"}
	if (Test-Connection -computername $server -count 1 -quiet) {$CPUs=Get-JADEntry -id $server -exact -pso -properties dnshostname,samaccountname,description}
	}

# We are going to user -like with *'s, so we cant to strip out any extraneous ID info like domains.
if ($ID.contains("janus_cap`\")) {$ID.replace("janus_cap`\","")}
if ($ID.contains("janus_cap`\")) {$ID.replace("janus.cap`\","")}
if ($ID.contains("janus_cap`\")) {$ID.replace("`@janus_cap","")}
if ($ID.contains("janus_cap`\")) {$ID.replace("`@janus.cap","")}
# Does the ID exist?
$test=Get-JADEntry -ID $ID -exact -properties SAMAccount name
if (($test -eq $null) -or($test -eq "")) {write-error "Cannot locate an AD account names $ID, exiting..." -erroraction stop}

if ($password -eq "12345678") {$password=Get-JPassword -nocrypt}
}

# The process section runs once for each object in the pipeline
process
{
# This will just run once if there is only one computer specified
foreach ($CPU IN $CPUs)
	{
write-host "Processing $($CPU.dnshostname)"
# Populate $tasks with all the relevant Scheduled Tasks on the current computer
	$tasks=Get-JScheduledTasks -computername $($CPU.dnshostname) | where {$_.RunAs -like "*$ID*"}
# Process each Scheduled Task
	foreach ($task IN $tasks)
		{
# TaskName is needed to change the password
		$taskName=$task.TaskName
# We need to remove the leading "\" in order for the cmdlet to work
		$taskName=$taskName.replace("`\","")
# This is in hte Janus Module, it does the actual work
		Set-JScheduledTaskPassword -TaskCredential $password -ComputerName $($CPU.dnshostname) -TaskName $taskName -Confirm:$false
		$success=$?
# write-host "Set-JScheduledTaskPassword -TaskCredential $password -ComputerName $($CPU.dnshostname) -TaskName $taskName -Confirm:`$false - success $success"
		}
	}
}

# The End section executes once regardless of how many objects are passed through the pipeline
end
{
# $tasks

# Clean up memory use
remove-variable tasks
remove-variable CPUs
}

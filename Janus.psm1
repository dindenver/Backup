# **********************************************************************************
#
# Script Name: Janus.psm1
# Version: 1.8
# Author: Dave M and Dennis K
# Date Created: ‎June ‎14, ‎2011
# _______________________________________
#
# MODIFICATIONS: 
# Date Modified: 12-20-2011
# Modified By: Dave M
# Reason for modification: Added Standard Header and Logging
# What was modified: Header and functions
#
# Date Modified: 9-28-2012
# Modified By: Dave M
# Reason for modification: Added Inheritance functions
#
# Date Modified: 11-8-2012
# Modified By: Dave M
# Reason for modification: Added AD Attribue Error checking to Get-JADEntry
#
# Date Modified: 12-28-2012
# Modified By: Dave M
# Reason for modification: Added Add-JHolidays
#
# Date Modified: 4-22-2013
# Modified By: Dave M
# Reason for modification: Added Show-JFailureAuditEvents
#
# `: This Module contains cmdlets for use at Janus.
#
# Usage:
# Import-Module Janus
#
# **********************************************************************************

$ErrorActionPreference = "SilentlyContinue"
$WarningActionPreference = "SilentlyContinue"

$error.clear()

set-alias -name cat			get-content
set-alias -name cd			set-location
set-alias -name cls			clear-host
set-alias -name kill			stop-process
set-alias -name echo			write-output

set-alias -name copy			copy-item
set-alias -name del			remove-item
set-alias -name dir			get-childitem
set-alias -name move			move-item
set-alias -name set			set-variable
set-alias -name type			get-content
set-alias -name grep			Get-JStringMatches

# Legacy Janus Module Aliases
set-alias -name Show-JMessageTracking -value Get-JMessageTracking -Scope Global
set-alias -name Add-Delegate -value Add-JEWSDelegate -Scope Global
set-alias -name Get-PF -value Show-JPublicFolderInfo -Scope Global
set-alias -name Move-ADEntry -value Move-JADEntry -Scope Global
set-alias -name Add-Contact -value Add-JContact -Scope Global
set-alias -name Set-Pass -value Set-JPassword -Scope Global
set-alias -name Move-Bulk -value Move-JMailboxesInBulk -Scope Global
set-alias -name Out-Excel -value Out-JExcel -Scope Global
set-alias -name Report-ASUser -value Show-JActiveSyncUserInfo -Scope Global
set-alias -name Add-CRDelegate -value Add-JCRDelegate -Scope Global
set-alias -name Report-OOF -value Show-JOOFSettings -Scope Global
set-alias -name SendTo-Zip -value ConvertTo-JZip -Scope Global
set-alias -name Fix-Tentative -value Update-JCalendarTentativeSettings -Scope Global
set-alias -name Execute-RemoteCMD -value Send-JRemoteCMD -Scope Global
set-alias -name Control-Service -value Repair-JService -Scope Global

set-alias -name Check-Para -value Protect-JParameter -Scope Global
set-alias -name Get-ADEntry -value Get-JADEntry -Scope Global
set-alias -name OctetToGUID -value ConvertFrom-JOctetToGUID -Scope Global
set-alias -name Set-ADEntry -value Set-JADEntry -Scope Global
set-alias -name Change-SMTPAddress -value Update-JPrimarySMTPAddress -Scope Global

function help
{
    $test=@()
    [string]$filter=""
    [string]$cmdlets=""
    $test=get-help -full $args[0]
    $cmdlets=$test
    if ($test.count -lt 1)
	{
	$filter='*' + $args[0] + '*'
	$filter=$filter.replace('**','*')
	write-host "Search Filter: $filter"
	get-help -full $($filter)
	}
    elseif (($test.count -eq 2) -and($cmdlets.contains("Microsoft.Exchange.Management.PowerShell.E2010")))
	{
	$filter='Microsoft.Exchange.Management.PowerShell.E2010\' + $args[0]
	write-host "Search Filter: $filter"
	get-help -full $filter
	}
    else
	{
	get-help -full $args[0]
	}
} # function help

function prompt
{
    "PS " + $(get-location) + "> "
}

& {
    for ($i = 0; $i -lt 26; $i++) 
    { 
        $funcname = ([System.Char]($i+65)) + ':'
        $str = "function global:$funcname { set-location $funcname } " 
        invoke-expression $str 
    }
} # function prompt

function Get-JMessageTracking
{
<#
	.SYNOPSIS
		Searches all mailbox and Hub servers for the message specified.

	.DESCRIPTION
		Generates a list of all Mailbox and Transport Servers and searches
		for the message requested. It returns all Events that match the
		specified criteria.

	.PARAMETER  from
		Enter the smtp address of the Sender you want to track.

	.PARAMETER  to
		Enter the smtp address of the recipient you want to track.
		NOTE: Includes To, CC and BCC fields.

	.PARAMETER  subject
		Enter the Subject line you are searching for.
		NOTE: Message tracking automatically treats the string you
		enter as "*$subject*". This is a feature of message tracking
		and cannot be changed.

	.PARAMETER  start
		Enter the earliest date for the message you want to track.
		NOTE: Will default to 30 days ago if not specified.

	.PARAMETER  end
		Enter the latest date for the message you want to track.
		NOTE: Will default to the moment of execution if not specified.

	.PARAMETER  report
		This switch instructs the cmdlet to sort and display the entire
		message details by date.

	.EXAMPLE
		Get-JMessageTracking -from Randy.Moore@janus.com -to rgmoore@hotmail.com -subject "Test e-mail" -start "6/8/2011 9:00:05 am" -end "6/8/2011 9:01:38 am"
		
		Description
		-----------
		Returns any messages sent after 9:00:05 AM on 6/8/2011 and before
		9:01:38 AM on 6/8/2011 with the subject of "Test e-mail" sent from
		Randy.Moore@janus.com to rgmoore@hotmail.com
		
	.NOTES
		Requires the Loading of the Exchange Module.
		Show-JMessageRecipients uses the same parameters, but returns
		the confirmed recipients of the message.

#>

# Initialization Section
# Parameters called so that they can be passed into the script via the command line.
	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $false)]
		[String]
$from=$null,
        [Parameter(Position = 1, Mandatory = $false)]
		[String]
$to=$null,
        [Parameter(Position = 2, Mandatory = $false)]
		[String]
$subject=$null,
        [Parameter(Position = 3, Mandatory = $false)]
		[system.datetime]
$start,
        [Parameter(Position = 4, Mandatory = $false)]
		[system.datetime]
$end,
        [Parameter(Position = 4, Mandatory = $false)]
		[switch]
$report=$false
)

$results=@()
$pass=0
$case=0

# Main script
if (-not ($global:ExchangeSnapIn))
	{
	[string]$errorstring="WARNING: Exchange module is not loaded, some cmdlets will not work correctly, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Set the Start Date if the parameter is not set
if ($start -eq $null)
	{
	$start=(get-date).adddays(-30)
	} else
	{
	$start=get-date($start)
	}

# Set the End Date if the parameter is not set
if ($end -eq $null)
	{
	$end=get-date
	} else
	{
	$end=get-date($end)
	}

if ($from -ne "")
	{
	$case=$case+1
	}

if ($to -ne "")
	{
	$case=$case+2
	}

if ($subject -ne "")
	{
	$case=$case+4
	}

# $case is used to identify which get-messagetrackinglog paramters are populated
switch ($case)
	{
	"0"
		{
# This retreives objects representing all of the hub (and edge) servers
		$servers=Get-TransportServer
# This retreives objects representing all of the mailbox servers
		$servers+=Get-MailboxServer
		foreach ($server in $servers)
				{
				write-host("Processing $server...")
# $results holds all of the message events
				$results+=get-messagetrackinglog -Server $server -Start $start -End $end -resultsize unlimited
				$pass=$pass+1
				$matches=$results.length
# something happens with $matches, we need to reduce it by the number of iterations in the Foreach loop
				$matches=$matches - $pass
				write-host ("$matches result(s) so far...")
				}
# If you use the -report switch, formats and sorts the events
		if($report)
			{
			write-host ("Sorting result(s)...")
			$results | sort Timestamp | fl | write-output
			} else
# Or else it just returns the collection of events for manipulation
			{
			return $results
			}
		}
	
	"1"
		{
		$servers=Get-TransportServer
		$servers+=Get-MailboxServer
		foreach ($server in $servers)
				{
				write-host("Processing $server...")
				$results+=get-messagetrackinglog -Sender $from -Server $server -Start $start -End $end -resultsize unlimited
				$pass=$pass+1
				$matches=$results.length
				$matches=$matches - $pass
				write-host ("$matches result(s) so far...")
				}
		if($report)
			{
			write-host ("Sorting result(s)...")
			$results | sort Timestamp | fl | write-output
			} else
			{
			return $results
			}
		}
	
	"2"
		{
		$servers=Get-TransportServer
		$servers+=Get-MailboxServer
		foreach ($server in $servers)
				{
				write-host("Processing $server...")
				$results+=get-messagetrackinglog -Recipient $to -Server $server -Start $start -End $end -resultsize unlimited
				$pass=$pass+1
				$matches=$results.length
				$matches=$matches - $pass
				write-host ("$matches result(s) so far...")
				}
		if($report)
			{
			write-host ("Sorting result(s)...")
			$results | sort Timestamp | fl | write-output
			} else
			{
			return $results
			}
		}
	
	"3"
		{
		$servers=Get-TransportServer
		$servers+=Get-MailboxServer
		foreach ($server in $servers)
				{
				write-host("Processing $server...")
				$results+=get-messagetrackinglog -Sender $from -Recipient $to -Server $server -Start $start -End $end -resultsize unlimited
				$pass=$pass+1
				$matches=$results.length
				$matches=$matches - $pass
				write-host ("$matches result(s) so far...")
				}
		if($report)
			{
			write-host ("Sorting result(s)...")
			$results | sort Timestamp | fl | write-output
			} else
			{
			return $results
			}
		}
	
	"4"
		{
		$servers=Get-TransportServer
		$servers+=Get-MailboxServer
		foreach ($server in $servers)
				{
				write-host("Processing $server...")
				$results+=get-messagetrackinglog -Server $server -messagesubject "$subject" -Start $start -End $end -resultsize unlimited
				$pass=$pass+1
				$matches=$results.length
				$matches=$matches - $pass
				write-host ("$matches result(s) so far...")
				}
		if($report)
			{
			write-host ("Sorting result(s)...")
			$results | sort Timestamp | fl | write-output
			} else
			{
			return $results
			}
		}
	
	"5"
		{
		$servers=Get-TransportServer
		$servers+=Get-MailboxServer
		foreach ($server in $servers)
				{
				write-host("Processing $server...")
				$results+=get-messagetrackinglog -Sender $from -Server $server -messagesubject "$subject" -Start $start -End $end -resultsize unlimited
				$pass=$pass+1
				$matches=$results.length
				$matches=$matches - $pass
				write-host ("$matches result(s) so far...")
				}
		if($report)
			{
			write-host ("Sorting result(s)...")
			$results | sort Timestamp | fl | write-output
			} else
			{
			return $results
			}
		}
	
	"6"
		{
		$servers=Get-TransportServer
		$servers+=Get-MailboxServer
		foreach ($server in $servers)
				{
				write-host("Processing $server...")
				$results+=get-messagetrackinglog -Recipient $to -Server $server -messagesubject "$subject" -Start $start -End $end -resultsize unlimited
				$pass=$pass+1
				$matches=$results.length
				$matches=$matches - $pass
				write-host ("$matches result(s) so far...")
				}
		if($report)
			{
			write-host ("Sorting result(s)...")
			$results | sort Timestamp | fl | write-output
			} else
			{
			return $results
			}
		}
	
	"7"
		{
		$servers=Get-TransportServer
		$servers+=Get-MailboxServer
		foreach ($server in $servers)
				{
				write-host("Processing $server...")
				$pass=$pass+1
				$matches=$results.length
				$matches=$matches - $pass
				$matches=$matches - 1
				write-host ("$matches result(s) so far...")
				}
		if($report)
			{
			write-host ("Sorting result(s)...")
			$results | sort Timestamp | fl | write-output
			} else
			{
			return $results
			}
		}
	
	default
		{
		[string]$errorstring="PARAMETER ERROR, exiting...`n"
		$Error.add($errorstring)
		write-error ($errorstring) -erroraction stop
		}
	}
remove-variable servers -ErrorAction SilentlyContinue
remove-variable results -ErrorAction SilentlyContinue
remove-variable matches -ErrorAction SilentlyContinue
remove-variable pass -ErrorAction SilentlyContinue
} # function Get-JMessageTracking

function Show-JMessageRecipients
{
<#
	.SYNOPSIS
		Searches all mailbox and Hub servers for the recipients of the 
		message specified.

	.DESCRIPTION
		Generates a list of all Mailbox and Transport Servers and searches
		for the message requested. It returns all reipients that match the
		specified criteria.

	.PARAMETER  from
		Enter the smtp address of the Sender you want to track.

	.PARAMETER  to
		Enter the smtp address of the recipient you want to track.
		NOTE: Includes To, CC and BCC fields.

	.PARAMETER  subject
		Enter the Subject line you are searching for.

	.PARAMETER  start
		Enter the earliest date for the message you want to track.
		NOTE: Will default to one week ago if not specified.

	.PARAMETER  end
		Enter the latest date for the message you want to track.
		NOTE: Will default to the moment of execution if not specified.

	.EXAMPLE
		Show-MessageRecipients -from Randy.Moore@janus.com -to rgmoore@hotmail.com -subject "Test e-mail" -start "6-8-2011 9:00 AM" -end "6-8-2011 9:01 AM"
		
		Description
		-----------
		Returns the recipients of any messages sent after 9:00:05 AM on
		6/8/2011 and before 9:01:38 AM on 6/8/2011 with the subject of
		"Test e-mail" sent from Randy.Moore@janus.com to rgmoore@hotmail.com
		
	.NOTES
		Requires the Loading of the Exchange Module.

#>

# Initialization Section
# Parameters called so that they can be passed into the script via the command line.
	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $false)]
		[String]
$from=$null,
        [Parameter(Position = 1, Mandatory = $false)]
		[String]
$to=$null,
        [Parameter(Position = 2, Mandatory = $false)]
		[String]
$subject=$null,
        [Parameter(Position = 3, Mandatory = $false)]
		[system.datetime]
$start,
        [Parameter(Position = 4, Mandatory = $false)]
		[system.datetime]
$end)

[string[]]$rcpts=@()

$case=0

if (-not ($global:ExchangeSnapIn))
	{
	[string]$errorstring="WARNING: Exchange module is not loaded, some cmdlets will not work correctly, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

if ($start -eq $null)
	{
	$start=(get-date).adddays(-30)
	}

if ($end -eq $null)
	{
	$end=get-date
	}

if ($from -ne "")
	{
	$case=$case+1
	}

if ($to -ne "")
	{
	$case=$case+2
	}

if ($subject -ne "")
	{
	$case=$case+4
	}

# $case is used to determine which get-messagetrackinglog parameters are specified
switch ($case)
	{
	"0"
		{
# collects all of the transport servers (hubs and edge if present)
		$servers=Get-TransportServer
# Collects all of the mailbox servers
		$servers+=Get-MailboxServer
		foreach ($server in $servers)
				{
				write-host ("Searching Logs on $server...")
# DELIVER event is the event that identifies that a message made it to the users' mailbox
				$collection=get-messagetrackinglog -Server $server -EventId "DELIVER" -Start $start -End $end -resultsize unlimited
				foreach ($msg in $collection)
					{
# $msg.Recipients has the recipients that the DELIVER event applies to
					$msgarray=$msg.Recipients
					foreach ($rcpt IN $msgarray)
						{
						$rcpts+=$rcpt
						}
					}
				}
		}
	
	"1"
		{
		$servers=Get-TransportServer
		$servers+=Get-MailboxServer
		foreach ($server in $servers)
				{
				write-host ("Searching Logs on $server...")
				$collection=get-messagetrackinglog -Sender $from -Server $server -EventId "DELIVER" -Start $start -End $end -resultsize unlimited
				foreach ($msg in $collection)
					{
					$msgarray=$msg.Recipients
					foreach ($rcpt IN $msgarray)
						{
						$rcpts+=$rcpt
						}
					}
				}
		}
	
	"2"
		{
		$servers=Get-TransportServer
		$servers+=Get-MailboxServer
		foreach ($server in $servers)
				{
				write-host ("Searching Logs on $server...")
				$collection=get-messagetrackinglog -Recipient $to -Server $server -EventId "DELIVER" -Start $start -End $end -resultsize unlimited
				foreach ($msg in $collection)
					{
					$msgarray=$msg.Recipients
					foreach ($rcpt IN $msgarray)
						{
						$rcpts+=$rcpt
						}
					}
				}
		}
	
	"3"
		{
		$servers=Get-TransportServer
		$servers+=Get-MailboxServer
		foreach ($server in $servers)
				{
				write-host ("Searching Logs on $server...")
				$collection=get-messagetrackinglog -Sender $from -Recipient $to -Server $server -EventId "DELIVER" -Start $start -End $end -resultsize unlimited
				foreach ($msg in $collection)
					{
					$msgarray=$msg.Recipients
					foreach ($rcpt IN $msgarray)
						{
						$rcpts+=$rcpt
						}
					}
				}
		}
	
	"4"
		{
		$servers=Get-TransportServer
		$servers+=Get-MailboxServer
		foreach ($server in $servers)
				{
				write-host ("Searching Logs on $server...")
				$collection=get-messagetrackinglog -Server $server -EventId "DELIVER" -messagesubject "$subject" -Start $start -End $end -resultsize unlimited
				foreach ($msg in $collection)
					{
					$msgarray=$msg.Recipients
					foreach ($rcpt IN $msgarray)
						{
						$rcpts+=$rcpt
						}
					}
				}
		}
	
	"5"
		{
		$servers=Get-TransportServer
		$servers+=Get-MailboxServer
		foreach ($server in $servers)
				{
				write-host ("Searching Logs on $server...")
				$collection=get-messagetrackinglog -Sender $from -Server $server -EventId "DELIVER" -messagesubject "$subject" -Start $start -End $end -resultsize unlimited
				foreach ($msg in $collection)
					{
					$msgarray=$msg.Recipients
					foreach ($rcpt IN $msgarray)
						{
						$rcpts+=$rcpt
						}
					}
				}
		}
	
	"6"
		{
		$servers=Get-TransportServer
		$servers+=Get-MailboxServer
		foreach ($server in $servers)
				{
				write-host ("Searching Logs on $server...")
				$collection=get-messagetrackinglog -Recipient $to -Server $server -EventId "DELIVER" -messagesubject "$subject" -Start $start -End $end -resultsize unlimited
				foreach ($msg in $collection)
					{
					$msgarray=$msg.Recipients
					foreach ($rcpt IN $msgarray)
						{
						$rcpts+=$rcpt
						}
					}
				}
		}
	
	"7"
		{
		$servers=Get-TransportServer
		$servers+=Get-MailboxServer
		foreach ($server in $servers)
				{
				write-host ("Searching Logs on $server...")
				$collection=get-messagetrackinglog -Sender $from -Recipient $to -Server $server -EventId "DELIVER" -messagesubject "$subject" -Start $start -End $end -resultsize unlimited
				foreach ($msg in $collection)
					{
					$msgarray=$msg.Recipients
					foreach ($rcpt IN $msgarray)
						{
						$rcpts+=$rcpt
						}
					}
				}
		}
	
	default
		{
		[string]$errorstring="PARAMETER ERROR, exiting...`n"
		$Error.add($errorstring)
		write-error ($errorstring) -erroraction stop
		}
	}

# De-dupes the list of recipients
$rcpts=$rcpts | Sort-Object -Unique
$rcpts | ft -auto -wrap

remove-variable servers -ErrorAction SilentlyContinue
remove-variable rcpts -ErrorAction SilentlyContinue
} # function Show-JMessageRecipients

function Add-JEWSDelegate
{
<#
	.SYNOPSIS
		Adds a delegate to a mailbox.

	.DESCRIPTION
		Adds the specified Delegate to the specified folder to a mailbox.

	.PARAMETER  mb
		Enter the SMTP address of the mailbox being delegated

	.PARAMETER  viewer
		Enter the SMTP address of the mailbox the will be able to view new
		folders.

	.PARAMETER  folder
		Specify the name of the Foder that the Delegate will have access to.
		This will accept a comma-seperated list for multiple folders.
		NOTE: This defaults to "Calendar" if it is not specified.

	.PARAMETER  perms
		Specify the Permissions to grant from the following list:
		None, Contributor, Reviewer, Author, Editor.
		NOTE: This defaults to "Reviewer" if it is not specified.

	.PARAMETER  copy
		Specify if Delegate should receive a copy of Meeting-related Messages.

	.EXAMPLE
		Add-JEWSDelegate -mb JournalSocialNe@janus.com -viewer Randy.Moore@janus.com -folder Calendar
		
		Description
		-----------
		Allows Randy Moore to be able to view Calendar of thr Journal Social
		Networking Mailbox.
		
	.NOTES
		Requires EWS to be installed on the computer that executes this
		script.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("Identity")]
        [Alias("ID")]
		[String]
$mb=$(Throw "You must specify a Calendar to add a delegate to (e.g., JournalSocialNe@janus.com)."),
        [Parameter(Position = 1, Mandatory = $true)]
        [Alias("user")]
		[String]
$viewer=$(Throw "You must specify a delegate to add (e.g., Randy.Moore@janus.com)."),
        [Parameter(Position = 2, Mandatory = $false)]
		[String[]]
		[ValidateSet(
			'Cal',
			'Calendar',
			'Contact',
			'Contacts',
			'Inbox',
			'Note',
			'Notes',
			'Task',
			'Tasks'
		)]
$folder="Calendar",
        [Parameter(Position = 3, Mandatory = $false)]
        [Alias("permissions")]
        [Alias("permission")]
		[String]
		[ValidateSet(
			'Editor',
			'Author',
			'Reviewer',
			'Contributor',
			'None'
		)]
$perms="Reviewer",
        [Parameter(Position = 4, Mandatory = $false)]
        [Alias("ReceiveCopiesOfMeetingRequests")]
        [Alias("ReceiveCopiesOfMeetingMessages")]
		[switch]
$copy
)

# Initialization
# Sets up access to the EWS DLL
$EWSFile=$null
if (test-path "D:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll") {$EWSFile="D:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"}
if (test-path "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll") {$EWSFile="C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"}
if (test-path "D:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll") {$EWSFile="D:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll"}
if (test-path "C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll") {$EWSFile="C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll"}
$test=Get-Module -ListAvailable | where {$_.name -like "EWSMail"}

if (($test -ne $null) -and($EWSFile -ne $null)) {
	$dllpath = $EWSFile
	[void][Reflection.Assembly]::LoadFile($dllpath)
	Add-Type -Path $EWSFile
	Remove-Module EWSMail -erroraction silentlyContinue
	Import-Module EWSMail -erroraction silentlyContinue
	$global:EWSModule=$?
	}

# Initates an object with the EWS services as methods
$service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

# Filters and Functions
# switch is case sensitive, converting $folder to lower case makes the switch statement less complicated
$folder=$folder.tolower()

# $global:ExchangeSnapIn is a variable that is created and set to true in the Janus Module if the snapin is loaded
if (-not ($global:ExchangeSnapIn))
	{
	[string]$errorstring="WARNING: Exchange module is not loaded, some cmdlets will not work correctly, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# $global:EWSModule is a variable that is created and set to true in the Janus Module if EWS is loaded
if (-not ($global:EWSModule))
	{
	[string]$errorstring="WARNING: EWS module is not loaded, come cmdlets will not work correctly...`nExiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Error checking
# does $mb exist?
$test=get-mailbox -id $mb
$successful=$?
if ($successful -eq $false)
	{
	[string]$errorstring="Error in Mailbox ID, cannot locate mailbox, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	} else
	{
	$mb=$test.PrimarySmtpAddress.tostring()
	}

# does $viewer exist?
$test=get-jadentry -id $viewer -exact -pso -properties msexchrecipientdisplaytype,mail
$successful=$?
if (($successful -eq $false) -or($test.msexchrecipientdisplaytype=$null) -or($test.msexchrecipientdisplaytype=""))
	{
	[string]$errorstring="Error in Delegate ID, cannot locate mailbox, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	} else
	{
	$viewer=$test.mail
	}

# Main Script
# Try allows you to catch an error. In this case, an error is generated if p-uccas03 is returned because it is a re-direct in http
try
	{
	$service.AutodiscoverUrl($mb)
	}


# Manually assigns the url if there is an error
Catch [system.exception]
 {
	$URI='https://p-ucusxhc04.janus.cap/EWS/Exchange.asmx'
	$service.URL = New-Object Uri($URI)
	write-output ("Caught an Autodiscover URL exception, recovering...")
 }

# Sets the EWS service to impersonate the mailbox you are modifying
$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $mb);

# Sets up a EWS Mailbox Object
$mbMailbox = new-object Microsoft.Exchange.WebServices.Data.Mailbox($mb)
# Sets up a EWS Delegate User Object
$dgUser = new-object Microsoft.Exchange.WebServices.Data.DelegateUser($viewer)

# process for Calendars
If ($folder -like "*cal*")
	{
	Write-Output ("Updating Calendar")
# Sets the permissions for the Delegate
	$dgUser.Permissions.CalendarFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::$perms
# If the Delegate can act on the Calendar messages, they will receive a copy
	if (($copy -eq $true) -and($perms -like "Editor")) {$dgUser.ReceiveCopiesOfMeetingMessages = $true}
	else  {$dgUser.ReceiveCopiesOfMeetingMessages = $false}
	}
# Process for Inbox
If ($folder -like "*inbox*")
	{
	Write-Output ("Updating Inbox")
	$dgUser.Permissions.InboxFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::$perms
	}
Process for Tasks folder
If ($folder -like "*task*")
	{
	Write-Output ("Updating Tasks")
	$dgUser.Permissions.TasksFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::$perms
	}
# Process for Contacts folder
If ($folder -like "*contact*")
	{
	Write-Output ("Updating Contacts")
	$dgUser.Permissions.ContactsFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::$perms
	}
# Prodcess for Notes Folder
If ($folder -like "*note*")
	{
	Write-Output ("Updating Notes")
	$dgUser.Permissions.NotesFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::$perms
	}

# Don't give delegates access to private items
$dgUser.ViewPrivateItems = $false

# Initiates an empty array
$dgArray = new-object Microsoft.Exchange.WebServices.Data.DelegateUser[] 1
# Assigned the Delegate user to the first element of the array
$dgArray[0] = $dgUser

# Sets Meeting inbvites to default behavior if they have enough permissions
if (($perms -like "Editor") -and($copy)) {$service.AddDelegates($mbMailbox, [Microsoft.Exchange.WebServices.Data.MeetingRequestsDeliveryScope]::DelegatesAndSendInformationToMe, $dgArray);}
# Makes sure that delegate does not receive meeting invites if they do not have sufficient permissions to act on those invites
else  {$service.AddDelegates($mbMailbox, [Microsoft.Exchange.WebServices.Data.MeetingRequestsDeliveryScope]::NoForward, $dgArray);}

remove-variable service
remove-variable dgArray
remove-variable dgUser
remove-variable mbMailbox
remove-variable ACEUser
remove-variable WindowsIdentity
} # function Add-JEWSDelegate

function Move-JMailbox
{
<#
	.SYNOPSIS
		Move an Exchange Mailbox.

	.DESCRIPTION
		Move the designated mailbox to the specified mailbox store. This
		sends an e-mail when it begins and finishes. It also sends a MOM
		alert if there is an error during the move. This also Logs to the
		C: drive.

	.PARAMETER  user
		Enter the Mailbox to move.
		NOTE: Accepts the following data:
		* GUID
		* Distinguished name (DN)
		* Domain\Account
		* User principal name (UPN)
		* LegacyExchangeDN
		* SmtpAddress
		* Alias

	.PARAMETER  store
		Specify the target Mailbox Database Store.

	.EXAMPLE
		Move-JMailbox -user "Moore, Randy" -Store "ucmailv03\Mailbox Store 3-1"
		
		Description
		-----------
		Moves Randy's mailbox to the Mailbox Store 3-1 on the ucmailv03
		server,
		
	.NOTES
		Requires the Loading of the Exchange Module.

#>

# Initialization Section
# Parameters called so that they can be passed into the script via the command line.
	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("Identity")]
        [Alias("ID")]
		[String]
$user,
        [Parameter(Position = 1, Mandatory = $true)]
        [Alias("mailboxstore")]
        [Alias("database")]
		[String]
$store)

$movesuccess=$null
$mrtype=$null
$sname=$null
$DC=(Get-JDCs -gc).Name

# Functions and Filters Section
# Makes sure the Exchange SnapIn is loaded
if (-not ($global:ExchangeSnapIn))
	{
	[string]$errorstring="WARNING: Exchange module is not loaded, some cmdlets will not work correctly, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# MAIN Script
# Does $user exist
$testmb=get-mailbox -id $user
$success=$?
if ($successful -eq $false)
	{
	[string]$errorstring="Error locating mailbox named " + $user + ", exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Does $store exist
$testdb=Get-MailboxDatabase -id $store
$success=$?
if ($successful -eq $false)
	{
	[string]$errorstring="Error locating mailbox store named " + $store + ", exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Moves mailbox with reporting turned on and with the best defaults to ensure speed and success
New-MoveRequest -identity $user -TargetDatabase "$store" -DomainController $DC -AcceptLargeDataLoss -baditemlimit unlimited -AllowLargeItems -confirm:$false
$movesuccess=$?
if ($movesuccess -eq $false)
    {
# Alert when move fails
    Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Alert: High Resolution state: New - Mailbox Move Request Failure" -Body "The move request for mailbox $user just failed."
    }
    else
        {
# Notification when move is successful
        Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Mailbox Move success" -Body "The move request for mailbox $user to move to $store was created successfully."
        }

} # function Move-JMailbox

function Show-JDistributionListMembers
{
<#
	.SYNOPSIS
		Reports an alphabetized list of DL members.

	.DESCRIPTION
		Reports all members of a DL (including sub-DLs) in alphabetical
		order. Can also be used to see if a specific SMTP address is a member.

	.PARAMETER  list
		Enter the Display Name of a Distribution Group (aka Distribution List).
		NOTE: Also accepts:
		* GUID
		* Distinguished name (DN)
		* User principal name (UPN)
		* LegacyExchangeDN
		* Domain\Account name
		* Alias

	.PARAMETER  ID
		Enter the display name of the AD account that you want to see are a
		member of this List/Group.

	.EXAMPLE
		Show-JDistributionListMembers  -list "1MISC-Exchange Administrators"
		
		Description
		-----------
		Returns all members of the 1MISC-Exchange Administrators
		Distribution Group and all of its sub-groups.
		
	.EXAMPLE
		Show-JDistributionListMembers  -list "1MISC-Exchange Administrators" -id "Moore, Randy"
		
		Description
		-----------
		Returns whether or not Moore, Randy is a member of the
		1MISC-Exchange Administrators Distribution Group.
		
	.NOTES
		Requires the Loading of the Exchange Module.

#>

	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true)]
		[String]
        [Alias("Group")]
        [Alias("DL")]
$list=$(Throw "You must specify a distribution group (e.g., IT Messaging Services)."),
        [Parameter(Position = 1, Mandatory = $false)]
        [Alias("Identity")]
		[String]
$ID)

# Initializes $hits as an array.
$script:hits=@()

# Is the Exchange Snapin Loaded?
if (-not ($global:ExchangeSnapIn))
	{
	[string]$errorstring="WARNING: Exchange module is not loaded, some cmdlets will not work correctly, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Functions
# This function is the meat of the procedure. Since we don't know how many times it will be called, we had to break it out of the main script logic
function DLCheck
{
Param([string]$group,[string]$check)
# Creates a collection of objects that represent the members of the list
$members=Get-DistributionGroupMember -id $group
Foreach ($member in $members)
	{
# $check tells the function to compare the members to a specific address
	If ($check -eq $null -or $check -eq "")
		{
# We have to use a script level variable in order for the main function to access this data
		$script:hits+=$member
# If one of the members is a group, we have to start all over again with another instance of DLCheck
		if ($member.recipienttype -like "*group")
			{
			$grp = Get-DynamicDistributionGroup -identity $($member.Name)
			$successful=$?
			if ($successful)
			{
# This is the routine to enumerate DDLs or Query-based lists
			$script:hits+=get-Recipient -RecipientPreviewFilter $grp.LdapRecipientFilter
			write-output ("Note: Results for Dynamic Lists are subject to changes in AD...")
			} else
			{
			DLCheck -group $member -check $check
			} # if ($successful)
			} # if ($member.recipienttype -like "*group")
		} else
		{
		if ($member.recipienttype -like "*group")
			{
			DLCheck -group $member -check $check
			} else
			{
			if (($member.displayname -like "*$check*") -or($member.SamAccountName -like "*$check*") -or ($member.Alias -like "*$check*") -or ($member.PrimarySmtpAddress -like "*$check*"))
				{
				$script:hits+=$member
				}
			} # if ($member.recipienttype -like "*group")
		} # If ($check -eq $null -or $check -eq "")
	} # Foreach ($member in $members)
remove-variable members
} # function DLCheck

# Main script

# Error checking
# Does the group exist and if so is it a DDL
$group = Get-DynamicDistributionGroup -identity $list
$successful=$?
if ($successful)
{
# If the group is a DDL, then all the mebers will be returned by this command
get-Recipient -RecipientPreviewFilter $group.LdapRecipientFilter | select displayname,PrimarySmtpAddress,name | sort displayname
write-output ("Note: Results for Dynamic Lists are subject to changes in AD...")
} else
{
# If it is not a DDL, you will need to process it
$group = Get-DistributionGroup -id $list
$successful=$?
if (-not ($successful))
	{
	[string]$errorstring="Error validating group, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	} # if (-not ($successful))
DLCheck -group $list -check $ID
Write-Output ("Enumerated list of DL members")
$script:hits | select displayname,PrimarySmtpAddress,SamAccountName | ft -auto -wrap
Write-Output ("De-duped list of DL members")
$script:hits | select displayname,PrimarySmtpAddress,SamAccountName -unique | sort displayname | ft -auto -wrap
} # if ($successful)

remove-variable members -ErrorAction SilentlyContinue
remove-variable hits -ErrorAction SilentlyContinue
} # function Show-JDistributionListMembers

function Show-JPublicFolderInfo
{
<#
	.SYNOPSIS
		Returns data on the requested PF.

	.DESCRIPTION
		Returns all attributes of the specified Public Folder.

	.PARAMETER  pf
		Enter the sub-string that should appear anywhere in the PF Identity.

	.EXAMPLE
		Show-JPublicFolderInfo -pf "IT Operations"
		
		Description
		-----------
		Returns all the attributes of any Public Folder with *IT Operations^
		in the Identity.
		
	.NOTES
		Requires the Loading of the Exchange Module.

#>

	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("publicfolder")]
        [Alias("ID")]
		[String]
$pf=$(Throw "You must specify a Public Folder (e.g., `"IT Operations`").")
)

# Is the Exchange Snap In loaded?
if (-not ($global:ExchangeSnapIn))
	{
	[string]$errorstring="WARNING: Exchange module is not loaded, some cmdlets will not work correctly, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Error checking
[string]$Server=""
# This routine discovers PF Servers in order
$successful=$true
do
	{
	if($successful -eq $false) {Write-Output ("Looking for Public Folder Server...")}
	$successful=test-path -path "\\P-UCSGXM01\c$"
	if($successful) {$Server="P-UCSGXM01";break}
	$successful=test-path -path "\\P-UCUKXM01\c$"
	if($successful) {$Server="P-UCUKXM01";break}
	$successful=test-path -path "\\P-UCJPXM01\c$"
	if($successful) {$Server="P-UCJPXM01";break}
	$successful=test-path -path "\\B-UCUSXJ01\c$"
	if($successful) {$Server="B-UCUSXJ01";break}
	$successful=test-path -path "\\P-UCUSXJ01\c$"
	if($successful) {$Server="P-UCUSXJ01";break}
	} while ($server -eq "")

Get-PublicFolder -identity \ -Server $server -recurse -ResultSize Unlimited | where {$_.identity -like "*$pf*"} | fl
} # function Show-JPublicFolderInfo

function Show-KVSTemp
{
<#
	.SYNOPSIS
		Returns all AD Accounts in the KVS Temp OU.

	.DESCRIPTION
		Returns Display Name, SAM Account NAme and SMTP address of all AD
		accounts in the Janus.cap\Disabled Users\KVS _Temp OU.

	.EXAMPLE
		Show-KVSTemp
		
		Description
		-----------
		Returns Display Name, SAM Account NAme and SMTP address of all AD
		accounts in the Janus.cap\Disabled Users\KVS _Temp OU.
		
	.NOTES
		Requires the Loading of the Active Directory Module.

#>

# List all the AD entries in the KVS_Temp OU
$KVST=Get-JADEntry -id "," -InOU "OU=KVS _Temp,OU=Disabled Users,OU=Janus,DC=janus,DC=cap" -properties displayname,SamAccountName,mail,modifytimestamp -pso -ErrorAction SilentlyContinue
$Display=$KVST | select displayname, SamAccountName, mail, modifytimestamp | sort -property displayname

$Display | select DisplayName, SamAccountName, modifytimestamp | ft -auto -wrap

remove-variable Display
remove-variable KVST
} # function Show-KVSTemp

function Update-JDisabledUser
{
<#
	.SYNOPSIS
		Processes and AD Account and Mailbox.

	.DESCRIPTION
		Moves the AD user to the KVS Archive OU Disables the user and disables
		ActiveSync Access.

	.PARAMETER  ID
		Enter the Name of the AD Account to process.
		NOTE: AD Name, Display Name, SamAccountName

	.EXAMPLE
		Update-JDisabledUser -ID JM27253
		
		Description
		-----------
		Move the JM27253 AD Account to the KVS Archive OU, disables it,
		hides it from the GAL and disables their Active Sync access.
		
	.NOTES
		Requires the Loading of the Exchange Module.

#>

	[CmdletBinding()]
Param(
        [Alias("Identity")]
[string]
$ID=$(Throw "You must specify an AD ID to process (e.g., JM27253)"))

[string]$CMD=""
$ADO=$null
$mailbox=$null
$logger=New-JSyslogger -dest_host "p-ucslog02.janus.cap"

if (-not ($global:ExchangeSnapIn))
	{
	[string]$errorstring="WARNING: Exchange module is not loaded, some cmdlets will not work correctly, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

new-item "\\p-ucadm01\d$\Scripts\Logs\process-KVSTemp.LOG" -type file -force

# The following Routine moves users that are in the KVS Archive OU.
$ADO=get-adentry -ID $ID -exact -InOU "OU=KVS _Temp,OU=Disabled Users,OU=Janus,DC=janus,DC=cap" -properties mail -pso

# Ensures the move will only affect one account
if ((($ADO | measure-object).count) -ne 1)
	{
	[string]$errorstring="Error retrieving AD Object " + $ID + ", exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

$mailbox=get-mailbox -id $ADO.mail
$successful=$?
if ($successful -eq $false)
	{
	Write-Output "ERROR - Could not locate mailbox for $mailbox"
	}
else
	{
	Set-JADEntry -id $ID -useraccountcontrol 514
	Set-mailbox -id $mailbox -HiddenFromAddressListsEnabled $true
	Set-CASMailbox -identity $mailbox -ActiveSyncEnabled:$False -ActiveSyncMailboxPolicy "Default Block"
	$CMD="Move-JADEntry -id $ADO -targetpath `"OU=KVS Archive,OU=Disabled Users,OU=Janus,DC=janus,DC=cap`""
	$logger.send($CMD)
	Move-JADEntry -id $ID -ou "OU=KVS Archive,OU=Disabled Users,OU=Janus,DC=janus,DC=cap"
	}
remove-variable mailbox
remove-variable ADO
remove-variable logger
remove-variable CMD
remove-variable successful
} # function Update-JDisabledUser

function Move-JADEntry
{
<#
	.SYNOPSIS
		Moves the specified User to the requested OU.

	.DESCRIPTION
		Moves the specified AD User to the requested OU.

	.PARAMETER  ID
		Enter the Name of the AD Account to be moved.
		NOTE: Compares against AD Name or Display Name

	.PARAMETER  ou
		Enter the name of the AD OU that the account will be moved to. The
		can be a fragment of the name.

	.EXAMPLE
		Move-JADEntry -ID JM27253 -OU KVSTemp
		
		Description
		-----------
		Moves the JM27253 account to the KVSTemp OU.
		
	.NOTES
		Requires the Loading of the Janus Module.

#>

	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("Identity")]
		[String]
$ID=$(Throw "You must specify an AD ID to move (e.g., JM27253)."),
        [Parameter(Position = 1, Mandatory = $true)]
        [Alias("organizationalunit")]
        [Alias("targetpath")]
		[String]
$OU=$(Throw "You must specify an AD OU to move the ID to (e.g., KVSTemp)."))

# Error checking
$OU=$OU.replace("_"," _")
$OU=$OU.replace("  _"," _")
$ADE=get-adentry -id $ID -exact -pso -properties canonicalname
$successful=$?
if (-not ($successful)) {$errorstring="Error retreiving AD Entry, exiting...";write-output $errorstring;write-error $errorstring -erroraction stop}
[string]$ADS=$ADE.adspath
write-output ("ID: $ADS")
if ($ADE.adspath -eq $null) {$ADE | fl;$errorstring="Error retreiving AD Entry, exiting...";write-output $errorstring;write-error $errorstring -erroraction stop}
(($ADE | measure-object).count)
if ((($ADE | measure-object).count) -ne 1) {$errorstring="$ID does not match a single AD Entry, exiting...";write-output $errorstring;write-error $errorstring -erroraction stop}
$OUE=get-adentry -id $OU -exact -ou -pso -properties name
$successful=$?
if (-not ($successful)) {$errorstring="Error retreiving AD OU Entry, exiting...";write-output $errorstring;write-error $errorstring -erroraction stop}
[string]$OUS=$OUE.adspath
write-output ("OU: $OUS")
if ($OUE.adspath -eq $null) {$OUE | fl;$errorstring="Error retreiving AD OU Entry, exiting...";write-output $errorstring;write-error $errorstring -erroraction stop}
if ((($OUE | measure-object).count) -ne 1) {$errorstring="$OU does not match a single AD OU, exiting...";write-output $errorstring;write-error $errorstring -erroraction stop}

# Main Script
$ADO=[ADSI]$ADE.adspath
$OUO=[ADSI]$OUE.adspath
$ADO.psbase.MoveTo($OUO)
remove-variable ADO
remove-variable OUO
remove-variable ADE
remove-variable OUE
remove-variable ADS
} # function Move-JADEntry

function Get-JMailboxCalendarDelegate
{
<#
	.SYNOPSIS
		Lists Calendars where the specified AD Account is a Delegate.

	.DESCRIPTION
		Lists Calendars where the specified AD Account is a Delegate.

	.PARAMETER  ID
		Enter the AD Name of the Delegate you are looking for.
		NOTE: Must match at least part of the AD name.

	.EXAMPLE
		Get-JMailboxCalendarDelegate -id JM27253
		
		Description
		-----------
		Lists all mailboxes that JM27253 is a delegate of.
		
	.NOTES
		Requires the Loading of the Exchange Module.

#>

	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("Identity")]
		[String]
$ID=$(Throw "You must specify an AD ID to search for delegates. (e.g., JM27253)"))

if (-not ($global:ExchangeSnapIn))
	{
	[string]$errorstring="WARNING: Exchange module is not loaded, some cmdlets will not work correctly, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Pull all mailboxes and assisgn their attributes to $CalendarSettings then we select out just the data we need.
$CalendarSettings = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Get-CalendarProcessing | Select Identity,DisplayName,ResourceDelegates -ErrorAction SilentlyContinue

# Use Where-Object to locate any Delegate that matches the query stored in $ID.

$DelegateFindings = $CalendarSettings | Where-Object {$_.ResourceDelegates -like "*$ID*"} | FT -auto -wrap
$DelegateFindings += $CalendarSettings | Where-Object {$_.BookInPolicy -like "*$ID*"} | FT -auto -wrap

# Display the results

return $DelegateFindings

# mem cleanup
remove-variable CalendarSettings -ErrorAction SilentlyContinue
remove-variable DelegateFindings -ErrorAction SilentlyContinue
} # function Get-JMailboxCalendarDelegate

function Add-JContact
{
<#
	.SYNOPSIS
		Adds a Contact to AD.

	.DESCRIPTION
		Adds an SMTP Contact to to the correct OU in AD. Correctly formats
		the Displayname as well.

	.PARAMETER  SMTP
		Enter the SMTP address that mail sent to that contact will be
		delivered to.

	.PARAMETER  first
		Enter the First (or Given) Name of the new Contact.

	.PARAMETER  last
		Enter the last (or Surname or Family) Name of the new Contact.

	.EXAMPLE
		Add-JContact -SMTP rgmoore@hotmail.com -first Randy -last Moore
		
		Description
		-----------
		Sets up a new Contact with the Display name of Moore, Randy that
		delivers to rgmoore@hotmail.com.
	
	.NOTES
		Requires the Loading of the Exchange Module.

#>

	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("externaladdress")]
		[String]
$smtp=$(Throw "You must specify an SMTP Address for this contact (e.g., rgmoore@hotmail.com)."),
        [Parameter(Position = 1, Mandatory = $true)]
        [Alias("givenname")]
        [Alias("gn")]
		[String]
$first=$(Throw "You must specify a Given Name for this contact (e.g., Randy)."),
        [Parameter(Position = 2, Mandatory = $true)]
        [Alias("surname")]
        [Alias("sn")]
		[String]
$last=$(Throw "You must specify a Surname for this contact (e.g., Moore)."))

if (-not ($global:ExchangeSnapIn))
	{
	[string]$errorstring="WARNING: Exchange module is not loaded, some cmdlets will not work correctly, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Sets up the Display Name
[string]$alias=$last + $first

#Sets up the Alias
do
	{
	$alias=$alias.replace(" ","")
	} while ($alias.contains(" "))
do
	{
	$alias=$alias.replace("`n","")
	} while ($alias.contains("`n"))
do
	{
	$alias=$alias.replace("`r","")
	} while ($alias.contains("`r"))
do
	{
	$alias=$alias.replace("`!","")
	} while ($alias.contains("`!"))

do
	{
	$alias=$alias.replace("`#","")
	} while ($alias.contains("`#"))

do
	{
	$alias=$alias.replace("`$","")
	} while ($alias.contains("`$"))

do
	{
	$alias=$alias.replace("`%","")
	} while ($alias.contains("`%"))

do
	{
	$alias=$alias.replace("`&","")
	} while ($alias.contains("`&"))

do
	{
	$alias=$alias.replace("`'","")
	} while ($alias.contains("`'"))

do
	{
	$alias=$alias.replace("`*","")
	} while ($alias.contains("`*"))

do
	{
	$alias=$alias.replace("`+","")
	} while ($alias.contains("`+"))

do
	{
	$alias=$alias.replace("`/","")
	} while ($alias.contains("`/"))

do
	{
	$alias=$alias.replace("`=","")
	} while ($alias.contains("`="))

do
	{
	$alias=$alias.replace("`?","")
	} while ($alias.contains("`?"))

do
	{
	$alias=$alias.replace("`^","")
	} while ($alias.contains("`^"))

do
	{
	$alias=$alias.replace("`_","")
	} while ($alias.contains("`_"))

do
	{
	$alias=$alias.replace("``","")
	} while ($alias.contains("``"))

do
	{
	$alias=$alias.replace("`{","")
	} while ($alias.contains("`{"))

do
	{
	$alias=$alias.replace("`(","")
	} while ($alias.contains("`("))

do
	{
	$alias=$alias.replace("`)","")
	} while ($alias.contains("`)"))

do
	{
	$alias=$alias.replace("`|","")
	} while ($alias.contains("`|"))

do
	{
	$alias=$alias.replace("`}","")
	} while ($alias.contains("`}"))

do
	{
	$alias=$alias.replace("`~","")
	} while ($alias.contains("`~"))

# Creates the Contact
New-MailContact -ExternalEmailAddress "SMTP:$smtp" -Name $alias -Alias $alias -OrganizationalUnit 'janus.cap/janus/contacts' -FirstName $first -Initials '' -LastName $last
remove-variable alias
} # function Add-JContact

function Revoke-JMailboxAccess
{
<#
	.SYNOPSIS
		Processes an account for access removal.

	.DESCRIPTION
		Hides a mailbox and removes access for the specified access.

	.PARAMETER  mb
		Enter the Display Name of a Distribution Group (aka Distribution List).
		NOTE: Will accept: GUID, ADObjectID, Distinguished name (DN),
		Domain\Account, User principal name (UPN), LegacyExchangeDN,
		SmtpAddress or Alias

	.PARAMETER  viewer
		Enter the AD name of the account that is having access revoked.

	.EXAMPLE
		Revoke-JMailboxAccess -mb ITSM -viewer JM27253
		
		Description
		-----------
		JM27253 will no longer be able to read content in the ITSM mailbox.
		ITSM will be hidden.
		
	.NOTES
		Requires the Loading of the Exchange Module.

#>

# Initialization Section

# Parameters called so that they can be passed into the script via the command line.
	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("Identity")]
        [Alias("ID")]
		[String]
$mb=$(Throw "You must specify a mailbox to revoke access from (e.g., ITSM)"),
        [Parameter(Position = 1, Mandatory = $true)]
        [Alias("user")]
		[String]
$viewer=$(Throw "You must specify an account to revoke access from (e.g., JM27253)"))

if (-not ($global:ExchangeSnapIn))
	{
	[string]$errorstring="WARNING: Exchange module is not loaded, some cmdlets will not work correctly, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Hides mailbox
set-mailbox -id $mb -HiddenFromAddressListsEnabled $true
$successful=$?
if ($successful -eq $false)
	{
	[string]$errorstring="Failed at hiding " + $mb + ", exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Grants the actual permission to the user
Remove-MailboxPermission $mb -User $viewer -AccessRights FullAccess -confirm:$false
$successful=$?
if ($successful -eq $false)
	{
	[string]$errorstring="Failed at removing permissions from " + $mb + ", exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}
} # function Revoke-JMailboxAccess

function Add-JMailboxAccess
{
<#
	.SYNOPSIS
		Grants access to a mailbox.

	.DESCRIPTION
		This un-hides the mailbox, grants access Updates the Description and
		CA9 and sends an e-mail to the Viewer that they have access and how
		to access it.

	.PARAMETER  mb
		Enter the Display Name of a Distribution Group (aka Distribution List).
		NOTE: Will accept: GUID, ADObjectID, Distinguished name (DN),
		Domain\Account, User principal name (UPN), LegacyExchangeDN,
		SmtpAddress or Alias

	.PARAMETER  viewer
		Enter the Display Name of a Distribution Group (aka Distribution List).
		NOTE: Will accept: Distinguished Name, GUID (objectGUID), Security
		Identifier (objectSid) or SAM account name  (sAMAccountName)

	.PARAMETER  expires
		Enter the date that the access is supposed to expire on.

	.EXAMPLE
		Add-JMailboxAccess -mb ITSM -viewer JM27253 -expires 12-21-2012
		
		Description
		-----------
		Grants access for the ITSM mailbox to the JM27253 account. Indicates
		that JM27253 should not have access after 12-21-2012.
		
	.NOTES
		Requires the Loading of the Exchange and ActiveDirectory Modules.

#>

# Initialization Section

# Parameters called so that they can be passed into the script via the command line.
	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("Identity")]
        [Alias("ID")]
		[String]
$mb=$(Throw "You must specify a mailbox to grant access to (e.g., ITSM)"),
        [Parameter(Position = 1, Mandatory = $true)]
        [Alias("user")]
		[String]
$viewer=$(Throw "You must specify an account to give access to (e.g., JM27253)"),
        [Parameter(Position = 2, Mandatory = $true)]
$expires=$(Throw "You must specify an Expiration date (e.g., 12-1-2012)"))

[string]$DC=Get-JDCs -GC

# Error checking
if (-not ($global:ExchangeSnapIn))
	{
	[string]$errorstring="WARNING: Exchange module is not loaded, some cmdlets will not work correctly, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

$test=get-mailbox -id $mb -domaincontroller $DC
$successful=$?
if (-not($successful))
	{
	[string]$errorstring="Mailbox $mb could not be located, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop

	}

$test=get-jadentry -id $viewer -exact -properties name -pso
$success=$?
if ((-not($successful)) -or($test.name -eq "") -or($test.name -eq $null))
	{
	[string]$errorstring="AD ID $viewer could not be located, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Unhides mailbox so that Outloook can map to it.
set-mailbox -id $mb -HiddenFromAddressListsEnabled $false
$successful=$?
if ($successful -eq $false)
	{
	[string]$errorstring="Failed at unhiding " + $mb + ", exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

[string]$ADO=(get-mailbox -id $mb).SamAccountName

# Gets the old description, so we don't lose any information
$olddescription=(get-jadentry -id $ADO -exact -properties description -pso).description
$successful=$?
if ($successful -eq $false)
	{
	[string]$errorstring="Failed at getting description for " + $ADO + ", exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Gets first and last name so they can be added to the Description
$ADN=(get-adentry -id $viewer -exact -pso -properties givenname).givenname + " " + (get-adentry -id $viewer -exact -pso -properties sn).sn
$successful=$?
if ($successful -eq $false)
	{
	[string]$errorstring="Failed at finding AD User for " + $viewer + ", exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Generates the new string that will be used in the Description
$description= $olddescription + " - Access granted to " + $ADN + " through " + $expires
Set-JADEntry -id $ADO -description $description
$successful=$?
if ($successful -eq $false)
	{
	[string]$errorstring="Failed at setting description for " + $ADO + ", exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Sets Custom Attribute 9
Set-JADEntry -id $ADO -extensionattribute9 $expires
$successful=$?
if ($successful -eq $false)
	{
	[string]$errorstring="Failed at setting CA9 for AD Object " + $dn + ", exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Grants the actual permission to the user
Add-MailboxPermission -id $mb -User $viewer -AccessRights FullAccess
$successful=$?
if ($successful -eq $false)
	{
	[string]$errorstring="Failed at seetting mailbox permissions for " + $mb + ", exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Notify Users
[string]$smtp=$((get-jadentry -id $mb -pso -properties mail).mail)
$body="$ADN,`n  The best way to access this resource is to follow this procedure:`n - Open Outlook.`n - Click on File, then click on the big Account Settings button.`n - Click on the Account settings menu option, then double click on your e-mail address.`n - Click on More settings, then the Advanced Tab.`n - Click Add then type $smtp in the window that pops up and click OK.`n - On the same Advanced Tab if Download shared folders option is checked please uncheck it.`n  - Click OK, OK, Next, Finish, Close.`n - You may have to restart Outlook for this change to take effect.`n - You are done!`nI know this procedure is a little complicated, but you only have to do it once, then it stays on your computer until you get a new computer or you get a new profile in Outlook.`n  Please feel free to contact me or anyone on my team if you have any questions or concerns."

$rcpt=(get-adentry -id $viewer -exact -properties mail -pso).mail

Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_ExchangeAdmin@janus.com" -To $rcpt -Subject "Notification - Access Granted" -Body $body
remove-variable rcpt
remove-variable body
remove-variable ADN
remove-variable ADO
remove-variable test
} # function Add-JMailboxAccess

function Set-JPassword
{
<#
	.SYNOPSIS
		Sets the password of an AD Account and unlocks it.

	.DESCRIPTION
		Sets the password of an AD Account and unlocks it.

	.PARAMETER  ID
		Enter the AD Account you want to modify.
		NOTE: Will accept: Distinguished Name, GUID (objectGUID), Security
		Identifier (objectSid) or SAM account name  (sAMAccountName)

	.PARAMETER  pass
		Indicates the password the account will have.
		NOTE: Command will fail if passwword does not conform to the password
		GPO for this account.

	.EXAMPLE
		Set-JPassword -ID JM27253 -pass 1etmein!
		
		Description
		-----------
		Sets the password for JM27253 to "1etmein!" and unlocks the account.
		
	.NOTES
		Requires the Loading of the ActiveDirectory Module.

#>

	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("Identity")]
		[String]
$ID,
        [Parameter(Position = 1, Mandatory = $true)]
        [Alias("passwd")]
        [Alias("pwd")]
        [Alias("password")]
		[String]
$pass)

[INT]$UAC=528

$ADE=Get-JADEntry -id $ID -exact -pso -properties useraccountcontrol
$ADO=[ADSI]$ADE.adspath

# Error checking
if (($ADE -eq "") -or($ADE -eq $null))
	{
	[string]$errorstring="$ID does not match a single account, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

$ADO.psbase.invoke("SetPassword",$pass)
$successful=$?
Write-Output ("Did the password reset succeed? $successful")

$ADS=(Get-ADEntry -id $ID -exact -properties useraccountcontrol -pso).useraccountcontrol
$array=$ADS.split("`n")
$UAC=$array[0]
$lockout=$UAC -band 16
if ($lockout -eq 16) {$UAC=$UAC-16}
Set-JADEntry -id $ID -useraccountcontrol $UAC
$successful=$?
Write-Output ("Did the AD Account unlock succeed? $successful")
remove-variable ADS
remove-variable array
remove-variable UAC
remove-variable lockout
remove-variable ADO
remove-variable ADE
} # function Set-JPassword

function Update-JPrimarySMTPAddress
{
<#
	.SYNOPSIS
		This script will change the primary SMTP address for a user.

	.DESCRIPTION
		This script will change the primary SMTP address for a user.
		It will add the address if it does not already exist.

	.PARAMETER  id
		Enter the Name of the Mailbox you want to modify.
		NOTE: Accepts the following data:
		* GUID
		* Distinguished name (DN)
		* Domain\Account
		* User principal name (UPN)
		* LegacyExchangeDN
		* SmtpAddress
		* Alias

	.PARAMETER  address
		Enter the SMTP address that will be the new primary
		address.

	.EXAMPLE
		Update-JPrimarySMTPAddress -id JM27253 -address randy@janusintech.com

		Description
		-----------
		Set the Primary smtp address for JM27253 to randy@janusintech.com.

	.NOTES
		Requires the Loading of the Exchange Module.

#>

	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("Identity")]
		[String]
$id,
        [Parameter(Position = 1, Mandatory = $false)]
        [Alias("mail")]
        [Alias("primarysmtpaddress")]
		[String]
$address)

# Initialization
[string]$ProxyAddresses=""
[Microsoft.Exchange.Data.SmtpProxyAddress]$smtp=$address

if (-not ($global:ExchangeSnapIn))
	{
	[string]$errorstring="WARNING: Exchange module is not loaded, some cmdlets will not work correctly, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Error checking
$mb=get-mailbox -id $ID -erroraction silentlycontinue -warningaction silentlycontinue
$success=$?
if ($success -eq $false)
	{
	[string]$errorstring="Cannot locate $id in the Address Book, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

if ((($mb | measure-object).count) -ne 1)
	{
	[string]$errorstring="$id does not match a single mailbox, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

$amb=get-mailbox -id $address -erroraction silentlycontinue -warningaction silentlycontinue

if (($amb -ne $null) -and($amb.samaccountname -ne $mb.samaccountname))
	{
	[string]$errorstring="$address exists on another Mailbox, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

if((-not($address.contains("@"))) -and(-not($address.contains("."))))
	{
	[string]$errorstring="$address is not in a valid format, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

set-mailbox -Identity $id -EmailAddressPolicyEnabled $False -erroraction silentlycontinue -warningaction silentlycontinue

$mb.EmailAddresses += $address

$mb | set-mailbox -erroraction silentlycontinue -warningaction silentlycontinue

set-mailbox -Identity $id -EmailAddressPolicyEnabled $False -PrimarySmtpAddress $smtp -erroraction silentlycontinue -warningaction silentlycontinue
$success=$?
Write-output ("Was execution successful: $success")
} # function Update-JPrimarySMTPAddress

function Move-JMailboxesInBulk
{
<#
	.SYNOPSIS
		Moves mailboxes in the specified CSV file.

	.DESCRIPTION
		Moves the mailboxes in the specified CSV. Also checkled the specified
		drive for space before moving.

	.PARAMETER  file
		Enter the file Name of the CSV file to use for Moving. The CSV should
		be formatted as follows:
		Mailbox,StoreDB,Server,LogDrive
		NOTE: Defaults to "\\p-ucadm01\d$\Scripts\bulk-mover.csv"

	.EXAMPLE
		Move-JMailboxesInBulk
		
		Description
		-----------
		Moves all the Mailboxes in \\p-ucadm01\d$\Scripts\bulk-mover.csv.
		
	.EXAMPLE
		Move-JMailboxesInBulk -CSV "c:\move.csv"
		
		Description
		-----------
		Moves all the Mailboxes in the "c:\move.csv file.
		
	.NOTES
		Requires the Loading of the Exchange Module.

#>

	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $false)]
        [Alias("filepath")]
		[String]
$file="\\p-ucadm01\d$\Scripts\bulk-mover.csv")

if (-not ($global:ExchangeSnapIn))
	{
	[string]$errorstring="WARNING: Exchange module is not loaded, some cmdlets will not work correctly, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

[string]$target=$null
[string]$db=$null
[string]$server=$null
[string]$olddb=$null
[string]$oldserver=$null
[string]$drive=$null
$csv=$null
$mb=0

$logger=New-JSyslogger -dest_host "p-ucslog02.janus.cap"

$csv=Import-Csv -Path "$file"
$successful=$?
if($successful -eq $false)
	{
	[string]$errorstring="Error Reading File " + $file + ", exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	} # if($successful -eq $false)

:line foreach ($line IN $csv)
	{
	if ($line -ne "Mailbox,StoreDB,Server,LogDrive")
	{
	Write-output ("Processing: $line")
	$target=$line.Mailbox
	$target=$target.Replace(" ","")
	$db=$line.StoreDB
	$server=$line.Server
	$drive=$line.LogDrive
	$logger.send($line)
	Write-Output "$line"
	$oldserver=(Get-Mailbox -id $target).ServerName
	write-output("Target Mailbox: $target")
	write-output("Old Server: $oldserver")
	write-output("Target Server: $server")
	$mb=(Get-MailboxStatistics -id $target).TotalItemSize.value.tobytes() / 1024 / 1024 / 1024
	Write-output ("MB Size (gb): $mb")
	$mb=$mb + 5
	$temp=Get-WMIObject Win32_Volume -computer $server | where {$_.Name -like $drive} | select freespace
	$logspace=$temp.freespace / 1GB
	Write-output ("Log Drive Free Space (gb): $logspace")
	if ($mb -gt $logspace)
		{Write-output ($mb + " is greater than " + $logspace + ", breaking out of loop...");break line}
	Write-Output "Move-MB -user $target -store $db"
	if (($oldserver -like "*ucmailukv01*") -or($server -like "*ucmailukv01*") -or($server -like "*ucmailsgv01*") -or($server -like "*ucmailsgv01*"))
		{
		if (($oldserver -like "*ucmailukv01*") -or($server -like "*ucmailukv01*"))
			{
			out-file -file "\\ucmailukv01\R$\Scripts\Logs\scheduledmove.log" -inputobject $db
			out-file -file "\\ucmailukv01\R$\Scripts\Logs\scheduledmove.log" -inputobject $target -append
			$task=Get-ScheduledTask -computername ucmailukv01 -name "Scheduled Mailbox Move"
			Start-ScheduledTask -task $task
			} else
			{
			out-file -file "\\ucmailsgv01\M$\Scripts\Logs\scheduledmove.log" -inputobject $db
			out-file -file "\\ucmailsgv01\M$\Scripts\Logs\scheduledmove.log" -inputobject $target -append
			$task=Get-ScheduledTask -computername ucmailsgv01 -name "SGMove Test"
			Start-ScheduledTask -task $task
			} # if (($oldserver -like "*ucmailukv01*") -or($server -like "*ucmailukv01*"))
		} else # if (($oldserver -like "*ucmailukv01*") -or($server -like "*ucmailukv01*") -or($server -like "*ucmailsgv01*") -or($server -like "*ucmailsgv01*"))
		{
		Move-JMailbox -user $target -store $db
		$success=$?
		do
		{
		$complete=(Get-MoveRequest -id $target).Status
		if ($success -eq $false) {$complete=$null}
		} while (($Complete -notlike "*Completed*") -and($complete -ne $null))
		} # if (($oldserver -like "*ucmailukv01*") -or($oldserver -like "*ucmailukv01*"))
	} # if ($line -ne "Mailbox,StoreDB,Server,LogDrive")
	} # :line foreach ($line IN $csv)
remove-variable csv -ErrorAction SilentlyContinue
remove-variable logger -ErrorAction SilentlyContinue
remove-variable logspace
remove-variable mb
remove-variable server
remove-variable oldserver
remove-variable target
remove-variable db
remove-variable olddb
remove-variable drive
remove-variable temp
} # function Move-JMailboxesInBulk

function Out-JExcel
{

<#
	.SYNOPSIS
		Exports the object fed to it into Excel.

	.DESCRIPTION
		Exports the object fed to it into Excel.

	.PARAMETER  property
		Accepts a piped Object from Powershell.

	.PARAMETER  raw
		Indicates to export wiothout object conversion.

	.PARAMETER  file
		Specifies the file to export to.

	.EXAMPLE
		Get-WmiObject win32_bios -computer (cat c:\servers.txt) | select __server,name,@{label='Release Date';expression={$_.ConvertToDateTime($_.releasedate)}} | Out-JExcel
		
		Description
		-----------
		Looks up the stats for the servers listed in C:\servers.txt and
		exports them to Excel.
		
	.NOTES
		Requires Excel to be installed on the computer it is run from.

	.LINK
		http://kentfinkle.com/PowershellAndExcel.aspx

	.LINK
		http://www.ithassle.nl/2010/11/error-using-excel-as-dcom-object-in-a-scheduled-task/
#>
	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true)]
		[string[]]
$property,
        [Parameter(Position = 1, Mandatory = $false)]
		[switch]
$raw,
        [Parameter(Position = 2, Mandatory = $false)]
		[string]
$file=$null)

begin
	{
# start Excel and open a new workbook
	$Excel = New-Object -Com Excel.Application
	$successful=$?
	if (-not($successful))
		{
		[string]$errorstring="WARNING: Error loading Excel`n Exiting...`n"
		$Error.add($errorstring)
		write-error ($errorstring) -erroraction stop
		}
	$Excel.visible = $True
	$Excel = $Excel.Workbooks.Add()
	$Sheet = $Excel.Worksheets.Item(1)
# initialize our row counter and create an empty hashtable
# which will hold our column headers
	$Row = 1
	$HeaderHash = @{}
	} # begin

process
	{
	if ($_ -eq $null)
		{
		return
		}
	if ($Row -eq 1)
		{
# when we see the first object, we need to build our header table
		if (-not $property)
			{
# if we haven't been provided a list of properties,
# we'll build one from the object's properties
			$property=@()
			if ($raw)
				{
				$_.properties.PropertyNames | %{$property+=@($_)
				}
			} else
				{
				$_.PsObject.get_properties() | % {$property += @($_.Name.ToString())
				}
			}
		}

	$Column = 1
	foreach ($header in $property)
		{
# iterate through the property list and load the headers into the first row
# also build a hash table so we can retrieve the correct column number
# when we process each object
		$HeaderHash[$header] = $Column
		$Sheet.Cells.Item($Row,$Column) = $header.toupper()
		$Column ++
		}
# set some formatting values for the first row
	$WorkBook = $Sheet.UsedRange
	$WorkBook.Interior.ColorIndex = 19
	$WorkBook.Font.ColorIndex = 11
	$WorkBook.Font.Bold = $True
	$WorkBook.HorizontalAlignment = -4108
	}
	$Row ++
	foreach ($header in $property)
		{
# now for each object we can just enumerate the headers, find the matching property
# and load the data into the correct cell in the current row.
# this way we don't have to worry about missing properties
# or the “ordering” of the properties
		if ($thisColumn = $HeaderHash[$header])
			{
		 	if ($raw)
				{
			        $Sheet.Cells.Item($Row,$thisColumn) = [string]$_.properties.$header
				}
				else
				{
				$Sheet.Cells.Item($Row,$thisColumn) = [string]$_.$header
				}
			}
		}
	} # process

end
	{
# now just resize the columns and we're finished
	if ($Row -gt 1) { [void]$WorkBook.EntireColumn.AutoFit() }

	if($file -eq $null)
		{
		[string]$tempp=gc env:temp
		$file=$tempp + "\Temp.xlsx"
		}

Remove-Item $file
$excel.saveas($file,"51")
$excel.quit()
remove-variable excel -ErrorAction SilentlyContinue
remove-variable sheet -ErrorAction SilentlyContinue
remove-variable workbook -ErrorAction SilentlyContinue
	} # end
} # function Out-JExcel

function Show-JActiveSyncUserInfo
{
<#
	.SYNOPSIS
		Reports ActiveSync stats for the specified user.

	.DESCRIPTION
		Reports whether the account is ActiveSync Enabled, which ActiveSync
		Mailbox Policy in in place, if the account has ActiveSync Device
		Partnership(s), ActiveSync Allowed Device IDs, whether the account is
		a member of GRP_jmobile_iDevice, whether the account is a member of
		GRP_Device_Wireless, the Date of the password was last set, the Date
		of the Last Bad Password Attempt, whether the AD account is Locked
		Out, And the Lockout Time of their Account.

	.PARAMETER  ID
		Enter the Name of the account to check.
		NOTE: Distinguished Name, GUID (objectGUID), or SAM account name
		(sAMAccountName)

 	.PARAMETER  Global
		Check Password dates against all DCs.

	.EXAMPLE
		Show-ActiveSyncUserInfo -id jm27253
		
		Description
		-----------
		Reports various stats related to ActiveSync for the JM27253 AD Account.
		
	.NOTES
		Requires the Loading of the Exchange Module.

#>

	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("Identity")]
		[String]
$ID,
        [Parameter(Position = 1, Mandatory = $false)]
		[Switch]
$Global)

if (-not ($global:ExchangeSnapIn))
	{
	[string]$errorstring="WARNING: Exchange module is not loaded, some cmdlets will not work correctly, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Main Script
[string]$string=""
$array=@()
[System.DateTime]$pwdlastset=get-date

get-adentry -id $ID -exact | out-null
$successful=$?
if ($successful -eq $false)
	{
	[string]$errorstring="Unable to locate " + $ID + " in AD, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

$ADO=get-adentry -id $ID -exact -properties displayname,memberof,msexchmobilemailboxpolicylink,msexchomaadminwirelessenable,pwdlastset,lockoutTime,useraccountcontrol,badpasswordtime,badpwdcount -pso

$name=$ADO.displayname

if ($ADO.msexchomaadminwirelessenable -eq 0) {$enabled=$false} else {$enabled=$true}

$string=$ADO.msexchmobilemailboxpolicylink
$array=$string.split("CN=")
$Policy=$Array[3]
$Policy=$Policy.replace(",","")

$string=$ADO.memberof
$JMobile=$string.contains("jmobile")

$BBMobile=$string.contains("bbmobile")

$Wireless=$string.contains("GRP_Device_Wireless")

$pwdlastset=$ADO.pwdlastset

# Initialize valuies checked against multiple DCs
$LockedOut=$False
$LockoutTime=get-date ($ADO.lockoutTime)
$BadPasswordTime=get-date ($ADO.badpasswordtime)
$BadPasswordCount=get-date ($ADO.badpwdcount)

# Check against all DCs
if ($Global)
	{
	$DCs=Get-JDCs
	} else
	{
	$DCs=(Get-JDCs -GC).name
	}

foreach ($DC IN $DCs)
	{
	$ADO=get-adentry -id $ID -exact -dc $DC -properties pwdlastset,lockoutTime,useraccountcontrol,badpasswordtime,badpwdcount -pso

	$NewLO=$ADO.useraccountcontrol -band 16
	if ($NewLO -eq 16) {$LockedOut=$True}

	$NewLockoutTime=get-date ($ADO.lockoutTime)
	if($NewLockoutTime -gt $LockoutTime) {$LockoutTime=$NewLockoutTime}

	$NewBadPasswordTime=get-date ($ADO.badpasswordtime)
	if($NewBadPasswordTime -gt $BadPasswordTime) {$BadPasswordTime=$NewBadPasswordTime}

	$NewBadPasswordCount=$ADO.badpwdcount
	if ($NewBadPasswordCount -gt $BadPasswordCount) {$BadPasswordCount=$NewBadPasswordCount}
	}

Write-Output("Name: $name")
Write-Output("ActiveSync Enabled: $enabled")
Write-Output("ActiveSync Policy: $Policy")
Write-Output("Are they in GRP_jmobile_iDevice (Required for Mobile Iron): $JMobile")
Write-Output("Are they in GRP_BBmobile (Used for BB10 Reporting): $BBMobile")
Write-Output("Are they in GRP_Device_Wireless (Required for Janus WiFi): $Wireless")
Write-Output("Newest AD Account Password Last Set onall DCs: $pwdlastset")
if ($global)
	{
	Write-Output("Are they locked out: $LockedOut")
	Write-Output("Newest AD Lockout Time across all DCs: $LockoutTime")
	Write-Output("Highest Last Bad Password Count for all DCs: $BadPasswordCount")
	Write-Output("Newest Last Bad Password Attempt on all DCs: $BadPasswordTime")
	} else
	{
	Write-Output("Are they locked out on $DCs : $LockedOut")
	Write-Output("Newest AD Lockout Time on $DCs : $LockoutTime")
	Write-Output("Highest Last Bad Password Count on $DCs : $BadPasswordCount")
	Write-Output("Newest Last Bad Password Attempt on $DCs : $BadPasswordTime")
	}
if ($ADO) {remove-variable ADO}
remove-variable DCs
} # function Show-JActiveSyncUserInfo

function Get-Janus
{
<#
	.SYNOPSIS
		Reports cmdlets Loaded from the Janus Module.

	.DESCRIPTION
		Reports cmdlets Loaded from the Janus Module.

	.EXAMPLE
		Check-Janus
		
		Description
		-----------
		Reports cmdlets Loaded from the Janus Module.
		
	.NOTES
		Requires the Loading of the Janus Module.

#>

# Main Script
write-host("`nThe following cmdlets are loaded:`nAdd-JADGroup`nAdd-JContact`nAdd-JCRDelegate`nAdd-JEWSDelegate`nAdd-JHolidays`nAdd-JMailboxAccess`nAdd-JSerenaTicket`nCompare-JFiles`nConvertFrom-JOctetToGUID`nConvertTo-JHereString`nConvertTo-JZip`nExport-JPST`nExport-PSOToCSV`nGet-JADEntry`nGet-JADIDfromSID`nGet-Janus`nGet-JConstructors`nGet-JDCs`nGet-JDirInfo`nGet-JDSACLs`nGet-JEWSDelegates`nGet-JFileEncoding`nGet-JGroupMembers`nGet-JHotFixbyDate`nGet-JInheritance`nGet-JLogonSessions`nGet-JMailboxCalendarDelegate`nGet-JMessageTracking`nGet-JModule`nGet-JPassword`nGet-JRDPSession`nGet-JScheduledTasks`nGet-JStringMatches`nGet-JWMILogonSessions`nImport-JPST`nImport-JSCOM`nMove-JADEntry`nMove-JMailbox`nMove-JMailboxesInBulk`nNew-JSyslogger`nOut-JExcel`nProtect-JParameter`nRemove-JADGroupMembership`nRemove-JCRDelegate`nRemove-JCRMeetingbyOrganizer`nRemove-JRDPSession`nRepair-JService`nRevoke-JMailboxAccess`nSend-JRemoteCMD`nSend-JTCPRequest`nSet-JADEntry`nSet-JDSACLs`nSet-JGroupMaintenanceMode`nSet-JIISPass`nSet-JInheritance`nSet-JMailboxForwarding`nSet-JMaintenanceMode`nSet-JOOFMessage`nSet-JPassword`nSet-JRandomADPassword`nShow-JActiveSyncUserInfo`nShow-JADLockoutStatus`nShow-JASDevices`nShow-JCRPermissions`nShow-JDiskInfo`nShow-JDistributionListMembers`nShow-JETPMessageStatus`nShow-JFailureAuditEvents`nShow-JIISSettings`nShow-JLyncUserInfo`nShow-JMailboxPermissions`nShow-JMessageRecipients`nShow-JOOFSettings`nShow-JPublicFolderInfo`nShow-JRDPSessions`nShow-JSearchEstimate`nShow-JServices`nShow-JStatus`nShow-KVSTemp`nStart-SOAPRequest`nUpdate-JCalendarTentativeSettings`nUpdate-JDisabledUser`nUpdate-JEWSDelegate`nUpdate-JPrimarySMTPAddress`nUpdate-KVSTempUser`n")

# Test if bpp_exchange or bps_exchange is available
$admin=get-adentry -id "_Exchange" -pso -person -properties mail,distinguishedname -ErrorAction SilentlyContinue

$extest=get-mailbox -identity $admin.mail -ErrorAction SilentlyContinue
$global:ExchangeSnapIn=$?

if (-not($global:ExchangeSnapIn))
	{
	write-output("WARNING: Exchange module is not loaded, some cmdlets will not work correctly...`n")
	write-output $error
	}

remove-variable extest -ErrorAction SilentlyContinue

# Test to see if the EWS Mail Module is loaded and working
$ewstest=Get-EWSMailMessage -ResultSize 1 -Mailbox $admin.mail -ErrorAction SilentlyContinue
$global:EWSModule=$?
if (-not ($global:EWSModule))
	{
	write-output("WARNING: EWS module is not loaded, come cmdlets will not work correctly...`n")
	write-output $error
	}

remove-variable ewstest -ErrorAction SilentlyContinue
remove-variable admin -ErrorAction SilentlyContinue

$global:JanusPSModule=$true
} # function Get-Janus

function Get-JModule
{
<#
	.SYNOPSIS
		Reports cmdlets Loaded from the specified Module.

	.DESCRIPTION
		Reports cmdlets Loaded from the specified Module.

	.PARAMETER  Module
		Enter the name of the module to report the functions.

	.EXAMPLE
		Get-JModule -Module EWSMAIL
		
		Description
		-----------
		Reports cmdlets Loaded from the EWS Module.
	

#>


	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
		[String]$module)

# Main Script
begin
	{
	
	}

process
{
$mod = Get-Module $module
$successful = $?
if ($successful -eq $false)
	{
	[string]$errorstring="Cannot Locate Module " + $module + ", exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

$mod.exportedfunctions
}

end
	{
	remove-variable mod
	}
} # function Get-JModule

function Add-JCRDelegate
{
<#
	.SYNOPSIS
		Adds the specified Delegate to the indicated Conference Room.

	.DESCRIPTION
		Add the AD ID to th Book In Policy Attribute of a AutoAccept Agent
		Settings for the Mailbox specified.

	.PARAMETER  cal
		Enter the Display Name of a Conference Room. Note, this will have no
		effect if Auto-Accept is not working for the Calendar for this Mailbox.
		NOTE: Accepts any of the following: GUID, Distinguished name (DN),
		Domain\Account, User principal name (UPN), LegacyExchangeDN,
		SmtpAddress or Alias

	.PARAMETER  del
		Enter the display name of the Mailbox that will have Delegate Access
		Granted.
		NOTE: Accepts any of the following: GUID, Distinguished name (DN),
		Domain\Account, User principal name (UPN), LegacyExchangeDN,
		SmtpAddress or Alias

	.PARAMETER  reviewer
		This switch instructs PowerShell to make the user in the -del
		paramter a Read-Only Delegate.
		NOTE: This overrides -editor

	.PARAMETER  editor
		This switch instructs PowerShell to make the user in the -del
		paramter an Editor-Only Delegate.
		NOTE: This is overridden by -editor

	.EXAMPLE
		Add-JCRDelegate -cal "!CR-ACP-04-Gemini Peak-CAP20" -del jm27253
		
		Description
		-----------
		Adds the JM27253 AD ID to the Book In Policy Attribute of the
		"!CR-ACP-04-Gemini Peak-CAP20" Calendar AutoAccept Settings.
		
	.NOTES
		Requires the Loading of the Exchange Module.

#>

	[CmdletBinding()]
Param(
        [Parameter(Position = 0, ValueFromPipelineByPropertyName=$true, Mandatory = $true)]
        [Alias("Identity")]
        [Alias("calendar")]
        [Alias("mailbox")]
        [Alias("cal")]
		[String]
$ID=$(Throw "You must specify a Calendar to modify (e.g., `"!CR-ACP-04-Gemini Peak-CAP20`")."),
        [Parameter(Position = 1, Mandatory = $true)]
        [Alias("del")]
		[String]
$User=$(Throw "You must specify adelegate (e.g., JM27253)."),
        [Parameter(Position = 2, Mandatory = $false)]
        [Alias("reveiwer")]
		[Switch]
$reviewer,
        [Parameter(Position = 2, Mandatory = $false)]
        [Alias("editer")]
        [Alias("owner")]
		[Switch]
$editor
)

if (-not ($global:ExchangeSnapIn))
	{
	[string]$errorstring="WARNING: Exchange module is not loaded, some cmdlets will not work correctly, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

[array]$newdel=@()
$User=$User.tolower()
$DC=(Get-JDCs -gc).Name
$successful=Test-Connection $DC -count 1 -ErrorAction silentlycontinue
if($successful -eq $null)
	{
	$DC="p-jcdcd05.janus.cap"
	$successful=Test-Connection $DC -count 1 -ErrorAction silentlycontinue
	if($successful -eq $null)
		{
		[string]$errorstring="Error locating a DC, exiting...`n"
		$Error.add($errorstring)
		write-error ($errorstring) -erroraction stop
		}
	}

# Error checking
$MB=Get-Mailbox -id $ID -DomainController $DC
$ID=$MB.PrimarySmtpAddress
$smtp=$MB.PrimarySmtpAddress

if ((($MB | measure-object).count) -ne 1)
	{
	[string]$errorstring=$ID + " does not match a single Mailbox, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

$DMB=Get-JADEntry -id $User -exact -pso -properties mail

if ((($DMB | measure-object).count) -ne 1)
	{
	[string]$errorstring=$User + " does not match a single Mailbox, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

$User=$DMB.mail

$CMB=Get-CalendarProcessing -id $ID -DomainController $DC
$successful=$?
if ($successful -eq $false)
	{
	Write-Output ("Calendar lookup error.`n")
	$CMB=Get-CalendarProcessing -resultzise unlimited -DomainController $DC | where {$_.Identity -like "*$ID*"}
	$successful=$?
	if ($successful -eq $false)
		{
		[string]$errorstring="Unrecoveravble error, exiting...`n"
		$Error.add($errorstring)
		write-error ($errorstring) -erroraction stop
		}
	}

if ((($CMB | measure-object).count) -ne 1)
	{
	[string]$errorstring=$ID + " does not match a single Mailbox, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Assign Delegate Permissions
if ($editor)
	{
	write-output ("Adding Editor Delegation...")
	$cdels=Get-JEWSDelegates -mb $ID -folder Calendar
	$update=$false
	foreach ($cdel IN $cdels)
		{
		if ($smtp -like $cdel.DelegatePrimarySmtpAddress) {$update=$true}
		}
	if ($update)
		{
		Update-JEWSDelegate -mb $ID -viewer $User -folder Calendar -perms Editor
		} else
		{
		Add-JEWSDelegate -mb $ID -viewer $User -folder Calendar -perms Editor
		}
	}

if ($reviewer)
	{
	write-output ("Adding Reviewer Delegation...")
	$cdels=Get-JEWSDelegates -mb $ID -folder Calendar
	$update=$false
	foreach ($cdel IN $cdels)
		{
		if ($smtp -like $cdel.DelegatePrimarySmtpAddress) {$update=$true}
		}
	if ($update)
		{
		Update-JEWSDelegate -mb $ID -viewer $User -folder Calendar -perms Reviewer
		} else
		{
		Add-JEWSDelegate -mb $ID -viewer $User -folder Calendar -perms Reviewer
		}
	}

# Save old settings
$olddel=$CMB.bookinpolicy
# write-host("Old Delegates: $olddel")
$stringOldDel=[string]$olddel
$stringOldDel=$stringOldDel.tolower()

if ($stringOldDel -like "*$User*")
	{
	[string]$errorstring=$User + " is already a Delegate, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}
$newdel=$User
foreach ($mbd IN $olddel)
	{
# write-host("Old Delegate: $MBD")
# $MBD | gm
	if ($mbd -ne $null) {$mb=(Get-JADEntry -id $($mbd.Name) -pso -exact -properties mail).mail} else {$mb = $null}
	if ($mb -ne $null) {$newdel+=$mbd}
	}

Set-CalendarProcessing -id $ID -bookinpolicy $newdel -DomainController $DC -erroraction silentlyContinue
$successful=$?
if ($successful -eq $false)
	{
	[string]$errorstring="Error adding Delegate, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

Write-Output ("Execution successful.`n")

remove-variable mb
remove-variable newdel
remove-variable olddel
remove-variable DMB
remove-variable CMB
remove-variable stringOldDel
remove-variable DC
remove-variable CDrive
remove-variable temp
remove-variable successful
remove-variable smtp
} # function Add-JCRDelegate

function Show-JOOFSettings
{
<#
	.SYNOPSIS
		Reports the Server-Side OOF Settings of the specified user.

	.DESCRIPTION
		Reports the Start, Stop, internal reply and external reply for the
		Out of Office Settings on the specified mailbox.

	.PARAMETER  PrimarySmtpAddress
		Enter the SMTP Address of the mailbox to check.

	.EXAMPLE
		Show-JOOFSettings -PrimarySmtpAddress Randy.Moore@Janus.com
		
		Description
		-----------
		Returns the OOF Settigns for Randy.Moore@Janus.com.
		
	.NOTES
		Requires the Loading of the Exchange Module.

#>
    [CmdletBinding()]
    param(
        [Parameter(Position=0, ValueFromPipelineByPropertyName=$true, Mandatory=$true)]
        [System.String]
        [Alias("Identity")]
        [Alias("ID")]
        [Alias("mail")]
        $PrimarySmtpAddress,
        [Parameter(Position=1, Mandatory=$false)]
        [System.String]
        $ver = "Exchange2007_SP1"
        )

    begin {

if (-not ($global:EWSModule))
	{
	[string]$errorstring="WARNING: EWS module is not loaded, come cmdlets will not work correctly...`nExiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	} # if (-not ($global:EWSModule))

        $sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        $user = [ADSI]"LDAP://<SID=$sid>"
    } # begin

    process {
        $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -Arg $ver
        $service.AutodiscoverUrl($PrimarySmtpAddress)    

        if($PrimarySmtpAddress -notmatch "@") {
          $PrimarySmtpAddress = (Get-Recipient $PrimarySmtpAddress).PrimarySMTPAddress.ToString()
        } # if($PrimarySmtpAddress -notmatch "@")

        $oof = $service.GetUserOofSettings($PrimarySmtpAddress)
        New-Object PSObject -Property @{
            State = $oof.State
            ExternalAudience = $oof.ExternalAudience
            StartTime = $oof.Duration.StartTime
            EndTime = $oof.Duration.EndTime
            InternalReply = $oof.InternalReply
            ExternalReply = $oof.ExternalReply
            AllowExternalOof = $oof.AllowExternalOof
            Identity = (Get-Recipient $PrimarySmtpAddress).Identity
        } # New-Object PSObject
    } # process
end
	{
remove-variable oof
remove-variable service
remove-variable user
remove-variable sid
	} # end
} # function Show-JOOFSettings

function ConvertTo-JZip {
<# 
 
.SYNOPSIS  
   Add Files or Folders to a ZIP file using the native feature in Windows. 
    
.DESCRIPTION  
   Add Files or Folders to a ZIP file using the native feature in Windows    
   Will create the zip file if it does not already exist 
    
   There is zip file support for powershell using things like DotNetZip 
   But in my opinion it is neater to use pure code (if you can)  
 
.Parameter source 
   Enter file or folder to be zipped - Use full path.
   Note: this will not work if a full path is not specified.

.Parameter zipFileName 
   Enter zip file name - if full path is not entered then parent of source is used. 

.EXAMPLE  
   ConvertTo-JZip -source "c:\scripts\notes.txt" -zipFileName "x.zip" 
 
   Description 
   ----------- 
   Add file "c:\scripts\notes.txt" to "c:\scripts\x.zip" 
 
   (Notice that ZIP file is placed in same folder as the source file) 
.EXAMPLE  
   ConvertTo-JZip -source "c:\scripts\notes.txt" -zipFileName "c:\x.zip" 
 
   Description 
   ----------- 
   Add file "c:\scripts\notes.txt" to "c:\x.zip" 
 
   (Notice that ZIP file is placed in explicit path specified) 
.EXAMPLE  
   ConvertTo-JZip -source "c:\scripts\" -zipFileName "x.zip" 
 
   Description 
   ----------- 
   Add folder "c:\scripts\" to "c:\x.zip" 
 
   (Notice that ZIP file is placed in source folder's parent folder) 
.EXAMPLE  
   ConvertTo-JZip -source "c:\scripts\" -zipFileName "x.zip" 
 
   Description 
   ----------- 
   Add folder "c:\scripts\" to "c:\x.zip" 
 
   (Notice that ZIP file is placed in source folder's parent folder) 
.EXAMPLE  
   ConvertTo-JZip c:\scripts x.zip 
 
   Description 
   ----------- 
   Add folder c:\scripts to c:\x.zip 
 
   (Notice this is a simplified command line) 
   - The ZIP file is placed in source folder's parent folder 
   - There is no trailing slash on source folder  
   - No parameter names specified 
   - No quotes because there are no spaces in path 
    
.link 
   http://gallery.technet.microsoft.com/ScriptCenter/en-us/e093e2d0-672b-402b-a33f-568610cc4bb7 
.link 
   http://msdn.microsoft.com/en-us/library/ms723207(VS.85).aspx 
 
#> 
 
[CmdletBinding()]  
   Param ( 
	[Parameter(ValueFromPipeline=$False,Mandatory=$true,HelpMessage="Enter file or folder to be zipped - Use full path")]
        [Alias("filepath")]
	[string]
	$source,
	[Parameter(ValueFromPipeline=$False,Mandatory=$true,HelpMessage="Enter zip file name  - if full path is not entered then parent of source is used")] 
	[ValidatePattern('.\.zip$')]
        [Alias("zipfilepath")]
	[string]
	$zipFileName) 

# requires -version 2.0 
       
$zipHeader=[char]80 + [char]75 + [char]5 + [char]6 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 
 
If ( (TEST-PATH $zipFileName) -eq $FALSE )  
{ 
   $iPathEnd = ($zipFileName.lastindexof("\")) 
   if ($iPathEnd -le 0)  
   { 
# $zipFileName is not a full file path 
      If ( (TEST-PATH $source) -eq $FALSE )  
      { 
         return "Error reading Source: "+$source 
      } 
      else 
      { 
# Find parent folder of source to prepend to $zipFileName 
         if ([string](get-childitem $source | Where-Object { $_.PSIsContainer }) -eq "") 
         { 
# Source is a file 
# Extract just the folder path 
            $iPathEnd=($source.lastindexof("\")) 
            $prependPath=$source.substring(0,$iPathEnd+1)             
         } 
         else 
         { 
# Source is a folder 
# Add a trailing slash for consistancy  
            if ($source.substring($source.length-1,1) -eq "\") 
            { 
               $prependPath=$source 
            } 
            else 
            { 
               $prependPath=$source+"\" 
            } 
# Find parent folder  
            $iPathEnd=(($source.substring(0,$source.length-1)).lastindexof("\")) 
            $prependPath=$source.substring(0,$iPathEnd+1) 
         } 
      } 
 
      $zipFileName= $prependPath+$zipFileName 
   } 
} 
else 
{ 
   Remove-Item $zipFileName 
} 
Add-Content $zipFileName -value $zipHeader 
 
$ExplorerShell=NEW-OBJECT -comobject 'Shell.Application' 
$SendToZip=$ExplorerShell.Namespace($zipFileName).CopyHere($source)

# return "Created ZIP file "+$zipFileName 
return $SendToZip
} # function ConvertTo-JZip

function Update-JCalendarTentativeSettings
{
<#
	.SYNOPSIS
		Manages all settings for Calendar processing.

	.DESCRIPTION
		In defauilt mode it will enable "AutoUpdate" for Calendar processing. It
		also sets both Tentaive flags to False and allows for the processing of
		external Meeting invites.

	.PARAMETER  id
		Enter the Name of the Mailbox to modify.
		NOTE: Accepts the following data:
		* GUID
		* Distinguished name (DN)
		* Domain\Account
		* User principal name (UPN)
		* LegacyExchangeDN
		* SmtpAddress
		* Alias

	.PARAMETER  revert
		In this mode, it performs the following Account updates: enable "AutoUpdate"
		for Calendar processing. It also sets both Tentaive flags to True and allows
		for the processing of external Meeting invites.

	.EXAMPLE
		Update-CalendarTentativeSettings -id Randy.Moore@janus.com
		
		Description
		-----------
		Enables "AutoUpdate" for Calendar processing. It also sets both Tentaive
		flags to False and allows for the processing of external Meeting invites.
		
	.EXAMPLE
		Update-CalendarTentativeSettings -id Randy.Moore@janus.com -revert
		
		Description
		-----------
		Enable "AutoUpdate" for Calendar processing. It also sets both Tentaive
		flags to True and allows for the processing of external Meeting invites.
		
	.NOTES
		Requires the Loading of the Exchange Module.

#>
	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("Identity")]
        [Alias("mailbox")]
		[String]
	$id,
        [Parameter(Position = 1, Mandatory = $false)]
		[switch]
	$revert)

if (-not ($global:ExchangeSnapIn))
	{
	[string]$errorstring="WARNING: Exchange module is not loaded, some cmdlets will not work correctly, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Error Checking
$test=get-mailbox -id $id
$successful=$?
if ($test -eq $null) {$successful=$false}
if ($test -eq "") {$successful=$false}
if (($test | measure-object).count -ne 1) {$successful=$false}

if ($successful -eq $false) 
	{
	[string]$errorstring="Error locating Mailbox " + $id + ", exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

if ($reveert)
	{
	Set-CalendarProcessing -Identity $id -AutomateProcessing AutoUpdate -AddNewRequestsTentatively $true -ProcessExternalMeetingMessages $true -TentativePendingApproval $true  -Confirm:$false
	$successful=$?
	}
	else
	{
	Set-CalendarProcessing -Identity $id -AutomateProcessing AutoUpdate -AddNewRequestsTentatively $false -ProcessExternalMeetingMessages $true -TentativePendingApproval $false  -Confirm:$false
	$successful=$?
	}
write-host("Was Exchange Tentative Setting update successful: $successful")
remove-variable test
} # function Update-JCalendarTentativeSettings

function Send-JRemoteCMD
{
<#
	.SYNOPSIS
		Opens and runs a command in a CMD window on a remote computer.

	.DESCRIPTION
		Uses WMI to open a Command Prompt Session and execute the specified
		command.

	.PARAMETER  sys
		Enter the system name you want the command to run on.

	.PARAMETER  command
		Enter the Command to execute.

	.EXAMPLE
		Execute-RemoteCMD -sys p-ucadm02 -command "dir c:\ > c:\dir.txt"
		
		Description
		-----------
		Runs the "dir c:\ > c:\dir.txt" command from a Command Prompt on P-UCADM02.
		
	.EXAMPLE
		Execute-RemoteCMD -sys p-ucadm01 "C:\Windows\System32\WindowsPowerShell\v1.0\powwershell.exe -file `"D:\Scripts\Scheduled Tasks\BAS-Fix.ps1`""
		
		Description
		-----------
		Runs the BAS-Fix.ps1 Powershell Script on p-ucadm01.
		
	.NOTES
		Requires the Loading of the Janus Module.

#>
	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [Alias("computer")]
		[String]
$sys,
        [Parameter(Position = 1, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [Alias("scriptblock")]
		[String]
$command)

begin
	{ 
	[string]$cmd = "CMD.EXE /C " + $command
	}

process
	{
	$newproc = Invoke-WmiMethod -class Win32_process -name Create -ArgumentList ($cmd) -ComputerName $sys
	if ($newproc.ReturnValue -ne 0 )
		{
		Write-Output "Command $($command) failed to execute Sucessfully on $($sys)"
		}
	}
End
	{
	Write-Output "Remote Execution Successful."
remove-variable newproc
	}
} # function Send-JRemoteCMD

function Repair-JService
{
<#
	.SYNOPSIS
		Control all aspects of a service on a local or remove computer.

	.DESCRIPTION
		This allows you to stop or start services on any computer. It also allows
		you to change the credentials used by that service.

	.PARAMETER  svc
		Enter the Name of the Service being controlled.

	.PARAMETER  sys
		Enter the Computer that the controlled service is on.
		NOTE: If this is not specified, the local machine is used.

	.PARAMETER  pass
		Enter the new password that the service will use. This will not be
		changed if this is not specified.

	.PARAMETER  stop
		This switch instructs Powershell to stop the service only. The service will
		be restarted otherwise.

	.PARAMETER  start
		This switch instruct Powershell to start the service only. The service will
		be restarted otherwise.

	.EXAMPLE
		Repair-JService -svc spooler -sys p-ucadm02 -pass Iz34$
		
		Description
		-----------
		This will stop the spooler service on p-ucadm02, reset the password used
		by that service and start it again.
		
	.NOTES
		Requires WMI permissions

#>
	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("service")]
        [Alias("Identity")]
        [Alias("ID")]
		[string]
$svc,
        [Parameter(Position = 1, Mandatory = $false)]
        [Alias("computer")]
        [Alias("system")]
		[String]
$sys=$env:computername,
        [Parameter(Position = 2, Mandatory = $false)]
        [Alias("password")]
        [Alias("passwd")]
        [Alias("pwd")]
		[System.String]
$pass=$null
)

$start=$false
$stop=$false

$filter = 'Name=' + "'" + $svc + "'" + ''
$service = Get-WMIObject -ComputerName $sys -namespace "root\cimv2" -class Win32_Service -Filter $filter -EnableAllPrivileges  -Authentication 6

switch ($service.StartMode)
	{
	"Manual"
	{
	switch ($service.State)
	{
	"Stopped"
	{
	$start=$false
	$stop=$false
	} # "Stopped"
	"Running"
	{
	$start=$true
	$stop=$true
	} # "Running"
	default
	{
	$start=$true
	$stop=$false
	} # default
	} # switch ($service.State)
	} # "Manual"
	"Auto"
	{
	switch ($service.State)
	{
	"Stopped"
	{
	$start=$true
	$stop=$false
	} # "Stopped"
	"Running"
	{
	$start=$true
	$stop=$true
	} # "Running"
	default
	{
	$start=$true
	$stop=$false
	} # default
	} # switch ($service.State)
	} # "Auto"
	"Disabled"
	{
	switch ($service.State)
	{
	"Stopped"
	{
	$start=$false
	$stop=$false
	} # "Stopped"
	"Running"
	{
	$start=$false
	$stop=$false
	} # "Running"
	default
	{
	$start=$false
	$stop=$false
	} # default
	} # switch ($service.State)
	} # "Disabled"
	default
	{
	switch ($service.State)
	{
	"Stopped"
	{
	$start=$true
	$stop=$false
	} # "Stopped"
	"Running"
	{
	$start=$true
	$stop=$true
	} # "Running"
	default
	{
	$start=$false
	$stop=$false
	} # default
	} # switch ($service.State)
	} # default
	} # switch ($service.StartMode)

if ($stop -eq $true)
	{
	Write-Output ("Stopping service $svc on server $sys")
	$temp=$service.StopService()
	$successful=$temp.ReturnValue
	if ($successful -ne 0)
		{
		[string]$errorstring="Error stopping service " + $svc + " on system " + $sys + ", exiting...`n"
		$Error.add($errorstring)
		write-error ($errorstring) -erroraction stop
		}

	do
		{
		$test=$null
		Write-Output("Waiting for service $svc to stop on server $sys...")
		$temp=Get-WMIObject -ComputerName $sys -namespace "root\cimv2" -class Win32_Service -Filter $filter -EnableAllPrivileges  -Authentication 6
		[string]$test=$temp.State
		} while ($test -ne "Stopped")
	}

if ($pass -ne "")
	{
	Write-Output ("Resetting password $svc on server $sys")
	$temp=$service.Change($null,$null,$null,$null,$null,$null,$null,$pass,$null,$null,$null)
	$successful=$temp.ReturnValue
	if ($successful -ne 0)
		{
		[string]$errorstring="Error changing username and password for service " + $svc + " on system " + $sys + ", exiting...`n"
		$Error.add($errorstring)
		write-error ($errorstring) -erroraction stop
		}
	}

if ($start -eq $true)
	{
	Write-Output ("Starting service $svc on server $sys")
	$temp=$service.StartService()
	$successful=$temp.ReturnValue
	if ($successful -ne 0)
		{
		[string]$errorstring="Error starting service " + $svc + " on system " + $sys + ", exiting...`n"
		$Error.add($errorstring)
		write-error ($errorstring) -erroraction stop
		}
	}
remove-variable temp
remove-variable service
} # function Repair-JService

function Set-JIISPass
{
<#
	.SYNOPSIS
		Sets the password for all IIS Application Pools and directory
		authenticationthat use the specified account.

	.DESCRIPTION
		Searches the specified server for IIS Application Pools and directory
		Authentication that use the specified account. Resets the password on
		those Application Pools.

	.PARAMETER  sys
		This is the name of the computer where the IIS Application Pool is
		located.

	.PARAMETER  svcact
		This is the Service Account that PowerShell will search for.

	.PARAMETER  pass
		This is the password that PowerShell will reset to.

	.EXAMPLE
		Set-JIISPass -sys p-complyfs01 -svcact JANUS_CAP\bpp_exspaudit -pass Iz34$
		
		Description
		-----------
		Connects to p-complyfs01, locates all IIS Appl;ication Pools that use
		JANUS_CAP\bpp_exspaudit as a service account and resets the password they
		use to "Iz34$."
		
	.NOTES
		Requires WMI Permissions.

#>
	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("computer")]
        [Alias("system")]
		[String]
$sys,
        [Parameter(Position = 1, Mandatory = $true)]
        [Alias("Identity")]
        [Alias("ID")]
        [Alias("serviceaccount")]
		[String]
$svcact,
        [Parameter(Position = 2, Mandatory = $true)]
        [Alias("password")]
        [Alias("passwd")]
        [Alias("pwd")]
		[String]
$pass)

$tmp = $NULL

# Get all of the IIS Classes in WMI
$Classes=gwmi -namespace "root\microsoftiisv2" -computername $sys -list -authentication 6 | where {$_.name -like "*IIS*"}
foreach ($Class IN $Classes)
	{
	$ClassInstances = Get-WMIObject -class $($Class.name) -namespace "root\microsoftiisv2" -computer $sys -authentication 6
	foreach($Instance IN $ClassInstances)
		{
		if ($Instance.AnonymousUserName -like "*$svcact*")
			{
			$name=$Instance.RelativePath
			Write-Output ("Updating $name on server $sys...")
			$Instance.PsObject.Properties["AnonymousUserPass"].Value = "$pass"
			$Instance.Put()
			}
		if ($Instance.LogOdbcUserName -like "*$svcact*")
			{
			$name=$Instance.RelativePath
			Write-Output ("Updating $name on server $sys...")
			$Instance.PsObject.Properties["LogOdbcPassword"].Value = "$pass"
			$Instance.Put()
			}
		if ($Instance.UNCUserName -like "*$svcact*")
			{
			$name=$Instance.RelativePath
			Write-Output ("Updating $name on server $sys...")
			$Instance.PsObject.Properties["UNCPassword"].Value = "$pass"
			$Instance.Put()
			}
		if ($Instance.WAMUserName -like "*$svcact*")
			{
			$name=$Instance.RelativePath
			Write-Output ("Updating $name on server $sys...")
			$Instance.PsObject.Properties["WAMUserPass"].Value = "$pass"
			$Instance.Put()
			}
		} # foreach($Instance IN $ClassInstances)
	} # foreach ($Class IN $Classes)

Execute-RemoteCMD -sys $sys -command IISReset
$successful=$?
if ($successful -eq $false)
	{
	[string]$errorstring="Error restarting IIS, the change will not take affect until IIS is restarted using the IISReset /RESTART command, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

remove-variable Classes -ErrorAction SilentlyContinue
remove-variable ClassInstances -ErrorAction SilentlyContinue
} # function Set-JIISPass

function Compare-JFiles
{
<#
	.SYNOPSIS
		Compares the contents of two files.

	.DESCRIPTION
		Does a line-by-line comparison of the contents of two files (e.g.,
		compare-files "d:\scripts\x.txt" "d:\scripts\y.txt").

	.EXAMPLE
		Compare-JFiles "d:\scripts\x.txt" "d:\scripts\y.txt"
		
		Description
		-----------
		Shows an on-screen list of the differences between "d:\scripts\x.txt"
		and "d:\scripts\y.txt".
		

#>

# Error Checking
if ($args[0] -eq "")
	{
	[string]$errorstring="Missing File Name, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}
if ($args[1] -eq "")
	{
	[string]$errorstring="Missing File Name, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}
$test=Get-Content $args[0]
$successful=$?
if ($successful -eq $false)
	{
	[string]$errorstring="Error opening file: " + $args[0] + ", exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}
$test=Get-Content $args[1]
$successful=$?
if ($successful -eq $false)
	{
	[string]$errorstring="Error opening file: " + $args[1] + ", exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Main script
Compare-Object $(Get-Content $args[0]) $(Get-Content $args[1])
} # function Compare-JFiles

function Show-JIISSettings
{
<#
	.SYNOPSIS
		Reports all IIS settings.

	.DESCRIPTION
		Reports all IIS v7 settings of the specified systems.

	.PARAMETER  sys
		The name of the system to report on.

	.EXAMPLE
		Show-JIISSettings -sys p-Lynccwa01
		
		Description
		-----------
		Reports all IIS v7 settings of p-Lynccwa01.
		
	.NOTES
		Requires WMI Access.

#>
	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("computer")]
        [Alias("system")]
        [Alias("Identity")]
        [Alias("ID")]
		[String]
$sys=$env:computername)

write-output ("IIsACE Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsACE
write-output ("IIsAdminACL Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsAdminACL
write-output ("IIsApplicationPool Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsApplicationPool
write-output ("IIsApplicationPools Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsApplicationPools
write-output ("IIsApplicationPoolSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsApplicationPoolSetting
write-output ("IIsApplicationPoolsSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsApplicationPoolsSetting
write-output ("IIsCertMapper Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsCertMapper
write-output ("IIsCertMapperSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsCertMapperSetting
write-output ("IIsCompressionScheme Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsCompressionScheme
write-output ("IIsCompressionSchemes Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsCompressionSchemes
write-output ("IIsCompressionSchemeSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsCompressionSchemeSetting
write-output ("IIsCompressionSchemesSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsCompressionSchemesSetting
write-output ("IIsComputer Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsComputer
write-output ("IIsComputerSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsComputerSetting
write-output ("IIsCustomLogModule Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsCustomLogModule
write-output ("IIsCustomLogModuleSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsCustomLogModuleSetting
write-output ("IIsDirectory Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsDirectory
write-output ("IIsFilter Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsFilter
write-output ("IIsFilters Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsFilters
write-output ("IIsFilterSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsFilterSetting
write-output ("IIsFiltersSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsFiltersSetting
write-output ("IIsFtpInfo Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsFtpInfo
write-output ("IIsFtpInfoSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsFtpInfoSetting
write-output ("IIsFtpServer Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsFtpServer
write-output ("IIsFtpServerSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsFtpServerSetting
write-output ("IIsFtpService Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsFtpService
write-output ("IIsFtpServiceSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsFtpServiceSetting
write-output ("IIsFtpVirtualDir Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsFtpVirtualDir
write-output ("IIsFtpVirtualDirSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsFtpVirtualDirSetting
write-output ("IIsImapInfo Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsImapInfo
write-output ("IIsImapInfoSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsImapInfoSetting
write-output ("IIsImapRoutingSource Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsImapRoutingSource
write-output ("IIsImapRoutingSourceSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsImapRoutingSourceSetting
write-output ("IIsImapServer Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsImapServer
write-output ("IIsImapServerSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsImapServerSetting
write-output ("IIsImapService Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsImapService
write-output ("IIsImapServiceSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsImapServiceSetting
write-output ("IIsImapSessions Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsImapSessions
write-output ("IIsImapSessionsSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsImapSessionsSetting
write-output ("IIsImapVirtualDir Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsImapVirtualDir
write-output ("IIsImapVirtualDirSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsImapVirtualDirSetting
write-output ("IIsIPSecuritySetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsIPSecuritySetting
write-output ("IIsLogModule Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsLogModule
write-output ("IIsLogModules Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsLogModules
write-output ("IIsLogModuleSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsLogModuleSetting
write-output ("IIsLogModulesSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsLogModulesSetting
write-output ("IIsMimeMap Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsMimeMap
write-output ("IIsMimeMapSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsMimeMapSetting
write-output ("IIsNntpExpiration Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpExpiration
write-output ("IIsNntpExpirationSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpExpirationSetting
write-output ("IIsNntpExpire Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpExpire
write-output ("IIsNntpExpireSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpExpireSetting
write-output ("IIsNntpFeed Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpFeed
write-output ("IIsNntpFeeds Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpFeeds
write-output ("IIsNntpFeedSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpFeedSetting
write-output ("IIsNntpFeedsSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpFeedsSetting
write-output ("IIsNntpGroups Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpGroups
write-output ("IIsNntpGroupsSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpGroupsSetting
write-output ("IIsNntpInfo Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpInfo
write-output ("IIsNntpInfoSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpInfoSetting
write-output ("IIsNntpRebuild Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpRebuild
write-output ("IIsNntpRebuildSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpRebuildSetting
write-output ("IIsNntpServer Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpServer
write-output ("IIsNntpServerSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpServerSetting
write-output ("IIsNntpService Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpService
write-output ("IIsNntpServiceSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpServiceSetting
write-output ("IIsNntpSessions Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpSessions
write-output ("IIsNntpSessionsSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpSessionsSetting
write-output ("IIsNntpVirtualDir Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpVirtualDir
write-output ("IIsNntpVirtualDirSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsNntpVirtualDirSetting
write-output ("IIsObject Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsObject
write-output ("IIsObjectSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsObjectSetting
write-output ("IIsPop3Info Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsPop3Info
write-output ("IIsPop3InfoSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsPop3InfoSetting
write-output ("IIsPop3RoutingSource Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsPop3RoutingSource
write-output ("IIsPop3RoutingSourceSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsPop3RoutingSourceSetting
write-output ("IIsPop3Server Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsPop3Server
write-output ("IIsPop3ServerSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsPop3ServerSetting
write-output ("IIsPop3Service Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsPop3Service
write-output ("IIsPop3ServiceSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsPop3ServiceSetting
write-output ("IIsPop3Sessions Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsPop3Sessions
write-output ("IIsPop3SessionsSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsPop3SessionsSetting
write-output ("IIsPop3VirtualDir Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsPop3VirtualDir
write-output ("IIsPop3VirtualDirSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsPop3VirtualDirSetting
write-output ("IIsSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsSetting
write-output ("IIsSmtpAlias Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsSmtpAlias
write-output ("IIsSmtpAliasSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsSmtpAliasSetting
write-output ("IIsSmtpDL Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsSmtpDL
write-output ("IIsSmtpDLSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsSmtpDLSetting
write-output ("IIsSmtpDomain Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsSmtpDomain
write-output ("IIsSmtpDomainSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsSmtpDomainSetting
write-output ("IIsSmtpInfo Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsSmtpInfo
write-output ("IIsSmtpInfoSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsSmtpInfoSetting
write-output ("IIsSmtpRoutingSource Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsSmtpRoutingSource
write-output ("IIsSmtpRoutingSourceSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsSmtpRoutingSourceSetting
write-output ("IIsSmtpServer Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsSmtpServer
write-output ("IIsSmtpServerSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsSmtpServerSetting
write-output ("IIsSmtpService Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsSmtpService
write-output ("IIsSmtpServiceSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsSmtpServiceSetting
write-output ("IIsSmtpSessions Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsSmtpSessions
write-output ("IIsSmtpSessionsSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsSmtpSessionsSetting
write-output ("IIsSmtpUser Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsSmtpUser
write-output ("IIsSmtpUserSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsSmtpUserSetting
write-output ("IIsSmtpVirtualDir Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsSmtpVirtualDir
write-output ("IIsSmtpVirtualDirSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsSmtpVirtualDirSetting
write-output ("IIsStructuredDataClass Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsStructuredDataClass
write-output ("IIsUserDefinedLogicalElement Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsUserDefinedLogicalElement
write-output ("IIsUserDefinedSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsUserDefinedSetting
write-output ("IIsWebDirectory Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsWebDirectory
write-output ("IIsWebDirectorySetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsWebDirectorySetting
write-output ("IIsWebFile Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsWebFile
write-output ("IIsWebFileSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsWebFileSetting
write-output ("IIsWebInfo Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsWebInfo
write-output ("IIsWebInfoSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsWebInfoSetting
write-output ("IIsWebServer Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsWebServer
write-output ("IIsWebServerSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsWebServerSetting
write-output ("IIsWebService Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsWebService
write-output ("IIsWebServiceSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsWebServiceSetting
write-output ("IIsWebVirtualDir Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsWebVirtualDir
write-output ("IIsWebVirtualDirSetting Objects")
Get-WMIObject -namespace "root\microsoftiisv2" -computer $sys -authentication 6 -class IIsWebVirtualDirSetting
} # function Show-JIISSettings

function ConvertTo-JHereString
{
<#
	.SYNOPSIS
		This is a simple function, convert a here-string to an array. You can
		comment out any items you do not want included in the output array.

	.DESCRIPTION
		This is a simple function, convert a here-string to an array. You can
		comment out any items you do not want included in the output array.

	.PARAMETER  HString
		Enter the HereString or HereString Variable you want converted..

	.EXAMPLE
$HS=@"
SERVERI01
SERVERI02
SERVERI05
#SERVERI06
"@

ConvertTo-JHereString $HS

Output:

SERVERI01
SERVERI02
SERVERI05

* Notice SERVERI06 is not included since it has the #

The Output is an array, so you can select individual items or a range etc.

	.EXAMPLE
ConvertTo-JHereString -HString $HString

	.EXAMPLE
ConvertTo-JHereString | Convert-HString

	.EXAMPLE
$HString | ConvertTo-JHereString | Test-Online | Get-QADComputer |
Select-Object Name,OSName,OSVersion |
Export-Csv -Path C:\ps\OSVersions.csv -NoTypeInformation

* Test-Online is available here:

http://gallery.technet.microsoft.com/scriptcenter/2789c120-48cc-489b-8d61-c1602e954b24

	.NOTES
		Requires -Version 2.0.

#>

[CmdletBinding()]
 Param
   (
    [Parameter(Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
    [String]
$HString
   ) # End Param

Begin
{
    Write-Verbose "Converting Here-String to Array"
} # Begin
Process
{
    $HString -split "`n" | ForEach-Object {

        $line = $_.trim()
        if ($line -notmatch "#")
            {
                $line
            }
        }
} # Process
End
{
# Nothing to do here.
} # End

} # function ConvertTo-JHereString

function Protect-JParameter
{
# Reads $global:val removes illegal characters and then updates $global:val
	[string]$return = ""
	$input=$global:val
	$return=[regex]::Replace($input,"[^a-z\x20_\!\#\$\&\(\)\*\,\.\/\:\?\@\\\{\}\'\-\+\=A-Z0-9]","")
	$successful = $?
	if(-not $successful) {write-output $error;Write-error "RegEx Error, exiting...";exit}
	$global:val = $return
remove-variable return
} # function Protect-JParameter

function Get-JADEntry
{
<#
.DESCRIPTION
     Name: Get-ADEntry
     Version: 0.9
     AUTHOR: Dave M
     DATE  : 8/29/2011

.SYNOPSIS
     Retrieves and unwraps a distinct set of properties using
     the [adsisearcher] method.

.DESCRIPTION
     Retrieves and unwraps a distinct set of properties using
     the [adsisearcher] method.

.PARAMETER  ID
	Enter the Name of the AD Entry you are looking for. No attempt is made to
	find near misses. some queries will seach both Name and DisplayName.
	NOTE: AD Name, Display Name or SMTP Address

.PARAMETER  exact
	Use this switch if you do not want Get-ADEntry to perform a fuzzy search.

.PARAMETER  properties
	Enter the AD Properties to return on the requested AD Entries.
	NOTE: E.G,  "name","DisplayName","EMail"

.PARAMETER  InOU
	Enter the OU that you want to perform the search in. By Default, it will search
	in "dc=janus, dc=cap".
	NOTE: Must be in Distinguished name or "LDAP://dc=janus, dc=cap" format.

.PARAMETER  norecurse
	Use this switch to prevent the search from including Sub-OUs.

.PARAMETER  dc
	Specify which DC to use.

.PARAMETER  person
	Use this switch to only return User Entries. Does not combine with OU, computer,
	nmb, Lync, DL or group Switches. If more than one is required, do not use any
	of the following switches: OU, computer, nmb, Lync, DL, group, extension, user, schema.
	NOTE: Superceded by: OU, computer, Lync, DL, group, extension, user, schema

.PARAMETER  group
	Use this switch to only return Group Entries. Does not combine with OU, computer,
	nmb, Lync, DL or user Switches. If more than one is required, do not use any
	of the following switches: OU, computer, nmb, Lync, DL, group, extension, user, schema.
	NOTE: Superceded by: OU, computer, Lync, DL, Supercedes: extension, user, schema

.PARAMETER  DL
	Use this switch to only return Mail-enabled Group Entries. Does not combine with OU, computer,
	nmb, Lync, group or user Switches. If more than one is required, do not use any
	of the following switches: OU, computer, nmb, Lync, DL, group, extension, user, schema.
	NOTE: Superceded by: OU, computer, Lync, Supercedes: group, extension, user, schema

.PARAMETER  computer
	Use this switch to only return Computer Entries. Does not combine with OU,
	nmb, Lync, DL, group or user Switches. If more than one is required, do not use any
	of the following switches: OU, computer, nmb, Lync, DL, group, extension, user, schema.
	NOTE: Superceded by: OU, Supercedes: Lync, DL, group, extension, user, schema

.PARAMETER  OU
	Use this switch to only return OU Entries. Does not combine with computer,
	nmb, Lync, DL, group or user Switches. If more than one is required, do not use any
	of the following switches: OU, computer, nmb, Lync, DL, group, extension, user, schema.
	NOTE:  Supercedes: computer, Lync, DL, group, extension, user, schema

.PARAMETER  Lync
	Use this switch to only return Lync/Lync-enabled Entries. Does not combine with OU, computer,
	nmb, DL, group or user Switches. If more than one is required, do not use any
	of the following switches: OU, computer, nmb, Lync, DL, group, extension, user, schema.
	NOTE: Superceded by: OU, computer, Supercedes: DL, group, extension, user, schema

.PARAMETER  Schema
	Use this switch to only search for Schema AD Entries. If more than one is required, do not use any
	of the following switches: OU, computer, nmb, Lync, DL, group, extension, user, schema.
	NOTE: Superceded by: OU, computer, Supercedes: Lync, DL, group, extension, user, Description

.PARAMETER  Description
	Use this switch to only search the Description Attribute. If more than one is required, do not use any
	of the following switches: OU, computer, nmb, Lync, DL, group, extension, user, schema.
	NOTE: Superceded by: OU, computer, Supercedes: Lync, DL, group, extension, user

.PARAMETER  pso
	Specifies a PSObject should be returned.

.PARAMETER  d
	Use this switch to output debug data for scripting troubleshooting of this
	command.

.EXAMPLE
	Get-JADEntry -id jm27253 -exact
	
	Description
	-----------
	Finds every entry type (except OUs) with the name of "jm27253"

.EXAMPLE
	Get-JADEntry -id jm27253
	
	Description
	-----------
	Finds every entry type (except OUs) with the name or displayname like "*jm27253*"

.EXAMPLE
	Get-JADEntry -id user -OU
	
	Description
	-----------
	Finds every OU with the name like "*user*"

.EXAMPLE
	Get-JADEntry -id "p-ucbes0" -computer
	
	Description
	-----------
	Finds every computer entry with the name like "*p-ucbes0*"

.EXAMPLE
	Get-JADEntry -id intech -nmb
	
	Description
	-----------
	Finds every contact, mail-enabled (but not mailbox-enabled) AD acount, mail-
	enabled public folder, etc. with the name or displayname like "*intech*"

.EXAMPLE
	Get-JADEntry -id randy.moore@janus.com -Lync
	
	Description
	-----------
	Finds every entry type (except OUs) with the SIP address like
	"*randy.moore@janus.com*"

.EXAMPLE
	Get-JADEntry -id messaging -DL
	
	Description
	-----------
	Finds every mail-enabled group with the name or displayname like "*messaging*"

.EXAMPLE
	Get-JADEntry -id adm_ex -group
	
	Description
	-----------
	Finds every AD group with the name or displayname like "*adm_ex*"

.EXAMPLE
	Get-JADEntry -id jm27253 -user
	
	Description
	-----------
	Finds every User entry  with the name or displayname like "*jm27253*"

.NOTES
	Uses WMI and .net framework (both of which are required to run Powershell). Uses
	the current credentials.

.LINK
    http://gallery.technet.microsoft.com/scriptcenter/Extract-arbitrary-list-of-6f59d3b4#
#>
# requires -version 2
	[CmdletBinding()]
Param(
[Parameter(Position = 0,Mandatory=$true,ValueFromPipeline=$true,HelpMessage='An AD name or Display Name to search for.')]
        [Alias("Identity")]
        [Alias("Computer")]
        [Alias("User")]
        [Alias("SAMAccountName")]
        [Alias("DisplayName")]
        [Alias("Mail")]
        [Alias("PrimarySMTPAddress")]
        [Alias("EmailAddress")]
        [Alias("WindowsEmailAddress")]
	[string]
$ID,
[Parameter(Position = 1,Mandatory=$false,HelpMessage='Use this switch if you do not want Get-ADEntry top do a fuzzy search.')]
	[switch]
$exact,
[Parameter(Position = 2,Mandatory=$false,HelpMessage='An array of property names')]
	[string[]]
$properties=@("*",""),
[Parameter(Position = 3,Mandatory=$false,HelpMessage='DN or LDAP path of container to search in')]
	[string]
$InOU="LDAP://dc=janus, dc=cap",
[Parameter(Position = 4,Mandatory=$false,HelpMessage='This switch instructs Powershell not to search sub-OUs.')]
	[switch]
$norecurse,
[Parameter(Position = 5,Mandatory=$false,HelpMessage='This switch instructs Powershell to only return User Objects.')]
	[string]
$dc,
[Parameter(Position = 6,Mandatory=$false,HelpMessage='This switch instructs Powershell to only return User Objects.')]
	[switch]
$person,
[Parameter(Position = 7,Mandatory=$false,HelpMessage='This switch instructs Powershell to only return Group Objects.')]
	[switch]
$group,
[Parameter(Position = 8,Mandatory=$false,HelpMessage='This switch instructs Powershell to only return Distribution Group Objects.')]
	[switch]
$DL,
[Parameter(Position = 9,Mandatory=$false,HelpMessage='This switch instructs Powershell to only return Computer Objects.')]
	[switch]
$system,
[Parameter(Position = 10,Mandatory=$false,HelpMessage='This switch instructs Powershell to only return OU Objects.')]
	[switch]
$OU,
[Parameter(Position = 11,Mandatory=$false,HelpMessage='This switch instructs Powershell to only return Lync-enabled Objects.')]
	[switch]
$Lync,
[Parameter(Position = 12,Mandatory=$false,HelpMessage='This switch instructs Powershell to only return AD Schema Objects.')]
	[switch]
$description,
[Parameter(Position = 13,Mandatory=$false,HelpMessage='This switch instructs Powershell to only return AD Schema Objects.')]
	[switch]
$schema,
[Parameter(Position = 14,Mandatory=$false,HelpMessage='This switch instructs Powershell to only search for AD Entries by Telephone Extension.')]
	[switch]
$extension,
[Parameter(Position = 15,Mandatory=$false,HelpMessage='Specifies a PSObject should be returned.')]
	[switch]
$PSO,
[Parameter(Position = 16,Mandatory=$false,HelpMessage='Display diagnostic info on screen.')]
	[switch]
$d
)
begin
{
# Results is an array of hash tables
$results=@(@{},@{})
$PSOresults=@()
$hash=@{}
[string]$value=""
[string]$filter=""
[array]$attributes=@()
[array]$schema=@()

# Error checking

if ($properties -contains '*')
	{
if ($d) {write-host("Executing * Properties")}
	$properties="*"
	} else
	{
if ($d) {write-host("Checking  Properties")}
# Gets the local domain info
	[string]$localdomaininfo = get-wmiobject -class "Win32_NTDomain" -namespace "root\CIMV2"

# Is this Stage?
	if ($localdomaininfo.contains("JANUSDEV")) {$InSOU="LDAP://" + "CN=Schema,CN=Configuration,DC=tjanusadmin,DC=net"} else {$InSOU="LDAP://" + "CN=Schema,CN=Configuration,DC=janusadmin,DC=net"}

	[string]$ObjSFilter='(&(objectclass=attributeSchema)(DistinguishedName=*))'
if ($d) {$ObjSFilter | out-string}
	$objSSearch = New-Object System.DirectoryServices.DirectorySearcher
if ($d) {$objSSearch | out-string}
	$objSDir = New-Object System.DirectoryServices.DirectoryEntry($InSOU)
if ($d) {$objSDir | out-string}
	$objSSearch.SearchRoot = $objSDir
	$objSSearch.SearchScope = "OneLevel"
if ($d) {$objSSearch.SearchScope | out-string}
	$objSSearch.PropertiesToLoad.Add("ldapdisplayname") | out-null
if ($d) {$objSSearch.PropertiesToLoad | out-string}
	$objSSearch.PageSize = 10000
if ($d) {$objSSearch.PageSize | out-string}
	$objSSearch.Filter = $ObjSFilter
if ($d) {$objSSearch.Filter | out-string}
if ($d) {$objSSearch | out-string}
	$attributes = $objSSearch.FindAll()
if ($d) {"Attributes Count"}
if ($d) {write-host ("$($attributes.count)")}
if ($d) {"Sorting Attributes"}
	$attributes = $attributes | select –ExpandProperty Properties
	$attributes | where {$ldapdisplaynames += $_.item("ldapdisplayname")}
if ($d) {"LDAPDisplayName Count"}
if ($d) {$ldapdisplaynames.count}
# if ($d) {$ldapdisplaynames | out-string}

	foreach ($property IN $properties)
		{
if ($d) {write-host("Checking Property $property")}
#		$filter=$property
if ($d) {write-host("Using Filter $property")}
		if (($ldapdisplaynames -like $property).count -eq 0) {$errorstring="$property is not a valid AD Attribute. To find out more, execute the following command: Get-JADEntry $property -schema -inou `"CN=Schema,CN=Configuration,DC=janusadmin,DC=net`" -properties ldapdisplayname -pso, exiting...";write-error $errorstring -erroraction stop}
if ($d) {write-host("Count - $(($ldapdisplaynames -like $filter).count)")}
		}
	}
if ($d) {$properties}
}

Process
{
# Normalize GUIDs
$guidtest=$ID.split("-")
$dguidtest=$ID.split(" ")
if (($ID.length -eq 36) -and($guidtest.count -eq 5))
	{
	$string=$ID
	$string=$string.replace("-","")
	$guid="\" + $string[6] + $string[7] + "\" + $string[4] + $string[5] + "\" + $string[2] + $string[3] + "\" + $string[0] + $string[1] + "\" + $string[10] + $string[11] + "\" + $string[8] + $string[9] + "\" + $string[14] + $string[15] + "\" + $string[12] + $string[13] + "\" + $string[16] + $string[17] + "\" + $string[18] + $string[19] + "\" + $string[20] + $string[21] + "\" + $string[22] + $string[23] + "\" + $string[24] + $string[25] + "\" + $string[26] + $string[27] + "\" + $string[28] + $string[29] + "\" + $string[30] + $string[31]
	write-output("Octet GUID: $guid")
	} elseif (($ID.length -ge 31) -and($dguidtest.count -eq 16))
	{
	$guid="\" + [Convert]::ToString($dguidtest[0], 16) + "\" + [Convert]::ToString($dguidtest[1], 16) + "\" + [Convert]::ToString($dguidtest[2], 16) + "\" + [Convert]::ToString($dguidtest[3], 16) + "\" + [Convert]::ToString($dguidtest[4], 16) + "\" + [Convert]::ToString($dguidtest[5], 16) + "\" + [Convert]::ToString($dguidtest[6], 16) + "\" + [Convert]::ToString($dguidtest[7], 16) + "\" + [Convert]::ToString($dguidtest[8], 16) + "\" + [Convert]::ToString($dguidtest[9], 16) + "\" + [Convert]::ToString($dguidtest[10], 16) + "\" + [Convert]::ToString($dguidtest[11], 16) + "\" + [Convert]::ToString($dguidtest[12], 16) + "\" + [Convert]::ToString($dguidtest[13], 16) + "\" + [Convert]::ToString($dguidtest[14], 16) + "\" + [Convert]::ToString($dguidtest[15], 16)
	write-output("Decimal GUID: $guid")
	} else
	{
	$guid="$ID"
	}

# Normalize $InOU
if ($InOU)
	{
	if (-not($InOU.contains("LDAP://"))) {$InOU="LDAP://" + $InOU}
	}

# Gets the local domain info
[string]$localdomaininfo = get-wmiobject -class "Win32_NTDomain" -namespace "root\CIMV2"
# Is this Stage?
if (($localdomaininfo.contains("JANUSDEV")) -and($InOU -eq "LDAP://dc=janus, dc=cap")) {$InOU="LDAP://dc=tjanus, dc=cap"}

# Remove the 2 if looking for an extension
if (($extension) -and($exact))
	{
	$length=($ID.length)-1
	if (($ID[0] -eq "2") -and($length -eq 4)) {$ID=$ID.substring(1,$length)}
	} elseif ($extension)
	{
	$length=($ID.length)-1
	if (($ID[0] -eq "2") -and($length -eq 4)) {$ID=$ID.substring(1,$length)}
	}

if ($d)
	{
	write-output ("ID: $ID")
	write-output ("Exact Search Only: $exact")
	write-output ("Properties: $properties")
	write-output ("Search Root: $InOU")
	write-output ("Non-recursive Search: $norecurse")
	write-output ("User-only search: $person")
	write-output ("Group-only search: $group")
	write-output ("DL-only search: $DL")
	write-output ("Lync-enabled only search: $Lync")
	write-output ("Computer-only search: $system")
	write-output ("OU-only search: $OU")
	write-output ("Extension search: $extension")
	write-output ("Schema-only search: $schema")
	write-output ("LDAP GUID version: $guid")
	write-output ("Debug mode: $d")
	}

if ($exact)
{
if ($d)
	{
	write-output ("Performing Exact search for $ID...")
	}
if ($OU)
	{
	[string]$filter='(&(&(ou>="")(|(name=' + $ID + ')(DistinguishedName=' + $ID + '))))'
	}
elseif($system)
	{
	[string]$filter='(&(&(sAMAccountType=805306369)(name=' + $ID + ')))'
	}
elseif($Lync)
	{
	[string]$filter='(&(&(|(objectCategory=user)(objectCategory=inetOrgPerson)(objectCategory=contact))(!(msRTCSIP-ApplicationOptions=*))(msRTCSIP-PrimaryUserAddress=sip:' + $ID + ')))'
	}
elseif($DL)
	{
	[string]$filter='(&(&(|(&(objectCategory=person)(objectSid=*)(!samAccountType:1.2.840.113556.1.4.804:=3))(&(objectCategory=person)(!objectSid=*))(&(objectCategory=group)(groupType:1.2.840.113556.1.4.804:=14)))(anr=' + $ID + ')(& (mailnickname=*) (| (objectCategory=group) ))))'
	}
elseif($group)
	{
	[string]$filter='(&(&(&(|(&(objectCategory=person)(objectSid=*)(!samAccountType:1.2.840.113556.1.4.804:=3))(&(objectCategory=person)(!objectSid=*))(&(objectCategory=group)(groupType:1.2.840.113556.1.4.804:=14))))(objectCategory=group)(|(name=' + $ID + ')(displayname=' + $ID + ')(samaccountname=' + $ID + '))))'
	}
elseif($extension)
	{
	[string]$filter='(&(objectCategory=user)(telephonenumber=*' + $ID + '))'
	}
elseif($person)
	{
	[string]$filter='(&(objectCategory=user)(|(DisplayName=' + $ID + ')(name=' + $ID + ')(MailnickName=' + $ID + ')(distinguishedname=' + $ID + ')(msexchmailboxguid=' + $ID + ')(objectguid=' + $ID + ')(msexchmailboxguid=' + $guid + ')(objectguid=' + $guid + ')(samaccountname=' + $ID + ')))'
	}
elseif($description)
	{
	$ID=$ID.replace("`\","\5C")
	$ID=$ID.replace(" ","\20")
	$ID=$ID.replace("`"
","\22")
	$ID=$ID.replace("`&","\26")
	$ID=$ID.replace("`'","\27")
	$ID=$ID.replace("`,","\2C")
	$ID=$ID.replace("`.","\2E")
	$ID=$ID.replace("`/","\2F")
	$ID=$ID.replace("`:","\3A")
	$ID=$ID.replace("`;","\3B")
	[string]$filter='(description=' + $ID + ')'
	}
elseif($schema)
	{
	[string]$filter='(&(objectclass=attributeSchema)(|(DistinguishedName=' + $ID + ')(cn=' + $ID + ')(schemaIDGUID=' + $guid + ')(ObjectGUID=' + $guid + ')(schemaIDGUID=' + $ID + ')(ObjectGUID=' + $ID + ')(Name=' + $ID + ')))'
	}
else
	{
	[string]$filter='(&(&(objectCategory=*)(objectClass=*)) (|(DistinguishedName=' + $ID + ')(name=' + $ID + ')(displayname=' + $ID + ')(mail=' + $ID + ')(objectguid=' + $ID + ')(msexchmailboxguid=' + $ID + ')(objectguid=' + $guid + ')(msexchmailboxguid=' + $guid + ')(samaccountname=' + $ID + ')))'
	}
}
else
{
if ($d)
	{
	write-output ("Performing fuzzy search for $ID...")
	}
if ($OU)
	{
	[string]$filter='(&(&(ou>="")(|(name=*' + $ID + '*)(DistinguishedName=*' + $ID + '*))))'
	}
elseif($system)
	{
	[string]$filter='(&(&(sAMAccountType=805306369)(name=*' + $ID + '*)))'
	}
elseif($Lync)
	{
	[string]$filter='(&(&(|(objectCategory=user)(objectCategory=inetOrgPerson)(objectCategory=contact))(!(msRTCSIP-ApplicationOptions=*))(msRTCSIP-PrimaryUserAddress=sip:*' + $ID + '*)))'
	}
elseif($DL)
	{
	[string]$filter='(&(&(&(& (mailnickname=*) (| (objectCategory=group)(objectCategory=msExchDynamicDistributionList) )))(objectCategory=group)(|(name=*' + $ID + '*)(mailNickname=*' + $ID + '*)(samaccountname=*' + $ID + '*))))'
	}
elseif($group)
	{
	[string]$filter='(&(&(&(|(&(objectCategory=person)(objectSid=*)(!samAccountType:1.2.840.113556.1.4.804:=3))(&(objectCategory=person)(!objectSid=*))(&(objectCategory=group)(groupType:1.2.840.113556.1.4.804:=14))))(objectCategory=group)(|(name=*' + $ID + '*)(displayname=*' + $ID + '*)(samaccountname=*' + $ID + '*))))'
	}
elseif($extension)
	{
	[string]$filter='(&(objectCategory=user)(telephonenumber=*' + $ID + '*))'
	}
elseif($person)
	{
	[string]$filter='(&(objectCategory=user)(|(DisplayName=*' + $ID + '*)(name=*' + $ID + '*)(MailnickName=*' + $ID + '*)(distinguishedname=*' + $ID + '*)(msexchmailboxguid=*' + $ID + '*)(objectguid=*' + $ID + '*)(msexchmailboxguid=*' + $guid + '*)(objectguid=*' + $guid + '*)(samaccountname=*' + $ID + '*)))'
	}
elseif($description)
	{
	$ID=$ID.replace("`\","\5C")
	$ID=$ID.replace(" ","\20")
	$ID=$ID.replace("`"
","\22")
	$ID=$ID.replace("`&","\26")
	$ID=$ID.replace("`'","\27")
	$ID=$ID.replace("`,","\2C")
	$ID=$ID.replace("`.","\2E")
	$ID=$ID.replace("`/","\2F")
	$ID=$ID.replace("`:","\3A")
	$ID=$ID.replace("`;","\3B")
	[string]$filter='(description=*' + $ID + '*)'
	}
elseif($schema)
	{
	[string]$filter='(&(objectclass=attributeSchema)(|(DistinguishedName=*' + $ID + '*)(cn=*' + $ID + '*)(schemaIDGUID=*' + $guid + '*)(ObjectGUID=*' + $guid + '*)(schemaIDGUID=*' + $ID + '*)(ObjectGUID=*' + $ID + '*)(Name=*' + $ID + '*)))'
	}
else
	{
	[string]$filter='(&(&(objectCategory=*)(objectClass=*)) (|(DistinguishedName=' + $ID + ')(name=*' + $ID + '*)(displayname=*' + $ID + '*)(mail=*' + $ID + '*)(objectguid=*' + $ID + '*)(msexchmailboxguid=*' + $ID + '*)(objectguid=*' + $guid + '*)(msexchmailboxguid=*' + $guid + '*)(samaccountname=*' + $ID + '*)))'
	}
}

$ObjFilter = $filter

if ($d)
	{
	write-output ($ObjFilter)
	}

if ($dc)
	{
	$ObjDC = "LDAP://"+$dc.split(".")[0]
	$objDir = New-Object System.DirectoryServices.DirectoryEntry $ObjDC
	} else
	{
	$objDir = New-Object System.DirectoryServices.DirectoryEntry($InOU)
	}

if ($d)
	{
	write-output ($objDir)
	}

$objSearch = New-Object System.DirectoryServices.DirectorySearcher
if ($d)
	{
	write-output ($objSearch)
	}

$objSearch.SearchRoot = $objDir
if ($norecurse)
	{
	$objSearch.SearchScope = "OneLevel"
	} else
	{
	$objSearch.SearchScope = "Subtree"
	}

if ($d)
	{
	write-output ($objSearch.SearchScope)
	}

if ($d)
	{
	write-output("Loading Properties...")
	}

if ($d)
	{
	foreach ($property in $properties)
		{
		$objSearch.PropertiesToLoad.Add($property)
		}
	} else
	{
	foreach ($property in $properties)
		{
		$objSearch.PropertiesToLoad.Add($property) | out-null
		}
	}

if ($d)
	{
	write-output("Properties Loaded...")
	}

if ($d)
	{
	write-output ($objSearch.PropertiesToLoad)
	}

$objSearch.PageSize = 10000
$objSearch.Filter = $ObjFilter
$AllObj = $objSearch.FindAll()
if ($d)
	{
	$count=($AllObj | measure-object).count
	write-output ("Hits: $count")
	}
}

end
{
if ($PSO)
{
foreach ($Obj in $AllObj)
	{
	$hash=@{}
	foreach($property in $Obj.Properties)
		{
		$names=$property.propertyNames|sort
		foreach($name in $names)
			{
			$value = $($Obj.Properties[$name])
			if (($name -like "badpasswordtime") -or($name -like "lastlogon") -or($name -like "lastlogontimestamp") -or($name -like "pwdlastset") -or($name -like "lockouttime"))
				{
				$value=[datetime]::fromfiletime($value)
				}
			if ($name -like "useraccountcontrol")
				{
				[string]$uacvalue="$value`n"
				if (($value -band 1) -ne 0) {$uacvalue += "ADS_UF_SCRIPT`n"}
				if (($value -band 2) -ne 0) {$uacvalue += "ADS_UF_ACCOUNTDISABLE`n"}
				if (($value -band 8) -ne 0) {$uacvalue += "ADS_UF_HOMEDIR_REQUIRED`n"}
				if (($value -band 16) -ne 0) {$uacvalue += "ADS_UF_LOCKOUT`n"}
				if (($value -band 32) -ne 0) {$uacvalue += "ADS_UF_PASSWD_NOTREQD`n"}
				if (($value -band 64) -ne 0) {$uacvalue += "ADS_UF_PASSWD_CANT_CHANGE`n"}
				if (($value -band 128) -ne 0) {$uacvalue += "ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED`n"}
				if (($value -band 256) -ne 0) {$uacvalue += "ADS_UF_TEMP_DUPLICATE_ACCOUNT`n"}
				if (($value -band 512) -ne 0) {$uacvalue += "ADS_UF_NORMAL_ACCOUNT`n"}
				if (($value -band 1024) -ne 0) {$uacvalue += "ADS_UF_INTERDOMAIN_TRUST_ACCOUNT`n"}
				if (($value -band 4096) -ne 0) {$uacvalue += "ADS_UF_WORKSTATION_TRUST_ACCOUNT`n"}
				if (($value -band 8192) -ne 0) {$uacvalue += "ADS_UF_SERVER_TRUST_ACCOUNT`n"}
				if (($value -band 65536) -ne 0) {$uacvalue += "ADS_UF_DONT_EXPIRE_PASSWD`n"}
				if (($value -band 131072) -ne 0) {$uacvalue += "ADS_UF_MNS_LOGON_ACCOUNT`n"}
				if (($value -band 262144) -ne 0) {$uacvalue += "ADS_UF_SMARTCARD_REQUIRED`n"}
				if (($value -band 524288) -ne 0) {$uacvalue += "ADS_UF_TRUSTED_FOR_DELEGATION`n"}
				if (($value -band 1048576) -ne 0) {$uacvalue += "ADS_UF_NOT_DELEGATED`n"}
				if (($value -band 2097152) -ne 0) {$uacvalue += "ADS_UF_USE_DES_KEY_ONLY`n"}
				if (($value -band 4194304) -ne 0) {$uacvalue += "ADS_UF_DONT_REQUIRE_PREAUTH`n"}
				if (($value -band 8388608) -ne 0) {$uacvalue += "ADS_UF_PASSWORD_EXPIRED`n"}
				if (($value -band 16777216) -ne 0) {$uacvalue += "ADS_UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION`n"}
				$value=$uacvalue
				} # if ($name -like "useraccountcontrol")
			if ($d) {write-ouput("$name = $value")}
			if ($d) {write-ouput("$hash")}
			$hash.add($name,$value)
			} # foreach($name in $names)
		} # foreach($property in $Obj.Properties)
	$PSObj=New-Object PSObject -Property $hash
	if ($d) {write-ouput("$PSObj")}
	$PSOresults+=$PSObj
	} # foreach ($Obj in $AllObj)
return $PSOresults
} else
{
$ObjCount=0
foreach ($Obj in $AllObj)
	{
	if ($ObjCount -gt 1) {$hash=@{};$results+=$hash}
	$results[$ObjCount]=@{}
	($Obj.Properties).length
	foreach($property in $Obj.Properties)
		{
		$names=$property.propertyNames|sort
		foreach($name in $names)
			{
			$value = $($Obj.Properties[$name])
			if (($name -like "badpasswordtime") -or($name -like "lastlogon") -or($name -like "lastlogontimestamp") -or($name -like "pwdlastset") -or($name -like "lockouttime"))
				{
				$value=[datetime]::fromfiletime($value)
				} # if (($name -like "badpasswordtime") -or($name -like "lastlogon") -or($name -like "lastlogontimestamp") -or($name -like "pwdlastset") -or($name -like "lockouttime"))
			if ($name -like "useraccountcontrol")
				{
				[string]$uacvalue="$value`n"
				if (($value -band 1) -ne 0) {$uacvalue += "ADS_UF_SCRIPT`n"}
				if (($value -band 2) -ne 0) {$uacvalue += "ADS_UF_ACCOUNTDISABLE`n"}
				if (($value -band 8) -ne 0) {$uacvalue += "ADS_UF_HOMEDIR_REQUIRED`n"}
				if (($value -band 16) -ne 0) {$uacvalue += "ADS_UF_LOCKOUT`n"}
				if (($value -band 32) -ne 0) {$uacvalue += "ADS_UF_PASSWD_NOTREQD`n"}
				if (($value -band 64) -ne 0) {$uacvalue += "ADS_UF_PASSWD_CANT_CHANGE`n"}
				if (($value -band 128) -ne 0) {$uacvalue += "ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED`n"}
				if (($value -band 256) -ne 0) {$uacvalue += "ADS_UF_TEMP_DUPLICATE_ACCOUNT`n"}
				if (($value -band 512) -ne 0) {$uacvalue += "ADS_UF_NORMAL_ACCOUNT`n"}
				if (($value -band 1024) -ne 0) {$uacvalue += "ADS_UF_INTERDOMAIN_TRUST_ACCOUNT`n"}
				if (($value -band 4096) -ne 0) {$uacvalue += "ADS_UF_WORKSTATION_TRUST_ACCOUNT`n"}
				if (($value -band 8192) -ne 0) {$uacvalue += "ADS_UF_SERVER_TRUST_ACCOUNT`n"}
				if (($value -band 65536) -ne 0) {$uacvalue += "ADS_UF_DONT_EXPIRE_PASSWD`n"}
				if (($value -band 131072) -ne 0) {$uacvalue += "ADS_UF_MNS_LOGON_ACCOUNT`n"}
				if (($value -band 262144) -ne 0) {$uacvalue += "ADS_UF_SMARTCARD_REQUIRED`n"}
				if (($value -band 524288) -ne 0) {$uacvalue += "ADS_UF_TRUSTED_FOR_DELEGATION`n"}
				if (($value -band 1048576) -ne 0) {$uacvalue += "ADS_UF_NOT_DELEGATED`n"}
				if (($value -band 2097152) -ne 0) {$uacvalue += "ADS_UF_USE_DES_KEY_ONLY`n"}
				if (($value -band 4194304) -ne 0) {$uacvalue += "ADS_UF_DONT_REQUIRE_PREAUTH`n"}
				if (($value -band 8388608) -ne 0) {$uacvalue += "ADS_UF_PASSWORD_EXPIRED`n"}
				if (($value -band 16777216) -ne 0) {$uacvalue += "ADS_UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION`n"}
				$value=$uacvalue
				} # if ($name -like "useraccountcontrol")
			$results[$ObjCount].add($name,$value)
			} # foreach($name in $names)
		} # foreach($property in $Obj.Properties)
	$ObjCount=$ObjCount+1
	} # foreach ($Obj in $AllObj)
foreach ($element IN $results) {$element.GetEnumerator()  | sort name;write-output ("`n")}
} # if ($PSO)

remove-variable results
remove-variable AllObj
remove-variable PSOResults
remove-variable OBJSearch
}
} # function Get-JADEntry

Function ConvertFrom-JOctetToGUID ($Octet)
{
# Function to convert Octet value (byte array) into string GUID value.
    $GUID = "{" + "\" + [Convert]::ToString($Octet[3], 16) `
        + "\" + [Convert]::ToString($Octet[2], 16) `
        + "\" + [Convert]::ToString($Octet[1], 16) `
        + "\" + [Convert]::ToString($Octet[0], 16) + "-" `
        + "\" + [Convert]::ToString($Octet[5], 16) `
        + "\" + [Convert]::ToString($Octet[4], 16) + "-" `
        + "\" + [Convert]::ToString($Octet[7], 16) `
        + "\" + [Convert]::ToString($Octet[6], 16) + "-" `
        + "\" + [Convert]::ToString($Octet[8], 16) `
        + "\" + [Convert]::ToString($Octet[9], 16) + "-" `
        + "\" + [Convert]::ToString($Octet[10], 16) `
        + "\" + [Convert]::ToString($Octet[11], 16) `
        + "\" + [Convert]::ToString($Octet[12], 16) `
        + "\" + [Convert]::ToString($Octet[13], 16) `
        + "\" + [Convert]::ToString($Octet[14], 16) `
        + "\" + [Convert]::ToString($Octet[15], 16) + "}"
    Return $GUID
} # Function ConvertFrom-JOctetToGUID

function Set-JADEntry
{
<#
	.SYNOPSIS
		Changes AD Settings on the specified AD account.

	.DESCRIPTION
		Allows you to change any of the AD Properties of an AD Entry.

	.EXAMPLE
		Set-JADEntry -ID JM27253 -description "Test description" -givenName Randall
		
		Description
		-----------
		Changes the description and the first name of the JM27253 AD Account.
		
	.NOTES
		Requires the Loading of the Janus Module.
#>

$hash=@{}
[int]$p=0
[int]$limit=0
[string]$key=""
[string]$value=""
[string]$check=""
[bool]$renamed=$false

# True turns on debugging
# $d=$true
# False turns off debugging
$d=$false

# Setup for Error checking
if ($d) {write-host("Checking  Properties")}
# Gets the local domain info
[string]$localdomaininfo = get-wmiobject -class "Win32_NTDomain" -namespace "root\CIMV2"

# Is this Stage?
if ($localdomaininfo.contains("JANUSDEV")) {$InSOU="LDAP://" + "CN=Schema,CN=Configuration,DC=tjanusadmin,DC=net"} else {$InSOU="LDAP://" + "CN=Schema,CN=Configuration,DC=janusadmin,DC=net"}
[string]$ObjSFilter='(&(objectclass=attributeSchema)(DistinguishedName=*))'
if ($d) {$ObjSFilter | out-string}
$objSSearch = New-Object System.DirectoryServices.DirectorySearcher
if ($d) {$objSSearch | out-string}
$objSDir = New-Object System.DirectoryServices.DirectoryEntry($InSOU)
if ($d) {$objSDir | out-string}
$objSSearch.SearchRoot = $objSDir
$objSSearch.SearchScope = "OneLevel"
if ($d) {$objSSearch.SearchScope | out-string}
$objSSearch.PropertiesToLoad.Add("ldapdisplayname") | out-null
if ($d) {$objSSearch.PropertiesToLoad | out-string}
$objSSearch.PageSize = 10000
if ($d) {$objSSearch.PageSize | out-string}
$objSSearch.Filter = $ObjSFilter
if ($d) {$objSSearch.Filter | out-string}
if ($d) {$objSSearch | out-string}
$attributes = $objSSearch.FindAll()
if ($d) {"Attributes Count"}
if ($d) {write-host ("$($attributes.count)")}
if ($d) {"Sorting Attributes"}
$attributes = $attributes | select –ExpandProperty Properties
$attributes | where {$ldapdisplaynames += $_.item("ldapdisplayname")}
if ($d) {"LDAPDisplayName Count"}
if ($d) {$ldapdisplaynames.count}
# if ($d) {$ldapdisplaynames | out-string}
# [string]$Valid=$ldapdisplaynames | out-string

function ADOSet
{
param([string]$key,[string]$value)
[string]$olddn=""
[string]$newdn=""
if((($value -like "true") -or($value -eq $true)) -and(($script:ADO -like "False") -or($script:ADO -like "True") -or($script:ADO -like "") -or($script:ADO -like $null)))
	{
	$value="TRUE"
	}
if((($value -like "false") -or($value -eq $false)) -and(($script:ADO -like "False") -or($script:ADO -like "True") -or($script:ADO -like "") -or($script:ADO -like $null)))
	{
	$value="FALSE"
	}
if (($key -notlike "name") -and($key -notlike "cn") -and($key -notlike "samaccountname") -and($key -notlike "distinguishedName") -and($key -notlike "userPrincipalName"))
	{
	if ($key -like "accountexpires")
		{
		$datetest=(Get-Date $value -erroraction silentlycontinue)
		if($?)
			{
			$value=[string](Get-Date $value).ToFileTime()
			} # if($?)
		} # if (($key -like "badpasswordtime") -or($key -like "lastlogon") -or($key -like "lastlogontimestamp") -or($key -like "pwdlastset") -or($key -like "accountexpires"))
	$script:ADO.psbase.invokeset("$key","$value")
	} else # if (($key -notlike "name") -and($key -notlike "cn") -and($key -notlike "samaccountname") -and($key -notlike "distinguishedName") -and($key -notlike "userPrincipalName"))
	{
	if (-not $script:renamed)
		{
		$oldname=$script:ADO.samaccountname
		$newname=$value.replace("@janus.cap","")
		$temp=$newname.split(",")
		if ($temp[0] -like "cn=")
			{
			$newname=$temp[1]
			} else
			{
			$newname=$temp[0]
			}
		$newdn="CN=" + $newname
		$script:ADO.psbase.Rename("$newdn")
		$script:ADO.psbase.invokeset("samaccountname" , "$newname")
		$UPN=$newname + "@janus.cap"
		$script:ADO.psbase.invokeset("userPrincipalName" , "$UPN")
		$script:renamed=$true
		} # if (-not $script:renamed)
	} # else
$script:ADO.SetInfo()
} # function ADOSet

[int]$limit=$args.count

if ($args[0] -like "-id")
	{
	$ID=$args[1]
	$p=2
	}
else
	{
	$ID=$args[0]
	$p=1
	}

for ($p=$p; $p -lt $limit; $p=$p+2)	{
	if($args[$p].StartsWith("-")) {$args[$p] = $args[$p].Substring(1)}
	$check="$($args[$p])"
	if (($ldapdisplaynames -like $check).count -eq 0) {write-output("$($args[$p]) is not a valid Property, discarding this Property...")}
	else {$hash.add("$($args[$p])","$($args[$p+1])")}
	}

if ($hash.count -eq 0)
	{
	write-output ("No valid properties found, exiting...")
	[string]$errorstring="No valid properties found, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}
$props=$hash.keys

$ADID=get-jadentry -ID $ID -exact -pso -properties $props
[string]$adspath=$ADID.adspath
if (($adspath -ne $null) -and($adspath -ne ""))
	{
	$script:ADO=[ADSI]$adspath
	} else
	{
	write-output ("AD Lookup error, cannot find $ID, exiting...")
	[string]$errorstring="AD Lookup error, cannot find $ID, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}
$script:renamed=$false
$hash.GetEnumerator() | % {$key=$_.key;$value=$_.value;ADOSet -key $($key) -value $($value);write-output("$key modified...")}
$script:ADO.psbase.CommitChanges() 
$script:ADO.SetInfo()
remove-variable ADO -Scope Script
remove-variable ADID
remove-variable props
} # function Set-JADEntry

function Show-JLyncUserInfo
{
<#
.SYNOPSIS
	Returns Lync Stats.

.DESCRIPTION
	Retreives Stats from AD for the specified user. Including:
	SAMAccountName
	Display Name
	msRTCSIP-UserEnabled - Boolean - True=Lync Enable; False = Lync Disabled

	msRTCSIP-PrimaryUserAddress - String - User's Primary SIP address
	msRTCSIP-PrimaryHomeServer - String - Distinguished Name of Home Server
	RTCSIP-OptionFlags(1) - Integer - Bitmask representing Lync features

.PARAMETER  ID
	Enter the ID of the user to report on.
	NOTE: Accepts SAMAccountName, Display Name or SMTP Address.

.EXAMPLE
	Show-JLyncUserInfo -ID jm27253
		
	Description
	-----------
	Gets the Lync stats for the jm27253 AD Accont.

.NOTES
	Requires the Loading of the Janus Module.

#>
	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("Identity")]
$ID
)

$LyncU=get-jadentry -id $ID -exact -pso

$SAMAccountName=$($LyncU.name)
$DisplayName=$($LyncU.DisplayName)
$mail=$($LyncU.mail)
$LyncEnabled=$($LyncU.'msRTCSIP-UserEnabled')
$mail=$($LyncU.mail)
$LyncAddress=$($LyncU.'msRTCSIP-PrimaryUserAddress')
$LyncPool=$($LyncU.'msRTCSIP-PrimaryHomeServer')
$LyncOptions=$($LyncU.'msRTCSIP-OptionFlags')
$msRTCSIPEnableFederation=$($LyncU.'msRTCSIP-EnableFederation')
$msRTCSIPFederationEnabled=$($LyncU.'msRTCSIP-FederationEnabled')
$msRTCSIPInternetAccessEnabled=$($LyncU.'msRTCSIP-InternetAccessEnabled')
$msRTCSIPNumDevicesPerUser=$($LyncU.'msRTCSIP-NumDevicesPerUser')

Write-output("SAMAccountName: $SAMAccountName")
Write-output("DisplayName: $DisplayName")
Write-output("Primary SMTP Address: $mail")
Write-output("Lync Address: $LyncAddress")
Write-output("Lync Enabled: $LyncEnabled")
Write-output("Lync Internet Enabled: $msRTCSIPInternetAccessEnabled")
Write-output("Lync Pool: $LyncPool")
if ($msRTCSIPEnableFederation -ne $null) {Write-output("Lync Federation Enabled: $msRTCSIPEnableFederation")}
if ($msRTCSIPFederationEnabled -ne $null) {Write-output("Lync Federation Enabled: $msRTCSIPFederationEnabled")}
if ($msRTCSIPNumDevicesPerUser -ne $null) {Write-output("Lync Device Limit: $msRTCSIPNumDevicesPerUser")}
Write-output("Lync Options: $LyncOptions")

# AD Attribute Name    Type    Meaning
ms
# RTCSIP-UserEnabled    Boolean    True = Lync Enable; False = not Lync Enabled
ms
# RTCSIP-OptionFlags (1)    Integer    Bitmask representing Lync features.

# msRTCSIP-PrimaryUserAddress    String    The primary SIP address of the user.

# msRTCSIP-PrimaryHomeServer    String    The Distinguished Name of the Home Server.
remove-variable LyncU
} # function Show-JLyncUserInfo

Function Get-JDSACLs
{
<#
.SYNOPSIS
	Returns AD ACLs of the specified AD Entry.

.DESCRIPTION
	Returns AD ACLs of the specified AD Entry.

.PARAMETER  ID
	Enter the ID of the user to report on.
	NOTE: Accepts SAMAccountName, Display Name or SMTP Address.

.PARAMETER  inherited
	Specifies whether or not to display inherited permissions.

.EXAMPLE
	Get-JDSACLs -ID jm27253
		
	Description
	-----------
	Gets the ACLs for the jm27253 AD Accont.

.EXAMPLE
	Get-JDSACLs -ID jm27253 -inherited
		
	Description
	-----------
	Gets the ACLs (including inherited permissions) for the jm27253 AD Accont.

.NOTES
	Requires the Loading of the Janus Module.

#>
	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("Identity")]
$ID,
[Switch]$Inherited = $False
)
 
# Function to list permissions assigned to objects in Active Directory
# This function reports on organizationalUnits within the domain by default.

# Error checking
$ADE=get-adentry -id $ID -exact -pso -properties distinguishedname
[string]$SearchRoot=$ADE.distinguishedname

# Set the output field separator (default is " ")
$OFS = "\"

# Connect to RootDSE
$RootDSE = [ADSI]"LDAP://RootDSE"
# Connect to the Schema
$Schema = [ADSI]"LDAP://$($RootDSE.Get('schemaNamingContext'))"
# Connect to the Extended Rights container
$Configuration = $RootDSE.Get("configurationNamingContext")
$ExtendedRights = [ADSI]"LDAP://CN=Extended-Rights,$Configuration"

# Find objects based on $SearchRoot and $ObjectType
[string]$LdapFilter = "(&(&(objectClass=*)(objectCategory=*))(distinguishedname=" + $SearchRoot + "))"
$Searcher = New-Object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$SearchRoot", $LdapFilter)
$Searcher.FindAll() | %{
$Object = $_.GetDirectoryEntry()

# Retrieve all Access Control Entries from the AD Object
$ACL = $Object.PsBase.ObjectSecurity.GetAccessRules($True, $Inherited, [Security.Principal.NTAccount])

# Get interesting values
$ACL | Select-Object @{n='Name';e={ $Object.Get("name") }}, @{n='DN';e={ $Object.Get("distinguishedName") }}, @{n='ObjectClass';e={ $Object.Class }}, @{n='SecurityPrincipal';e={ $_.IdentityReference.ToString() }}, @{n='AccessType';e={ $_.AccessControlType }}, @{n='Permissions';e={ $_.ActiveDirectoryRights }}, @{n='AppliesTo';e={
#
# Change the values for InheritanceType to friendly names
#
Switch ($_.InheritanceType) {
"None"            { "This object only" }
"Descendents"     { "All child objects" }
"SelfAndChildren" { "This object and one level Of child objects" }
"Children"        { "One level of child objects" }
"All"             { "This object and all child objects" }
} }}, @{n='AppliesToObjectType';e={
If ($_.InheritedObjectType.ToString() -NotMatch "0{8}.*") {
#
# Search for the Object Type in the Schema
#
$LdapFilter = "(SchemaIDGUID=\$($_.InheritedObjectType.ToByteArray() | %{ '{0:X2}' -f $_ }))"
$Result = (New-Object DirectoryServices.DirectorySearcher($Schema, $LdapFilter)).FindOne()
$Result.Properties["ldapdisplayname"]
} Else { "All" } }}, @{n='AppliesToProperty';e={
If ($_.ObjectType.ToString() -NotMatch "0{8}.*") {
#
# Search for a possible Extended-Right or Property Set
#
$LdapFilter = "(rightsGuid=$($_.ObjectType.ToString()))"
$Result = (New-Object DirectoryServices.DirectorySearcher($ExtendedRights, $LdapFilter)).FindOne()
If ($Result) {
$Result.Properties["displayname"]
} Else {
#
# Search for the attribute name in the Schema
#
$LdapFilter = "(SchemaIDGUID=\$($_.ObjectType.ToByteArray() | %{ '{0:X2}' -f $_ }))"
$Result = (New-Object DirectoryServices.DirectorySearcher($Schema, $LdapFilter)).FindOne()
$Result.Properties["ldapdisplayname"]
}
} Else { "All" } }}, @{n='Inherited';e={ $_.IsInherited }}
}
if ($result) {remove-variable result}
remove-variable Object
remove-variable ACL
remove-variable Searcher
remove-variable ADE
remove-variable RootDSE
remove-variable Schema
remove-variable Configuration
remove-variable ExtendedRights
} # Function Get-JDSACLs

function Update-JEWSDelegate
{
<#
	.SYNOPSIS
		Updates a delegate to a mailbox.

	.DESCRIPTION
		Updates the specified Delegate to the specified folder to a mailbox.

	.PARAMETER  mb
		Enter the SMTP address of the mailbox being delegated

	.PARAMETER  viewer
		Enter the SMTP address of the mailbox the will be able to view new
		folders.

	.PARAMETER  folder
		Specify the name of the Foder that the Delegate will have access to.
		This will accept a comma-seperated list for multiple folders.
		NOTE: This defaults to "Calendar" if it is not specified.

	.PARAMETER  perms
		Specify the Permissions to grant from the following list:
		None, Contributor, Reviewer, Author, Editor.
		NOTE: This defaults to "Reviewer" if it is not specified.

	.PARAMETER  copy
		Specify if Delegate should receive a copy of Meeting-related Messages.

	.EXAMPLE
		Update-JEWSDelegate -mb JournalSocialNe@janus.com -viewer Randy.Moore@janus.com -folder Calendar
		
		Description
		-----------
		Allows Randy Moore to be able to view Calendar of thr Journal Social
		Networking Mailbox.
		
	.NOTES
		Requires EWS to be installed on the computer that executes this
		script.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("Identity")]
        [Alias("ID")]
        [Alias("mailbox")]
        [Alias("mail")]
		[String]
$mb=$(Throw "You must specify a Calendar to add a delegate to (e.g., JournalSocialNe@janus.com)."),
        [Parameter(Position = 1, Mandatory = $true)]
        [Alias("user")]
		[String]
$viewer=$(Throw "You must specify a delegate to add (e.g., Randy.Moore@janus.com)."),
        [Parameter(Position = 2, Mandatory = $false)]
		[String[]]
		[ValidateSet(
			'Cal',
			'Calendar',
			'Contact',
			'Contacts',
			'Inbox',
			'Note',
			'Notes',
			'Task',
			'Tasks'
		)]
$folder="Calendar",
        [Parameter(Position = 3, Mandatory = $false)]
        [Alias("accessrights")]
        [Alias("permissions")]
        [Alias("permission")]
		[String]
		[ValidateSet(
			'Editor',
			'Author',
			'Reviewer',
			'Contributor',
			'None'
		)]
$perms="Reviewer",
        [Parameter(Position = 4, Mandatory = $false)]
        [Alias("ReceiveCopiesOfMeetingRequests")]
        [Alias("ReceiveCopiesOfMeetingMessages")]
		[switch]
$copy
)

# Initialization
$dllpath = "D:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2)

# Filters and Functions
$folder=$folder.tolower()

# Error checking
$test=get-mailbox -id $mb
$successful=$?
if ($successful -eq $false)
	{
	[string]$errorstring="Error in Mailbox ID, cannot locate mailbox, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	} else
	{
	$mb=$test.PrimarySmtpAddress.tostring()
	}

$test=get-mailbox -id $viewer
$successful=$?
if ($successful -eq $false)
	{
	[string]$errorstring="Error in Delegate ID, cannot locate mailbox, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	} else
	{
	$viewer=$test.PrimarySmtpAddress.tostring()
	}

# Main Script
$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$service.AutodiscoverUrl($mb.ToString())
$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $mb);

$mbMailbox = new-object Microsoft.Exchange.WebServices.Data.Mailbox($mb)
$dgUser = new-object Microsoft.Exchange.WebServices.Data.DelegateUser($viewer)

If ($folder -like "*cal*")
	{
	Write-Output ("Updating Calendar")
	$dgUser.Permissions.CalendarFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::$perms
	if ($perms -like "Editor") {$dgUser.ReceiveCopiesOfMeetingMessages = $true}
	else  {$dgUser.ReceiveCopiesOfMeetingMessages = $false}
	}
If ($folder -like "*inbox*")
	{
	Write-Output ("Updating Inbox")
	$dgUser.Permissions.InboxFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::$perms
	}
If ($folder -like "*task*")
	{
	Write-Output ("Updating Tasks")
	$dgUser.Permissions.TasksFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::$perms
	}
If ($folder -like "*contact*")
	{
	Write-Output ("Updating Contacts")
	$dgUser.Permissions.ContactsFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::$perms
	}
If ($folder -like "*note*")
	{
	Write-Output ("Updating Notes")
	$dgUser.Permissions.NotesFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::$perms
	}

$dgUser.ViewPrivateItems = $false

$dgArray = new-object Microsoft.Exchange.WebServices.Data.DelegateUser[] 1
$dgArray[0] = $dgUser

if (($perms -like "Editor") -and($copy)) {$service.UpdateDelegates($mb, [Microsoft.Exchange.WebServices.Data.MeetingRequestsDeliveryScope]::DelegatesAndSendInformationToMe, $dgArray);}
else  {$service.UpdateDelegates($mb, [Microsoft.Exchange.WebServices.Data.MeetingRequestsDeliveryScope]::NoForward, $dgArray);}
remove-variable service
remove-variable dgArray
remove-variable dgUser
remove-variable mbmailbox
remove-variable ACEUser
remove-variable WindowsIdentity
} # function Update-JEWSDelegate

function Get-JEWSDelegates
{
<#
	.SYNOPSIS
		Looks up a delegate to a mailbox.

	.DESCRIPTION
		Looks up the specified Delegate to the specified folder to a mailbox.

	.PARAMETER  mb
		Enter the SMTP address of the mailbox being delegated

	.PARAMETER  del
		Enter the SMTP address of the delegate to look up

	.PARAMETER  folder
		Specify the name of the Foder that the Delegate will have access to.
		This will accept a comma-seperated list for multiple folders.
		NOTE: Accepts the following values: 'Cal', 'Calendar', 'Contact',
		'Contacts', 'Inbox', 'Note', 'Notes', 'Task', 'Tasks'
		NOTE: This defaults to All folders if it is not specified.

	.EXAMPLE
		Get-JEWSDelegates -mb JournalSocialNe@janus.com -folder Calendar
		
		Description
		-----------
		Looks up thge view Calendar of the Journal Social
		Networking Mailbox.
		
	.NOTES
		Requires EWS to be installed on the computer that executes this
		script.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("Identity")]
        [Alias("ID")]
        [Alias("mailbox")]
        [Alias("mail")]
		[String]
$mb=$(Throw "You must specify a Calendar to add a delegate to (e.g., JournalSocialNe@janus.com)."),
        [Parameter(Position = 1, Mandatory = $false)]
        [Alias("user")]
        [Alias("delegate")]
		[String]
$del,
        [Parameter(Position = 2, Mandatory = $false)]
		[String[]]
		[ValidateSet(
			'Cal',
			'Calendar',
			'Contact',
			'Contacts',
			'Inbox',
			'Note',
			'Notes',
			'Task',
			'Tasks'
		)]
$folder=@('Cal','Contact','Inbox','Note','Task')
)

# Initialization
$dllpath = "D:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

[array]$delarray=@()
[array]$results=@()

# Filters and Functions
# Normalize data
$folder=$folder.tolower()

# Error checking
# Exit if the EWS Module is not loaded
if (-not ($global:EWSModule))
	{
	[string]$errorstring="WARNING: EWS module is not loaded, come cmdlets will not work correctly...`nExiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Triggered if $del is populated
if ($del)
	{
# Does $del exist in Exchange?
	$DEXO=get-mailbox -id $del
	$successful=$?
	if ($successful -eq $false)
		{
		[string]$errorstring="Error in Mailbox ID $del, cannot locate mailbox, exiting...`n"
		$Error.add($errorstring)
		write-error ($errorstring) -erroraction stop
		}
# Does $del return only one object?	
	if ((($DEXO | measure-object).count) -ne 1)
		{
		[string]$errorstring="$del does not match a single mailbox, exiting...`n"
		$Error.add($errorstring)
		write-error ($errorstring) -erroraction stop
		}
# EWS requires an smtp address, this makes sure that whatever parameter was assigned to $del, it is now an smtp address	
	$del=$DEXO.PrimarySmtpAddress.tostring()
	[string]$delname=$DEXO.DisplayName
	}

# Does $mb exist in Exchange?
$EXO=get-mailbox -id $mb
$successful=$?
if ($successful -eq $false)
	{
	[string]$errorstring="Error in Mailbox ID $mb, cannot locate mailbox, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# Does $mb return only one object?	
if ((($EXO | measure-object).count) -ne 1)
	{
	[string]$errorstring="$mb does not match a single mailbox, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# EWS requires an smtp address, this makes sure that whatever parameter was assigned to $mb, it is now an smtp address	
$mb=$EXO.PrimarySmtpAddress.tostring()
[string]$target=$EXO.DisplayName

# Main Script
$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

# Try allows you to catch an error. In this case, an error is generated if p-uccas03 is returned because it is a re-direct in http
try
	{
	$service.AutodiscoverUrl($mb)
	}


# Manually assigns the url if there is an error
Catch [system.exception]
 {
	$URI='https://p-ucusxhc04.janus.cap/EWS/Exchange.asmx'
	$service.URL = New-Object Uri($URI)
	write-output ("Caught an Autodiscover URL exception, recovering...")
 }

# You have to use impoersonation in order to access another mailbox
$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $mb);

# We only chek this once, as it is set by calendar, not delegate
If ($folder -like "*cal*")
	{
	$MtgReqSetting=($service.GetDelegates($mb,$true)).MeetingRequestsDeliveryScope
	}

write-output ("`nProcessing: $target")

# GetDelegates() uses a different parameter list if you are returning one or all
if ($del) {$delarray=($service.GetDelegates($mb,$true,$del)).DelegateUserResponses} else
{$delarray=($service.GetDelegates($mb,$true)).DelegateUserResponses}

# $delarray is an array (or collection, I am not sure of the difference) of Delegate responses. The Delegate responses contains all the relevant data on the delegates in the array.
foreach ($delegate IN $delarray)
{
$Name=$delegate.DelegateUser.UserId.DisplayName
$smtp=$delegate.DelegateUser.UserId.PrimarySmtpAddress
write-output("Processing: $Name")
If ($folder -like "*cal*")
	{
	$CalPerms=$delegate.DelegateUser.Permissions.CalendarFolderPermissionLevel
	$RecCopy=$delegate.DelegateUser.ReceiveCopiesOfMeetingMessages
	}
If ($folder -like "*inbox*")
	{
	$InboxPerms=$delegate.DelegateUser.Permissions.InboxFolderPermissionLevel
	}
If ($folder -like "*task*")
	{
	$TaskPerms=$delegate.DelegateUser.Permissions.TasksFolderPermissionLevel
	}
If ($folder -like "*contact*")
	{
	$ContactPerms=$delegate.DelegateUser.Permissions.ContactsFolderPermissionLevel
	}
If ($folder -like "*note*")
	{
	$NotePerms=$delegate.DelegateUser.Permissions.NotesFolderPermissionLevel
	}

# We only check this once, as it is set by mailbox, not delegate
$private=$delegate.DelegateUser.ViewPrivateItems

# We'll create a PSObject with custom properties so that we can control the format of the output
$results+=New-Object PSObject -Property @{
	Mailbox = $EXO
	DelegateDisplayName = $Name
	DelegatePrimarySmtpAddress = $smtp
	MeetingRequestsDeliveryScope = $MtgReqSetting
	CalendarFolderPermissionLevel = $CalPerms
	ReceiveCopiesOfMeetingMessages = $RecCopy
	InboxFolderPermissionLevel = $InboxPerms
	TasksFolderPermissionLevel = $TaskPerms
	ContactsFolderPermissionLevel = $ContactPerms
	NotesFolderPermissionLevel = $NotePerms
	ViewPrivateItems = $private
	}
}
return $results
remove-variable service
remove-variable ACEUser
remove-variable WindowsIdentity
remove-variable delarray
remove-variable EXO
remove-variable DEXO
} # function Get-JEWSDelegates

function Get-JDCs
{
<#
	.SYNOPSIS
		Returns DCs in this domain.

	.DESCRIPTION
		Retreives the current forest and then locates all DCs for that
		forest.

	.PARAMETER  gc
		Instructs this cmdlet to only return Global Catalog Servers

	.EXAMPLE
		Get-DCs
		
		Description
		-----------
		Returns all DCs that are not in the Forest Root domain.
		
	.EXAMPLE
		Get-DCs -gc
		
		Description
		-----------
		Returns one GCS fromthe current domain.
		
	.NOTES
		Retreives the current forest and then locates all DCs for that
		forest.
#>

[CmdletBinding()]             
Param(
        [Alias("gcs")]
[switch]
$gc
)
 
$forest = [system.directoryservices.activedirectory.Forest]::GetCurrentForest()
if ($gc)
	{
	$context = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext("forest","$($forest.name)")
	[System.DirectoryServices.ActiveDirectory.GlobalCatalog]::FindOne($context,"Denver")
	} else
	{
	$forest.domains | where {$_.name -ne $_.forest} | Foreach-Object {$_.DomainControllers} | Foreach-Object {$_.Name}
	}
remove-variable forest
} # function Get-JDCs

function Get-JConstructors
{
<#
	.SYNOPSIS
		Returns all known Constructors of the specified .NET Class.

	.DESCRIPTION
		Returns all known Constructors of the specified .NET Class. This
		can be used to create new Objects or instances of Objects as
		variables.

	.PARAMETER  type
		Enter the Object Type or Class. Do not use the square brackets
		(e.g., []) normally associated with object classes.

	.EXAMPLE
		Get-JConstructors -type DateTime
		
		Description
		-----------
		Returns all known constructors for the DateTime object Class.
		
	.EXAMPLE
		Get-JConstructors -type Microsoft.Exchange.WebServices.Data.ExchangeService
		
		Description
		-----------
		Returns all known Constructors for the Microsoft.Exchange.WebServices.Data.ExchangeService Object Class.
		
	.NOTES
		Retreives the current forest and then locates all DCs for that
		forest.
#>

[CmdletBinding()]             
param(
	[type]
$type
)

    foreach ($ctor in $type.GetConstructors())
    {
        $type.Name + "("
        foreach ($param in $ctor.GetParameters())
        {
        "`t{0} {1}," -f $param.ParameterType.FullName, $param.Name
        }
        ")"
    }
} # function Get-JConstructors

Function Set-JDSACLs
{

<#
.SYNOPSIS
	Sets AD ACLs for the specified AD Entry.

.DESCRIPTION
	Uses ADSI to Set AD ACLs for the specified AD Entry.

.PARAMETER  DE
	Enter the AD ID of the AD Entry to moodify.
	NOTE: Accepts SAMAccountName, Display Name or SMTP Address.

.PARAMETER  ID
	Enter the AD ID of the AD Entry that is gaining access to the AD Entry
	specified in the -DE parameter.
	NOTE: Accepts SAMAccountName, Display Name or SMTP Address.

.PARAMETER  perms
	Specify the permissions to apply.
	NOTE: Accepts:
		CreateChild
		DeleteChild
		ListChildren
		Self
		ReadProperty
		WriteProperty
		DeleteTree
		ListObject
		ExtendedRight
		Delete
		ReadControl
		GenericExecute
		GenericWrite
		GenericRead
		WriteDacl
		WriteOwner
		GenericAll
		Synchronize
		AccessSystemSecurity
		FullControl
		All
		Read
		Write

.PARAMETER  allow
	This switch instructs powershell to allow the permissions granted in the
	-perms parameter.
	NOTE: If this switch is not part of the command line, the permissions in
	-perms will be denied.


.PARAMETER  extended
	Specifies the Extended permissions to grant.
	NOTE: This is used for -perms of ReadProperty, WriteProperty and
	ExtendedRight.

.EXAMPLE
	Set-JDSACLs -de "IT PQM List" -id dgd -perms Read,WriteProperty -allow -extended member
		
	Description
	-----------
	This grants the DGD AD ID permissions to write to the Member Property of the "IT PQM List" AD Object in AD.

.EXAMPLE
	Set-JDSACLs -de "Messaging Team" -id jm27253 -perms Write
		
	Description
	-----------
	This revokes the permission to Write to the "Messaging Team" AD Entry for the JM27253 AD ID.

.NOTES
	Requires the Loading of the Janus Module.

#>
	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("identity")]
        [Alias("directoryentry")]
        [Alias("adobject")]
	[string]
$DE,
        [Parameter(Position = 1, Mandatory = $true)]
        [Alias("user")]
	[string]
$ID,
        [Parameter(Position = 2, Mandatory = $true)]
        [Alias("permissions")]
        [Alias("permission")]
        [Alias("accessrights")]
	[string[]]
	[ValidateSet(
	'CreateChild',
	'DeleteChild',
	'ListChildren',
	'Self',
	'ReadProperty',
	'WriteProperty',
	'DeleteTree',
	'ListObject',
	'ExtendedRight',
	'Delete',
	'ReadControl',
	'GenericExecute',
	'GenericWrite',
	'GenericRead',
	'WriteDacl',
	'WriteOwner',
	'GenericAll',
	'Synchronize',
	'AccessSystemSecurity',
	'FullControl',
	'All',
	'Read',
	'Write'
		)]
$perms,
        [Parameter(Position = 3, Mandatory = $false)]
	[switch]
$allow=$false,
        [Parameter(Position = 4, Mandatory = $false)]
	[string]
$extended=$null
)
 
# Error checking
[System.DirectoryServices.ActiveDirectoryRights[]]$adRights=@()
[System.Security.AccessControl.AccessControlType]$AccessControlType='Deny'
if (($perms -notlike "*ReadProperty*") -and($perms -notlike "*WriteProperty*") -and($perms -notlike "*ExtendedRight*") -and($extended))
	{
	[string]$errorstring="Do not use -extended parameter with the following Perms: $perms, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

if ((($perms -like "*ReadProperty*") -or($perms -like "*WriteProperty*") -or($perms -like "*ExtendedRight*")) -and(!($extended)))
	{
	[string]$errorstring="You must use use -extended parameter with the following Perms: $perms, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

if($extended)
	{
	$schemaIDGUID=(get-adentry -id $extended -schema -exact -InOU "CN=Schema,CN=Configuration,DC=janusadmin,DC=net" -Properties schemaIDGUID -pso).schemaIDGUID
	$strArray=$schemaIDGUID.split(" ")
	$bytArray=[system.byte[]]$strArray
	[guid]$guid=[guid]$bytArray
	}

$ADE=get-adentry -id $DE -exact -pso -properties distinguishedname
$successful=$?
if(-not($successful))
	{
	[string]$errorstring="AD Search error, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

if ((($ADE | measure-object).count) -ne 1)
	{
	$count=(($ADE | measure-object).count)
	[string]$errorstring="$DE does not match a single AD account. $count found, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

$Object=[ADSI]$ADE.adspath

$ADID=get-adentry -id $ID -exact -pso -properties distinguishedname
$successful=$?
if(-not($successful))
	{
	[string]$errorstring="AD Search error, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

if ((($ADID | measure-object).count) -ne 1)
	{
	$count=(($ADID | measure-object).count)
	[string]$errorstring="$ID does not match a single AD account. $count found, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

$IDDE=[ADSI]$ADID.adspath | select @{n='SID';e={(New-Object Security.Principal.NTAccount $_.name).translate( [Security.Principal.Securityidentifier] ).Value }}
[System.Security.Principal.IdentityReference]$IDSID = [System.Security.Principal.SecurityIdentifier]$IDDE.SID

$permissions=$perms
foreach($perm IN $permissions)
	{
	if ($perm -like "Write") {$perm="GenericWrite";$adRights+=$perm}
	elseif($perm -like "Read") {$perm="GenericRead";$adRights+=$perm}
	elseif(($perm -like "All") -or($perm -like "FullControl")){$perm="GenericAll";$adRights+=$perm}
	else {$adRights+=$perm}
	}

if ($allow -eq $true)
	{
	$AccessControlType=[System.Security.AccessControl.AccessControlType]::"Allow"
	} else
	{
	$AccessControlType=[System.Security.AccessControl.AccessControlType]::"Deny"
	}

if ($extended)
	{
	write-output ("Applying Extended permission")
	$AccessRule = new-object System.DirectoryServices.ActiveDirectoryAccessRule($IDSID , $adRights , $AccessControlType , $guid)
	} else
	{
	write-output ("Applying permission")
	$AccessRule = new-object System.DirectoryServices.ActiveDirectoryAccessRule($IDSID , $adRights , $AccessControlType)
	}

$Object.psbase.get_ObjectSecurity().AddAccessRule($AccessRule)
$successful=$?
if($successful)
	{
	$Object.psbase.CommitChanges()
	}

write-output("Success: $successful")

write-output("Current ACLs:")
get-jdsacls -id $($Object.name)
remove-variable ADRights
remove-variable AccessControlType
remove-variable Object
remove-variable ADID
remove-variable IDDE
remove-variable IDSID
} # Function Set-JDSACLs

Function Get-JLogonSessions
{
<#
.SYNOPSIS
    Retrieves all user sessions from local or remote servers

.DESCRIPTION
    Retrieves all user sessions from local or remote server

.PARAMETER computer
    Name of computer to run session query against.

.NOTES
	Requires Janus Module

.LINK
    https://boeprox.wordpress.org

.EXAMPLE
Get-JLogonSessions -computer "p-ucadm01"

Description
-----------
This command will query all current user sessions on 'p-ucadm01'.

#>
[cmdletbinding(
	DefaultParameterSetName = 'session',
	ConfirmImpact = 'low'
)]
    Param(
        [Parameter(
            Mandatory = $True,
            Position = 0,
            ValueFromPipeline = $True)]
            [string[]]$computer
            )
Begin {
    $report = @()
    }
Process {
    ForEach($system in $computer) {
# Uses the MS provided Query.exe and Parses 'query session' and store in $sessions:
        $sessions = query session /server:$system
            1..($sessions.count -1) | % {
# Initalizes the columns in $session
                $session = "" | Select Computer,SessionName, Username, Id, State, Type, Device
# Records the computer name
                $session.Computer = $system
# Records the Session name
                $session.SessionName = $sessions[$_].Substring(1,18).Trim()
# Records the USer name
                $session.Username = $sessions[$_].Substring(19,20).Trim()
# Records Session ID
                $session.Id = $sessions[$_].Substring(39,9).Trim()
# Records Session State
                $session.State = $sessions[$_].Substring(48,8).Trim()
# Records Session Type (console, etc.)
                $session.Type = $sessions[$_].Substring(56,12).Trim()
# Records The Session Device
                $session.Device = $sessions[$_].Substring(68).Trim()
                $report += $session
            }
        }
    }
End {
    $report
	remove-variable report
	remove-variable sessions
    }
} # Function Get-JLogonSessions


Function Get-JWMILogonSessions
{
<#
.SYNOPSIS
    Retrieves all user sessions from local or remote servers

.DESCRIPTION
    Retrieves tall user sessions from local or remote servers

.PARAMETER computer
    Name of computer to run session query against.

.NOTES
	Requires Janus Module loaded.

.LINK
    https://boeprox.wordpress.org

.EXAMPLE
Get-WMILogonSessions -computer "p-ucadm01"

Description
-----------
This command will query all current user sessions on 'p-ucadm01'.

#>
[cmdletbinding(
	DefaultParameterSetName = 'session',
	ConfirmImpact = 'low'
)]
    Param(
        [Parameter(
            Mandatory = $True,
            Position = 0,
            ValueFromPipeline = $True)]
            [string[]]$computer
    )
Begin {
# Create empty report
    $report = @()
    }
Process {
# Iterate through collection of computers
    ForEach ($system in $computer) {
# Get explorer.exe processes
# Retreives all instances of The Task Engine
        $processes = gwmi win32_process -computer $system -Filter "Name = 'taskeng.exe'"
# Retreives all instances of PowerShell
        $processes += gwmi win32_process -computer $system -Filter "Name = 'powershell.exe'"
# Retreives all instances of PowerShell ISE
        $processes += gwmi win32_process -computer $system -Filter "Name = 'powershell_ISE.exe'"
# Retreives all instances of The Command Prompt
        $processes += gwmi win32_process -computer $system -Filter "Name = 'cmd.exe'"
# Retreives all instances of The Scheduled Task Application
        $processes += gwmi win32_process -computer $system -Filter "Name = 'SCHTASKS.exe'"
# Retreives all instances of MMC
        $processes += gwmi win32_process -computer $system -Filter "Name = 'mmc.exe'"
# Go through collection of processes
        ForEach ($process in $processes) {
# This sets up the formatting for those columns
            $result = "" | Select Caption,User
# The gathers the Caption of the EXE
            $result.Caption = $process.Caption
# The gathers the Username associated with the EXE
            $result.User = ($process.GetOwner()).User
            $report += $result
          }
        }
    }
End {
    $report
	remove-variable report
	remove-variable processes
    }
} # Function Get-JWMILogonSessions

function Show-JStatus
{
# Records the Computer name
[string]$computer=gc env:computername
# Records user name
[string]$user=gc env:username
# Current Script
[string]$bparam=$myInvocation.BoundParameters | out-string
[string]$uparam=$myInvocation.UnboundArguments | out-string
[string]$Script=$myInvocation.ScriptName  + " MyCommand Name: " + $myInvocation.MyCommand.Name + " Line: " + $myInvocation.Line + " InvocationName: " + $myInvocation.InvocationName + " BoundParameters: " + $bparam + " UnboundArguments: " + $uparam

$logger=New-JSyslogger -dest_host "p-ucslog02.janus.cap"
$log="Current system: $computer - Current User: $user - Current Script: $Script"
$logger.send($log)

} # function Show-JStatus

function Show-JETPMessageStatus
{
<#
	.SYNOPSIS
		Looks up the status of Blackberry Activation.

	.DESCRIPTION
		Looks up the status of Blackberry Activation.

	.PARAMETER  id
		Enter the SMTP address of the mailbox to check

	.EXAMPLE
		Show-JETPMessageStatus -id Randy.Moore@janus.com
		
		Description
		-----------
		Looks up the blackberry activation status.
		
	.NOTES
		Requires EWS to be installed on the computer that executes this
		script.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("identity")]
        [Alias("mail")]
        [Alias("mailbox")]
        [Alias("primarysmtpaddress")]
		[String]
$id=$(gc env:username)
)

# Maximum e-mails to read
$ResultSize=100
# Set $d to true for debug info
# $d=$true
# Set $d to false to omit debug info
$d=$false

[string]$msgid=""
[string]$sender=""
$etpcount=0
$oldetpcount=0

# Error checking
$ADID=Get-JADEntry -id $id -exact -pso -properties mail
if ($d) {write-host("AD ID = $ADID")}
if($ADID -eq $null -or $ADID -eq "")
	{
	[string]$errorstring="$id does not match a single ID, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}
# $mailbox has to be in smtp format
$mailbox=$ADID.mail
if ($d) {write-host("Mailbox = $mailbox")}

# Instatiate the EWS service object
$service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)
if ($d) {write-host("Service = $service")}
		
# Set the impersonated user id on the service object if required
$ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress),$Mailbox
if ($d) {write-host("Impersonated User ID = $ImpersonatedUserId")}
$service.ImpersonatedUserId = $ImpersonatedUserId
if ($d) {write-host("Impersonated User ID = $($service.ImpersonatedUserId)")}

# Determine the EWS end-point using Autodiscover
try
	{
        $service.AutodiscoverUrl($Mailbox)
	}

# Manually assigns the url if there is an error
Catch [system.exception]
 {
	$URI='https://p-ucusxhc04.janus.cap/EWS/Exchange.asmx'
	$service.URL = New-Object Uri($URI)
	write-output ("Caught an Autodiscover URL exception, recovering...")
 }
if ($d) {write-host("Autodiscover URL = $($service.AutodiscoverUrl)")}

[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]$folderID=[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::"Inbox"
if ($d) {write-host("Folder ID = $folderID")}
[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]$root=[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::"Root"
if ($d) {write-host("Root = $root")}
[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]$MsgFolderRoot=[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::"MsgFolderRoot"
if ($d) {write-host("MSG Folder Root = $MsgFolderRoot")}

# Create a view based on the $ResultSize parameter value
$view = New-Object Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList $ResultSize
if ($d) {write-host("View = $view")}
$fview = New-Object Microsoft.Exchange.WebServices.Data.FolderView -ArgumentList $ResultSize
if ($d) {write-host("FView = $fview")}
		
# Define which properties we want to retrieve from each message
$emailProps = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
if ($d) {write-host("Email Props = $emailProps")}
$emailProps.RequestedBodyType = "Text"
if ($d) {write-host("Email Props = $emailProps")}
$view.PropertySet = $emailProps
if ($d) {write-host("View = $view")}
# Use FindItems method for the specified folder, AQS query and number of messages
$items = $service.FindItems($folderID,$view)
if ($d) {write-host("Items = $($items.count)")}

# Initialize ETP variable
$etp=$false
if ($d) {write-host("etp true? $etp")}

foreach ($item IN $items)
	{
	$subject=$item.Subject
if ($d) {write-host("Subject = $subject")}
	$msgid=$item.InternetMessageId
if ($d) {write-host("MSG ID = $msgid")}
# If the Subject starts with RIM_ and the Internet MEssage ID contains blackberry.net, it is probably an ETP message
	if (($subject -like "RIM_*") -and($msgid -like "*.blackberry.net*"))
		{
# $etp is True when there is an ETP message in their Inbox
		$etp=$true
if ($d) {write-host("etp true? $etp")}
# $etpdate is the date of the newest attempt to activate
		$newetpdate=Get-date ($($item.DateTimeReceived))
if ($d) {write-host("New etp date = $newetpdate")}
		if ($newetpdate -gt $etpdate) {$etpdate=$newetpdate}
if ($d) {write-host("etp date = $etpdate")}
		$etpcount=$etpcount+1
if ($d) {write-host("etp count = $etpcount")}
		}
	}
# This finds the folder above the top of the information store for their mailbox
$folders=$service.findfolders($Root,$null,$fview)
if ($d) {write-host("Folders = $folders")}

# Initializes BBInfo
$BBInfo=$false
if ($d) {write-host("BB Info present? $BBInfo")}

foreach ($folder IN $folders)
	{
# If they have a BlackBerryHandheldInfo folder, then we want to report that. It means their old Sync settings and data are still on the server
	if ($folder.DisplayName -like "*BlackBerryHandheldInfo*") {$BBInfo=$true}
if ($d) {write-host("BB Info present? $BBInfo")}
	}

$rptCollection = @()

## Define Extended Properties

$PR_DELETED_ON = new-object  Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26255, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime)
$PR_DELETED_MSG_COUNT = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26176, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$PR_DELETED_MESSAGE_SIZE_EXTENDED = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26267, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long)
$PR_DELETED_FOLDER_COUNT = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26177, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$PR_Sender_Name = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26177, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)

## End Define Extended Properties

## Define Property Sets
## Folder Set

$fpsFolderPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$fpsFolderPropertySet.add($PR_DELETED_ON)
$fpsFolderPropertySet.add($PR_DELETED_MSG_COUNT)
$fpsFolderPropertySet.add($PR_DELETED_MESSAGE_SIZE_EXTENDED)
$fpsFolderPropertySet.add($PR_DELETED_FOLDER_COUNT)
if ($d) {write-host("fps Folder Property Set = $fpsFolderPropertySet")}

## Item Set
$ipsItemPropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$ipsItemPropertySet.RequestedBodyType = "Text"
$ipsItemPropertySet.add($PR_DELETED_ON)
$ipsItemPropertySet.RequestedBodyType = "Text"
if ($d) {write-host("ips Item Property Set = $ipsItemPropertySet")}

# End Set

# MesgFolderRoot is the top of Information Store for a mailbox
$rfRootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$MsgFolderRoot)
if ($d) {write-host("rf Root Folder = $rfRootFolder")}
$fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(10000);
if ($d) {write-host("fv Folder View = $fvFolderView")}
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow
if ($d) {write-host("fv Folder View = $fvFolderView")}
$fvFolderView.PropertySet = $fpsFolderPropertySet
if ($d) {write-host("fv Folder View = $fvFolderView")}
if ($d) {$service.traceenabled = $true}

# Initalizes $oldetp
$oldetp=$false
if ($d) {write-host("Old etp = $oldetp")}

# This gets the first 10,000 or less folders in the top of their mailbox
$ffResponse = $rfRootFolder.FindFolders($fvFolderView);
if ($d) {write-host("ff Response = $ffResponse")}
foreach ($ffFolder in $ffResponse.Folders){ # foreach loop
if ($d) {write-host("ff Folder = $ffFolder")}
# If it is not the Inbox, we don't want to process it
	if($ffFolder.DisplayName -eq "Inbox") { # In Inbox
	$dcDeleteItemCount = $null
if ($d) {write-host("dc Delete Item Count = $dcDeleteItemCount")}
	$fptProptest = $ffFolder.TryGetProperty($PR_DELETED_MSG_COUNT, [ref]$dcDeleteItemCount) 
if ($d) {write-host("fpt Prop test = $fptProptest")}
	if($fptProptest){ # If Proptest
# $dcDeleteItemCount contains the number of deleted items found
		if ($dcDeleteItemCount -ne 0){ # If delcount
			$ffFolder.DisplayName +  " - Number Items Deleted :" + $dcDeleteItemCount
# Needed to build the view
			$bcBatchCount = 0;
			$bcBatchSize = 100
			$ivItemView = new-object Microsoft.Exchange.WebServices.Data.ItemView($bcBatchSize, $bcBatchCount)
			$ivItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::SoftDeleted
			$ivItemView.PropertySet = $ipsItemPropertySet
if ($d) {write-host("iv Item View = $ivItemView")}
#			$service.traceenabled = $false
			$fiFindItems = $ffFolder.FindItems($ivItemView)
if ($d) {write-host("fi Find Items = $($fiFindItems.count)")}
			foreach ($item in $fiFindItems.Items)
			{ # Foreach Item
				$lnum ++
				write-progress "Processing message" $lnum
				$delon = $null
				$ptProptest = $item.TryGetProperty($PR_DELETED_ON, [ref]$delon) 
# Initalizes the $subject variable
				$subject=$item.Subject
if ($d) {write-host("AD ID = $ADID")}
				$msgid=$item.InternetMessageId
if ($d) {write-host("AD ID = $ADID")}
				$msg=$item
if ($d) {write-host("AD ID = $ADID")}
# If the Subject starts with RIM_ and the Internet MEssage ID contains blackberry.net, it is probably an ETP message
				if (($subject -like "RIM_*") -and($msgid -like "*.blackberry.net*"))
					{ # If ETP
# $oldetp is True when there is an ETP message in their Recovered Deleted Items View
					$oldetp=$true
if ($d) {write-host("Old etp = $oldetp")}
					$newoldetpdate=Get-date ($($item.DateTimeReceived))
if ($d) {write-host("New old etp date = $newoldetpdate")}
# $oldetpdate is the date of the newest attempt to activate
					if($newoldetpdate -gt $oldetpdate) {$oldetpdate=$newoldetpdate}
if ($d) {write-host("Old etp date = $oldetpdate")}
					$oldetpcount=$oldetpcount+1
if ($d) {write-host("Old etp count = $oldetpcount")}
					} # If ETP
			} # Foreach Item
		} # If delcount
	} # If Proptest
} # In Inbox
} # foreach loop

Write-output ("Is ETP message still present in Inbox: $etp")
Write-output ("Number of ETP message(s) still present in Inbox: $etpcount")
Write-output ("Date/Time of ETP message still present in Inbox: $etpdate")
Write-output ("Was an ETP message Processed by the BES: $oldetp")
Write-output ("Number of ETP message(s) Processed by the BES: $oldetpcount")
Write-output ("Date/Time of ETP message that was processed: $oldetpdate")
Write-output ("Is BB Info present: $BBInfo")

remove-variable service -erroraction silentlycontinue -warningaction silentlycontinue
remove-variable items -erroraction silentlycontinue -warningaction silentlycontinue
remove-variable folders -erroraction silentlycontinue -warningaction silentlycontinue
remove-variable ffResponse -erroraction silentlycontinue -warningaction silentlycontinue
remove-variable fiFindItems -erroraction silentlycontinue -warningaction silentlycontinue
remove-variable etp -erroraction silentlycontinue -warningaction silentlycontinue
remove-variable etpdate -erroraction silentlycontinue -warningaction silentlycontinue
remove-variable oldetp -erroraction silentlycontinue -warningaction silentlycontinue
remove-variable oldetpdate -erroraction silentlycontinue -warningaction silentlycontinue
remove-variable BBInfo -erroraction silentlycontinue -warningaction silentlycontinue
} # function Show-JETPMessageStatus

function Remove-JCRDelegate
{
<#
	.SYNOPSIS
		Removes Conference Room Delegates.

	.DESCRIPTION
		Removes Conference Room Delegates. It does this by removing the
		specified ID from the BookInPolicy Field of a Conference Room.

	.PARAMETER id
		Enter the mailbox to modify.

	.PARAMETER user
		Enter the user to remove from the mailbox BookInPolicy Settings.

	.EXAMPLE
		Remove-CRDelegate -id !CR-YY-99-RUTest -user jm27253
		
		Description
		-----------
		Removes JM27253 as a delegate of the !CR-YY-99-RUTest Conference Room.
		
	.NOTES
		Requires the Exchange Module.
#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("identity")]
        [Alias("cr")]
        [Alias("conferenceroom")]
        [Alias("cal")]
        [Alias("calendar")]
		[String]
$id=$(Throw "You must specify a Calendar to modify (e.g., `"!CR-ACP-04-Gemini Peak-CAP20`")."),
        [Parameter(Position = 1, Mandatory = $true)]
        [Alias("delegate")]
        [Alias("del")]
		[String]
$user=$(Throw "You must specify a delegate (e.g., JM27253)."))

if (-not ($global:ExchangeSnapIn))
	{
	[string]$errorstring="WARNING: Exchange module is not loaded, some cmdlets will not work correctly, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

[array]$newdel=@()
$user=$user.tolower()
[string]$dn=""
[string]$mbdn=""
$DC=(Get-DCs -gc).Name
$CDrive="\\" + $DC + "\c$"
$temp=test-path $CDrive
$successful=$?
if($successful -eq $false)
	{
	$DC="p-jcdcd05.janus.cap"
	$CDrive="\\" + $DC + "\c$"
	test-path $CDrive
	$successful=$?
	if($successful -eq $false)
		{
		[string]$errorstring="Error locating a DC, exiting...`n"
		$Error.add($errorstring)
		write-error ($errorstring) -erroraction stop
		}
	}

# Error checking
get-mailboxcalendarsettings -id $id -DomainController $DC | out-null
$successful=$?
if ($successful -eq $false)
	{
	Write-Output ("Calendar lookup error.`n")
	get-mailboxcalendarsettings -resultzise unlimited -DomainController $DC | where {$_.Identity -like "*$id*"}
	$successful=$?
	if ($successful -eq $false)
		{
		[string]$errorstring="Unrecoveravble error, exiting...`n"
		$Error.add($errorstring)
		write-error ($errorstring) -erroraction stop
		}
	$CMB=get-mailboxcalendarsettings -resultzise unlimited  -DomainController $DC | where {$_.Identity -like "*$id*"}
	} else
	{
	$CMB=get-mailboxcalendarsettings -id $id -DomainController $DC
	}

if ((($CMB | measure-object).count) -ne 1)
	{
	[string]$errorstring=$id + " does not match a single Mailbox, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# More error checking
get-mailbox -id $user -DomainController $DC | out-null
$successful=$?
if ($successful -eq $false)
	{
	Write-Output ("Delegate AD lookup error.`n")
	get-mailbox -resultzise unlimited -DomainController $DC | where {Identity -like "*$user*"}
	$successful=$?
	if ($successful -eq $false)
		{
		[string]$errorstring="Unrecoveravble error, exiting...`n"
		$Error.add($errorstring)
		write-error ($errorstring) -erroraction stop
		}
	$DMB=get-mailbox -resultzise unlimited  -DomainController $DC | where {Identity -like "*$user*"}
	} else
	{
	$DMB=get-mailbox -id $user -DomainController $DC 
	}

if ((($DMB | measure-object).count) -ne 1)
	{
	[string]$errorstring=$user + " does not match a single Mailbox, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

[string]$userID=$DMB.SAMAccountName

# Save old settings
$olddel=$CMB.bookinpolicy
$bookinpolicy=[string]$olddel
$bookinpolicy=$bookinpolicy.tolower()

if ($bookinpolicy -notlike "*$userID*")
	{
	[string]$errorstring=$user + " is not a Delegate, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

$newdel=@()

foreach ($mbd IN $olddel)
	{
	$disabled=$false
	$dn=[string]$mbd.DistinguishedName
	$booldisabled=$dn.contains("Disabled")
	if (-not($booldisabled)) {$mb=get-mailbox -id $mbd -DomainController $DC;$disabled=$false} else {$mbdn="disabled";$disabled=$true}
	if($mb -ne $null) {$mbdn=[string]$mb.DistinguishedName} else {$mbdn="disabled";$disabled=$true}
	$booldisabled=$mbdn.contains("Disabled")
	if (-not($booldisabled)) {$disabled=$true}
	if (($mb.samaccountname -notlike $userid) -and($disabled -ne $true)) {$newdel+=$mbd}
	}

Set-mailboxcalendarsettings -id $id -bookinpolicy $newdel -DomainController $DC -erroraction silentlyContinue
$successful=$?
if ($successful -eq $false)
	{
	[string]$errorstring="Error adding Delegate, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}
Write-Output ("Execution successful.`n")
remove-variable newdel
remove-variable olddel
remove-variable CMB
remove-variable DMB
remove-variable UserID
remove-variable BookinPolicy
} # function Remove-JCRDelegate

Function New-JSyslogger
{
<#
	.SYNOPSIS
		Creates a Syslog object you can send syslog messagess to.

	.DESCRIPTION
		Creates a Syslog object for the syslog server specified. You
		can use the .send() method of that object to send messages to
		the Syslog server.

	.PARAMETER  dest_host
		Specify the syslog server to send syslog messages to.

	.EXAMPLE
		$logger=New-JSyslogger -dest_host "p-ucslog02.janus.cap"
		[string]$MI=$MyInvocation | fl | out-string
		[string]$data="TEST $MI"
		$logger.Send("$data")
		
		Description
		-----------
		Sends information about the currently running script to p-ucslog02.janus.cap.
		
	.EXAMPLE
		$logger=New-JSyslogger -dest_host "p-ucslog02.janus.cap"
		[string]$data=$error
		$logger.Send("$data","p-ucadm01","cron","alert")

		Description
		-----------
		Sends the current PowerShell error log to p-ucslog02.janus.cap.
		
		NOTE: The Send() Method accepts the following paramters:
 -data <String>
	This is the actual Log Message. 

        Required?                    true
        Position?                    1
        Default value
        Accept pipeline input?       false
        Accept wildcard characters?  False

 -host <String>
	This is the source computer of the Logged Message. 

        Required?                    false
        Position?                    2
        Default value                FQDN of the localhost
        Accept pipeline input?       false
        Accept wildcard characters?  False

 -facility <String>
	This is the "Facility" that generated the log (Valid values are: 'kern',
	'user', 'mail', 'daemon', 'security', 'auth', 'syslog', 'lpr', 'news',
	'uucp', 'cron', 'authpriv', 'ftp', 'ntp', 'logaudit', 'logalert',
	'clock', 'local0', 'local1', 'local2', 'local3', 'local4', 'local5',
	'local6', or 'local7').

        Required?                    false
        Position?                    3
        Default value                "user"
        Accept pipeline input?       false
        Accept wildcard characters?  False

 -severity <String>
	This is the alert level placed in the log (Valid values are: 'emerg',
	'panic', 'alert', 'crit', 'error', 'err', 'warning', 'warn', 'notice',
	'info', or 'debug'). 

        Required?                    false
        Position?                    4
        Default value                "info"
        Accept pipeline input?       false
        Accept wildcard characters?  False
		
	.NOTES
		Requires Janus Module Loaded.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("computer")]
        [Alias("system")]
        [Alias("sys")]
        [Alias("DestinationHost")]
		[String]
$dest_host = "p-ucslog02.janus.cap"
)

# Creates a new Syslog Server Object    
$SSO = New-Object PSObject

# Adds a ._UdpClient Property to the Syslog Server Object    
$SSO | Add-Member -MemberType NoteProperty -Name _UdpClient -Value $null

# Adds a .init() Method to the Syslog Server Object    
$SSO | Add-Member -MemberType ScriptMethod -Name init -Value {
	param
	(
	[String]
	$dest_host = "p-ucslog02.janus.cap",
	[Int32]
	$dest_port = 514
	)
	$this._UdpClient = New-Object System.Net.Sockets.UdpClient
# the .init() method allows you to specify a new port
	$this._UdpClient.Connect($dest_host, $dest_port)
	}
    
# Adds a .send() Method to the Syslog Server Object    
    $SSO | Add-Member -MemberType ScriptMethod -Name Send -Value {
	param
	(
	[String]
	$data = $(throw "Error SyslogSenderUdp:init; Log data must be given."),
	[String]
	$hostname = $(gc env:computername),
	[String]
	[ValidateSet(
		'kern',
		'user',
		'mail',
		'daemon',
		'security',
		'auth',
		'syslog',
		'lpr',
		'news',
		'uucp',
		'cron',
		'authpriv',
		'ftp',
		'ntp',
		'logaudit',
		'logalert',
		'clock',
		'local0',
		'local1',
		'local2',
		'local3',
		'local4',
		'local5',
		'local6',
		'local7'
		)]
	$facility = "user",
	[String]
	[ValidateSet(
		'emerg',
		'panic',
		'alert',
		'crit',
		'error',
		'err',
		'warning',
		'warn',
		'notice',
		'info',
		'debug'
		)]
	$severity = "info"
	)

# Facility is defined by the Syslog Protocol spcification
	$facility_map = @{
	"kern" = 0;
	"user" = 1;
	"mail" = 2;
	"daemon" = 3;
	"security" = 4;
	"auth" = 4;
	"syslog" = 5;
	"lpr" = 6;
	"news" = 7;
	"uucp" = 8;
	"cron" = 9;
	"authpriv" = 10;
	"ftp" = 11;
	"ntp" = 12;
	#"logaudit" = 13;
	#"logalert" = 14;
	"clock" = 15;
	"local0" = 16;	
	"local1" = 17;
	"local2" = 18;
	"local3" = 19;
	"local4" = 20;
	"local5" = 21;
	"local6" = 21;
	"local7" = 23;
	}

# Secerity is defined by the Syslog Protocol spcification
	$severity_map = @{
	"emerg" = 0;
	"panic" = 0;
	"alert" = 1;
	"crit" = 2;
	"error" = 3;
	"err" = 3;
	"warning" = 4;
	"warn" = 4;
	"notice" = 5;
	"info" = 6;
	"debug" = 7;
	}

# Map the text to the decimal value
	if ($facility_map.ContainsKey($facility))
	{
	$facility_num = $facility_map[$facility]
	}
	else
	{
	$facility_num = $facility_map["user"]
	}
        
	if ($severity_map.ContainsKey($severity))
	{
	$severity_num = $severity_map[$severity]
	}
	else
	{
	$severity_num = $severity_map["user"]
	}

# Calculate the PRI code
	$pri = ($facility_num * 8) + $severity_num

# Get a properly formatted, encoded, and length limited data string
# Replaces partial carriage returns with space
	$data = $data.replace("`f"," ")
# Replaces carriage returns with space
	$data = $data.replace("`n"," ")
# Replaces partial carriage returns with space
	$data = $data.replace("`r"," ")
# Replaces tabs with space
	$data = $data.replace("`t"," ")
# Replaces vertical tabs with space
	$data = $data.replace("`v"," ")
# Removes double spaces
do
	{
	$data = $data.replace("  "," ")
	} while ($data.contains("  "))

# Formats message in Syslog format
	$message = "<{0}>{1} {2}" -f $pri, $hostname, $data

# Converts string to ascii bytes
	$enc     = [System.Text.Encoding]::ASCII        $message = $Enc.GetBytes($message)

# Cuts off log entry at 1024 characters
	if ($message.Length -gt 1024)
	{
	$message = $message.SubString(0, 1024)
	}
        
# Sends actual message
	$this._UdpClient.Send($message, $message.Length) | out-null
	}

# Initializes new Syslog Server Object
	$SSO.init($dest_host)
    
# Returns Syslog Server Object for use by Function Caller
	$SSO
remove-variable SSO
} # Function New-JSyslogger

function Remove-JADGroupMembership
{
<#
	.SYNOPSIS
		Removes an AD Account from a Group(s).

	.DESCRIPTION
		Removes an AD Account from a Group(s). Optionally removing
		every group from the specified user.

	.PARAMETER  id
		Specify the user to remove from the group.

	.PARAMETER  group
		Specify the group to remove the suer from.
		NOTE: If this parameter is not specified, this cmdlet will
		remove every group from the specified user.

	.EXAMPLE
		Remove-JADGroupMembership -id JM27253 -group "Messaging Team"
		
		Description
		-----------
		Removes the JM27253 account from the Messaging Team Group.
		
	.EXAMPLE
		Remove-JADGroupMembership -id jm27253
		
		Description
		-----------
		Removes the JM27253 account from every group they are a member of.
		
	.NOTES
		Requires Janus Module Loaded.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("user")]
        [Alias("member")]
        [Alias("members")]
[string]
$ID,
        [Parameter(Position = 1, Mandatory = $false)]
        [Alias("identity")]
[string]
$group=$null
)

# Initialization
[string]$GroupDN=""

# Error Checking
$User=get-jadentry -id $ID -exact -pso -properties memberof
$success=$?
if($success)
	{
	if ((($User | measure-object).count) -ne 1)
	{
	$success=$false
	} # if ((($User | measure-object).count) -ne 1)	
	} # if($success)

if(-not($success))
	{
	$errorstring="AD Lookup Error for $ID, exiting..."
	$error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

# More Error Checking
if (($group -eq $null) -or($group -eq ""))
	{
	$GroupDN=(get-jadentry -id $group -exact -pso -properties distinguishedname).distinguishedname
	$success=$?
	if($success)
		{
		if ((($GroupDN | measure-object).count) -ne 1)
		{
		$success=$false
		} # if ((($GroupDN | measure-object).count) -ne 1)
		} # if($success)
	
	if(-not($success))
		{
		$errorstring="AD Lookup Error for $group, exiting..."
		$error.add($errorstring)
		write-error ($errorstring) -erroraction stop
		}
	} # if (($group -eq $null) -or($group -eq ""))

if (($group -eq $null) -or($group -eq ""))
	{
	write-output("Removing $ID from all groups...")
	start-sleep -s 1
	$oldgroups=$User.memberOf | fl | out-string
	write-output ("Previous Group Membership:`n$oldgroups`n")
	start-sleep -s 1
	ForEach ($GroupDN In $User.memberOf)
		{
		$ADGrp = [ADSI]("LDAP://" + $GroupDN)
		$ADGrp.Remove($User.ADsPath)
		} # ForEach ($GroupDN In $User.memberOf)
	} else # if ($group -eq $null)
	{
	write-output("Remnoving $ID from $group")
	$GroupDN=(get-jadentry -id $group -exact -pso -properties distinguishedname).distinguishedname
	$GroupDN="LDAP://" + $GroupDN
	$ADGrp = [ADSI]$GroupDN
	$ADGrp.Remove($User.ADsPath)
	} # else
write-output("Operation Complete!")
remove-variable ADGrp
remove-variable User
} # function Remove-JADGroupMembership

Function Get-JPassword
{
[CmdletBinding()]
param(
[Parameter(Position = 0, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the file name")]
	[string]
$file="\\nas01.janus.cap\installs$\server\Messaging Team\Scripts\Logs\12k12mRM.log",
[Parameter(Position = 1, Mandatory = $false)]
	[switch]
$encrypt=$false,
[Parameter(Position = 1, Mandatory = $false)]
	[switch]
$nocrypt=$false
)

[array]$collection=(0..255)
[char[]]$password=""
[int]$lastchar=256
[int]$char=0
[int]$nextchar=0
[string[]]$newpassword=""
[string]$crypt=""

function pseudohash {
[CmdletBinding()]
param(
[Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the first parameter")]
	[int]
$key,
[Parameter(Position = 1, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the second parameter")]
	[char]
$value
)

$pseudohash=new-object PSObject
$pseudohash | add-member -membertype NoteProperty -name "Code" -value "$Key"
$pseudohash | add-member -membertype NoteProperty -name "Char" -value "$Value"
$pseudohash
} # function pseudohash

# Error checking
$test=test-path -path $file
if (!($test))
	{
	$errorstring="Cannot locate file ($file), exiting..."
	$error.add($errorstring)
	write-error("$errorstring") -erroraction stop
	}

$collection[9]=pseudohash 9 "`t"
$collection[10]=pseudohash 10 "`n"
$collection[13]=pseudohash 13 "`r"
$collection[32]=pseudohash 32 " "
$collection[33]=pseudohash 33 "!"
$collection[34]=pseudohash 34 "`""
$collection[35]=pseudohash 35 "#"
$collection[36]=pseudohash 36 "$"
$collection[37]=pseudohash 37 "%"
$collection[38]=pseudohash 38 "&"
$collection[39]=pseudohash 39 "'"
$collection[40]=pseudohash 40 '('
$collection[41]=pseudohash 41 ')'
$collection[42]=pseudohash 42 "*"
$collection[43]=pseudohash 43 "+"
$collection[44]=pseudohash 44 ","
$collection[45]=pseudohash 45 "-"
$collection[46]=pseudohash 46 "."
$collection[47]=pseudohash 47 "/"
$collection[48]=pseudohash 48 "0"
$collection[49]=pseudohash 49 "1"
$collection[50]=pseudohash 50 "2"
$collection[51]=pseudohash 51 "3"
$collection[52]=pseudohash 52 "4"
$collection[53]=pseudohash 53 "5"
$collection[54]=pseudohash 54 "6"
$collection[55]=pseudohash 55 "7"
$collection[56]=pseudohash 56 "8"
$collection[57]=pseudohash 57 "9"
$collection[58]=pseudohash 58 ":"
$collection[59]=pseudohash 59 ";"
$collection[60]=pseudohash 60 "<"
$collection[61]=pseudohash 61 "="
$collection[62]=pseudohash 62 ">"
$collection[63]=pseudohash 63 "?"
$collection[64]=pseudohash 64 "@"
$collection[91]=pseudohash 91 "["
$collection[92]=pseudohash 92 "\"
$collection[93]=pseudohash 93 "]"
$collection[94]=pseudohash 94 "^"
$collection[95]=pseudohash 95 "_"
$collection[96]=pseudohash 96 "``"
$collection[123]=pseudohash 123 "{"
$collection[124]=pseudohash 124 "|"
$collection[125]=pseudohash 125 "}"
$collection[126]=pseudohash 126 "~"
$collection[65]=pseudohash 65 "A"
$collection[66]=pseudohash 66 "B"
$collection[67]=pseudohash 67 "C"
$collection[68]=pseudohash 68 "D"
$collection[69]=pseudohash 69 "E"
$collection[70]=pseudohash 70 "F"
$collection[71]=pseudohash 71 "G"
$collection[72]=pseudohash 72 "H"
$collection[73]=pseudohash 73 "I"
$collection[74]=pseudohash 74 "J"
$collection[75]=pseudohash 75 "K"
$collection[76]=pseudohash 76 "L"
$collection[77]=pseudohash 77 "M"
$collection[78]=pseudohash 78 "N"
$collection[79]=pseudohash 79 "O"
$collection[80]=pseudohash 80 "P"
$collection[81]=pseudohash 81 "Q"
$collection[82]=pseudohash 82 "R"
$collection[83]=pseudohash 83 "S"
$collection[84]=pseudohash 84 "T"
$collection[85]=pseudohash 85 "U"
$collection[86]=pseudohash 86 "V"
$collection[87]=pseudohash 87 "W"
$collection[88]=pseudohash 88 "X"
$collection[89]=pseudohash 89 "Y"
$collection[90]=pseudohash 90 "Z"
$collection[97]=pseudohash 97 "a"
$collection[98]=pseudohash 98 "b"
$collection[99]=pseudohash 99 "c"
$collection[100]=pseudohash 100 "d"
$collection[101]=pseudohash 101 "e"
$collection[102]=pseudohash 102 "f"
$collection[103]=pseudohash 103 "g"
$collection[104]=pseudohash 104 "h"
$collection[105]=pseudohash 105 "i"
$collection[106]=pseudohash 106 "j"
$collection[107]=pseudohash 107 "k"
$collection[108]=pseudohash 108 "l"
$collection[109]=pseudohash 109 "m"
$collection[110]=pseudohash 110 "n"
$collection[111]=pseudohash 111 "o"
$collection[112]=pseudohash 112 "p"
$collection[113]=pseudohash 113 "q"
$collection[114]=pseudohash 114 "r"
$collection[115]=pseudohash 115 "s"
$collection[116]=pseudohash 116 "t"
$collection[117]=pseudohash 117 "u"
$collection[118]=pseudohash 118 "v"
$collection[119]=pseudohash 119 "w"
$collection[120]=pseudohash 120 "x"
$collection[121]=pseudohash 121 "y"
$collection[122]=pseudohash 122 "z"
$collection[128]=pseudohash 128 "€"
$collection[129]=pseudohash 129 ""
$collection[130]=pseudohash 130 "‚"
$collection[131]=pseudohash 131 "ƒ"
$collection[132]=pseudohash 132 "`„"
$collection[133]=pseudohash 133 "…"
$collection[134]=pseudohash 134 "†"
$collection[135]=pseudohash 135 "‡"
$collection[136]=pseudohash 136 "ˆ"
$collection[137]=pseudohash 137 "‰"
$collection[138]=pseudohash 138 "Š"
$collection[139]=pseudohash 139 "‹"
$collection[140]=pseudohash 140 "Œ"
$collection[141]=pseudohash 141 ""
$collection[142]=pseudohash 142 "Ž"
$collection[143]=pseudohash 143 ""
$collection[144]=pseudohash 144 ""
$collection[145]=pseudohash 145 "‘"
$collection[146]=pseudohash 146 "’"
$collection[147]=pseudohash 147 "`“"
$collection[148]=pseudohash 148 "`”"
$collection[149]=pseudohash 149 "•"
$collection[150]=pseudohash 150 "–"
$collection[151]=pseudohash 151 "—"
$collection[152]=pseudohash 152 "˜"
$collection[153]=pseudohash 153 "™"
$collection[154]=pseudohash 154 "š"
$collection[155]=pseudohash 155 "›"
$collection[156]=pseudohash 156 "œ"
$collection[157]=pseudohash 157 ""
$collection[158]=pseudohash 158 "ž"
$collection[159]=pseudohash 159 "Ÿ"
$collection[160]=pseudohash 160 " "
$collection[161]=pseudohash 161 "¡"
$collection[162]=pseudohash 162 "¢"
$collection[163]=pseudohash 163 "£"
$collection[164]=pseudohash 164 "¤"
$collection[165]=pseudohash 165 "¥"
$collection[166]=pseudohash 166 "¦"
$collection[167]=pseudohash 167 "§"
$collection[168]=pseudohash 168 "¨"
$collection[169]=pseudohash 169 "©"
$collection[170]=pseudohash 170 "ª"
$collection[171]=pseudohash 171 "«"
$collection[172]=pseudohash 172 "¬"
$collection[173]=pseudohash 173 "­"
$collection[174]=pseudohash 174 "®"
$collection[175]=pseudohash 175 "¯"
$collection[176]=pseudohash 176 "°"
$collection[177]=pseudohash 177 "±"
$collection[178]=pseudohash 178 "²"
$collection[179]=pseudohash 179 "³"
$collection[180]=pseudohash 180 "´"
$collection[181]=pseudohash 181 "µ"
$collection[182]=pseudohash 182 "¶"
$collection[183]=pseudohash 183 "•"
$collection[184]=pseudohash 184 "¸"
$collection[185]=pseudohash 185 "¹"
$collection[186]=pseudohash 186 "º"
$collection[187]=pseudohash 187 "»"
$collection[188]=pseudohash 188 "¼"
$collection[189]=pseudohash 189 "½"
$collection[190]=pseudohash 190 "¾"
$collection[191]=pseudohash 191 "¿"
$collection[192]=pseudohash 192 "À"
$collection[193]=pseudohash 193 "Á"
$collection[194]=pseudohash 194 "Â"
$collection[195]=pseudohash 195 "Ã"
$collection[196]=pseudohash 196 "Ä"
$collection[197]=pseudohash 197 "Å"
$collection[198]=pseudohash 198 "Æ"
$collection[199]=pseudohash 199 "Ç"
$collection[200]=pseudohash 200 "È"
$collection[201]=pseudohash 201 "É"
$collection[202]=pseudohash 202 "Ê"
$collection[203]=pseudohash 203 "Ë"
$collection[204]=pseudohash 204 "Ì"
$collection[205]=pseudohash 205 "Í"
$collection[206]=pseudohash 206 "Î"
$collection[207]=pseudohash 207 "Ï"
$collection[208]=pseudohash 208 "Ð"
$collection[209]=pseudohash 209 "Ñ"
$collection[210]=pseudohash 210 "Ò"
$collection[211]=pseudohash 211 "Ó"
$collection[212]=pseudohash 212 "Ô"
$collection[213]=pseudohash 213 "Õ"
$collection[214]=pseudohash 214 "Ö"
$collection[215]=pseudohash 215 "×"
$collection[216]=pseudohash 216 "Ø"
$collection[217]=pseudohash 217 "Ù"
$collection[218]=pseudohash 218 "Ú"
$collection[219]=pseudohash 219 "Û"
$collection[220]=pseudohash 220 "Ü"
$collection[221]=pseudohash 221 "Ý"
$collection[222]=pseudohash 222 "Þ"
$collection[223]=pseudohash 223 "ß"
$collection[224]=pseudohash 224 "à"
$collection[225]=pseudohash 225 "á"
$collection[226]=pseudohash 226 "â"
$collection[227]=pseudohash 227 "ã"
$collection[228]=pseudohash 228 "ä"
$collection[229]=pseudohash 229 "å"
$collection[230]=pseudohash 230 "æ"
$collection[231]=pseudohash 231 "ç"
$collection[232]=pseudohash 232 "è"
$collection[233]=pseudohash 233 "é"
$collection[234]=pseudohash 234 "ê"
$collection[235]=pseudohash 235 "ë"
$collection[236]=pseudohash 236 "ì"
$collection[237]=pseudohash 237 "í"
$collection[238]=pseudohash 238 "î"
$collection[239]=pseudohash 239 "ï"
$collection[240]=pseudohash 240 "ð"
$collection[241]=pseudohash 241 "ñ"
$collection[242]=pseudohash 242 "ò"
$collection[243]=pseudohash 243 "ó"
$collection[244]=pseudohash 244 "ô"
$collection[245]=pseudohash 245 "õ"
$collection[246]=pseudohash 246 "ö"
$collection[247]=pseudohash 247 "÷"
$collection[248]=pseudohash 248 "ø"
$collection[249]=pseudohash 249 "ù"
$collection[250]=pseudohash 250 "ú"
$collection[251]=pseudohash 251 "û"
$collection[252]=pseudohash 252 "ü"
$collection[253]=pseudohash 253 "ý"

$password=get-content $file

if($encrypt)
	{
	foreach ($character IN $password)
		{
		if($lastchar -eq 256) {$lastchar=0}
		foreach ($lookup IN $collection) {if ($lookup.char -ceq $character) {$char=$lookup.code} }
		$nextchar=$lastchar+$char
		if ($nextchar -gt 255) {$nextchar=$nextchar-255}
		$crypt = "{0:D3}" -f $nextchar
		$newpassword=[string]$newpassword+[string]$crypt
		$lastchar=$char
		out-file -file $file -inputobject $newpassword -force -confirm:$false
		} # foreach ($character IN $password)
	} else
	{
	$length=$password.length-1
	for ($character=0; $character -le $length; $character=$character+3)
		{
		if($lastchar -eq 256) {$lastchar=0}
		$char=([int][string]$password[$character] * 100) + ([int][string]$password[$character+1] * 10) + ([int][string]$password[$character+2])
		$nextchar=$char-$lastchar
		if ($nextchar -lt 0) {$nextchar=$nextchar + 255}
		$crypt=$collection[$nextchar].Char
		$newpassword=[string]$newpassword+[char]$crypt
		$lastchar=$nextchar
		} # for ($character=0; $character -le $length; $character=$character+3)
	} # if($encrypt)

if($nocrypt)
	{
	$newpassword
	} else
	{
	$spass=ConvertTo-SecureString -string $newpassword -asplaintext -force
	$spass
	} # if($nocrypt)

Remove-Variable collection
Remove-Variable password
Remove-Variable lastchar
Remove-Variable char
Remove-Variable nextchar
Remove-Variable newpassword
Remove-Variable crypt
Remove-Variable spass

# $collection |  select Code,Char, @{expression={$_.SyncRoot -join ";"};label="SyncRoot"} -excludeproperty SyncRoot | ft

} # Function Get-JPassword

function Start-SOAPRequest
{
<#
	.SYNOPSIS
		Send an XML SOAP request to the specified url.

	.DESCRIPTION
		Send an XML SOAP request to the specified url.

	.PARAMETER  SOAPRequest
		Properly formatted XML for the SOAP server you are connecting to.

	.PARAMETER  url
		URL od the SOAP server you are connecting to.

	.EXAMPLE
		Start-SOAPRequest -SOAPRequest $XMLRequest -url https://workflow.stg.myjonline.com/gsoap/gsoap_ssl.dll?sbmappservices72
		
		Description
		-----------
		Submits the XML in the $XMLRequest to the url as a SOAP request.
		
	.NOTES
		Requires Janus Module Loaded.

#>
[CmdletBinding()]
param(
[Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the XML for the request")]
	[xml]
$SOAPRequest,
[Parameter(Position = 1, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the file name")]
	[string]
$URL='https://workflow.myjonline.com/gsoap/gsoap_ssl.dll?sbmappservices72'
)
	write-output "Sending SOAP Request To Server: $URL"
	$soapWebRequest = [System.Net.WebRequest]::Create($URL)
	$soapWebRequest.Headers.Add("SOAPAction","`"`"")

	$soapWebRequest.ContentType = "text/xml;charset=`"utf-8`""
	$soapWebRequest.Accept = "text/xml"
	$soapWebRequest.Method = "POST"
	$soapWebRequest.UseDefaultCredentials = $true
	write-output "Initiating Send."
	$requestStream = $soapWebRequest.GetRequestStream()
	$SOAPRequest.Save($requestStream)
	$requestStream.Close()
	write-output "Send Complete, Waiting For Response."
	$resp = $soapWebRequest.GetResponse()
	$responseStream = $resp.GetResponseStream()
	$soapReader = [System.IO.StreamReader]($responseStream)
	$ReturnXml = $soapReader.ReadToEnd()
	$responseStream.Close()
	write-output "Response Received."
	return $ReturnXml
} # function Execute-SOAPRequest

function Add-JSerenaTicket
{
<#
	.SYNOPSIS
		Creates a ticket in Serena using SOAP and XML.

	.DESCRIPTION
		Creates a ticket in Serena using SOAP and XML.

	.PARAMETER  description
		This is the description of the issue being reported.

	.PARAMETER routing
		Specify the primary Routing Group for this ticket.

	.PARAMETER  product
		Specify the product for the issue being reported.

	.EXAMPLE
		Add-JSerenaTicket -description "Error retreiving IM Manager Data, please investigate."
		
		Description
		-----------
		Creates an Exchange ticket that is routed to the Messaging Team with the description above.
		
	.EXAMPLE
		Add-JSerenaTicket -description "Offline Adress Book Error, please investigate." -routing "IT Service Desk" -product Outlook
		
		Description
		-----------
		Creates an Outlook ticket that is routed to the IT Service Desk with the description above.
		
	.NOTES
		Requires Janus Module Loaded.

#>

[CmdletBinding()]
param(
[Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the XML for the request")]
	[string]
$description,
[Parameter(Position = 1, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the file name")]
	[string]
$routing="Messaging Team",
[Parameter(Position = 2, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the file name")]
	[string]
$product="Exchange"
)

$pass=get-JPassword -nocrypt

$XMLString=@"
<soapenv:Envelope
xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
xmlns:urn="urn:sbmappservices72">
<soapenv:Header/>
<soapenv:Body>

    <urn:CreatePrimaryItem>
        <!-- userId and password will change between environments -->
        <urn:auth>
            <urn:userId>janus_cap\bpp_exchangeadmin</urn:userId>         <!-- We'll create a specific ID for the process -->
            <urn:password>$pass</urn:password>
        </urn:auth>

        <urn:project>
            <urn:fullyQualifiedName>Incidents</urn:fullyQualifiedName>
        </urn:project>

        <urn:parentItem>
        </urn:parentItem>

            <urn:item>
                <urn:description>              <!-- enter the description here -->
		$description
                </urn:description>

                <urn:extendedField>
                    <urn:id>
                        <urn:dbName>CONTACT</urn:dbName>
                    </urn:id>
                    <urn:setValueBy>DISPLAY-VALUE</urn:setValueBy>
                    <urn:setValueMethod>REPLACE-VALUES</urn:setValueMethod>
                    <urn:value>
                        <urn:displayValue>Dave Michael</urn:displayValue>  <!-- User's Name goes here.  This can be anyone or the "System" name submitting the ticket -->
                    </urn:value>
                </urn:extendedField>

                <!-- -->
                <!-- The OPEN_AND_ASSIGN field is used if the PRIMARY_ROUTING_GROUP is known so the request gets automatically routed. -->
                <!-- If the PRIMARY_ROUTING_GROUP is known, then the displayValue = '(Checked)', otherwise '(Not Checked)'             -->
                <!-- -->
                <urn:extendedField>
                    <urn:id>
                        <urn:dbName>OPEN_AND_ASSIGN</urn:dbName>
                    </urn:id>
                    <urn:setValueBy>DISPLAY-VALUE</urn:setValueBy>
                    <urn:setValueMethod>REPLACE-VALUES</urn:setValueMethod>
                    <urn:value>
                        <urn:displayValue>(Checked)</urn:displayValue>   <!-- Update as specified above, either '(Checked)' or '(Not Checked)' -->
                    </urn:value>
                </urn:extendedField>

                <!-- -->
                <!-- This should always be 'Software', but could change depending on the TECHNOLOGY_SERVICE -->
                <!-- -->
                <urn:extendedField>
                    <urn:id>
                        <urn:dbName>TYPE_OF_TECHNOLOGY_SERVICE</urn:dbName>
                    </urn:id>
                    <urn:setValueBy>DISPLAY-VALUE</urn:setValueBy>
                    <urn:setValueMethod>REPLACE-VALUES</urn:setValueMethod>
                    <urn:value>
                        <urn:displayValue>Software</urn:displayValue>
                    </urn:value>
                </urn:extendedField>

                <urn:extendedField>
                    <urn:id>
                        <urn:dbName>TECHNOLOGY_SERVICE</urn:dbName>
                    </urn:id>
                    <urn:setValueBy>DISPLAY-VALUE</urn:setValueBy>
                    <urn:setValueMethod>REPLACE-VALUES</urn:setValueMethod>
                    <urn:value>
                        <urn:displayValue>$product</urn:displayValue>   <!-- Enter the APM name here -->
                    </urn:value>
                </urn:extendedField>

                <!-- -->
                <!-- The potential PRIMARY_ROUTING_GROUP values will need to be setup in Courion. -->
                <!-- If the Administrator of a system is not known, use the default value of 'Client Services'.  -->
                <!-- -->
                <urn:extendedField>
                    <urn:id>
                        <urn:dbName>PRIMARY_ROUTING_GROUP</urn:dbName>
                    </urn:id>
                    <urn:setValueBy>DISPLAY-VALUE</urn:setValueBy>
                    <urn:setValueMethod>REPLACE-VALUES</urn:setValueMethod>
                    <urn:value>
                        <urn:displayValue>$routing</urn:displayValue>  <!-- Enter the Primary Routing Group name here -->
                    </urn:value>
                </urn:extendedField>

                <urn:extendedField>
                    <urn:id>
                        <urn:dbName>SEND_UPDATE_TO_SUBMITTER</urn:dbName>
                    </urn:id>
                    <urn:setValueBy>DISPLAY-VALUE</urn:setValueBy>
                    <urn:setValueMethod>REPLACE-VALUES</urn:setValueMethod>
                    <urn:value>
                        <urn:displayValue>(Not Checked)</urn:displayValue>   <!-- Update as specified above, either '(Checked)' or '(Not Checked)' -->
                    </urn:value>
                </urn:extendedField>

            </urn:item>

        <urn:options>
            <urn:sections>SECTIONS-NONE</urn:sections>    <!-- If needed, look for issueId in return value to get SBM ID of incident-->
        </urn:options>
    </urn:CreatePrimaryItem>

</soapenv:Body>
</soapenv:Envelope>
"@

[Xml]$XMLTicket=$XMLString

$ReturnXML=Start-SOAPRequest -SOAPRequest $XMLTicket -URL 'https://workflow.stg.myjonline.com/gsoap/gsoap_ssl.dll?sbmappservices72'

$ReturnXML -match "`<ae`:issueId`>(?<content>.*)`<`/ae`:issueId`>" | out-null
$TicketNumber=$matches['content']
if($TicketNumber -eq $null -or $TicketNumber -eq "")
	{
	Send-MailMessage -SmtpServer mailman.janus.cap -From "bpp_appmom.bpp_appmom@janus.com" -To "bpp_ExchangeAdmin@janus.com" -Subject "Warning - Audit finding" -Body "Powershell failed to create a ticket in Serena."
	} else
	{
	Send-MailMessage -SmtpServer mailman.janus.cap -From "Exchange.Administrator@janus.com" -To "Exchange.Administrator@janus.com" -Subject "On Call Notification - Serena Ticket created" -Body "PowerShell created a Serena Ticker number INCDT-$TicketNumber. You should be receiving notification from Serena automatically."
	}
} # function Add-JSerenaTicket

function Set-JOOFMessage
{
<#
	.SYNOPSIS
		Sets the OOF Message of the specified user.

	.DESCRIPTION
		Sets the OOF Message of the specified user. It also sets Custom
		Attribue 8, and removes any delivery restrictions on this
		mailbox.

	.PARAMETER  ID
		Enter the Mailbox to modify.

	.PARAMETER  message
		Enter the message recipients will see.

	.PARAMETER  expires
		Enter the Date that this OOF message should stop being sent.

	.EXAMPLE
		Set-JOOFMessage -id jm27253 -message "Randy is in the office, but is testing out of office features" -expires 12-21-12
		
		Description
		-----------
		Sets the Out of Office Message for JM27253.
		
	.NOTES
		Requires the Loading of the Janus Module.
		Requires the Loading of the Exchange Module.
		Requires the Loading of the EWS Module.

#>
	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the ID of the mailbox to modify")]
		[String]
        [Alias("Identity")]
        [Alias("mail")]
        [Alias("mb")]
        [Alias("mailbox")]
$id=$(Throw "You must specify an AD ID (e.g., JM27253)."),
        [Parameter(Position = 1, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the message to send")]
		[String]
$message=$(Throw "You must specify an outgoing message (e.g., Randy is not currently available)."),
        [Parameter(Position = 2, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the third parameter")]
        [Alias("expirationdate")]
		[datetime]
$expires=$((get-date).adddays(31))
)

# Functions and Filters


# Main Script
# The Begin section executes once regardless of how many objects are passed through the pipeline
begin
{
if ($global:JanusPSModule -ne $true) {Import-Module Janus -erroraction stop}
Show-JStatus

# Initialization
[string]$log=""
# This is used by EWS to prevent it from trying to use the wrong feature set/overload
$ver = "Exchange2007_SP1"
$syslogobject=New-JSyslogger -dest_host "p-ucslog02.janus.cap"

# Error checking
# If EWS is not loaded, stop processsing
if (-not ($global:EWSModule))
	{
	[string]$errorstring="WARNING: EWS module is not loaded, come cmdlets will not work correctly...`nExiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	} # if (-not ($global:EWSModule))

# If the Exchange snapin is not loaded, stop processing
if (-not ($global:ExchangeSnapIn))
	{
	[string]$errorstring="WARNING: Exchange module is not loaded, some cmdlets will not work correctly, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	} # if (-not ($global:ExchangeSnapIn))

# Does the mailbox exist
$MB=Get-mailbox -id $ID -erroraction silentlycontinue
$success=$?
if ($success -eq $true -and ((($MB | measure-object).count) -eq 1))
	{
# If so, what is it's SMTP address
	$mail=$MB.PrimarySmtpAddress
	} else
	{
	write-error "$ID is not a valid mailbox name, exiting..." -erroraction stop
	}

# Is $expires a valid date
get-date($expires) | out-null
$success=$?
if ($success -eq $true)
	{
	$CA8 = get-date($expires)
	} else
	{
	write-error "$expires is not a valid date, exiting..." -erroraction stop
	}

# Used for EWS Validation
        $sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        $user = [ADSI]"LDAP://<SID=$sid>"
} # begin

# The process section runs once for each object in the pipeline
process
{
# Sets Custom Attribute 8
set-mailbox -id $MB -CustomAttribute8 $CA8 -ExternalOofOptions External -RejectMessagesFrom $null -RejectMessagesFromDLMembers $null -AcceptMessagesOnlyFromDLMembers $null -AcceptMessagesOnlyFrom $null -RequireSenderAuthenticationEnabled $false -Confirm:$false | out-null
$success=$?
if ($success -eq $false)
	{
	[string]$errorstring="Error setting Custom Attribute 8, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}
# Gets today's date
$now = [DateTime]::Today
# Creates a time window object that has the duration between today and the expiration date.
$timeWindow = new-object Microsoft.Exchange.WebServices.Data.TimeWindow($now, $CA8)

# Establishes a service connection to EWS
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -Arg $ver
# Try allows you to catch an error. In this case, an error is generated if p-uccas03 is returned because it is a re-direct in http
try
	{
	$service.AutodiscoverUrl($mail)
	}

# Manually assigns the url if there is an error
Catch [system.exception]
 {
	$URI='https://p-ucusxhc04.janus.cap/EWS/Exchange.asmx'
	$service.URL = New-Object Uri($URI)
	write-output ("Caught an Autodiscover URL exception, recovering...")
 }

# This is a custom object EWS uses to set Out of Office settings
$OOFSettings = new-object Microsoft.Exchange.WebServices.Data.OofSettings
# Scheduled State allows you to set a start and end date
$OOFSettings.State=[Microsoft.Exchange.WebServices.Data.OofState]::Scheduled
$OOFSettings.Duration =$timeWindow
$OOFSettings.Duration.EndTime = $expires.date
$OOFSettings.Duration.StartTime=$now
$OOFSettings.InternalReply=$message
$OOFSettings.ExternalReply=$message
$OOFSettings.ExternalAudience=[Microsoft.Exchange.WebServices.Data.OofExternalAudience]::All

# Actually sets the OOF
$oof = $service.SetUserOofSettings($mail,$OOFSettings)
} # process

# The End section executes once regardless of how many objects are passed through the pipeline
end
{
remove-variable now
remove-variable service
remove-variable OOFSettings
remove-variable oof -erroraction silentlycontinue
remove-variable log
remove-variable syslogobject
} # end
} # function Set-JOOFMessage

function Set-JMailboxForwarding
{
<#
	.SYNOPSIS
		Sets up forwarding of mail from one mailbox to another.

	.DESCRIPTION
		Sets up forwarding of mail from one mailbox to another. Also
		sets Custom Attribute 7, removes all Delivery Restrictions and
		sets it to keep a copy and forward a copy.

	.PARAMETER  ID
		Enter the mailbox to modify.

	.PARAMETER  user
		Enter the Mailbox to receive the forwarded mail.

	.PARAMETER  expires
		Enter the date that this forwarding should end.
		NOTE: This will default to 31 days from the time of execution
		if it is not specified.

	.EXAMPLE
		Set-JMailboxForwarding -id jm27253 -user DGD -expires 12-21-12
		
		Description
		-----------
		Sets the JM27253 mailbox to keep a copy of mail sent to it and
		forward a copy to the DGD mailbox.
		
	.NOTES
		Requires the Loading of the Janus Module.
		Requires the Loading of the Exchange Module.

#>
	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the ID of the mailbox to be modified")]
		[String]
        [Alias("Identity")]
        [Alias("mb")]
        [Alias("mail")]
        [Alias("mailbox")]
$id,
        [Parameter(Position = 1, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the ID of the mailbox to receive mail moving forward")]
		[String]
        [Alias("Foreward")]
        [Alias("fwd")]
$user,
        [Parameter(Position = 2, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the third parameter")]
        [Alias("expirationdate")]
		[datetime]
$expires=$((get-date).adddays(31))
)

# Functions and Filters


# Main Script
# The Begin section executes once regardless of how many objects are passed through the pipeline
begin
{
if ($global:JanusPSModule -ne $true) {Import-Module Janus -erroraction stop}
Show-JStatus

# Initialization
$now=get-date

# If the Exchange snapin is not loaded, stop processing
if (-not ($global:ExchangeSnapIn))
	{
	[string]$errorstring="WARNING: Exchange module is not loaded, some cmdlets will not work correctly, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	} # if (-not ($global:ExchangeSnapIn))

# Does the mailbox exist
$MB=Get-mailbox -id $ID -erroraction silentlycontinue
$success=$?
if ($success -eq $true -and ((($MB | measure-object).count) -eq 1))
	{
# If so, what is it's SMTP address
	$mail=$MB.PrimarySmtpAddress
	} else
	{
	write-error "$ID is not a valid mailbox name, exiting..." -erroraction stop
	}

# Does the mailbox exist
$FWD=Get-JADentry -id $user -properties mail,msexchrecipientdisplaytype -pso -exact -erroraction silentlycontinue
$success=$?
if (($success -eq $true) -and($FWD.msexchrecipientdisplaytype -ne $null) -and ((($FWD | measure-object).count) -eq 1))
	{
# If so, what is it's SMTP address
	$fwdmail=$FWD.mail
	} else
	{
	write-error "$user is not a valid e-mail object, exiting..." -erroraction stop
	}

# Is $expires a valid date
get-date($expires) | out-null
$success=$?
if ($success -eq $true -and $expires -gt $now)
	{
	$CA7 = get-date($expires)
	} else
	{
	write-error "$expires is not a valid date, exiting..." -erroraction stop
	}

} # begin

# The process section runs once for each object in the pipeline
process
{
set-mailbox -id $MB -CustomAttribute7 $CA7 -RejectMessagesFrom $null -RejectMessagesFromDLMembers $null -AcceptMessagesOnlyFromDLMembers $null -AcceptMessagesOnlyFrom $null -RequireSenderAuthenticationEnabled $false -DeliverToMailboxAndForward $true -ForwardingAddress $fwdmail -Confirm:$false | out-null
$success=$?
if ($success -eq $false)
	{
	[string]$errorstring="Error setting up forwarding, exiting...`n"
	$Error.add($errorstring)
	write-error ($errorstring) -erroraction stop
	}

} # process

# The End section executes once regardless of how many objects are passed through the pipeline
end
{

remove-variable now
remove-variable MB
remove-variable FWD
remove-variable mail
remove-variable fwdmail
remove-variable CA7
} # end
} # function Set-JMailboxForwarding

Function Show-JServices
{

# **********************************************************************************
#
# Script Name: Show-JServices.ps1
# Version: 1.0
# Author: Dennis Kendrick
# Date Created: 03-26-2012
# _______________________________________
#
# MODIFICATIONS:
# Date Modified: N/A
# Modified By: N/A
# Reason for modification: N/A
# What was modified: N/A
# Description: This script gathers the services that the startup is set to Automatic and displays the status.
#
# Usage:
# Show-JServices -servername "P-UCADM01"
# $MyServers = Get-Content D:\Scripts\ServerList.txt
# $MyServers | ./Show-JServices
# **********************************************************************************

<#
	.SYNOPSIS
		Shows services that are set to Automatic start that are not running.  When the srvaccount switch is used it will show the services using the matching account.

	.DESCRIPTION
		Gathers the services that the startup is set to Automatic and displays the status.

	.PARAMETER  servername
		Specify the computer to display the services of.

	.EXAMPLE
		Show-JServices -servername "P-UCADM01"

		Description
		===========
		Lists all services set to automatically start on P-UCADM01.

	.EXAMPLE
		$MyServers = Get-Content D:\Scripts\ServerList.txt
		$MyServers | ./Show-JServices
        Show-JServices -servername P-UCEV01 -srvaccount bpp		

		Description
		===========
		Lists all services set to automatically start on each server listed in the D:\Scripts\ServerList.txt file.

	.NOTES
		Requires the Janus Module.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the server name.")]
        [Alias("Identity")]
        [Alias("computer")]
		[string]$servername=$(gc env:computername),
        
        [Parameter(Mandatory=$false,ValueFromPipeline=$false)]
        [string[]]$srvaccount = $null,
        
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [array[]]$serverlist = $null
        )
# Functions and Filters


# Main Script
# The Begin section executes once regardless of how many objects are passed through the pipeline
begin
{
     
     #[array]$serverlist = $_ 
}


# Initialization


# The process section runs once for each object in the pipeline
process
{
     
    if($srvaccount -ne $null)
    
    {
        Test-Path -Path "\\$servername\c$"| Out-Null
        write-Host -ForegroundColor Green "=========[$servername]-Services running under Service Account-$srvaccount================================================"
        $srvaccount = "*" + $srvaccount + "*"
        $results = Get-WmiObject win32_Service -ComputerName $serverName  -EnableAllPrivileges -Authentication 6 | where {$_.startname -like $srvaccount }
        
            if($results -eq $null)
            {
                Write-Host -ForegroundColor Magenta "No results found for $srvaccount."
            }
            
            else
            {
                $results
            }
    
    }
    
    if($serverlist -ne $null)
    {
        foreach($server in $serverlist)
        {
            Test-Path -Path "\\$servername\c$"| Out-Null
            write-Host -ForegroundColor Green "=========[$servername] - Services set to Automatic but Stopped================================================"
            Get-WmiObject win32_Service -ComputerName $server  -EnableAllPrivileges -Authentication 6 | where {$_.StartMode -eq "Auto" -and $_.State -ne "Running"} |  Select DisplayName,Name,StartMode,State | ft -AutoSize -Wrap    
             
        }
    
    }
    
    else
    {
        Test-Path -Path "\\$servername\c$"| Out-Null
        Write-Host -ForegroundColor Green "=========[$servername] - Services set to Automatic but Stopped================================================"
        Get-WmiObject win32_Service -ComputerName $serverName  -EnableAllPrivileges -Authentication 6 | where {$_.StartMode -eq "Auto" -and $_.State -ne "Running"} |  Select DisplayName,Name,StartMode,State #| ft -AutoSize -Wrap
      
    }

}

# The End section executes once regardless of how many objects are passed through the pipeline
end
{    
    #Cleans up the variable $servername from memory
    remove-variable servername
}

} # Function Show-JServices

Function Show-JADLockoutStatus
{

<#
	.SYNOPSIS
		Shows lockout status of AD account across all DCs.

	.DESCRIPTION
		Shows lockout status of the specified AD account on each DC.

	.PARAMETER  id
		Specify the AD account to check.

	.EXAMPLE
		./Show-JADLockoutStatus.ps1 -id jm27253

		Description
		===========
		Shows the lockout status of the AD Account JM27253 on all DCs.
		
	.NOTES
		Required loading of the Janus Module.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the AD ID to check")]
		[String]
        [Alias("Identity")]
$ID,
		[Switch]
$NoEvents
)

# Functions and Filters


# Main Script
# The Begin section executes once regardless of how many objects are passed through the pipeline
begin
{

# Initialization
[string[]]$DCs=@()
[array]$Hits=@()

$DCs=Get-JDCs
}

# The process section runs once for each object in the pipeline
process
{
foreach ($DC IN $DCs)
	{
	$Hit=Get-JADEntry -id $ID -exact -pso -DC $DC -properties name,displayname,accountexpires,useraccountcontrol,modifytimestamp,lastlogon,lastlogoff,badpwdcount,badpasswordtime,lockouttime,pwdlastset
	$PSO=New-Object PSObject -Property @{
		DC = [string]$DC
		Name = $Hit.name
		DisplayName = $Hit.displayname
		AccountExpires = $Hit.accountexpires -as [system.datetime]
		UserAccountControl = $Hit.useraccountcontrol
		PasswordLastSet = $Hit.pwdlastset
		ModifyTimestamp = $Hit.modifytimestamp -as [system.datetime]
		LastLogon = $Hit.lastlogon -as [system.datetime]
		LastLogoff = $Hit.lastlogoff -as [system.datetime]
		BadPwdCount = $Hit.badpwdcount -as [INT]
		BadPasswordTime = $Hit.badpasswordtime -as [system.datetime]
		LockoutTime = $Hit.lockouttime -as [system.datetime]
        } # New-Object PSObject

	$Hits+=$PSO
	}
}

# The End section executes once regardless of how many objects are passed through the pipeline
end
{
write-output ("`n`nAccount: $($Hits[0].DisplayName)")
$Hits=$Hits | sort BadPasswordTime -descending
$Hits | fl

write-output ("Searching for Failure Audit Events on $($Hits[0].DC) for ID $ID ...")
if (!($NoEvents)) {Show-JFailureAuditEvents -DC $Hits[0].DC -ID $ID}

remove-variable DCs
remove-variable DC
remove-variable PSO
remove-variable Hit
remove-variable Hits
}
} # Function Show-JADLockoutStatus

Function Remove-JCRMeetingbyOrganizer
{

<#
	.SYNOPSIS
		This script Removes appointments where the Organizer is the same as the specified ID.

	.DESCRIPTION
		This script Removes appointments where the Organizer is the same as the specified ID.

	.PARAMETER  CR
		Specify the Conference Room account to check.

	.PARAMETER  Organizer
		Specify the Organizer to look for.

	.PARAMETER  d
		Request extra debug information.

	.EXAMPLE
		Remove-JCRMeetingbyOrganizer -CR !CR-YY-99-RUTest -organizer Randy.Moore@janus.com

		Description
		===========
		Removes all meetings from the !CR-YY-99-RUTest Conference Room where "Randy.Moore@janus.com" is the Organizer.
		
	.NOTES
		Requires loading of the Exchange and EWSMail Modules.

#>

	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the Conference Room.")]
        [Alias("Identity")]
        [Alias("id")]
        [Alias("Mailbox")]
[string]
$CR=$null,
        [Parameter(Position = 1, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the Organizer.")]
        [Alias("user")]
        [Alias("org")]
[string]
$Organizer,
        [Parameter(Position = 2, Mandatory = $false)]
[switch]
$d=$false
)

$error.clear()

[string]$admin="bpp_ExchangeAdmin@janus.com"

if (-not($global:ExchangeSnapIn))
	{
	write-output("WARNING: Exchange module is not loaded, some cmdlets will not work correctly...`n")
	write-output $error -erroraction stop
	}

if (-not ($global:EWSModule))
	{
	write-output("WARNING: EWS module is not loaded, come cmdlets will not work correctly...`n")
	write-output $error -erroraction stop
	}

# Initialization
# $target is a flag that indicates that the meeting should be deleted
[bool]$target=$false
if ($d){write-output "Initial Delete Flag: $target"}
[string]$id=""
[string]$temp=""
[string]$filter=""
$results=@()
$pending=@()
$resultscsv=@()

# Error checking
# Does this CR exist?
$test=get-mailbox -id $CR
$success=$?
if ($d){write-output "Conference Room Mailbox: $test"}
if ($success -eq $false) {write-output ("Cannot locate Conference room named $CR");write-error("Cannot locate Conference room named $CR") -erroraction stop}

# Make sure $CR only matches one mailbox
$hits=($test | measure-object).count
if ($d){write-output "Number of matching mailboxes: $hits.count"}
if ($hits -ne 1)  {write-output ("$CR does not match a single conference room");write-error("$CR does not match a single conference room") -erroraction stop}

# Make sure that $CR is a conference room
if (($test.RecipientTypeDetails -ne "RoomMailbox") -and ($test.RecipientTypeDetails -ne "EquipmentMailbox")) {write-output ("$CR is not a conference Room");write-error("$CR is not a conference Room or Equipment, exiting...") -erroraction stop}
$cal=$test
if ($d){write-output "Current Calendar: $cal"}

# Test Organizer
$test=get-mailbox -id $Organizer
$success=$?
if ($d){write-output "Organizer: $Organizer"}
if ($success -eq $false) {write-output ("Cannot locate Organizer mailbox named $Organizer");write-error("Cannot locate Organizer mailbox named $Organizer") -erroraction silentlycontinue}

# Make sure $Organizer only matches one mailbox
$hits=($test | measure-object).count
if ($d){write-output "Number of matching mailboxes: $hits"}
if ($hits -ne 1)  {write-output ("$Organizer does not match a single mailbox");write-error("$CR does not match a single conference room") -erroraction silentlycontinue}

if ($hits -eq 0)
	{
	$chars=[regex]::IsMatch($Organizer,"[a-z_\@\.\!\#\$\%\&\'\*\+\-\/\=\?\^\`\{\|\}\~\@
A-Z0-9]")
	if ($chars -eq $false) {write-error ("$Organizer contains illegal characters, exiting...") -erroraction stop}
	$array=$Organizer.split("@")
	if ($array.count -gt 2) {write-error ("$Organizer contains too many `"`@`" symbols, exiting...") -erroraction stop}
	if ($Organizer.contains("..")) {write-error ("$Organizer contains `"`.`.`", exiting...") -erroraction stop}
	if ($Organizer.StartsWith(".")) {write-error ("$Organizer starts with `"`.`", exiting...") -erroraction stop}
	if ($Organizer.EndsWith(".")) {write-error ("$Organizer ends with `"`.`", exiting...") -erroraction stop}
	$id=$Organizer
	} else
	{
	$id=$test.PrimarySmtpAddress
	}

$id=$id.tolower()

if ($d){write-output "ID: $id"}

# Main Script
# Creates a log file to log the scripts activities
new-item "D:\Scripts\Logs\Remove-CRMeetingbyOrganizer-$CR.LOG" -type file -force

# What is the CR's displayname
$name=$cal.displayname
write-host (write-output "Processing $name")

# We need the smtp address to access EWS functions
$mailbox=$cal.PrimarySmtpAddress
if ($d){write-output "SMTP: $mailbox"}

# This creates an array of meetings
$meetings=Get-EWSMailMessage -mailbox $mailbox -folder Calendar -ResultSize 100000
$successful=$?
if ($d){$count=$meetings.count;write-output "Meeting Count: $count"}
$LOG="Get-EWSMailMessage -mailbox $mailbox -folder Calendar -ResultSize 100000 executed, did it complete successfully: " + $successful
out-file -file "D:\Scripts\Logs\Remove-CRMeetingbyOrganizer-$CR.LOG" -inputobject $LOG -append

# Goes through each meeting in the specified calendar
foreach ($meeting IN $meetings)
	{
# $target is a flag that indicates that the meeting should be deleted
	[bool]$target=$false
if ($d){write-output "Initial Delete Flag: $target"}

if ($d){write-output "Meeting details: $meeting"}
	$ICalUid=$meeting.ICalUid
	$successful=$?
	write-output ("Processing ICalUid")
	$LOG="Processing ICalUid"
	out-file -file "D:\Scripts\Logs\Remove-CRMeetingbyOrganizer-$CR.LOG" -inputobject $LOG -append

# Algorithm to look up AD account of organizer is on the next 4 lines
	$temp=[string]$meeting.organizer
	$temp=$temp.tolower()
if ($d){write-output "Organizer String: $temp"}

if ($temp.contains("$id")) {$target=$true} else {$target=$false}

# If the $target flag is true, we want to act on this meeting.
	if ($target -eq $true)
	{
	$results+=$meeting
	$MDel="Room: " + $meeting.mailbox + "Organizer: " + $meeting.Organizer + "Sent: " + $meeting.Sent + "End Date: " + $meeting.End + "Reccurrance Pattern: " + $meeting.Recurrence + "Last Recurrance: " + $meeting.LastOccurrence + "Meeting Message ID: " + $meeting.ICalUid
	$LOG="Deleting $MDel"
	out-file -file "D:\Scripts\Logs\Remove-CRMeetingbyOrganizer-$CR.LOG" -inputobject $LOG -append
	} # if ($target -eq $true)
	} # foreach ($meeting IN $meetings)

# Generate a CSV of the items to be deleted
$results | select Mailbox,Organizer,Subject,End,Recurrence,LastOccurrence,Id,ICalUid,InternetMessageId, @{expression={$_.SyncRoot -join ";"};label="SyncRoot"} -excludeproperty SyncRoot | export-csv "D:\Scripts\Logs\DeletedMeetingsLog`($Organizer`).csv" -NoTypeInformation
$successful=$?
$LOG="export-csv `"D:\Scripts\Logs\DeletedMeetingsLog`($Organizer`).csv`" -NoTypeInformation - Successful: $successful"
out-file -file "D:\Scripts\Logs\Remove-CRMeetingbyOrganizer-$CR.LOG" -inputobject $LOG -append

# Report info
$meetings |  select Mailbox,Organizer,Subject,End,Recurrence,LastOccurrence,Id,ICalUid,InternetMessageId, @{expression={$_.SyncRoot -join ";"};label="SyncRoot"} -excludeproperty SyncRoot | export-csv "D:\Scripts\Logs\All_Meetings`($CR`).csv" -NoTypeInformation

$LOG="Error Log: " + $error
out-file -file "D:\Scripts\Logs\Remove-CRMeetingbyOrganizer-$CR.LOG" -inputobject $LOG -append

# This performs the actual deletion, we will have to remove the "#" in order to get it to work in the final version
$results | Remove-EWSMailMessage -Mailbox $mailbox -confirm:$false
} # Function Remove-JCRMeetingbyOrganizer

Function Show-JDiskInfo
{

<#
	.SYNOPSIS
		Shows drive information for logical drives and mouint points.

	.DESCRIPTION
		Shows drive information for logical drives and mouint points of the specified computer.

	.PARAMETER  computer
		Specify the Conference Room account to check.

	.EXAMPLE
		Show-JDiskInfo -computer p-ucusxm01

		Description
		===========
		Shows drive information for logical drives and mouint points of p-ucusxm01.
		
	.NOTES
		Requires loading of the Exchange and EWSMail Modules.

#>

	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the name of the system to inspect.")]
        [Alias("Identity")]
        [Alias("id")]
        [Alias("sys")]
        [Alias("ComputerName")]
[string]
$computer=$(gc env:computername)
)

#Get-WmiObject Win32_Volume -ComputerName $computer | Select Name, Capacity, FreeSpace, BootVolume, SystemVolume, FileSystem | FT -auto -wrap
Get-WmiObject Win32_Volume -filter "DriveType=3"-ComputerName $computer | 
Select SystemName,Name,@{Name="Size (GB)";Expression={"{0:N2}" -f($_.capacity/1gb)}},
@{Name="Free (GB)";Expression={"{0:N2}" -f($_.freespace/1gb)}},
@{Name="Used (GB)";Expression={"{0:N2}" -f(($_.capacity/1gb) - ($_.freespace/1gb))}},
@{Name="% Free";Expression={"{0:N2}" -f(($_.freespace/1gb)/($_.capacity/1gb)*100)}},
BootVolume,FileSystem | sort-object Name | ft -wrap -autosize


} # Function Show-JDiskInfo

function Get-JHotFixbyDate
{

<#
	.SYNOPSIS
		Shows Hot Fixes installed after the specified date.

	.DESCRIPTION
		Shows Hot Fixes installed on the specified computer after the specified date.

	.PARAMETER  computer
		Specify the Conference Room account to check.

	.PARAMETER  date
		Specify the install date to search after.

	.EXAMPLE
		Get-JHotFixbyDate -computer p-ucadm02 -date 2/29/2012

		Description
		===========
		Shows Hot Fixes installed on p-ucadm02 after 2/29/2012.
		
	.NOTES
		Requires loading of the Janus Module.

#>

	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the name of the system to inspect.")]
        [Alias("Identity")]
        [Alias("id")]
        [Alias("sys")]
        [Alias("ComputerName")]
[string]
$computer=$(gc env:computername),
        [Parameter(Position = 1, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the date to search after.")]
        [Alias("InstallDate")]
[string]
$date=$((get-date).adddays(-32)),
        [Alias("debugenabled")]
[switch]
$d
)

# Initialization
$HFarray=@()
[string]$strdate=""
$intdate=(get-date).tofiletime()
$installdate=(get-date).tofiletime()
if ($d) {write-output ("Requested Install Date: $installdate")}

# Error checking
$testpath=test-path \\$computer\c$
if ($d) {write-output ("Is computer ($computer) reachable: $testpath")}
if ($testpath -eq $false)
	{
	write-output ("Cannot locate $computer")
	write-error("Cannot locate $computer") -erroraction stop
	}

$testdate=get-date ($date) -erroraction silentlycontinue
$success=$?
if ($d) {write-output ("Is $date a valid date: $success")}
if ($success -eq $false)
	{
	write-output ("$date is invalid, exiting...")
	write-error("$date is invalid, exiting...") -erroraction stop
	}

# Main Code
$checkdate=(get-date("$date")).tofiletime()
$HFs=Get-HotFix -computer $computer
if ($d) {write-output ("Hot Fixes Found: $($HFs.count)")}
if ($d) {write-output ("$HFs")}
foreach($HF IN $HFs)
	{
	$strdate=$HF.psbase.properties["installedOn"].Value
if ($d) {write-output ("Intalled On as a String - $strdate")}
	$intdate=[Convert]::ToInt64("$strdate", 16)
if ($d) {write-output ("Installed On as an INT64 - $intdate")}
	$installdate=[datetime]::FromFileTime("$intdate")
if ($d) {write-output ("Installed On as a Date - $installdate")}
if ($d) {write-output ("Install Date - $intdate")}

	if($intdate -ge $checkdate)
		{
		$FixComments=$HF.psbase.properties["FixComments"].Value
		$Name=$HF.psbase.properties["Name"].Value
		$ServicePackInEffect=$HF.psbase.properties["ServicePackInEffect"].Value
		$Status=$HF.psbase.properties["Status"].Value

	        $HFO=New-Object PSObject -Property @{
                    Name=$Name
                    KBURL = $HF.Caption
                    ComputerName = $HF.CSName
                    Description = $HF.Description
                    FixComments = $FixComments
                    ID = $HF.HotFixID
                    InstallDate = $installdate
                    InstalledBy = $HF.InstalledBy
                    ServicePackInEffect = $ServicePackInEffect
                    Status = $Status
                    } # New-Object PSObject

		$HFArray+=$HFO
		remove-variable HFO
		if ($d) {write-output ("Relevant Hot Fix - ");$HFO | fl}
		} # if($intdate -ge $installdate)
	} # foreach($HF IN $HFs)
Remove-Variable HFs
Remove-Variable strdate
Remove-Variable intdate
Remove-Variable installdate
return $HFArray
} # function Get-JHotFixbyDate

function Show-JSearchEstimate
{
<#
	.SYNOPSIS
		Creates estimates on the specified search as well as a tracking spreadsheet
		and mailbox.

	.DESCRIPTION
		Takes input and return estimates on the specified search. IT displays this
		estimate on screen. It also creates a spreadsheet anmed after the Search ID
		(e.g., JAN12-1) with the estimate numbers and fields fro tracking the acrual
		results. Then it creates a Search mailbox using the User-Setup.ps1 script.

	.PARAMETER  mbs
		Specify the number of mailboxes searched.

	.PARAMETER  terms
		Use this switch to indicate if search terms are being used.

	.PARAMETER  report
		Use this switch to run the script normally, but not create a spreadsheet or mailbox.

	.PARAMETER  d
		Use this switch to specify that you want debug information output.

	.EXAMPLE
		Show-SearchEstimate -mbs 6 -start 1-1-1970 -end 12-21-12
		Description
		===========
		Returns the estimated sizes and times required to export 6 mailboxes from 1-1-1970 to 12-21-12.
		
	.EXAMPLE
		Show-JSearchEstimate -mbs 6 -start 1-1-1970 -end 12-21-12 -terms
		Description
		===========
		Returns the estimated sizes and times required to search 6 mailboxes with specific term from 1-1-1970 to 12-21-12.
		
	.NOTES

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the first parameter")]
		[INT]
        [Alias("MailboxCount")]
        [Alias("Mailboxes")]
        [Alias("Custodians")]
$MBS=1,
        [Parameter(Position = 1, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the second parameter")]
		[DateTime]
        [Alias("StartDate")]
$start=$null,
        [Parameter(Position = 2, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the second parameter")]
		[DateTime]
        [Alias("EndDate")]
$end=$null,
        [Parameter(Position = 3, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the second parameter")]
		[Switch]
        [Alias("LimitedSearch")]
        [Alias("SearchTerms")]
$terms=$false,
        [Parameter(Position = 4, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the second parameter")]
		[Switch]
$report=$false,
        [Parameter(Position = 5, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the second parameter")]
		[Switch]
        [Alias("DebugEnabled")]
$d=$false
)

# Functions and Filters


# Main Script
# The Begin section executes once regardless of how many objects are passed through the pipeline
begin
{
if ($global:JanusPSModule -ne $true) {Import-Module Janus -erroraction stop}

# Initialization
$error.clear()
$TW=$false
[datetime]$twstart=0
$twdays=0
$twbulk=0
$twcount=0
$twsize=0
$twtime=0
$EV=$false
[datetime]$evstart=0
$evdays=0
$evbulk=0
$evcount=0
$evize=0
$evtime=0
$OR=$false
[datetime]$orstart=0
$ordays=0
$orbulk=0
$orcount=0
$orsize=0
$ortime=0
$OE=$false
[datetime]$oestart=0
$oedays=0
$oebulk=0
$oecount=0
$oesize=0
$oetime=0
$ZD=$false
[datetime]$zdstart=0
$zddays=0
$zdbulk=0
$zdcount=0
$zdsize=0
$zdtime=0
$ddcount=0
$ddtime=0
$exsize=0
$extime=0
$totaltime=0
$totalhours=0
[INT]$intRow=1
[INT]$intCol=1
[string]$strDate=""
[string]$dates=""
[string]$numformat='_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
[INT]$searchNumber=0

# Set Search ID
$strDate=(get-date).ToLongDateString()
$arrayDate=$strDate.split(", ")
do
	{
	$SearchNumber=$searchnumber+1
	$search=$arrayDate[2].substring("0","3") + $arrayDate[5].substring(2,2) + "-" + $searchnumber
	$search=$search.toupper()
	[string]$filename=$search + ".xlsx"
	$filename=$filename.replace("`\","-")
	$filename=$filename.replace("`/","-")
	$filename=$filename.replace("`*","_")
	$filename=$filename.replace("`?","_")
	$filename=$filename.replace("`:","_")
	$filename=$filename.replace("`"","_")
	$filename=$filename.replace("`>","_")
	$filename=$filename.replace("`<","_")
	$filename=$filename.replace("`|","_")
	$filename="\\janus.cap\groups`$\IT\Exchange\Reports\SearchLogs\" + $filename
	$nextnumber=test-path $filename
	} while ($nextnumber -eq $true)

# Error Checking
if ($MBS -lt 1)
	{
	write-output ("$MBS is not a valid number of mailboxes, exiting...")
	write-error ("$MBS is not a valid number of mailboxes, exiting...") -erroraction stop
	}

if (($start -gt $end) -or($end -lt 0))
	{
	write-output ("$MBS is not a valid number of mailboxes, exiting...")
	write-error ("$MBS is not a valid number of mailboxes, exiting...") -erroraction stop
	} else
	{
	$days=($end-$start).days
	} # if (($start -gt $end) -or($end -lt 0))
} # begin

# The process section runs once for each object in the pipeline
process
{

# Tumbleweed section
$twstart=$(get-date("2/14/2001"))
$twend=$(get-date("8/31/2005"))
if (($start -le $twend) -and($start -ge $twstart))
	{
	$tw=$true
	if ($start -gt $twstart) {$twstart=$start}
	if ($end -lt $twend) {$twend=$end}
if ($d) {write-output ("TW Start: $twstart")}
if ($d) {write-output ("TW End: $twend")}
	$twdays=(($twend-$twstart).days)+1
if ($d) {write-output ("TW Window (Days): $twdays")}
	$twbulk=$mbs * $twdays
if ($d) {write-output ("TW Work (Days * Mailboxes): $twbulk")}
	if ($terms)
		{
		$twcount=$twbulk * 0.0005
		$twsize=$twbulk * 0.0005
		$twtime=($twsize / 60) + ($twbulk / 1000)
if ($d) {write-output ("TW time (minutes): $twtime")}
		} else
		{
		$twcount=$twbulk * 5
		$twsize=$twbulk * 0.5
		$twtime=($twsize / 60) + ($twbulk / 1000)
if ($d) {write-output ("TW time (minutes): $twtime")}
		}
if ($d) {write-output ("TW Hits: $twcount")}
if ($d) {write-output ("TW Size (MB): $twsize")}
# Normalize data
	$twcount=([System.Math]::Round((($twcount + 49) / 100), 0)) * 100
if ($d) {write-output ("TW Hits: $twcount")}
	$twsize=[System.Math]::Round(($twsize + 0.49), 0)
if ($d) {write-output ("TW Size (MB): $twsize")}
	$twtime=[System.Math]::Round((($twtime + 8) / 15), 0) * 15
if ($d) {write-output ("TW time (minutes): $twtime")}
	} # if ($start -le $twstart)

$totaltime=$totaltime+$twtime
if ($d) {write-output ("Total Time (minutes): $totaltime")}

# ORCH section
# Orchestria DMC Start Date
$orstart=$(get-date("9/14/2005"))
# Orchestria DMC End Date
$orend=$(get-date("4/24/2010"))

if (($start -lt $orend) -and($end -ge $orstart))
	{
	$or=$true
	if ($start -gt $orstart) {$orstart = $start}
if ($d) {write-output ("OR Start: $orstart")}
	if ($end -lt $orend) {$orend = $end}
if ($d) {write-output ("OR End: $orend")}
	$ordays=($orend - $orstart).days + 1
if ($d) {write-output ("OR Window (Days): $ordays")}
	$orbulk=$mbs * $ordays
if ($d) {write-output ("OR Work (Days * Mailboxes): $orbulk")}
	if ($terms)
		{
		$orcount=$orbulk * 0.04
		$orsize=$orbulk * 0.004
		$ortime=$orbulk + ($orsize / 60) + ($orbulk / 1000) + ([INT]($orcount / 400))
		} else
		{
		$orcount=$orbulk * 400
		$orsize=$orbulk * 40
		$ortime=$orbulk
		}
if ($d) {write-output ("OR Hits $orcount")}
if ($d) {write-output ("OR Size (MB): $orsize")}
if ($d) {write-output ("OR Time (minutes): $ortime")}
# Normalize data
	$orcount=([System.Math]::Round((($orcount + 49) / 100), 0)) * 100
if ($d) {write-output ("OR Hits $orcount")}
	$orsize=[System.Math]::Round(($orsize + 0.49), 0)
if ($d) {write-output ("OR Size (MB): $orsize")}
	$ortime=[System.Math]::Round((($ortime + 8) / 15), 0) * 15
if ($d) {write-output ("OR Time (minutes): $ortime")}
	} # if ($start -ge $orstart)

$totaltime=$totaltime+$ortime
if ($d) {write-output ("Total Time (minutes): $totaltime")}

# EV ORCH section
# Orchestria EV start date
$oestart=$(get-date("4/23/2010"))
# Zero-Day Archiving Start Date
$ZDstart=$(get-date("1/22/2011"))

if (($start -lt $ZDstart) -and($start -ge $oestart))
	{
	$oe=$true
	if ($end -le $ZDstart) {$oedays=(($end-$start).days) + 1;if ($d) {write-output ("EV OR Start: $start");("EV OR End: $end")}} else {$oedays=(($ZDstart-$start).days) + 1;if ($d) {write-output ("EV OR Start: $start");("EV OR End: $ZDstart")}}
if ($d) {write-output ("EV OR Window (Days): $oedays")}
	$oebulk=$mbs * $oedays
if ($d) {write-output ("EV OR Work (Days * Mailboxes): $oebulk")}
	if ($terms)
		{
		$oecount=$oebulk * 0.04
		$oesize=$oebulk * 0.004
		$oetime=($oesize / 60)
		} else
		{
		$oecount=$oebulk * 400
		$oesize=$oebulk * 40
		$oetime=($oesize / 60)
		}
if ($d) {write-output ("EV OR Hits: $oecount")}
if ($d) {write-output ("EV OR size (MB): $oesize")}
if ($d) {write-output ("EV OR time (minutes): $oetime")}
# Normalize data
	$oecount=([System.Math]::Round((($oecount + 49) / 100), 0)) * 100
if ($d) {write-output ("EV OR Hits: $oecount")}
	$oesize=[System.Math]::Round(($oesize + 0.49), 0)
if ($d) {write-output ("EV OR size (MB): $oesize")}
	$oetime=[System.Math]::Round((($oetime + 8) / 15), 0) * 15
if ($d) {write-output ("EV OR time (minutes): $oetime")}
	} # if ($start -le $oestart)

$totaltime=$totaltime+$oetime
if ($d) {write-output ("Total Time (minutes): $totaltime")}

# EV Archive section
$ZDstart=$(get-date("1/22/2011"))
if ($start -lt $ZDstart)
	{
	$EV=$true
	if ($end -le $ZDstart) {$evdays=(($end-$start).days)+1;if ($d) {write-output ("EV Start: $start");("EV End: $end")}} else {$evdays=(($ZDstart-$start).days)+1;if ($d) {write-output ("EV Start: $start");("EV End: $ZDstart")}}
if ($d) {write-output ("EV Window (Days): $evdays")}
	$evbulk=$mbs * $evdays
if ($d) {write-output ("EV Work (Days * Mailboxes): $evbulk")}
	if ($terms)
		{
		$evcount=$evbulk * 0.02
		$evsize=$evbulk * 0.002
		$evtime=($evsize / 60)
if ($d) {write-output ("EV Time (minutes): $evtime")}
		} else
		{
		$evcount=$evbulk * 200
		$evsize=$evbulk * 20
		$evtime=($evsize / 60)
if ($d) {write-output ("EV Time (minutes): $evtime")}
		}
if ($d) {write-output ("EV Hits: $evcount")}
if ($d) {write-output ("EV Size (MB): $evsize")}
# Normalize data
	$evcount=([System.Math]::Round((($evcount + 49) / 100), 0)) * 100
if ($d) {write-output ("EV Hits: $evcount")}
	$evsize=[System.Math]::Round(($evsize + 0.49), 0)
if ($d) {write-output ("EV Size (MB): $evsize")}
	$evtime=[System.Math]::Round((($evtime + 8) / 15), 0) * 15
if ($d) {write-output ("EV Time (minutes): $evtime")}
	} # if ($start -lt $ZDstart)

$totaltime=$totaltime+$evtime
if ($d) {write-output ("Total Time (minutes): $totaltime")}

# Zero Day section
$ZDstart=$(get-date("1/21/2011"))
if ($end -ge $zdstart)
	{
	$zd=$true
	if ($start -gt $zdstart) {$zdstart = $start}
if ($d) {write-output ("0-Day Start: $zdstart")}
if ($d) {write-output ("0-Day End: $end")}
	$zddays=(($end-$zdstart).days)+1
if ($d) {write-output ("0-Day Window (Days): $zddays")}
	$zdbulk=$mbs * $zddays
if ($d) {write-output ("0-Day Work (Days * Mailboxes): $zdbulk")}
	if ($terms)
		{
		$zdcount=$zdbulk * 0.12
		$zdsize=$zdbulk * 0.012
		$zdtime=($zdsize / 60)
if ($d) {write-output ("0-Day Time (minutes): $zdtime")}
		} else
		{
		$zdcount=$zdbulk * 1200
		$zdsize=$zdbulk * 120
		$zdtime=($zdsize / 60)
if ($d) {write-output ("0-Day Time (minutes): $zdtime")}
		} # if ($terms)
if ($d) {write-output ("0-Day Hits: $zdcount")}
if ($d) {write-output ("0-Day Size (MB): $zdsize")}
# Normalize data
	$zdcount=([System.Math]::Round((($zdcount + 49) / 100), 0)) * 100
if ($d) {write-output ("0-Day Hits: $zdcount")}
	$zdsize=[System.Math]::Round(($zdsize + 0.49), 0)
if ($d) {write-output ("0-Day Size (MB): $zdsize")}
	$zdtime=[System.Math]::Round((($zdtime + 8) / 15), 0) * 15
if ($d) {write-output ("0-Day Time (minutes): $zdtime")}
	} # if ($start -le $zdstart)

$totaltime=$totaltime+$zdtime
if ($d) {write-output ("Total Time (minutes): $totaltime")}

# De-dupe section
$totalcount=$twcount+$orcount+$oecount+$evcount+$zdcount
$ddtime=($totalcount / 3000)
if ($d) {write-output ("De-dupe Time (minutes): $ddtime")}
$ddtime=[System.Math]::Round((($ddtime + 8) / 15), 0) * 15
if ($d) {write-output ("De-dupe Time (minutes): $ddtime")}
$ddcount=[INT]($twcount * 0.7) + [INT]($orcount * 0.5) + [INT]($oecount * 0.5) + [INT]($evcount * 0.7) + [INT]($zdcount * 0.25)
if ($d) {write-output ("De-dupe hits: $ddcount")}
$ddcount=([System.Math]::Round((($ddcount + 49) / 100), 0)) * 100
if ($d) {write-output ("De-dupe hits: $ddcount")}

$totaltime=$totaltime+$ddtime
if ($d) {write-output ("Total Time (minutes): $totaltime")}

# Delivery section
$exsize=$ddcount / 12
if ($d) {write-output ("Delivery size (MB): $exsize")}
$exsize=[System.Math]::Round(($exsize + 0.49), 0)
if ($d) {write-output ("Delivery size (MB): $exsize")}
$exsizeGB=[System.Math]::Round(($exsize / 1024), 1)
if ($d) {write-output ("Delivery size (GB): $exsizeGB")}
$extime=($exsize / 60)
if ($d) {write-output ("Delivery Time (minutes): $extime")}
$extime=[System.Math]::Round((($extime + 8) / 15), 0) * 15
if ($d) {write-output ("Delivery Time (minutes): $extime")}

$totaltime=$totaltime+$extime
if ($d) {write-output ("Total Time (minutes): $totaltime")}

} # process

# The End section executes once regardless of how many objects are passed through the pipeline
end
{

$Totalhours=[INT]($totaltime / 60)
if ($d) {write-output ("Total Time (hours): $Totalhours")}

write-output ("Projected Tumbleweed Self Policing hits: {0:#,#0}" -f $twcount)
write-output ("Tumbleweed Self Policing Data Time (Export and Import - minutes): {0:#,#0}" -f $twtime)
write-output ("Projected Orchestria DMC hits: {0:#,#0}" -f $orcount)
write-output ("Orchestria DMC export size (GB):  {0:#,#0.#}" -f $([INT]($orsize / 1024)))
write-output ("Orchestria DMC time(Export and Import - minutes): {0:#,#0}" -f $ortime)
write-output ("Projected EV Orchestria hits: {0:#,#0}" -f $oecount)
write-output ("EV Orchestria export size (GB): {0:#,#0.#}" -f $([INT]($oesize / 1024)))
write-output ("EV Orchestria time (Export and Import - minutes): {0:#,#0}" -f $oetime)
write-output ("Projected EV Vault hits: {0:#,#0}" -f $evcount)
write-output ("EV Vault export size (GB):  {0:#,#0.#}" -f $([INT]($evsize / 1024)))
write-output ("EV Vault Time (Export and Import - minutes): {0:#,#0}" -f $evtime)
write-output ("Projected Zero-Day hits: {0:#,#0}" -f $zdcount)
write-output ("Zero-Day export size (GB):  {0:#,#0.#}" -f $([INT]($zdsize / 1024)))
write-output ("Zero-Day time (Export and Import - minutes): {0:#,#0}" -f $zdtime)
write-output ("Total hits (before de-dupe): {0:#,#0}" -f $totalcount)
write-output ("De-Dupe time (minutes): {0:#,#0}" -f $ddtime)
write-output ("Estimated Final hits: {0:#,#0}" -f $ddcount)
write-output ("Final Export time (minutes): {0:#,#0}" -f $extime)
write-output ("Estimated delivery size (MB): {0:#,#0}" -f $exsize)
write-output ("Estimated delivery size (GB): {0:#,#0.#}" -f $exsizeGb)

write-output ("Total Man-hours: {0:#,#}" -f $totaLhours)
if ($report -eq $false)
	{
# Excel section
	$oldfile=test-path "$filename"
	if ($oldfile) {Remove-item "$filename" -force -confirm:$false -erroraction silentlycontinue}
	$Excel = New-Object -ComObject Excel.Application
	if ($d) { $Excel.visible = $True }
	$WKBS = $Excel.Workbooks.Add()
	$Sheet = $WKBS.Worksheets.Item(1)
	$sheet.columns.item('D').NumberFormat = "$numformat"
	$sheet.columns.item('E').NumberFormat = "$numformat"
	$sheet.columns.item('F').NumberFormat = "$numformat"
	$sheet.columns.item('G').NumberFormat = "$numformat"
	$sheet.columns.item('H').NumberFormat = "$numformat"
	$sheet.columns.item('I').NumberFormat = "$numformat"
	$sheet.columns.item('J').NumberFormat = "$numformat"
	$sheet.columns.item('K').NumberFormat = "$numformat"
	$sheet.columns.item('M').NumberFormat = "$numformat"
	$sheet.columns.item('N').NumberFormat = "$numformat"
	$sheet.columns.item('O').NumberFormat = "$numformat"
	$sheet.columns.item('P').NumberFormat = "$numformat"
	$sheet.columns.item('Q').NumberFormat = "$numformat"

#Counter variable for rows
	$intRow = 1
	$intCol = 1

	$range = $Sheet.Range("A1","Z10000")
	$range.HorizontalAlignment = 3

	$range = $Sheet.Range("A1","A6")
	$range.HorizontalAlignment = 1

	$range = $Sheet.Range("A8","A10000")
	$range.HorizontalAlignment = 1
	
	$Sheet.Cells.Item($intRow,$intCol) = "Search:"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$Search"
	$intRow++
	$intCol = 1
	
	$Sheet.Cells.Item($intRow,$intCol) = "Requester:"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	$intRow++
	$dates=[string]($start.toshortdatestring()) + " to " + [string]($end.toshortdatestring())
	$Sheet.Cells.Item($intRow,$intCol) = "Dates:"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$dates"
	$intRow++
	$intCol = 1
	$Sheet.Cells.Item($intRow,$intCol) = "Mailboxes:"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	for ($lines=1; $lines -le $mbs; $lines++)
	{$intRow++}
	
	$Sheet.Cells.Item($intRow,$intCol) = "Terms:"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	$intRow++
	$Sheet.Cells.Item($intRow,$intCol) = "File Path:"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "\\janus.cap\groups`$\IT\Exchange\Reports\SearchLogs$search"
	$intRow++
	$intCol = 1
	$intRow++

	$Sheet.Cells.Item($intRow,$intCol) = "Name"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	$intCol++

	$Sheet.Cells.Item($intRow,$intCol) = "SMTP Address(es)"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	$intCol++
	
	$Sheet.Cells.Item($intRow,$intCol) = "AD ID"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	$intCol++
	
	$Sheet.Cells.Item($intRow,$intCol) = "Orchestria DMC"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	$intCol++

	$Sheet.Cells.Item($intRow,$intCol) = "Orchestria DA"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	$intCol++
	
	$Sheet.Cells.Item($intRow,$intCol) = "EV - Vault"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	$intCol++
	
	$Sheet.Cells.Item($intRow,$intCol) = "EV - zero-Day"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	$intCol++

	$Sheet.Cells.Item($intRow,$intCol) = "Total"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	$intCol++

	$Sheet.Cells.Item($intRow,$intCol) = "Duplicates"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	$intCol++

	$Sheet.Cells.Item($intRow,$intCol) = "Exported"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	$intCol++

	$Sheet.Cells.Item($intRow,$intCol) = "TW Data"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	$intCol++

	$Sheet.Cells.Item($intRow,$intCol) = "Bloomberg Address"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	$intCol++

	$Sheet.Cells.Item($intRow,$intCol) = "Bloomberg"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	$intCol++

	$Sheet.Cells.Item($intRow,$intCol) = "Lync"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	$intCol++

	$Sheet.Cells.Item($intRow,$intCol) = "SMS"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	$intCol++
	
	$Sheet.Cells.Item($intRow,$intCol) = "IM"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	$intCol++

	$Sheet.Cells.Item($intRow,$intCol) = "File Data"
	$Sheet.Cells.Item($intRow,$intCol).Font.Bold = $True
	$intCol++

	$range = $Sheet.UsedRange
	[void]$range.EntireColumn.AutoFit()
	
	$Sheet.Name = “$search Checklist”

	$Sheet = $WKBS.Worksheets.Item(2)
	$intRow=1
	$intCol=1

	$Sheet.Cells.Item($intRow,$intCol) = "Projected Tumbleweed Self Policing hits:"
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$twcount"
	$intRow++
	$intCol--

	$Sheet.Cells.Item($intRow,$intCol) = "Tumbleweed Self Policing Data Time (Export and Import - minutes):"
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$twtime"
	$intRow++
	$intCol--

	$Sheet.Cells.Item($intRow,$intCol) = "Projected Orchestria DMC hits:"
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$orcount"
	$intRow++
	$intCol--

	$Sheet.Cells.Item($intRow,$intCol) = "Orchestria DMC export size (GB):"
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$([INT]($orsize/1024))"
	$intRow++
	$intCol--

	$Sheet.Cells.Item($intRow,$intCol) = "Orchestria DMC time(Export and Import - minutes):"
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$ortime"
	$intRow++
	$intCol--

	$Sheet.Cells.Item($intRow,$intCol) = "Projected EV Orchestria hits:"
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$oecount"
	$intRow++
	$intCol--

	$Sheet.Cells.Item($intRow,$intCol) = "EV Orchestria export size (GB):"
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$([INT]($oesize/1024))"
	$intRow++
	$intCol--

	$Sheet.Cells.Item($intRow,$intCol) = "EV Orchestria time (Export and Import - minutes):"
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$oetime"
	$intRow++
	$intCol--

	$Sheet.Cells.Item($intRow,$intCol) = "Projected EV Vault hits:"
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$evcount"
	$intRow++
	$intCol--
	
	$Sheet.Cells.Item($intRow,$intCol) = "EV Vault export size (GB):"
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$([INT]($evsize/1024))"
	$intRow++
	$intCol--
	
	$Sheet.Cells.Item($intRow,$intCol) = "EV Vault Time (Export and Import - minutes):"
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$evtime"
	$intRow++
	$intCol--
	
	$Sheet.Cells.Item($intRow,$intCol) = "Projected Zero-Day hits:"
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$zdcount"
	$intRow++
	$intCol--
	
	$Sheet.Cells.Item($intRow,$intCol) = "Zero-Day export size (GB):"
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$([INT]($zdsize/1024))"
	$intRow++
	$intCol--
	
	$Sheet.Cells.Item($intRow,$intCol) = "Export and Import - minutes):"
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$zdtime"
	$intRow++
	$intCol--
	
	$Sheet.Cells.Item($intRow,$intCol) = "Total hits (before de-dupe):"
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$totalcount"
	$intRow++
	$intCol--

	$Sheet.Cells.Item($intRow,$intCol) = "De-Dupe time (minutes):"
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$ddtime"
	$intRow++
	$intCol--

	$Sheet.Cells.Item($intRow,$intCol) = "Estimated Final hits:"
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$ddcount"
	$intRow++
	$intCol--
	
	$Sheet.Cells.Item($intRow,$intCol) = "Final Export time (minutes):"
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$extime"
	$intRow++
	$intCol--
	
	$Sheet.Cells.Item($intRow,$intCol) = "Estimated delivery size (MB):"
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$exsize"
	$intRow++
	$intCol--
	
	$Sheet.Cells.Item($intRow,$intCol) = "Estimated delivery size (GB):"
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$exsizeGB"
	$intRow++
	$intCol--

	$Sheet.Cells.Item($intRow,$intCol) = "Total Man-hours:"
	$intCol++
	$Sheet.Cells.Item($intRow,$intCol) = "$Totalhours"
	$intRow++
	$intCol--
	
	$Sheet.Name = “$search Estimates”
	
	if ($d) {write-output ("$numformat")}
	
	$range = $Sheet.UsedRange
	[void]$range.EntireColumn.AutoFit()
	$sheet.columns.item('B').NumberFormat = "$numformat"

	if ($d) {write-output ("$filename")}
	$WKBS.SaveAs("$filename",51)

	$excel.quit()

	$pass=Get-JPassword -nocrypt

	cd "\\janus.cap\groups$\IT\ITSM and Exchange\"
	./User-Setup.ps1  -first $search -last ZZAdmin -ID $search -pass $pass -location Search

	$UCPSTLoc="\\NAS01.janus.cap\UCPSTSearch01$\" + $search

	$test=test-path $UCPSTLoc
	if ($test -eq $true)
		{
		write-output ("Search PSTs should be save to: $UCPSTLoc")
		} else
		{
		new-item $UCPSTLoc -type directory
		write-output ("Search PSTs should be save to: $UCPSTLoc")
		}

	write-output ("Search Checklist located at: $filename")

	write-output ("Search Mailbox name: $search")
} # if ($report -eq $false)


remove-variable TW
remove-variable twstart
remove-variable twdays
remove-variable twbulk
remove-variable twcount
remove-variable twsize
remove-variable twtime
remove-variable EV
remove-variable evstart
remove-variable evdays
remove-variable evbulk
remove-variable evcount
remove-variable evize
remove-variable evtime
remove-variable OR
remove-variable orstart
remove-variable ordays
remove-variable orbulk
remove-variable orcount
remove-variable orsize
remove-variable ortime
remove-variable OE
remove-variable oestart
remove-variable oedays
remove-variable oebulk
remove-variable oecount
remove-variable oesize
remove-variable oetime
remove-variable ZD
remove-variable zdstart
remove-variable zddays
remove-variable zdbulk
remove-variable zdcount
remove-variable zdsize
remove-variable zdtime
remove-variable ddcount
remove-variable ddtime
remove-variable exsize
remove-variable extime
remove-variable totaltime
remove-variable totalhours
remove-variable excel
remove-variable sheet
remove-variable range
remove-variable WKBS
remove-variable test

if ($d) {$error | out-string}
} # end
} # function Show-JSearchEstimate

Function Get-JDirInfo
{
Param([string]$Folder = $(Get-Location).path)
$objFSO = New-Object -com  Scripting.FileSystemObject
$fld = $objFSO.GetFolder("$Folder")
$sizemb = "{0:N0}" -f ($fld.size / 1MB)
write-output ("$Folder" + " = " + $sizemb + " MB")
} # Function Get-JDirInfo

Function Set-JInheritance
{
<#
	.SYNOPSIS
		Turns on Inheritance

	.DESCRIPTION
		Checks the "nclude inheritable permissions from this objects parent" checkbox for a user's AD account.

	.PARAMETER  ID
		Sepcify the AD ID to change.

	.PARAMETER  Uncheck
		Us this switch to uncheck this checkbox.

	.EXAMPLE
		Set-JInheritance -ID JM27253

	Description
	===========
	Checks the "Include inheritable permissions from this objects parent" checkbox for the JM27253 AD Account.
		
	.EXAMPLE
		Set-JInheritance -ID JM27253 -Uncheck

	Description
	===========
	Unchecks the "nclude inheritable permissions from this objects parent" checkbox for the JM27253 AD Account.
		
	.NOTES
		Requires te Janus Module Loaded.

	.LINK
		http://social.technet.microsoft.com/Forums/en-US/winserverpowershell/thread/7d6346d4-bbab-463a-a594-528806a57630/

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the AD ID to modify")]
		[String]
        [Alias("Identity")]
$ID=$(Throw "You must specify an AD ID (e.g., JM27253)."),
        [Parameter(Position = 1, Mandatory = $false)]
		[Switch]
$uncheck
)

# Functions and Filters


# Main Script
# The Begin section executes once regardless of how many objects are passed through the pipeline
begin
{
Stores the DN used for an ADSI lookup
[string]$userDN=""

if ($uncheck)
	{
# allows inheritance
	$isProtected = $true

# preserve inherited rules
	$preserveInheritance = $false
	} else
	{
# allows inheritance
	$isProtected = $false

# preserve inherited rules
	$preserveInheritance = $true
	}
}

# The process section runs once for each object in the pipeline
process
{
# Error checking
$ADE=get-jadentry -id $ID -exact -pso -properties distinguishedname
$success=$?
if ($success -ne $true)
	{
	write-error ("Cannot locate user in Active Directory, exirting...") -erroraction stop
	}

# Retreive the DDN
$userDN=$ADE.distinguishedname
# PRepend DN with LDAP:// needed for ADSI operations
$userDN='LDAP://' + $userDN
# Cast user as an ADSI object
$ADO = [ADSI]"$userDN"
# Grab the ACLs for this AD Object
$acl = $ADO.objectSecurity

# Set the Inheritance on our copy of the object in AD
$acl.SetAccessRuleProtection($isProtected, $preserveInheritance)
# Commit changes made on our object to AD
$ADO.commitchanges()
write-host("Execution successful: $?")
}

# The End section executes once regardless of how many objects are passed through the pipeline
end
{

remove-variable isProtected
remove-variable preserveInheritance
remove-variable ADE
remove-variable userDN
remove-variable ADO
remove-variable acl
}
} # Function Set-JInheritance

Function Get-JInheritance
{
<#
	.SYNOPSIS
		Shows AD Inheritance

	.DESCRIPTION
		Displays the "Include inheritable permissions from this objects parent" checkbox for a user's AD account.

	.PARAMETER  ID
		Sepcify the AD ID to change.

	.EXAMPLE
		Set-JInheritance -ID JM27253

	Description
	===========
	Checks the "Include inheritable permissions from this objects parent" checkbox for the JM27253 AD Account.
		
	.NOTES
		Requires te Janus Module Loaded.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the AD ID to modify")]
		[String]
        [Alias("Identity")]
$ID=$(Throw "You must specify an AD ID (e.g., JM27253)."),
        [Parameter(Position = 1, Mandatory = $false)]
		[Switch]
$report
)

# Functions and Filters


# Main Script
# The Begin section executes once regardless of how many objects are passed through the pipeline
begin
{
Stores the DN used for an ADSI lookup
[string]$userDN=""
}

# The process section runs once for each object in the pipeline
process
{
# Error checking
$ADE=get-jadentry -id $ID -exact -pso -properties distinguishedname
$success=$?
if ($success -ne $true)
	{
	write-error ("Cannot locate user in Active Directory, exirting...") -erroraction stop
	}

# Retreive the DDN
$userDN=$ADE.distinguishedname
# PRepend DN with LDAP:// needed for ADSI operations
$userDN='LDAP://' + $userDN
# Cast user as an ADSI object
$ADO = [ADSI]"$userDN"
# Grab the ACLs for this AD Object
$acl = $ADO.objectSecurity
$checked=$acl.psbase.AreAccessRulesProtected
}

# The End section executes once regardless of how many objects are passed through the pipeline
end
{
if ($report)
	{
	write-output("Inheritance preserved: $checked")
	} else
	{
	return $checked
	}

remove-variable ADE
remove-variable userDN
remove-variable ADO
remove-variable acl
}
} # Function Get-JInheritance

Function Get-JGroupMembers
{
<#
	.SYNOPSIS
		Gets the members of a Security Group

	.DESCRIPTION
		Looks up the members of an AD group and returns them to the requestor.

	.PARAMETER  ID
		Indicate the group to enumerate.

	.PARAMETER  Report
		Use this switch to format the command to output to screen. Otherwise, ti returns a PSObject.

	.EXAMPLE
		Get-JGroupMembers -id "Messaging Team"

	  Description
	  ===========
	  Returns a Collection of PSObject containing all the members of the ADM_ExchangeAdmins AD Group.
		
	.EXAMPLE
		Get-JGroupMembers -id "Messaging Team" -report

	  Description
	  ===========
	  Returns tale on screen containing all the members of the ADM_ExchangeAdmins AD Group.
		
	.NOTES
		REquires the Janus Modiule Loaded.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the first parameter")]
		[String]
        [Alias("Identity")]
$ID=$(Throw "You must specify a parameter (e.g., Value1)."),
        [Parameter(Position = 1, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the second parameter")]
		[switch]
$report
)

# Functions and Filters


# Main Script
# The Begin section executes once regardless of how many objects are passed through the pipeline
begin
{
if ($global:JanusPSModule -ne $true) {Import-Module Janus -erroraction stop}

# Initialization
[string[]]$distinguishednames=@()
[array]$members=@()
[string[]]$temparray=@()
[string]$tempdn=""
[int]$index=0

# Error checking
$test=get-jadentry -id $ID -exact -pso -properties member
if ($test -eq $null)
	{
	$errorstring="Cannot locate an AD object matching $ID, exiting..."
	write-error $errorstring -erroraction stop
	}

if ($test.member -eq $null)
	{
	$errorstring="$ID does not have members, exiting..."
	write-error $errorstring -erroraction stop
	}
} # begin

# The process section runs once for each object in the pipeline
process
{
$memberattribute=(get-jadentry -id $ID -exact -pso -properties member).member
$array=$memberattribute.split(",")
foreach ($element IN $array)
{
if ( ($element.contains("CN=")) -and ($element.contains("DC=cap")) )
	{
	$temparray=$element.split(" ")
	$tempdn=$tempdn + "," + $temparray[0]
# write-host("$tempdn")
	$distinguishednames+=$tempdn
# write-host("$distinguishednames")
	$tempdn=""
# write-host("$tempdn")
	for ($index=1; $index -le ($temparray.count -1); $index++)
		{
		if ($index -eq 1)
			{
			$tempdn=$temparray[$index]
			$tempstring=$tempdn.Substring(3)
# write-host("$tempdn")
			} else
			{
			$tempdn=$tempdn + " " + $temparray[$index]
			$tempstring=$tempstring + " " + $temparray[$index]
# write-host("$tempdn")
			}
		$tempADE=get-jadentry -id $tempstring -exact -pso -properties name,displayname,mail,samaccountname,distinguishedname,homemdb,whencreated,whenchanged,description,memberof,member,canonicalname,physicaldeliveryofficename
		$members+=$tempADE
		}
	} elseif ($element.contains("CN=")) # if ( ($element.contains("CN=")) -and ($element.contains("DC=cap")) )
	{
	$tempdn=$element
	$tempstring=$tempdn.Substring(3)
	$tempADE=get-jadentry -id $tempstring -exact -pso -properties name,displayname,mail,samaccountname,distinguishedname,homemdb,whencreated,whenchanged,description,memberof,member,canonicalname,physicaldeliveryofficename
	$members+=$tempADE
# write-host("$tempdn")
	} elseif ($element -eq $array[-1]) # elseif ($element.contains("CN="))
	{
	$tempdn=$tempdn + "," + $element
# write-host("$tempdn")
	$distinguishednames+=$tempdn
# write-host("$distinguishednames")
	} else # elseif ($element -eq $array[-1])
	{
	$tempdn=$tempdn + "," + $element
# write-host("$tempdn")
	} # else
} # foreach ($element IN $array)
} # process

# The End section executes once regardless of how many objects are passed through the pipeline
end
{

if ($report)
	{
	$members | ft name,displayname,mail -auto -wrap
	} else
	{
	return $members
	}

remove-variable temparray
}
} # Function Get-JGroupMembers

Function Get-JADIDfromSID
{
<#
	.SYNOPSIS
		Returns the SAMAccountName associated with a SID.

	.DESCRIPTION
		Takes a SID and looks up the SAMAccountName in AD.

	.PARAMETER  ID
		Indicate the group to enumerate.

	.EXAMPLE
		Get-JADIDfromSID S-1-5-21-8915387-589958402-134157935-69260

	  Description
	  ===========
	  Returns the SAMAccontName associated with the SID S-1-5-21-8915387-589958402-134157935-69260.
		
	.NOTES
		REquires the Janus Modiule Loaded.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the first parameter")]
		[String]
        [Alias("Identity")]
$ID=$(Throw "You must specify a parameter (e.g., Value1).")
)

# Functions and Filters


# Main Script
# The Begin section executes once regardless of how many objects are passed through the pipeline
begin
{

# Initialization
[string]$SAMAccountName=""

# Error checking
if ($sid.length -ne 42)
	{
	$errorstring="$SID is not the correct length, exiting..."
	write-error $errorstring -erroraction stop
	}

if ($sid[0] -notlike "S")
	{
	$errorstring="$SID is not a valid format, exiting..."
	write-error $errorstring -erroraction stop
	}

if ($sid[1] -ne "-")
	{
	$errorstring="$SID is not a valid format, exiting..."
	write-error $errorstring -erroraction stop
	}

if ($sid[3] -ne "-")
	{
	$errorstring="$SID is not a valid format, exiting..."
	write-error $errorstring -erroraction stop
	}

if ($sid[5] -ne "-")
	{
	$errorstring="$SID is not a valid format, exiting..."
	write-error $errorstring -erroraction stop
	}

if ($sid[8] -ne "-")
	{
	$errorstring="$SID is not a valid format, exiting..."
	write-error $errorstring -erroraction stop
	}

if ($sid[16] -ne "-")
	{
	$errorstring="$SID is not a valid format, exiting..."
	write-error $errorstring -erroraction stop
	}

if ($sid[26] -ne "-")
	{
	$errorstring="$SID is not a valid format, exiting..."
	write-error $errorstring -erroraction stop
	}

if ($sid[36] -ne "-")
	{
	$errorstring="$SID is not a valid format, exiting..."
	write-error $errorstring -erroraction stop
	}

} # begin

# The process section runs once for each object in the pipeline
process
{
$objSID = New-Object System.Security.Principal.SecurityIdentifier("$id")
$objUser = $objSID.Translate( [System.Security.Principal.NTAccount])

} # process

# The End section executes once regardless of how many objects are passed through the pipeline
end
{
$SAMAccountName=$objUser.Value
$SAMAccountName=$samaccountname.substring(10)
return $SAMAccountName

remove-variable temparray
}
} # Function Get-JADIDfromSID

Function Add-JHolidays
{
<#
	.SYNOPSIS
		Updates Mailbox Calendars with Janus US Holidays.

	.DESCRIPTION
		Updates the specified Mailbox Calendar with Janus US Holidays. The dates are extracted from here: \\p-ucslog03.janus.cap\d$\Scripts\Scheduled Tasks\Holiday Project\Holidays.csv.

	.PARAMETER  ID
		Specify the AD ID to be updated. Accepts SAMAccountName, SMTP Address or Distinguished Name.

	.EXAMPLE
		./Add-JHolidays.ps1 -ID JM27253

		Description
		===========
		Updates the Calendar associated with JM27253 with the Janus US Holidays.
		
	.NOTES
		Requires the Janus Module and the EWS API Loaded.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the Identity of the Mailbox to Modify.")]
	[String]
        [Alias("Identity")]
        [Alias("DistinguishedName")]
        [Alias("SAMAccountName")]
        [Alias("mail")]
        [Alias("PrimarySMTPAddress")]
        [Alias("WindowsEmailAddress")]
$ID,
        [Parameter(Position = 1, Mandatory = $false, ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false, HelpMessage="Use this Switch to activate Debug Mode.")]
	[Switch]
$d
)

Begin
{
#Initialization
[string]$Subject = ""
[string]$MessageID = ""
[string]$UserAccount = ""
[string]$EmailAddress = ""
$ErrorActionPreference="SilentlyContinue"

#Check to see if the Holiday CSV is found.
$TestCsv = Test-Path "\\p-ucslog03.janus.cap\d$\Scripts\Scheduled Tasks\Holiday Project\Holidays.csv"
if ($d) {Write-Output "Test CSV value: $TestCsv"}

if ($TestCsv -eq $False)
{
    Write-Error "The CSV is not found!" -ErrorAction Stop
}

#Load the CSV file
$HolidayCSV = Import-Csv -Path "\\p-ucslog03.janus.cap\d$\Scripts\Scheduled Tasks\Holiday Project\Holidays.csv"
if ($d) {Write-Output "Holiday CSV values: $TestCsv"}
if($? -ne $True)
{ Write-Error "Could not load file." -ErrorAction Stop }

#Import the Janus module for added functionality
if ($global:JanusPSModule -ne $true) {Import-Module Janus -erroraction stop}

#Checks to see if the EWS Module is functional for the below code that will check the mailbox.
Write-Output ("Checking for EWS Module...")
$EWSFile=$null
if (test-path "D:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll") {$EWSFile="D:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"}
if (test-path "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll") {$EWSFile="C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"}
if (test-path "D:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll") {$EWSFile="D:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll"}
if (test-path "C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll") {$EWSFile="C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll"}
if ($d) {Write-Output "EWSFile value: $EWSFile"}

if ($EWSFile -ne $null)
	{
	$dllpath = $EWSFile
	Add-Type -Path $EWSFile
	[void][Reflection.Assembly]::LoadFile($dllpath)
	} # if (($test -ne $null) -and($EWSFile -ne $null))
if ($d) {Write-Output "dllpath value: $dllpath"}

# Create Service Object. We only need Exchange 2007 schema for creating calendar items (this will work with Exchange>=12)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)
if ($d) {Write-Output "Service value:`n";$service | fl}

#Use the credentials of the scheduled task
$service.UseDefaultCredentials = $true;

} # Begin

Process
{  
# Main Script
$UserAccount = $ID
write-output("Processing $userAccount ...")
if ($d) {Write-Output "UserAccount value: $UserAccount"}
$EmailAddress = (Get-JADEntry -ID $UserAccount -exact -pso -properties mail).mail
if ($d) {Write-Output "EmailAddress value: $EmailAddress"}
# Check email address
if ((($EmailAddress | measure-object).count) -ne 1)
{
    [string]$errorstring="SMTP Address does not match required criteria." + $EmailAddress
    $Error.add($errorstring)
    Write-Output ($errorstring)
}

try
{
    $service.AutodiscoverUrl($EmailAddress)
}

# Manually assigns the url if there is an error
Catch [system.exception]
{
    $URI='https://p-ucusxhc04.janus.cap/EWS/Exchange.asmx'
    $service.URL = New-Object Uri($URI)
    write-output ("Caught an Autodiscover URL exception, recovering...")
}

if ($d) {Write-Output "service.URL value: $($service.URL)"}

#Impersonation will allow the script to access other mailboxes
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress);
if ($d) {Write-Output "service.ImpersonatedUserId value: ";$service.ImpersonatedUserId | out-string}

foreach ($Holiday in $HolidayCSV)
{
if ($d) {Write-Output "Holiday values: $Holiday"}
	$Appointment = New-Object Microsoft.Exchange.WebServices.Data.Appointment($service);
	$Appointment.Subject=$Holiday."Subject";
if ($d) {Write-Output "Appointment.Subject value: $($Appointment.Subject)"}
if ($d) {Write-Output "Holiday.StartDate value: $($Holiday.StartDate)"}
if ($d) {Write-Output "Holiday.StartTime value: $($Holiday.StartTime)"}
	$StartDate=Get-Date($Holiday.StartDate + " " + $Holiday.StartTime);
	$Appointment.Start=$StartDate;
if ($d) {Write-Output "Appointment.Start value: $($Appointment.Start)"}
if ($d) {Write-Output "Holiday.EndDate value: $($Holiday.EndDate)"}
if ($d) {Write-Output "Holiday.EndTime value: $($Holiday.EndTime)"}
	$EndDate=Get-Date($Holiday.EndDate + " " + $Holiday.EndTime);
	$Appointment.End=$EndDate;
if ($d) {Write-Output "Appointment.End value: $($Appointment.End)"}
	$Appointment.IsAllDayEvent=$False;
    $Appointment.LegacyFreeBusyStatus="Free"
if ($d) {Write-Output "Appointment.LegacyFreeBusyStats value: $($Appointment.LegacyFreeBusyStatus)"}
if ($d) {Write-Output "Appointment.IsAllDayEvent value: $($Appointment.IsAllDayEvent)"}

	#Save the Holiday.
	$Appointment.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar);
if ($d) {Write-Output "Appointment.Save Success: $?"}
} # End of CalendarItem Processing Foreach

} # ProcessDone

End
{
Remove-Variable TestCsv
Remove-Variable HolidayCSV
Remove-Variable EWSFile
Remove-Variable dllpath
Remove-Variable service
Remove-Variable UserAccount
Remove-Variable EmailAddress
Remove-Variable Appointment
Remove-Variable StartDate
Remove-Variable EndDate
} # End
} # Function Add-JHolidays

function Get-JFileEncoding
{
<#
	.SYNOPSIS
		Displays File Encoding.

	.DESCRIPTION
		Returns the File Encoding used on the specified File.

	.PARAMETER  file
		Enter the sub-string that should appear anywhere in the PF Identity.

	.EXAMPLE
		Get-JFileEncoding -file Janus.psm1
		
		Description
		-----------
		Returns the File Encoding used on the Janus.psm1 File.
		
	.NOTES
		Requires the Loading of the Janus Module.

	.LINK
		http://www.leeholmes.com/guide

#>

	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("Identity")]
        [Alias("ID")]
        [Alias("Path")]
        [Alias("FilePath")]
		[String]
$file
)

# Error checking
$test=Test-Path $file
if ($Test -eq $false)
	{
	write-error ("Unable to locate $file, exiting...") -erroraction stop
	}

# Main Script
## The hashtable used to store our mapping of encoding bytes to their
## name. For example, "255-254 = Unicode"
$encodings = @{}

## Find all of the encodings understood by the .NET Framework. For each,
## determine the bytes at the start of the file (the preamble) that the .NET
## Framework uses to identify that encoding.
$encodingMembers = [System.Text.Encoding] |
    Get-Member -Static -MemberType Property

$encodingMembers | Foreach-Object {
    $encodingBytes = [System.Text.Encoding]::($_.Name).GetPreamble() -join '-'
    $encodings[$encodingBytes] = $_.Name
}

## Find out the lengths of all of the preambles.
$encodingLengths = $encodings.Keys | Where-Object { $_ } |
    Foreach-Object { ($_ -split "-").Count }

## Assume the encoding is UTF7 by default
$result = "UTF7"

## Go through each of the possible preamble lengths, read that many
## bytes from the file, and then see if it matches one of the encodings
## we know about.
foreach($encodingLength in $encodingLengths | Sort -Descending)
{
    $bytes = (Get-Content -encoding byte -readcount $encodingLength $file)[0]
    $encoding = $encodings[$bytes -join '-']

    ## If we found an encoding that had the same preamble bytes,
    ## save that output and break.
    if($encoding)
    {
        $result = $encoding
        break
    }
}

## Finally, output the encoding.
return [System.Text.Encoding]::$result

} # function Get-JFileEncoding

Function Get-JStringMatches
{
<#
	.SYNOPSIS
		Looks for string matches in files.

	.DESCRIPTION
		Loks for the specified String/String Pattern in the specified files.

	.PARAMETER  string
		Enter the Stroing pattern to match.

	.PARAMETER  file
		Specify the File(s) to search.

	.PARAMETER  folder
		Specify the name of the Foder that the Delegate will have access to.

	.PARAMETER  s
		Indicates to look in all sub-directories for he requested match.

	.PARAMETER  i
		Indicates the the search should ignore case.
		NOTE: Use -i:$false to run a Case-Sensitive Ssearch.

	.EXAMPLE
		Get-JStringMatches hf0 r*.txt -s
		
		Description
		-----------
		Searches for all instances of *hf0* in files that match r*,txt in the current directory and all sub-directories.
		
	.EXAMPLE
		Get-JStringMatches hf0 r*.txt -s -i:$false
		
		Description
		-----------
		Searches for all case-sensitive instances of *hf0* in files that match r*,txt in the current directory and all sub-directories.
		
	.NOTES
		Requires the Janus Modiule to be loaded.

#>

	[CmdletBinding()]
Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [Alias("pattern")]
		[String]
$string,
        [Parameter(Position = 1, Mandatory = $false)]
        [Alias("Path")]
		[String]
$file="*.*",
        [Parameter(Mandatory = $false)]
        [Alias("recursive")]
        [Alias("includesubdirectories")]
        [Alias("Subtree")]
		[Switch]
$s,
        [Parameter(Mandatory = $false)]
        [Alias("IgnoreCase")]
        [Alias("CaseInsensitive")]
		[Switch]
$i=$true
)
if ($s)
	{
	if ($i)
	{
	get-childitem .\ -include $file -rec | select-string -pattern $string
	} else
	{
	get-childitem .\ -include $file -rec | select-string -pattern $string -casesensitive
	}
	} else
	{
	if ($i)
	{
	select-string -path $file -pattern $string
	} else
	{
	select-string -path $file -pattern $string -casesensitive
	}
	}
} # Function Get-JStringMatches

function Send-JTCPRequest
{
<#
	.SYNOPSIS
		Sends a packet to a computer using TCP and returns the response.

	.DESCRIPTION
		Sends a packet to the specified computer using TCP and returns the response as a string.

	.PARAMETER  ComputerName
		Specify the server to send the TCP Packet to.
		NOTE: This may have to match the name on the certificate if using SSL.

	.PARAMETER  Test
		Use this Switch to test the TCP copnnection.

	.PARAMETER  Port
		Specify the TCP Port to connect with.

	.PARAMETER  UseSSL
		Use this Switch to indicate SSL is required.

	.PARAMETER  InputObject
		Specify the information to send to the specified computer.

	.PARAMETER  Delay
		Specify the delay, in miliseconds, between sends and receives.

	.EXAMPLE
		Send-JTCPRequest -computername mdm.stg.myjonline.com -port 443 -usessl -inputobject "GET /hastatus.html HTTP/1.1"
		
		Description
		-----------
		Copnnects to mdm.stg.myjonline.com using SSL over port 443 and sends "GET /hastatus.html HTTP/1.1" and returns the response.
	
	.EXAMPLE
		[Byte[]] $ping = 0x45, 0x00 , 0x00, 0x3c, 0x3a, 0xff, 0x00, 0x00, 0x80, 0x01, 0x5c, 0x55, 0xc0, 0xa8, 0x92, 0x16, 0xc0, 0xa8, 0x90, 0x05, 0x08, 0x00, 0x4d, 0x35, 0x00, 0x01, 0x00, 0x26, 0x61, 0x62, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68, 0x69, 0x6a, 0x6b, 0x6c, 0x6d, 0x6e, 0x6f, 0x70, 0x71, 0x72, 0x73, 0x74, 0x75, 0x76, 0x77, 0x61, 0x62, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68, 0x69
		Send-JTCPRequest -computername p-ucadm01 -port 7 -inputobject $ping
		
		Description
		-----------
		Sends a ping packet to P-UCADM01.
	

	.NOTES
		Requires the Janus Modiule to be loaded.

#>

	[CmdletBinding()]
param(
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the The computer to connect to:")]
	[String]
        [Alias("Identity")]
        [Alias("Name")]
        [Alias("Host")]
        [Alias("Computer")]
        [Alias("System")]
        [Alias("Server")]
        [Alias("FQDN")]
        [Alias("DNSHostName")]
$ComputerName = "localhost",
        [Parameter(Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Use this switch to determine if you just want to test the connection")]
	[switch]
$Test,
        [Parameter(Position = 2, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Specify the port to use:")]
	[int]
$Port = 80,
        [Parameter(Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Use this switch to determine if the connection should be made using SSL:")]
	[switch]
$UseSSL,
        [Parameter(Position = 1, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Specify The input string to send to the remote host:")]
	[string]
$InputObject,
        [Parameter(Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Specify the delay, in milliseconds, to wait between commands:")]
	[int]
$Delay = 100
)

Set-StrictMode -Version Latest

[string] $SCRIPT:output = ""

## Store the input into an array that we can scan over. If there was no input,
## then we will be in interactive mode.
$currentInput = $inputObject
if(-not $currentInput)
{
    $currentInput = @($input)
}
$scriptedMode = ([bool] $currentInput) -or $test

function Main
{
    ## Open the socket, and connect to the computer on the specified port
    if(-not $scriptedMode)
    {
        write-host "Connecting to $computerName on port $port"
    }

    try
    {
        $socket = New-Object Net.Sockets.TcpClient($computerName, $port)
    }
    catch
    {
        if($test) { $false }
        else { Write-Error "Could not connect to remote computer: $_" }

        return
    }

    ## If we're just testing the connection, we've made the connection
    ## successfully, so just return $true
    if($test) { $true; return }

    ## If this is interactive mode, supply the prompt
    if(-not $scriptedMode)
    {
        write-host "Connected.  Press ^D followed by [ENTER] to exit.`n"
    }

    $stream = $socket.GetStream()

    ## If we wanted to use SSL, set up that portion of the connection
    if($UseSSL)
    {
        $sslStream = New-Object System.Net.Security.SslStream $stream,$false
        $sslStream.AuthenticateAsClient($computerName)
        $stream = $sslStream
    }

    $writer = new-object System.IO.StreamWriter $stream

    while($true)
    {
        ## Receive the output that has buffered so far
        $SCRIPT:output += GetOutput

        ## If we're in scripted mode, send the commands,
        ## receive the output, and exit.
        if($scriptedMode)
        {
            foreach($line in $currentInput)
            {
                $writer.WriteLine($line)
                $writer.Flush()
                Start-Sleep -m $Delay
                $SCRIPT:output += GetOutput
            }

            break
        }
        ## If we're in interactive mode, write the buffered
        ## output, and respond to input.
        else
        {
            if($output)
            {
                foreach($line in $output.Split("`n"))
                {
                    write-host $line
                }
                $SCRIPT:output = ""
            }

            ## Read the user's command, quitting if they hit ^D
            $command = read-host
            if($command -eq ([char] 4)) { break; }

            ## Otherwise, Write their command to the remote host
            $writer.WriteLine($command)
            $writer.Flush()
        }
    }

    ## Close the streams
    $writer.Close()
    $stream.Close()

    ## If we're in scripted mode, return the output
    if($scriptedMode)
    {
        $output
    }
}

## Read output from a remote host
function GetOutput
{
    ## Create a buffer to receive the response
    $buffer = new-object System.Byte[] 1024
    $encoding = new-object System.Text.AsciiEncoding

    $outputBuffer = ""
    $foundMore = $false

    ## Read all the data available from the stream, writing it to the
    ## output buffer when done.
    do
    {
        ## Allow data to buffer for a bit
        start-sleep -s 1

        ## Read what data is available
        $foundmore = $false
        $stream.ReadTimeout = 1000

        do
        {
            try
            {
                $read = $stream.Read($buffer, 0, 1024)

                if($read -gt 0)
                {
                    $foundmore = $true
                    $outputBuffer += ($encoding.GetString($buffer, 0, $read))
                }
            } catch { $foundMore = $false; $read = 0 }
        } while($read -gt 0)
    } while($foundmore)

    $outputBuffer
}

. Main
}

function Global:Set-JMaxWindowSize 
{

    if ($Host.Name -match "console")
       {
        $MaxHeight = $host.UI.RawUI.MaxPhysicalWindowSize.Height
        $MaxWidth = $host.UI.RawUI.MaxPhysicalWindowSize.Width

        $MyBuffer = $Host.UI.RawUI.BufferSize
        $MyWindow = $Host.UI.RawUI.WindowSize

        $MyWindow.Height = ($MaxHeight)
        $MyWindow.Width = ($Maxwidth-2)

        $MyBuffer.Height = (9999)
        $MyBuffer.Width = ($Maxwidth-2)

        $host.UI.RawUI.set_bufferSize($MyBuffer)
        $host.UI.RawUI.set_windowSize($MyWindow)
       }

    $CurrentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $CurrentUserPrincipal = New-Object Security.Principal.WindowsPrincipal $CurrentUser
    $Adminrole = [Security.Principal.WindowsBuiltinRole]::Administrator
    If (($CurrentUserPrincipal).IsInRole($AdminRole)){$Elevated = "Administrator"}

    $Title = $Elevated + " $ENV:USERNAME".ToUpper() + ": $($Host.Name) " + $($Host.Version) + " - " + (Get-Date).toshortdatestring()
    $Host.UI.RawUI.set_WindowTitle($Title)

} # Set-JMaxWindowSize

Function Show-JFailureAuditEvents
{
<#
	.SYNOPSIS
		Shows Failure Audits in the Security Event Log for a given ID.

	.DESCRIPTION
		Shows Failure Audits in the Security Event Log of the specified DC for a given ID.

	.PARAMETER  DC
		Description of what Paramter1 is.

	.PARAMETER  Parameter2
		Description of what Paramter2 is.

	.EXAMPLE
		Show-JFailureAuditEvents -DC P-JCDCD05.janus.cap -ID jm27253
		
	.NOTES
		Requires the Janus Module.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the first parameter")]
		[String]
        [Alias("DomainController")]
        [Alias("Server")]
        [Alias("Computer")]
        [Alias("ComputerName")]
$DC=$(Throw "You must specify a parameter (e.g., Value1)."),
        [Parameter(Position = 1, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the second parameter")]
		[String]
        [Alias("Identity")]
        [Alias("User")]
        [Alias("Name")]
        [Alias("SAMAccountName")]
$ID=$null
)

# Functions and Filters


# Main Script
# The Begin section executes once regardless of how many objects are passed through the pipeline
begin
{
if ($global:JanusPSModule -ne $true) {Import-Module Janus -erroraction stop}

# Initialization
[string]$log=""
$now=get-date
$after=$now.addminutes(-60)
$test=test-connection $DC -quiet -count 1
if ($test -eq $null)
	{
	write-error ("Cannot locate $DC, exiting...") -erroraction stop
	}
$IDError=$ID
$ID=(get-jadentry $ID -exact -pso -properties samaccountname).samaccountname
$success=$?
if ($success -ne $true)
	{
	write-error ("Cannot locate $IDError, exiting...") -erroraction stop
	}
}

# The process section runs once for each object in the pipeline
process
{
# $Events=Get-EventLog -LogName Security -ComputerName $DC -After $after  -EntryType FailureAudit | where {$_.Message -like "*$ID*"}
$Events=Get-EventLog -LogName Security -ComputerName $DC -After $after -EntryType FailureAudit -Message "*$ID*" -newest 10
$Events=$Events | sort TimeGenerated -descending | select Message,TimeGenerated,InstanceId,Source,ReplacementStrings -First 10
}

# The End section executes once regardless of how many objects are passed through the pipeline
end
{
$Events | fl

remove-variable Events
remove-variable IDError
remove-variable now
remove-variable after
remove-variable test
remove-variable log
}
} # Function Show-JFailureAuditEvents

Function Import-JPST
{

<#
	.SYNOPSIS
		Imports PSTs into a mailbox.

	.DESCRIPTION
		Imports all PST files from the specified location into the specified mailbox.
		Will let you specify the PST folder to start inmporting from and the Mailbox
		Folder to import into.

	.PARAMETER  Mailbox
		Specify the Exchange Mailbox to import the PST content into.

	.PARAMETER  PSTFolderPath
		Specofy the File path where the PST files are located.
		NOTE: Defaults to the current location

	.PARAMETER  PSTMessageRootFolder
		Specify the top-level folder to import the PST messages from.
		NOTE: Defaults to "Root Folder" used by EV PST files

	.PARAMETER  MailboxMessageDestinationFolder
		Specify the Exchange Mailbox Folder to import the PST messages into.
		NOTE: Defaults to "Search Hits"

	.EXAMPLE
		Import-JPST -mailbox DEC12-21 -PSTFolderPath "\\nas01\ucpstsearch01$\DEC12-21\DEC12-21 EV\Import" -PSTMessageRootFolder "Root Items" -MailboxMessageDestinationFolder "EV Hits"

	Description
	===========
		Imports all PST files located at \\nas01\ucpstsearch01$\DEC12-21\DEC12-21 EV\Import. It will start importing at the Root Items folder and blow. It will put all messages into the EV Hits mailbox folder.

	.NOTES
		Requires the Exchange Module and the Janus Module.

#>

param (
        [Parameter(Position = 0, Mandatory = $true, HelpMessage="Enter the Mailbox to Import Messages into")]
		[String]
        [Alias("Identity")]
        [Alias("SearchMailbox")]
        [Alias("DistinguishedName")]
        [Alias("mail")]
        [Alias("GUID")]
        [Alias("Display name")]
        [Alias("SAMAccountNAme")]
        [Alias("Userprincipalname")]
        [Alias("LegacyExchangeDN")]
        [Alias("SmtpAddress")]
        [Alias("PrimarySmtpAddress")]
        [Alias("Alias")]
$Mailbox,
        [Parameter(Position = 1, Mandatory = $false)]
		[String]
        [Alias("FolderPath")]
        [Alias("FilePath")]
        [Alias("FileFolderPath")]
$PSTFolderPath=$((Get-Location).Path),
        [Parameter(Position = 2, Mandatory = $false)]
		[String]
        [Alias("RootFolder")]
        [Alias("PSTRootFolder")]
        [Alias("PSTFolder")]
        [Alias("SourceFolder")]
$PSTMessageRootFolder="Root Folder",
        [Parameter(Position = 3, Mandatory = $false)]
		[String]
        [Alias("MailboxFolder")]
        [Alias("OutlookFolder")]
        [Alias("ExchangeFolder")]
        [Alias("DestinationFolder")]
$MailboxMessageDestinationFolder="Search Hits"
)

# Functions and Filters


# Main Script
# The Begin section executes once regardless of how many objects are passed through the pipeline
begin
{
# Initialization
# Checing for Janus Module
if ($global:JanusPSModule -ne $true) {Import-Module Janus -erroraction stop}

# Checking for Exchange Snapin
Write-host("Checking for Exchange 2010 Module.")
$test=Get-DatabaseAvailabilityGroup
$global:ExchangeSnapIn=$?
if ($global:ExchangeSnapIn -ne $true)
	{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
	$test=Get-DatabaseAvailabilityGroup
	$success=$?
	If ($success -ne $true)
		{
		write-error ("Unable to load Exchange 2010 Module, exiting...") -erroraction stop
		} else # If ($success -ne $true)
		{
		$global:ExchangeSnapIn = $true
		} # else
	} # if ($global:ExchangeSnapIn -ne $true)
Write-host("Exchange 2010 Module Loaded.")

# Getting a DC
[string]$DC=Get-JDCs | where {$_ -like "*dcd*"} | get-random
do
	{
# Test that DC is reachable
	$ping=Test-Connection $DC -Count 1
	$test=$?
# Get a new DC if it is not
	if ($test -eq $false)
		{
# Filter out old DC from new DC list
		$DC=Get-JDCs | where {$_ -like "*dca*"} |  where {$_ -like $DC}  | get-random
		}
	} while ($test -eq $false)

# Checking for File path
$test=test-path $pstfolderpath
if ($test -ne $true)
	{
	$errorstring="Unable to confirm that $pstfolderpath is a valid path, exiting..."
	Write-host ("$errorstring")
	write-error ("$errorstring") -erroraction stop
	}

# Checking to see mailbox exists
$oldmailbox=$Mailbox
$Mailbox=(get-mailbox $Mailbox).PrimarySmtpAddress
$test=$?
if ($test -ne $true)
	{
	$errorstring="Unable to confirm that $oldmailbox is a valid mailbox, exiting..."
	Write-host ("$errorstring")
	write-error ("$errorstring") -erroraction stop
	}

# load the file names of all PST files into this variable
$Items = Get-ChildItem -path $pstfolderpath -filter *.pst | sort name -descending
}

# The process section runs once for each object in the pipeline
process
{
# Run once for each PST file
foreach ($item in $items)
{
# Format the file locatin to include the folder location and file name.
	$filelocation = $pstfolderpath + "`\" + $item.name
	do {
# $name is used to name the import job
		$name=$item.name
# Get a random CAS to manage the import job (this should ensure near even distribution across all CAS servers
		$server=(Get-TransportServer | where {$_.name -like "b-ucus*"} | Get-Random).name
# Let user know what we are doing
		$command = "New-MailboxImportRequest -Mailbox $Mailbox -name $name -DomainController $DC -FilePath $filelocation -SourceRootFolder $PSTMessageRootFolder -TargetRootFolder $MailboxMessageDestinationFolder -ConflictResolutionOption KeepSourceItem -BadItemLimit Unlimited -AcceptLargeDataLoss -MRSServer $server -confirm:`$false"
		Write-host $command
# Atual import command
		New-MailboxImportRequest -Mailbox $Mailbox -name $name -DomainController $DC -FilePath  $filelocation -SourceRootFolder $PSTMessageRootFolder -TargetRootFolder $MailboxMessageDestinationFolder -ConflictResolutionOption KeepAll -BadItemLimit Unlimited -AcceptLargeDataLoss -MRSServer $server -confirm:$false
	} while (!($?))
}

}

# The End section executes once regardless of how many objects are passed through the pipeline
end
{
# Variable cleanup
remove-variable oldmailbox
remove-variable test
remove-variable Items
remove-variable name
remove-variable server
remove-variable command
Write-host ("Execution complete, exiting...")
}
} # Function Import-JPST

Function Export-JPST
{
<#
	.SYNOPSIS
		Exports PSTs based on criteria specified.

	.DESCRIPTION
		Used for legal/Compliance Searches, Exports e-mails from a given date range and mailbox to a specific folder.

	.PARAMETER  mb
		Specify the Mailbox to export PST files from.

	.PARAMETER  file
		Specify the file naming convention to use for the PST files.

	.PARAMETER  folder
		Specify the folder in the mailbox to export (it will export this folder and all subfolders).

	.PARAMETER  UNC
		Specify the folder path to save the PST files at.

	.PARAMETER  start
		Specify the sent date to start exporting messages from.

	.PARAMETER  end
		Specify the sent date to start exporting messages until.

	.PARAMETER  d
		specify that you want debug information on the screen.

	.EXAMPLE
		Export-JPST.ps1 -mb zDEC12-21 -UNC \\nas01.janus.cap\UCPSTSearch01$\DEC12-21 -filename DEC12-21-Final.PST -folder "Final Results" -start "January 1, 1970" -end "December 21, 2012"
		
	.NOTES
		Requires the Janus Module Loaded.
		Requires the Exchange Snap In Loaded.

#>

param (
        [Parameter(Position = 0, Mandatory = $true, HelpMessage="Enter the mailbox to export the PST files from")]
		[String]
        [Alias("Identity")]
        [Alias("Mailbox")]
        [Alias("Alias")]
        [Alias("MailNickname")]
$mb,
        [Parameter(Position = 1, Mandatory = $true, HelpMessage="Enter the filenaming convention to use for the PSTs")]
		[String]
        [Alias("Filename")]
        [Alias("PSTFilename")]
        [Alias("PSTFile")]
$File,
        [Parameter(Position = 1, Mandatory = $true, HelpMessage="Enter the mailbox folder to export to the PSTs")]
		[String]
        [Alias("FolderName")]
        [Alias("MailboxRootFolder")]
        [Alias("MailboxFolder")]
$Folder,
        [Parameter(Position = 2, Mandatory = $true, HelpMessage="Enter the folder path to save the PST files at")]
		[String]
        [Alias("FolderPath")]
        [Alias("PSTFolderPath")]
$UNC,
        [Parameter(Position = 3, Mandatory = $true, HelpMessage="Enter the start date of the export")]
		[System.DateTime]
        [Alias("StartDate")]
        [Alias("SearchStartDate")]
$start,
        [Parameter(Position = 4, Mandatory = $true, HelpMessage="Enter the end date of the export")]
		[System.DateTime]
        [Alias("EndDate")]
        [Alias("SearchEndDate")]
$end,
        [Parameter(Position = 5, Mandatory = $false)]
		[Switch]
$d
)

# Functions and Filters


# Main Script
# The Begin section executes once regardless of how many objects are passed through the pipeline
begin
{
# Checking for Janus Module
if ($global:JanusPSModule -ne $true) {Import-Module Janus -erroraction stop}

# Checking for Exchang Snap In
if ($D) {Write-host("Checking for Exchange 2010 Snap In.")}
$test=Get-DatabaseAvailabilityGroup
$global:ExchangeSnapIn=$?
if ($global:ExchangeSnapIn -ne $true)
	{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
	$test=Get-DatabaseAvailabilityGroup
	$success=$?
	If ($success -ne $true)
		{
		write-error ("Unable to load Exchange 2010 Module, exiting...") -erroraction stop
		} else # If ($success -ne $true)
		{
		$global:ExchangeSnapIn = $true
		} # else
	} # if ($global:ExchangeSnapIn -ne $true)
if ($d) {Write-host("Exchange 2010 Snap In Loaded.")}

# Initialization
# This is the start date of the current month of exports
[System.DateTime]$currentstartdate=$start.adddays($(-1*($start.day -1)))
if ($d) {write-host("Current Start DAte: $currentstartdate")}
# This is the end date for the current month of exports
[System.DateTime]$currentenddate=$currentstartdate.addmonths(1)
if ($d) {write-host("Current End Date: $currentenddate")}
[string]$strmonth=""
[string]$stryear=""
[string]$strfile=""

# error checking
# Does the mailbox exist?
$test=get-mailbox $mb
$success=$?
if ($success -ne $true)
	{
	$errrorstring="Could not locate a mailbox with the name $mb , exiting..."
	write-host("$errorstring")
	write-error("$errorstring") -erroraction stop
	}
if ($d) {write-host("Mailbox located successfully.")}

# Is the file path valid?
$test=test-path $UNC
$success=$?
if ($success -ne $true)
	{
	$errrorstring="Could not locate a folder with the name $UNC , exiting..."
	write-host("$errorstring")
	write-error("$errorstring") -erroraction stop
	}
if ($d) {write-host("UNC Path Verified successfully.")}

# .contains() is case sensitive, normalize the data
$file=$file.toupper()
# In order for the file naming convention to work, the .pst must be omitted, normalize the date
if ($file.contains(".PST"))
	{
	$file=$file.substring("0","$($file.length -4)")
	}
if ($d) {write-host("File Naming Convention used: $file-XXXXX.PST")}

# Downstream logic relies on the last character being a \, normalize the data
if ($UNC[-1] -notlike '\')
	{
	$UNC=$UNC + '\'
	}
if ($d) {write-host("Current UNC Path saving PSTs to: $UNC")}

} # begin

# The process section runs once for each object in the pipeline
process
{
# Use a loop to execute versus all possible dates needed for the export
for ($loop=$currentstartdate;$loop -lt $end;$loop=$loop.addmonths(1))
{
if ($d) {write-host("Processing mail sent after $loop ...")}
# strmonth and year are used for the file naming
$strmonth=($currentstartdate.month).tostring()
if ($strmonth.length -eq 1) {$strmonth= "0" + $strmonth}
$stryear=($currentstartdate.year).tostring()
if ($stryear.length -ne 2) {$stryear=$stryear.substring("2","2")}
$strfile=$file + "-" + $strmonth + $stryear + ".PST"
if ($d) {write-host("File Name for this export: $strfile")}
$path=$UNC + $strfile
if ($d) {write-host("File Path for this export: $path")}
# Get a random CAS to manage the import job (this should ensure near even distribution across all CAS servers
$server=(Get-TransportServer | where {$_.name -like "p-ucus*"} | Get-Random).name
if ($d) {write-host("MRS Server used: $server")}
$CMD="New-MailboxExportRequest -ContentFilter `{`(Sent -ge $currentstartdate`) -and `(Sent -lt $currentenddate`)`} -Mailbox $mb -Name $strfile –FilePath $path -suspend -MRSServer $server -ExcludeDumpster -AcceptLargeDataLoss -BadItemLimit Unlimited -ConflictResolutionOption KeepSourceItem -SourceRootFolder $folder -Confirm:`$false"
write-host("Executing:`n$CMD")
# Actual commanded needed
New-MailboxExportRequest -ContentFilter {(Sent -ge $currentstartdate) -and (Sent -lt $currentenddate)} -Mailbox $mb -Name $strfile –FilePath $path -suspend -MRSServer $server -ExcludeDumpster -AcceptLargeDataLoss -BadItemLimit Unlimited -ConflictResolutionOption KeepSourceItem -SourceRootFolder $folder -Confirm:$false
$success=$?
# Error processing for export request
if ($success -ne $true)
	{
	$errrorstring="Error executing $CMD , exiting..."
	write-host("$errorstring")
	write-error("$errorstring") -erroraction stop
	}
if ($d) {write-host("Export Request Sent successfully.")}
# Set up next month of exports
$currentstartdate=$currentenddate
if ($d) {write-host("Current Start Date: $currentstartdate")}
# Set up next month of exports
$currentenddate=$currentenddate.addmonths(1)
if ($d) {write-host("Current End Date: $currentenddate")}
} # for ($loop=$start;$loop -lt $end;$loop.addmonths(1))
} # process

# The End section executes once regardless of how many objects are passed through the pipeline
end
{
remove-variable test
remove-variable currentstartdate
remove-variable currentenddate
remove-variable strmonth
remove-variable stryear
remove-variable strfile
remove-variable path
write-host("Export complete.")
write-host("Type `"Get-MailboxExportRequest | Resume-MailboxExportRequest`" to begin exporting.")
} # end
} # Function Export-JPST

Function Import-JSCOM
{
if (test-path "C:\Program Files\System Center Operations Manager 2012\Powershell\OperationsManager")
{
if (test-connection p-scom02.janus.cap) {$server="p-scom02.janus.cap"}
if (test-connection p-scom03.janus.cap) {$server="p-scom03.janus.cap"}
if (test-connection p-scom04.janus.cap) {$server="p-scom04.janus.cap"}
if (test-connection b-scom01.janus.cap) {$server="b-scom01.janus.cap"}
if (test-connection b-scom02.janus.cap) {$server="b-scom02.janus.cap"}
Add-PSSnapin Microsoft.EnterpriseManagement.OperationsManager.Client -erroraction silentlycontinue
. "C:\Program Files\System Center Operations Manager 2012\Powershell\OperationsManager\Functions.ps1"
Start-OperationsManagerClientShell -ManagementServerName: $server -PersistConnection: $true -Interactive: $true
}
} # Function Import-JSCOM

Function Set-JMaintenanceMode
{
<#
	.SYNOPSIS
		Puts a server in maintenance mode in mode in SCOM.

	.DESCRIPTION
		Puts the specified server in maintenance mode in SCOM.

	.PARAMETER  ID
		Enter the name of the server to put in Maintenance mode.

	.PARAMETER  End
		Enter the Date and time that maintenance mode should end.
		NOTE: This will default to 30 minutes after the cmdlet is executed.

	.PARAMETER  Reason
		Enter the Reason for the Maintenance Mode that Administrators will see in SCOM.
		NOTE: Valid Values are: PlannedOther, UnplannedOther, PlannedHardwareMaintenance, UnplannedHardwareMaintenance, PlannedHardwareInstallation, UnplannedHardwareInstallation, PlannedOperatingSystemReconfiguration, UnplannedOperatingSystemReconfiguration, PlannedApplicationMaintenance, ApplicationInstallation, ApplicationUnresponsive, ApplicationUnstable, SecurityIssue, LossOfNetworkConnectivity. 
		NOTE: Reason will default to PlannedApplicationMaintenance

	.PARAMETER  Comment
		Enter the message Administrators will see in SCOM for this maintenance mode.

	.PARAMETER  d
		Specify that Debug information should be displayed.

	.EXAMPLE
		Set-JMaintenanceMode -id P-UCADM01 -end "12/21/2012 1:00 am" -Comment "Testing"
		
		Description
		-----------
		Sets P-UCADM01 in Maintenance Mode in SCOM through 12-21-2012 1:00 am.
		
	.NOTES
		Requires the Loading of the Janus Module.
		Requires the Loading of the SCOM Module.

#>
	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the name of the comuter to put in maintenence mode")]
		[String]
        [Alias("Identity")]
        [Alias("Name")]
        [Alias("Host")]
        [Alias("Computer")]
        [Alias("System")]
        [Alias("Server")]
        [Alias("FQDN")]
        [Alias("DNSHostName")]
$ID=$(Throw "You must specify an AD ID (e.g., P-UCADM01)."),
        [Parameter(Position = 1, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the date and time that the maintenance mode is supposed to end")]
        [Alias("expires")]
		[system.datetime]
$End=$((get-date).addminutes(30)),
        [Parameter(Position = 3, Mandatory = $false, HelpMessage="Enter the reason for the maintenence mode")]
		[String]
$Reason="PlannedApplicationMaintenance",
        [Parameter(Position = 2, Mandatory = $false, HelpMessage="Enter the maintenence mode comment")]
		[String]
        [Alias("Explanation")]
        [Alias("Note")]
$Comment="IT Maintenance",
		[switch]
$d
)

begin
{
if ($d) {write-host("Entering Function...")}
[system.datetime]$now=get-date
if ($d) {write-host("Checking Date...")}
if ($end -le $now)
	{
	$errorsting="$end is an invalid date/time, Exiting..."
	Write-output("$errorsting")
	$error | out-string
	Write-Error("$errorsting") -erroraction stop
	}
if ($d) {write-host("Confirmed Valid End Date.")}
	
if ($d) {write-host("Testing to see if SCOM Module is loaded.")}
$test=Get-SCOMCommand -erroraction silentlycontinue
if ($test -eq $null)
	{
	Import-JSCOM
	if ($? -ne $true) {write-error ("Error Loading SCOM Module, exiting...") -erroraction stop}
	}
if ($d) {write-host("Confirmed that SCOM Module is loaded.")}
}

Process
{
if ($d) {write-host("Processing Server $ID ...")}
$test=Test-Connection $ID -count 1 -erroraction silentlycontinue
if ($test -ne $null)
	{
if ($d) {write-host("Server $ID confirmed.")}
if ($d) {write-host("Retreiving Server Instance...")}
	$Instance=Get-SCOMClassInstance -name "$ID*"  | where {$($_.GetMonitoringClasses()).Name -like "*Microsoft.Windows.Computer*"}
	$success=$?
if ($d) {write-host("Instance created for $ID - success: $success")}
	if ($success -ne $true)
		{
		$errorsting="Could not locate $ID in SCOM, Exiting..."
		Write-output("$errorstring")
		$error | out-string
		Write-Error("$errorstring") -erroraction stop
		}
if ($d) {write-host("Checking to see if server $ID is already in Mainenance Mode.")}
	$MME=$Instance | Get-SCOMMaintenanceMode
	if ($MME -eq $null)
		{
if ($d) {write-host("No Maintenance Mode found, creating new Mainenance Mode.")}
		$Instance | Start-SCOMMaintenanceMode -EndTime $End -Reason $Reason -Comment $Comment -erroraction silentlycontinue
		$success=$?

		if ($success -ne $true)
			{
			$errorsting="Failed to Start Maintenence Mode  for $ID in SCOM, Exiting..."
			Write-output("$errorstring")
			$error | out-string
			Write-Error("$errorstring") -erroraction stop
			}
		write-host("Maintenance Mode Started.")
		} else
		{
if ($d) {write-host("Maintenance Mode found, updating Mainenance Mode.")}
		$MME | Set-SCOMMaintenanceMode -EndTime $End -Reason $Reason -Comment $Comment -erroraction silentlycontinue
		$success=$?
		if ($success -ne $true)
			{
			$errorsting="Failed to update Maintenence Mode  for $ID in SCOM, Exiting..."
			Write-output("$errorstring")
			$error | out-string
			Write-Error("$errorstring") -erroraction stop
			}
		write-host("Maintenance Mode Updated.")
		} # if ($MME -eq $null)
	} else
	{
	$errorsting="Failed to locate $ID on the network, Exiting..."
	Write-output("$errorstring")
	$error | out-string
	Write-Error("$errorstring") -erroraction stop
	} # if ($test -ne $null)
} # Process

end
{
remove-variable test
remove-variable MME
remove-variable Now
remove-variable Instance
remove-variable Success
} # end

} # Function Set-JMaintenanceMode

function Add-JADGroup
{
<#
	.SYNOPSIS
		Adds a group to AD

	.DESCRIPTION
		Adds the specified Security Group in AD.

	.PARAMETER  ID
		Description of what Paramter1 is.

	.EXAMPLE
		Add-JADGRoup -ID ADM_Exchange_Admins
		
	.NOTES
		Put notes here.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the AD ID of this group")]
		[String]
        [Alias("Identity")]
        [Alias("Name")]
        [Alias("CN")]
        [Alias("SAMAccountName")]
$ID=$(Throw "You must specify the AD ID of the new Group (e.g., ADM_Exchange_Admins)."),
        [Parameter(Position = 1, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the OU for this group (e.g., OU=Security Groups,OU=Janus,DC=janus,DC=cap)")]
		[String]
        [Alias("OrganizationalUnit")]
        [Alias("ParentOrganizationalUnit")]
        [Alias("ParentOU")]
        [Alias("Container")]
        [Alias("ParentContainer")]
$OU="OU=Security Groups,OU=Janus,DC=janus,DC=cap",
        [Parameter(Position = 2, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Universal Group (true/false)")]
		[Switch]
$Universal=$false,
        [Parameter(Position = 3, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the Maanger of this group")]
		[String]
$Owner=$(gc env:username),
        [Parameter(Position = 4, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the Description of this Group")]
		[String]
$Description="",
        [Parameter(Position = 5, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enable Debug info (true/false)")]
		[Switch]
$d=$false
)

# Initialization
$ADS_PROPERTY_APPEND = 3
$ADS_GROUP_TYPE_GLOBAL_GROUP =     0x00000002
$ADS_GROUP_TYPE_LOCAL_GROUP =      0x00000004
$ADS_GROUP_TYPE_UNIVERSAL_GROUP =  0x00000008
$ADS_GROUP_TYPE_SECURITY_ENABLED = 0x80000000
[array]$test=@()

# Error checking
if ($d) {write-host("Checking ID: $ID")}
$test=get-jadentry -pso -exact -id $ID -properties displayname
if ($d) {write-host("$($test.count) found for $ID")}
if ($test.count -ne 0)
	{
	write-error("$ID may already exist, exiting...") -erroraction stop
	}

if ($d) {write-host("Checking $Owner")}
$test=get-jadentry -pso -exact -id $Owner -properties displayname,givenname,sn,DistinguishedName
if ($d) {write-host("$($test.count) found for $Owner")}
if ($test.count -ne 1)
	{
	write-error("$Owner does not match a single AD entry, exiting...") -erroraction stop
	}
$manager=$test
if ($d) {write-host("Manager:");$manager | out-string}

if ($d) {write-host("Checing $OU")}
$test=get-jadentry -pso -exact -id $OU -properties DistinguishedName
if ($d) {write-host("$($test.count) found for $OU")}
if ($test.count -ne 1)
	{
	write-error("$OU does not match a single AD OU, exiting...") -erroraction stop
	}
$targetOU=[adsi]"LDAP://$($test.DistinguishedName)"
if ($d) {write-host("OU:");$targetOU | out-string}

# MaIN Script
# Creates a Directory Object to access AD methods and objects
$objDomain = New-Object System.DirectoryServices.DirectoryEntry
if ($d) {write-host("Domain:");$objDomain | fl}

# Creates the group
$objGroup = $objDomain.Create("group", "CN=" + $ID)
if ($d) {write-host("Group:");$objGroup | fl}

# Set the group type
if ($Universal)
	{
	$objGroup.Put("groupType", $($ADS_GROUP_TYPE_UNIVERSAL_GROUP -bor $ADS_GROUP_TYPE_SECURITY_ENABLED))
if ($d) {write-host("GroupType=$($ADS_GROUP_TYPE_UNIVERSAL_GROUP -bor $ADS_GROUP_TYPE_SECURITY_ENABLED) - $?")}
	} else
	{
	$objGroup.Put("groupType", $($ADS_GROUP_TYPE_GLOBAL_GROUP -bor $ADS_GROUP_TYPE_SECURITY_ENABLED))
if ($d) {write-host("GroupType=$($ADS_GROUP_TYPE_GLOBAL_GROUP -bor $ADS_GROUP_TYPE_SECURITY_ENABLED) - $?")}
	} # if ($Universal)

# Set SAMAccountName
$objGroup.Put("SAMAccountName", $ID )
if ($d) {write-host("SAMAccountName = $ID - $?")}
# Set the Description and Info Fields
$objGroup.Put("Description","Owned by $($manager.givenname) $($manager.sn) - $description")
if ($d) {write-host("Description = Owned by $($manager.givenname) $($manager.sn) - $description - $?")}
$objGroup.Put("Info","Owned by $($manager.givenname) $($manager.sn) - $description")
if ($d) {write-host("Info = Owned by $($manager.givenname) $($manager.sn) - $description - $?")}
# Set the Manager
$objGroup.Put("ManagedBy",$($manager.DistinguishedName))
if ($d) {write-host("ManagedBy = $($manager.DistinguishedName)" - $?)}
# The settings above are not committed until the setinfo() method is called
$objGroup.SetInfo()
if ($d) {write-host("SetInfo() successful: $?")}
# Move Group to correct OU
$objGroup.MoveTo($targetOU)
if ($d) {write-host("MoveTo() successful: $?")}
$objGroup.SetInfo()
if ($d) {write-host("SetInfo() successful: $?")}

# Cleanup
remove-variable ADS_PROPERTY_APPEND
remove-variable ADS_GROUP_TYPE_GLOBAL_GROUP
remove-variable ADS_GROUP_TYPE_LOCAL_GROUP
remove-variable ADS_GROUP_TYPE_UNIVERSAL_GROUP
remove-variable ADS_GROUP_TYPE_SECURITY_ENABLED
remove-variable test
remove-variable manager
remove-variable targetOU
remove-variable objDomain
remove-variable objGroup
} # function Add-JADGRoup

Function Get-JRDPSession
{

<#
.SYNOPSYS
Queries a local or remote computer for existing RDPSessions

.DESCRIPTION
Uses Qwintsa to query local or remote computers for existing RDP Sessions, reutrns a custom PSObject and can be used in conjuction with Remove-RDPSession

.EXAMPLE
Get-RDPSession -ComputerName TESTVM

Returns information about existing sessions on TESTVM:

SESSIONNAME  : services
USERNAME     : 
ID           : 0
STATE        : Disc
TYPE         :
DEVICE       :
ComputerName : TESTVM

SESSIONNAME  : console
USERNAME     :
ID           : 1
STATE        : Conn
TYPE         :
DEVICE       :
ComputerName : TESTVM

SESSIONNAME  : rdp-tcp#0
USERNAME     : Jason
ID           : 2
STATE        : Active
TYPE         :
DEVICE       :
ComputerName : TESTVM

SESSIONNAME  : rdp-tcp
USERNAME     :
ID           : 65536
STATE        : Listen
TYPE         :
DEVICE       :
ComputerName : TESTVM

.EXAMPLE
Get-JRDPSession -ComputerName TESTVM | where {$_.username -like 'Jason'}

Finds the session for the user named 'Jason'.

.NOTES
Requires PowerShell Version 3
#>

[cmdletBinding()]
Param
    (
	[Parameter(ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[String[]]
	$ComputerName = "$env:COMPUTERNAME"
    )
Begin {}
Process
    {
write-host ("Processing $ComputerName")
        $Results = qwinsta /Server:$ComputerName
        $Props = ($Results[0].Trim(" ") -replace ("\b *\B")).Split(" ")
        $Sessions = $Results[1..$($Results.Count -1)]
        Foreach ($Session in $Sessions)
            {
                $hash = [ordered]@{
                        $Props[0] = $Session.Substring(1,18).Trim()
                        $Props[1] = $Session.Substring(19,22).Trim()
                        $Props[2] = $Session.Substring(41,7).Trim()
                        $Props[3] = $Session.Substring(48,8).Trim()
                        $Props[4] = $Session.Substring(56,12).Trim()
                        $Props[5] = $Session.Substring(68,8).Trim()
                        'ComputerName' = "$ComputerName"
                    }
                New-Object -TypeName PSObject -Property $hash
            }
    }
End {}
} # Function Get-JRDPSession

Function Remove-JRDPSession
{
<#
.SYNOPSYS
Removes an existing RDP session on a remote workstation
 
.DESCRIPTION
Uses Rwinsta to remove a remote session by ID number, accepts input from Get-RDPSession or through parameters
 
.EXAMPLE
Remove-RDPSession -Computername TESTVM -ID 2

Removes the session 2 from TESTVM

.EXAMPLE
Get-RDPSession -ComputerName TESTVM | where {$_.username -like 'Jason'} | Remove-RDPSession

Finds the session for the user named 'Jason', and closes it.
 
.NOTES
Requires PowerShell Version 3
#>
[cmdletBinding()]
Param 
    (
        [Parameter(
        Mandatory=$true,
        ValueFromPipelineByPropertyName=$True
        )]
        [String]
$ComputerName,
        [Parameter(
        Mandatory=$true,
        ValueFromPipelineByPropertyName=$True
        )]
        [int]
$ID
    )
Begin {}
Process 
    {
        rwinsta /Server:$ComputerName $ID
    }
End {}
} # Function Remove-JRDPSession


Function Set-JRandomADPassword{

<#
.DESCRIPTION
     Name: Set-JRandomADPassword.ps1
     Version: 1.0
     AUTHOR: Dennis Kendrick
     DATE  : 01/08/2014

.SYNOPSIS
     Changes an Active Directory Password with a Random Generated Password.

.DESCRIPTION
    Changes an Active Directory Password with a Random Generated Password.

.PARAMETER  ID
	Enter the Name of the AD account you are looking for.  Preferably SamAccount or the DisplayName.
	NOTE: AD Name, Display Name

.PARAMETER  length
	Enter a substring of the name or Displayname. This will find all near misses.
	NOTE: An integer

.PARAMETER  pattern
	Specify pattern Ex: LUNSLUNS This would set the password for an 8 character password.  L-Lowercase U-Uppercase N-Number S-Symbol'
	NOTE: LUNS Characters

.PARAMETER  d
	Use this switch to output debug data for scripting troubleshooting of this
	command.

.PARAMETER testaccounts
    Use this switch if you just want to change the passwords of our test accounts in prod.

.NOTES
	

.LINK
    
#>
#requires -version 2
Param(
[Parameter(Position=0,Mandatory = $false,ValueFromPipelineByPropertyName=$true,HelpMessage='AD Username Needed.')]
[Alias("Identity")]
[string]$ID,

[Parameter(Position=1,Mandatory = $false,HelpMessage='Enter a number for password length.')]
[int]$length,

[Parameter(Position=1,Mandatory = $false,HelpMessage='Specify pattern Ex: LUNSLUNS This would set the password for an 8 character password.  L-Lowercase U-Uppercase N-Number S-Symbol')]
[string]$pattern,

[switch]$d,

[switch]$testaccounts

)

# The ADModule is required
# Import-Module activedirectory -ErrorAction Stop


# File name for Powershell transcript. (LOG LOCATION)
$log="D:\Scripts\Logs\Set-JRandomADPassword.log"

# Start a new Transcript
Start-Transcript -append -path $log -Force

Write-Output "Resetting password on $(Get-Date)"

#This function returns a random password.
function RandomPassword
{
    param (
        [int]$length,
        [string]$pattern # optional
    )
 
    # Define classes of character pools, there are four classes
    # by default: L - lowercase, U - uppercase, N - numeric,
    # S - special
    $pattern_class = @("L", "U", "N", "S")
 
    # Character pools for classes defined above
    $charpool = @{ 
        "L" = "abcdefghijklmnopqrstuvwxyz";
        "U" = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        "N" = "1234567890";
        "S" = "!#%&$@*()_-+={}?"
    }
 
    $rnd = New-Object System.Random
 
    # Introduce a random delay to avoid same random seed
    # during frequent calls
    Start-Sleep -milliseconds $rnd.Next(500) 
 
    # Create a random pattern if pattern is not defined or
    # fill the remaining if the pattern length is less than
    # password length
    if (!$pattern -or $pattern.length -lt $length) {
 
        if (!$pattern)
        {
            $pattern = ""
            $start = 0
        } else {
            $start = $pattern.length - 1
        }
 
        # Create a random pattern
        for ($i=$start; $i -lt $length; $i++)
        {
            $pattern += $pattern_class[$rnd.Next($pattern_class.length)]
        }
 
        
     }
 
     $password = ""
 
     for ($i=0; $i -lt $length; $i++)
     {   
        $wpool = $charpool[[string]$pattern[$i]]      
        $password += $wpool[$rnd.Next($wpool.length)]    
     }                
 
     return $password
}

#This switch is encountered first to check if -testaccounts was used.  If it is invoked we run this code block which retreives our test accounts in the OU below where the SamAccountName contains zTest and XM0.
#The statement to grab the accounts can be modified to include additional accounts.
if($testaccounts){
$DBTestAccounts = Get-JADEntry -ID "zMailflow" -pso -properties SAMAccountName
    if($d) {
    Write-Host "Here is a list of accounts being changed: $DBTestAccounts"}
#Invoke the RandomPassword function and pass to it the length and pattern it will return a password.
$RandomPass = RandomPassword -length 10 -pattern LUNSLUNSLU
        if($d){
            Write-Host "Password used is: $RandomPass" -ForegroundColor Cyan
}
#Foreach loop sets the password on the accounts.
    foreach($account in $DBTestAccounts){
    $PassOutcome = $NULL
    Set-JPassword -Id $account.SAMAccountName -Pass $RandomPass
    $PassOutcome = $?
        if ($d){
            Write-Host "Password Attempt Success:$PassOutcome for $account`n" -ForegroundColor Yellow
    }
        

}

Stop-Transcript
BREAK

}


#FROM HERE DOWN IS ONLY USED IF TESTACCOUNTS WAS NOT SPECIFIED

#This series checks if a variable is set and if not it asks for input.
if (!$ID){ $ID = Read-Host "Please enter an ID (SAMAccount)"}
if ($d){ Write-host ("The ID used: $ID") -ForegroundColor Yellow}
if (!$length){ [int]$length = Read-Host "Please supply an integer value for the length of the password."}
if ($d) { Write-Host -ForegroundColor Yellow "The length will be $length characters"}
if(!$pattern) { $pattern = Read-Host "Specify pattern Ex: LUNSLUNS This would set the password for an 8 character password.  L-Lowercase U-Uppercase N-Number S-Symbol"}
#If the pattern length does not equal the length specified the error is displayed.
if ($pattern.length -ne $length) { Write-Host -ForegroundColor Yellow -BackgroundColor Red "ERROR: The length $length is not the same length as the pattern $pattern.  This should match EX: if your length is 10 then your pattern needs 10 characters LUNSLUNSLU." ; BREAK}
if ($d) { Write-Host -ForegroundColor Yellow "The pattern used will be $pattern."}

#Checks to make sure the length and pattern are set then generates a password and sets it.
if (($length) -and ($pattern)){

if ($d) { Write-Host -ForegroundColor Yellow "The length of the password will be $length characters. `n The pattern used is :$pattern" }
$RandomPass = RandomPassword -length $length -pattern $pattern
if ($d) {Write-Host -ForegroundColor Yellow "The password used is: $RandomPass"}
Set-JPassword -Id $ID.SAMAccountName -Pass $RandomPass
$PassOutcome = $?

}


if ($d){
Write-Host "Password Attempt Success:$PassOutcome" -ForegroundColor Yellow
}#End if

Stop-Transcript

} #End Set-JRandomADPassword

Function Show-JASDevices
{

<#
	.SYNOPSIS
		Reports all ActiveSync Devices.

	.DESCRIPTION
		Reports all ActiveSync Devices. Reports Device ID (needed to remove/modify ActiveSync Devices), Device Friendly Name (e.g., Identity,DeviceFriendlyName,DeviceOS,FirstSyncTime,LastPolicyUpdateTime,LastSuccessSync,LastSyncAttemptTime,DevicePolicyApplied,DeviceAccessState) ,Device OS (e.g., BlackBerry 10), First Sync Time, Last Policy Update Time, Last Success Sync, Last Sync Attempt Time, Device Policy Applied, Device Access State (e.g., Blocked or Allowed).

	.PARAMETER  ID
		Specify ID to report on.
		NOTE: If no ID is specified, will run against every user.
		NOTE: Will Accept Display Name, SAM Account Name, SMTP Address or Distingqushed Name.

	.PARAMETER  D
		Indicate that a "debug" level of output is requested.

	.EXAMPLE
		Report-ASDevices -id jm27253

		Description
		-----------
		Reports all Active Sync Device Pairings for the user ID jm27253.

	.EXAMPLE
		Report-ASDevices

		Description
		-----------
		Reports all Active Sync Device Pairings for all users.

	.NOTES
		Requires both Excahnge and Janus Modules.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
		[String]
        [Alias("Identity")]
        [Alias("DisplayName")]
        [Alias("SAMAccountName")]
        [Alias("mail")]
        [Alias("PrimarySMTPAddress")]
        [Alias("WindowsEmailAddress")]
        [Alias("DistinguishedName")]
$ID='*',
        [Parameter(Position = 1)]
		[Switch]
$D
)

# Functions and Filters


# Main Script
# The Begin section executes once regardless of how many objects are passed through the pipeline
begin
{
$error.clear()
if ($d) {write-host("Checking for Exchange 2010 Snap-In...")}
if ($global:E2010SnapIn -ne $true)
	{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -erroraction stop
	$global:E2010SnapIn=$?
	$global:ExchangeSnapIn=$global:E2010SnapIn
	. 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1' -erroraction stop; Connect-ExchangeServer -auto -AllowClobber -erroraction stop
	}
if ($d) {write-host("Exchange 2010 Snap-In Loaded.")}

if ($d) {write-host("Checking for Janus Module...")}
if ($global:JanusPSModule -ne $true) {Import-Module Janus -erroraction stop}
if ($d) {write-host("Janus Module Lodaded.")}

# Initialization
[string]$log=""
if ($ID -ne '*')
	{
if ($d) {write-host("Looking up $ID in AD...")}
	$ID=(get-jadentry -id $ID -exact -pso -properties samaccountname).samaccountname
if ($d) {write-host("AD Lookup complete.")}
	}
}

# The process section runs once for each object in the pipeline
process
{
if ($ID -ne '*')
	{
if ($d) {write-host("Looking up Single User...")}
	Get-ActiveSyncDevice -Mailbox $ID | Get-ActiveSyncDeviceStatistics | select Identity,DeviceFriendlyName,DeviceOS,FirstSyncTime,LastPolicyUpdateTime,LastSuccessSync,LastSyncAttemptTime,DevicePolicyApplied,DeviceAccessState
if ($d) {write-host("User Lookup Complete.")}
	} else
	{
if ($d) {write-host("Looking up Organization...")}
	Get-ActiveSyncDevice -ResultSize unlimited | Get-ActiveSyncDeviceStatistics | select Identity,DeviceFriendlyName,DeviceOS,FirstSyncTime,LastPolicyUpdateTime,LastSuccessSync,LastSyncAttemptTime,DevicePolicyApplied,DeviceAccessState | sort LastSyncAttemptTime -descending
if ($d) {write-host("Organization Lookup Complete.")}
	}

}

# The End section executes once regardless of how many objects are passed through the pipeline
end
{
$log=$error | out-string
if ($d) {write-host("`nError Log:`n$log")}

if ($d) {write-host("Cleaning up Variables...")}
remove-variable log
if ($d) {write-host("Variable Clean up complete.")}
}

} # Function Show-JASDevices

Function Show-JRDPSessions
{
<#
	.SYNOPSIS
		Returns RDP Sessions for UC Servers.

	.DESCRIPTION
		Returns RDP Sessions logged on with a specific ID for all specified servers.

	.PARAMETER  ID
		Specify ID to search for.
		NOTE: This must be their ActiveDirectory SAM Account Name (e.g., JM27253).
		NOTE: Will return all IDs if none is specified.

	.PARAMETER  Server
		Specify server(s) to check against.
		NOTE: Will return all Servers that have an AD Description of "UC Server" if none is specified.

	.PARAMETER  D
		Specify that a debug level of out put is requested.

	.EXAMPLE
		Show-JRDPSessions

		Description
		===========
		Lists all RDP Sessions for all servers with "UC Server" in their AD Description.
		
	.EXAMPLE
		Show-JRDPSessions -ID jm27253 -server "p-ucusxm01","p-ucusxm02","p-ucusxm03","p-ucusxm04","b-ucusxm01","b-ucusxm02","b-ucusxm03","b-ucusxm04"

		Description
		===========
		Lists all RDP Session logged on by JM27253 logged into p-ucusxm01, p-ucusxm02, p-ucusxm03, p-ucusxm04, b-ucusxm01, b-ucusxm02, b-ucusxm03 or b-ucusxm04.
		
	.NOTES
		Put notes here.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
		[String]
        [Alias("Identity")]
$ID='*',
        [Parameter(Position = 1, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
		[String[]]
$Server=@('*'),
        [Parameter(Position = 2, Mandatory = $false)]
		[Switch]
$D
)

# Import the Active Directory module for the Get-ADComputer CmdLet
if ($d) {write-output("Loading Janus Module...")}
Import-Module Janus -erroraction stop
if ($d) {write-output("Janus Module Loaded...")}

[array]$SessionList=@()
[array]$Servers=@()
[string]$ServerName=""

if ($d) {write-output("Server(s): $server");write-output("ID: $ID")}

# Query Active Directory for computers running a Server operating system
if ($d) {write-output("Looking up Servers in AD...")}
if ($Server[0] -eq '*')
	{
	$Servers = Get-JADEntry -ID "uc server" -Description -pso -properties name
	} else
	{
	foreach ($CPU IN $server)
		{
		$Servers += Get-JADEntry -ID $CPU -exact -pso -properties name
		}
	}
if ($d) {write-output("Server Lookup Complete...")}
if ($d) {write-output("Server(s): $Servers")}

# Loop through the list to query each server for login sessions
ForEach ($System in $Servers) {
	$ServerName = $System.Name

if ($d) {write-output("Processing $ServerName ...")}

	# Run the qwinsta.exe and parse the output
	$queryResults = (qwinsta /server:$ServerName | foreach { (($_.trim() -replace "\s+",","))} | ConvertFrom-Csv) 
	
	# Pull the session information from each instance
	ForEach ($queryResult in $queryResults)
		{
if ($d) {write-output("Processing $($queryResult.USERNAME) ...")}
		$result=New-Object PSObject -Property @{
RDPServerName = $System.Name
UserName = $queryResult.USERNAME
SessionName = $queryResult.SESSIONNAME
SessionState = $queryResult.STATE
SessionID = $queryResult.ID
		} # End of Object Creation
		# We only want to display where a "person" is logged in. Otherwise unused sessions show up as USERNAME as a number
		If (($result.username -match "[a-z]") -and ($result.username -ne $NULL))
			{ 
			# When running interactively, uncomment the Write-Host line below to show the output to screen
			if ($ID -ne '*')
				{
				if ("$ID" -like "*$($queryResult.USERNAME)*") {$SessionList += $result;if ($d) {write-output("Adding Session Matching ID $ID ...")}}
				} else
				{
				$SessionList += $result
				}
			}
		}
}

# When running interactively, uncomment the Write-Host line below to see the full list on screen
return $SessionList
} # Function Show-JRDPSessions

Function Show-JCRPermissions
{
<#
	.SYNOPSIS
		Report Permissions for the specified CR.

	.DESCRIPTION
		Report Permissions for the specified Conference Room. this checks Book In Policy, Delegates and Full Mailbox Access.

	.PARAMETER  ID
		Specify the Conference Room to Report on.

	.EXAMPLE
		Show-CRPermissions -ID !CR-YY-99-RUTest@janus.com

		Description
		===========
		Lists all users with permissions to this mailbox
		
	.NOTES
		Requires Exchange Module to be loaded.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="You must specify a Conference Room (e.g., !CR-YY-99-RUTest@janus.com).")]
		[String]
        [Alias("Identity")]
$ID=$(Throw "You must specify a Conference Room (e.g., !CR-YY-99-RUTest@janus.com).")
)

# Functions and Filters


# Main Script
# The Begin section executes once regardless of how many objects are passed through the pipeline
begin
{
# Test for Janus Module (Needed for reading Outlook-level Delegates
if ($global:JanusPSModule -ne $true) {Import-Module Janus -erroraction stop}

# Test for Exchange Module
Write-host("Checking for Exchange 2010 Module.")
$test=Get-DatabaseAvailabilityGroup
$global:ExchangeSnapIn=$?
if ($global:ExchangeSnapIn -ne $true)
	{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
	$test=Get-DatabaseAvailabilityGroup
	$success=$?
	If ($success -ne $true)
		{
		write-error ("Unable to load Exchange 2010 Module, exiting...") -erroraction stop
		} else # If ($success -ne $true)
		{
		$global:ExchangeSnapIn = $true
		} # else
	} # if ($global:ExchangeSnapIn -ne $true)
Write-host("Exchange 2010 Module Loaded.")

# Initialization
# $results is the final list of names
[string[]]$results=@()
# Holds details for processing (removal of "JANUS_CAP\" etc.
[string]$temp=""
# Stores the Delegate SAMAccountName IDs
[array]$DIDs=@()

# Error checking, does this mailbox exist?
$test=get-mailbox -id $id
if (($? -eq $true) -and($test -ne $null))
	{
	$id=$test.PrimarySmtpAddress
	} else
	{
	write-error "Error retrieving mailbox data for $ID, exiting..." -erroraction stop
	}
} # Begin

# The process section runs once for each object in the pipeline
process
{
write-host "Gathering users with Book In Policy permissions..."
# Book In Policy reurns an array of SAMAccountNames
$bips=(Get-CalendarProcessing $ID).bookinpolicy
foreach ($bip IN $bips)
	{
	$DIDs+=$bip.Name
	}

write-host "Gathering users with Fll Mailbox Access permissions..."
# Mailbox Permissions returns an array of ADObjects, we can extract the AD ID (in JANUS_CAP\JMXXXXX format) and then normalize
$mbps=Get-MailboxPermission $ID | where {$_.AccessRights -like "*fullaccess*"}
foreach ($mbp IN $mbps)
	{
# Set $temp to the ID value
	$temp=$mbp.User.RawIdentity
# Remove the Janus_CAP\ portion of the ID
	$temp=$temp.replace("JANUS_CAP\","")
# Grab valid relegate (not SIDs, built-in accounts, etc.)
	if (($temp -notlike "*S-1-5*") -and($temp -notlike "*NT AUTHORITY*") -and($temp -notlike "*`$*") -and($temp -notlike "*!jm*") -and($temp -notlike "*bps*") -and($temp -notlike "*bpp*") -and($temp -notlike "*JANUSADMIN*") -and($temp -notlike "*GRP_IT_*") -and($temp -notlike "*ADM_*") -and($temp -notlike "*AFADM*") -and($temp -notlike "*IT Messaging Services*") -and($temp -notlike "*Servers*") -and($temp -notlike "*EVucusxm*"))
		{
		$DIDs+=$temp
		}
	}

write-host "Gathering users with Outlook Delegate permissions..."
# This uses EWS to retrieve Mailbox Delegates (the kind visible in Outlook)
$dels=Get-JEWSDelegates $ID -folder Calendar
foreach ($del IN $dels)
	{
# Only grab Delegaes who can book meetings
	if (($del.DelegateDisplayName -ne $null) -and($del.CalendarFolderPermissionLevel -notlike "*None*") -and($del.CalendarFolderPermissionLevel -notlike "*Reviewer*"))
		{
		$DIDs+=$del.DelegateDisplayName
		}
	}

write-host "Resolving Groups..."
# We need to expand groups into individual IDs.
# NOTE: This DOES NOT recurse.
foreach ($DID in $DIDs)
	{
# Return an ADEntry Object so we can see if it is a Group
	$result=Get-JADEntry -id $DID -exact -PSO -properties member,displayname,canonicalname
# Skip it if it is not a Group
	if ($result.member -ne $null)
		{
# Get an array of Group Members
		$members=$result.member
# Look up each member and import it into out list of Delegates
		foreach ($member IN $members)
			{
			$temp=$member
			$temp=$temp.replace("JANUS_CAP\","")
# Filter out SIDs, Built-In Accounts, etc.
			if (($temp -notlike "*S-1-5*") -and($temp -notlike "*NT AUTHORITY*") -and($temp -notlike "*`$*") -and($temp -notlike "*!jm*") -and($temp -notlike "*bps*") -and($temp -notlike "*bpp*") -and($temp -notlike "*JANUSADMIN*") -and($temp -notlike "*GRP_IT_*") -and($temp -notlike "*ADM_*") -and($temp -notlike "*AFADM*") -and($temp -notlike "*IT Messaging Services*") -and($temp -notlike "*Servers*") -and($temp -notlike "*EVucusxm*"))
				{
				$DIDs+=$temp
				}
			}
		}
	}

write-host "Gathering Display Names..."
# Convert the list of AD IDs to Displaynames to ease comprehension
foreach ($DID in $DIDs)
	{
# Get an AD Entry Object for each ID
	$result=Get-JADEntry -id $DID -exact -PSO -properties displayname, canonicalname,SAMAccountName
# Filter out Disabled IDs, etc.
	if (($result.canonicalname -notlike "*disabled*") -and ($result.canonicalname -like "*Janus`/User*"))
		{
		$results+="$($result.DisplayName) ($($result.SAMAccountName))"
		}
	}
} # Process

# The End section executes once regardless of how many objects are passed through the pipeline
end
{
# You can copy the results to the screen or pipe the output to a TXT file.
Write-Output "`n`nThe following mailboxes have control of the Calendar for $ID`:"
$results | sort -unique | ft -auto -wrap

# Variable cleanup
Remove-Variable test
Remove-Variable results
Remove-Variable result
Remove-Variable DIDs
Remove-Variable temp
Remove-Variable members
Remove-Variable dels
Remove-Variable mbps
Remove-Variable bips
Remove-Variable success
} # End
} # Function Show-CRPermissions

Function Update-KVSTempUser
{
<#
	.SYNOPSIS
		Checks to see if user is 0-day archived.

	.DESCRIPTION
		Checks to see if user is 0-day archived. Will consider it archived if it
		can retrieve 10 messages without finding an IPM Noate and it can find at
		least one EV Shortcut. If it determines that Archivng is comlete, moves
		user to KVS Archive OU. Otherwise it enables account and unhides them
		from the Exchange Address Book.

	.PARAMETER  ID
		Specify user to process.

	.PARAMETER  d
		Specify debug level of output.

	.EXAMPLE
		Update-KVSTempUser -id jm27253

		DESCRIPTION
		===========
		Checks to see if JM27253 is 0-day archived. Will consider it archived if
		it can retrieve 10 messages without finding an IPM Noate and it can find
		at least one EV Shortcut. If it determines that Archivng is comlete, moves
		user to KVS Archive OU. Otherwise it enables account and unhides them
		from the Exchange Address Book.
		
	.NOTES
		Requires Janus and EWS Modules.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the first parameter")]
		[String]
        [Alias("Identity")]
$ID,
        [Parameter(Position = 1, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Specify Debug level of output is desired.")]
		[Switch]
$d
)

if ($global:JanusPSModule -ne $true) {Import-Module Janus -erroraction stop}

# Initialization
[bool]$complete=$false
[bool]$EVStubs=$false
[string]$report="Mailboxes in KVS_TEMP OU for more than 3 days:`n================================================`n"
[string]$entry=""
[string]$activity=@()
[array]$messages=@()

# Get all AD Entries in KVS_Temp
$KVST=get-jadentry -id $ID -inou "LDAP://OU=KVS _Temp,OU=Disabled Users,OU=Janus,DC=janus,DC=cap" -pso -properties name,mail
$CMD="Executing: `$KVST=get-jadentry -id $ID -inou `"LDAP://OU=KVS _Temp,OU=Disabled Users,OU=Janus,DC=janus,DC=cap`" -pso -properties name,mail - Successful: $?"
if ($d) {write-host "$CMD"}

# Checking that there are results
if ($KVST.count -eq 0) {write-error ("No Accounts found, exiting...");$LOG="No Accounts found, exiting...";out-file -file "\\p-ucadm01\d$\Scripts\Logs\Archive-EV.LOG" -inputobject $LOG -append;write-error ("No Accounts found, exiting...") -erroraction stop}
# ADO is AD Object
foreach ($ADO IN $KVST)
	{
# Reset Mailbox Tracking Flags
	Set-JDSACLs $($ADO.name) bpp_exspaudit -perms genericall -allow
	Set-JDSACLs $($ADO.name) bpp_kvsadmin -perms genericall -allow
	$complete=$true
	$EVStubs=$false
	$mb=$ADO.mail
	$ID=$ADO.name
	$messages=Get-EWSMailMessage -Mailbox $mb -resultsize 10 | select messageclass,Sent
	$EWSSuccessful=$?
	$CMD="Executing: `$messages=Get-EWSMailMessage -Mailbox $mb -resultsize 10 `| select messageclass,Sent - Successful: $EWSSuccessful"
	if ($d) {write-host "$CMD"}
# If there aren't enough messages in the inbox, get some more
	if (($messages.count -lt 10) -and($EWSSuccessful -eq $true)) {$messages+=Get-EWSMailMessage -Mailbox $mb -resultsize 10 -folder "SentItems" | select messageclass,Sent;$messages+=Get-EWSMailMessage -Mailbox $mb -resultsize 10 -folder "DeletedItems" | select messageclass,Sent}

# If you can't retrieve any messages, the archiving is not complete
	if (-not ($EWSSuccessful)) {$complete=$false}

# If the message retreival command was successful but we have no e-mails, Archiving is done
	if (($EWSSuccessful) -and($messages.count -eq 0)) {$complete=$true;$EVStubs=$true}

	foreach($mail IN $messages)
		{
		if ($complete -eq $false) {break}
# If the message is not a draft or a stub, archiving is not complete
		if (($mail.messageclass -eq "IPM.note") -and($mail.Sent -ne "") -and($mail.messageclass -notlike "*enterprisevault*")) {$complete=$false}

# We use EVStubs to confirm that we found at least one Stub
		if ($mail.messageclass -like "*enterprisevault*") {$EVStubs=$true}
		} # foreach($mail IN $messages)
	write-output ("Is $mb Complete: $complete")
# Run the Process-EVUser cmdlet if there was at least one stub and no non-archived mail
	if (($complete) -and($EVStubs))
		{
		Update-JDisabledUser -ID $ID -erroraction silentlycontinue
		$CMD="Executing: Update-JDisabledUser -ID $ID - Successful: $?"
		if ($d) {write-host "$CMD"}
		$activity+=$CMD
		$activity+="`n`n"
		} # if (($complete) -and($EVStubs))
	} # foreach ($ADO IN $KVST)

start-sleep -s 15

# Looks up the accounts left in KVS_Temp
$KVST=get-jadentry -id $ID -inou "LDAP://OU=KVS _Temp,OU=Disabled Users,OU=Janus,DC=janus,DC=cap" -pso -properties name,mail,useraccountcontrol,msexchhidefromaddresslists,whenchanged
$CMD="Executing: `$KVST=get-jadentry -id $ID -inou `"LDAP://OU=KVS _Temp,OU=Disabled Users,OU=Janus,DC=janus,DC=cap`" -pso -properties name,mail,useraccountcontrol,msexchhidefromaddresslists,whenchanged - Successful: $?"
if ($d) {write-host "$CMD"}
foreach ($ADO IN $KVST)
	{
	$mb=$ADO.mail
	$ID=$ADO.name
	$now=get-date
	$threshhold=$now.adddays(-3)
# If $hidden is true, the account is not visible in the GAL
	$hidden=$ADO.msexchhidefromaddresslists
# Disabled will be 0 if the account is enabled, 2 if disabled
	$array=$ADO.useraccountcontrol.split("`n`n")
	$disabled=$array[0] -band 2
	$ADO.whenchanged
# Converts the $ADO.whenchanged to a date format, silently continue is specified in case this value is empty
	$modified=get-date($ADO.whenchanged) -erroraction silentlycontinue
	if(($hidden) -or($disabled))
		{
# This command enables the account and unhides it from the GAL
		set-Jadentry -id $ID -useraccountcontrol 512 -msexchhidefromaddresslists FALSE
		$log="set-Jadentry -id $ID -useraccountcontrol 512 -msexchhidefromaddresslists `$false - Successful: $?"
		if ($d) {write-host "$log"}
		$activity+=$log
		$activity+="`n`n"
		}
# If the account has not been modified in 3 days, we need to report it to the UC Team
	elseif($modified -lt $threshhold){$entry=$ADO.name + " - " + $ADO.mail + "`n`n";$report=$report+$entry}
	} # foreach ($ADO IN $KVST)

write-output ("Process Complete!")
write-output "KVS_Temp Activity Report`n`n$activity"

remove-variable KVST -erroraction silentlycontinue
remove-variable messages -erroraction silentlycontinue
remove-variable report -erroraction silentlycontinue
remove-variable activity -erroraction silentlycontinue
} # Function Update-KVSTempUser

#Function Update-JDisabledUser
Function Update-JArchiveUser
{
<#
	.SYNOPSIS
		Processes a single user in Disabled Users OU and puts in KVS_Temp OU when done.

	.DESCRIPTION
		Processes a single user in Disabled Users OU enabling accout and unhiding
		from Exchange Address Book and puts in KVS_Temp OU when done. Puts user in
		zDisabled Users if the user does not have a mailbox.

	.PARAMETER  ID
		Specify user to Process.

	.PARAMETER  D
		Specify debug-level of output.

	.EXAMPLE
		Update-JDisabledUsers -ID JM27253

		DESCRIPTION
		===========
		If JM27253 does not have a mailbox, moves it to zDisabled Users OU. Otherwise, enables user, unhides user from Exchange Address Book and moves it to KVS_Temp OU.
		
	.NOTES
		Requires Janus and Exchange Modules.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the first parameter")]
		[String]
        [Alias("Identity")]
$ID,
        [Parameter(Position = 1, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the second parameter")]
		[Switch]
$d
)

if ($global:JanusPSModule -ne $true) {Import-Module Janus -erroraction stop}

if ($global:ExchangeSnapIn -ne $true) {Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -erroraction silentlyContinue}

# Initialization
$CSV=@("Mailbox,StoreDB,Server,LogDrive")
[string]$CMD=""
$now=get-date
[string]$report=""

# The following Routine moves users that are in the Disabled Users OU that do not have mailboxes to the zdisabled IDs OU. $DU is Diabled USers
$DU=Get-JADEntry -id $ID -exact -InOU "OU=Disabled Users,OU=Janus,DC=janus,DC=cap" -properties name,extensionAttribute7,extensionAttribute8,extensionAttribute9,altRecipient,description -norecurse -pso -ErrorAction:SilentlyContinue

if ($DU -ne $null)
{
# $ADO is ADObject
foreach ($ADO in $DU)
	{
write-host("Processing $($ADO.name) ...")
	Set-JDSACLs $($ADO.name) bpp_exspaudit -perms genericall -allow
	Set-JDSACLs $($ADO.name) bpp_kvsadmin -perms genericall -allow
	$mailbox=get-mailbox -id $ADO.name
	$success=$?
write-host("$($mailbox.count) mailbox(es) found...")
	if ($success -eq $false)
		{
write-host("Processing $($ADO.name) for zDisabled Users...")
		Move-JADEntry -id $ADO.name -ou "OU=zDisabled IDs,DC=janus,DC=cap"
		$log="Move-JADEntry -id $ADO.name -ou `"OU=zDisabled IDs,DC=janus,DC=cap`" - Successful: $?"
		if ($d) {write-host "$log"}
		$report+=$log
		$report+="`n`n"
		}
	else
		{
write-host("Processing $($ADO.name) for KVS_Temp...")
# Sets password
		Set-JPassword -id $ADO.name -Pass "T3stingin$tage"
		$log="Set-JPassword -id $ADO.name -Pass `"T3stingin$tage`" - Successful: $?"
		if ($d) {write-host "$log"}
		$report+=$log
		$report+="`n`n"
# Enable account
		Set-JADEntry -id $ADO.name -useraccountcontrol 512
		$log="Set-JADEntry -id $ADO.name -useraccountcontrol 512 - Successful: $?"
		if ($d) {write-host "$log"}
		$report+=$log
		$report+="`n`n"
# Unhide mailbox
		Set-mailbox -id $mailbox -HiddenFromAddressListsEnabled $false
		$log="Set-mailbox -id $mailbox -HiddenFromAddressListsEnabled `$false - Successful: $?"
		if ($d) {write-host "$log"}
		$report+=$log
		$report+="`n`n"
# Disable ACtive Sync
		Set-CASMailbox -identity $mailbox -ActiveSyncEnabled:$False -ActiveSyncMailboxPolicy "Default Block"
		$log="Set-CASMailbox -identity $mailbox -ActiveSyncEnabled:`$False -ActiveSyncMailboxPolicy `"Default Block`" - Successful: $?"
		if ($d) {write-host "$log"}
		$report+=$log
		$report+="`n`n"
# Use Blockmail to determine if their mail should be disabled
		$blockmail=$true
		$CA7=get-date($ADO.extensionAttribute7) -erroraction silentlyContinue
		$CA8=get-date($ADO.extensionAttribute8) -erroraction silentlyContinue
		$CA9=get-date($ADO.extensionAttribute9) -erroraction silentlyContinue
		if((($CA7 -ne "") -and($CA7 -ne $null)) -and($CA7 -lt $now))
			{
			$blockmail=$false
			}
		if((($CA8 -ne "") -and($CA8 -ne $null)) -and($CA8 -lt $now))
			{
			$blockmail=$false
			}
		if((($CA9 -ne "") -and($CA9 -ne $null)) -and($CA9 -lt $now))
			{
			$blockmail=$false
			}
		if ((($ADO.altRecipient) -ne "") -and(($ADO.altRecipient) -ne ""))
			{
			$blockmail=$false
			}
# If we are blocking mail, perform these cmdlets
		if ($blockmail -eq $true)
			{
# Only accept mail from bpp_exchangeamdin
			set-mailbox -id $mailbox  -AcceptMessagesOnlyFrom BPP_ExchangeAdmin
			$log="set-mailbox -id $mailbox  -AcceptMessagesOnlyFrom BPP_ExchangeAdmin - Successful: $?"
			if ($d) {write-host "$log"}
			$report+=$log
			$report+="`n`n"
# Require authorization from all mail
			set-mailbox -id $mailbox -RequireSenderAuthenticationEnabled $true
			$log="set-mailbox -id $mailbox -RequireSenderAuthenticationEnabled `$true - Successful: $?"
			if ($d) {write-host "$log"}
			$report+=$log
			$report+="`n`n"
			}
# Move AD Object to KVS Temp
		Move-JADEntry -id $ADO.name -ou "KVS_Temp"
		$log="Move-JADEntry -id $ADO.name -ou `"KVS _Temp`" - Successful: $?"
		if ($d) {write-host "$log"}
		$report+=$log
		$report+="`n`n"
		}
	}

remove-variable DU
} # if ($DU -ne $null)
} # Function Update-JDisabledUser

Function Export-PSOToCSV
{
<#
	.SYNOPSIS
		Converts a PS Object to CSV File

	.DESCRIPTION
		Converts a the specified PS Object (including from the pipeline) to CSV File.

	.PARAMETER  PSO
		Specify the PSObject Variable to be converted.

	.PARAMETER  file
		Specify the file name of the CSV File.

	.EXAMPLE
		Export-PSOToCSV -pso $test -file "D:\Scripts\Temp\Test.CSV"
		
		Description
		-----------
		Reads the $test PSO variable and outputs the contents to a CSV file named D:\Scripts\Temp\Test.CSV.
		
	.EXAMPLE
		$test | Export-PSOToCSV -file "D:\Scripts\Temp\Test.CSV"
		
		Description
		-----------
		Reads the $test PSO variable and outputs the contents to a CSV file named D:\Scripts\Temp\Test.CSV.
		
	.NOTES
		Requires the Loading of the Janus Module.
		http://stackoverflow.com/questions/16407239/exporting-collection-of-hashtable-data-to-csv

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the name of a PSObject for conversion")]
		[PSCustomObject]
        [Alias("PSCustomObject")]
        [Alias("PSObject")]
$PSO,
        [Parameter(Position = 1, Mandatory = $true, ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false, HelpMessage="Enter the filename to convert to")]
		[String]
$file
)
begin{}
process
{
# Export-CSV is expecting a Collection, not an array of hashtables.
[System.Collections.ArrayList]$PSC=$PSO
# This works with Export-CSV
$PSC | Export-Csv -Path $file -NoTypeInformation -Force
}
end
{
Remove-Variable PSC
}
} # Function Export-PSOToCSV

Function Show-JMailboxPermissions
{
<#
	.SYNOPSIS
		Report Mailbox Permissions at every level.

	.DESCRIPTION
		Report Mailbox Permissions at every level by checking Mailbox Permissions, Send As Permissions, Send on Behalf Permissions, Delegates and Folder Permissions.

	.PARAMETER  ID
		Specify the ID to check.

	.PARAMETER  Report
		Specify that a CSV report of permissions discovered should be save at the following location:
		\\p-ucadm01\d$\Scripts\Logs\<MailboxID>-Mailbox-Permissions-<MMDDYY>.csv

	.EXAMPLE
		Show-JMailboxPermissions -id JM27253

		DESCRIPTION
		===========
		Who has access to the mailbox identified as JM27253.
		
	.EXAMPLE
		Show-JMailboxPermissions -id JM27253 -report

		DESCRIPTION
		===========
		Save a report of who has access to the mailbox identified as JM27253 to a CSV file named \\p-ucadm01\d$\Scripts\Logs\JM27253-Mailbox-Permissions-072214.csv
		
	.NOTES
		Require Janus, Exchange and EWS Modules.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the first parameter")]
		[String]
        [Alias("Identity")]
$ID,
		[Switch]
$Report,
		[Switch]
$d
)

# Functions and Filters


# Main Script
# The Begin section executes once regardless of how many objects are passed through the pipeline
begin
{
$error.clear()
if ($global:JanusPSModule -ne $true) {Import-Module Janus -erroraction stop}

Write-host("Checking for Exchange 2010 Module...")
$test=Get-DatabaseAvailabilityGroup
$global:ExchangeSnapIn=$?
if ($global:ExchangeSnapIn -ne $true)
	{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
	$test=Get-DatabaseAvailabilityGroup
	$success=$?
	If ($success -ne $true)
		{
		write-error ("Unable to load Exchange 2010 Module, exiting...") -erroraction stop
		} else # If ($success -ne $true)
		{
		$global:ExchangeSnapIn = $true
		} # else
	} # if ($global:ExchangeSnapIn -ne $true)
Write-host("Exchange 2010 Module Loaded!")

# Initialization
[string]$errorstring=$error | out-string
[string]$datecode=get-date -DisplayHint Date -Format MMddyy
$MsgFolderRoot = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot
[string]$User=""
[string]$Name=""
[String]$CSV=""
[String]$file="\\p-ucadm01\d$\Scripts\Logs\$ID-Mailbox-Permissions-$datecode.csv"
[String]$path=""
[String]$strout=""
}

# The process section runs once for each object in the pipeline
process
{
# Error chacking
if ($d) {write-host "Looking up $ID"}
$test=get-mailbox $ID
$successful=$?
if ($d) {write-host "Was Lookup successful: $successful"}
if (-not $successful)
	{
	$errorstring="Cannot locate mailbox with the folloiwng ID: $ID`nError Loge:`n"
	$errorstring+=$error | out-string
	write-error "$errorstring" -erroraction stop
	}
$ID=$test.PrimarySmtpAddress
if ($d) {write-host "Parameter=$ID"}
$Displayname=$test.Displayname
if ($d) {write-host "Display Name: $Displayname"}
$CSV="`"Name`",`"$Displayname`"`n"
if ($d) {write-host "CSV: `n$CSV"}
$CSV+="`"Address`",`"$ID`"`n"
if ($d) {write-host "CSV: `n$CSV"}
$SAMAccountName=$test.SAMAccountName
if ($d) {write-host "ID: $SAMAccountName"}
$CSV+="`"Account`",`"$SAMAccountName`"`n"
$CSV+="`n"
$CSV+="AD Access`n"
if ($d) {write-host "CSV: `n$CSV"}
$CSV+="`"Name`",`"Account`",`"Permissions`"`n"
if ($d) {write-host "CSV: `n$CSV"}

if ($d) {write-host "Looking up permissions"}
$MBPs=Get-MailboxPermission -id $ID | where {$_.User -notlike "*JANUSADMIN*"} | where {$_.User -notlike "*S-1-5*"} | where {$_.User -notlike "*NT AUTHORITY*"} | where {$_.User -notlike "*bpp*"} | where {$_.User -notlike "*bps*"} | where {$_.User -notlike "*exchange*"} | where {$_.User -notlike "*!jm*"} | where {$_.User -notlike "*B-UC*"} | where {$_.User -notlike "*P-UC*"} | where {$_.User -notlike "*ADM_*"} | where {$_.User -notlike "*EVuc*"} | where {$_.User -notlike "*AFUsers*"} | where {$_.User -notlike "*AFAdmin*"} | where {$_.User -notlike "*GRP_IT_Security_Requests*"} | sort User -unique
if ($d) {write-host "Was lookup successful: $?"}
foreach ($MBP IN $MBPs)
	{
if ($d) {write-host "Processing $MBP"}
	$User=$MBP.User
if ($d) {write-host "Found $User"}
	$User=$User.replace("JANUS_CAP`\","")
if ($d) {write-host "User: $User"}
	$Account=$User
if ($d) {write-host "ID: $Account"}
	$Perms=$MBP.AccessRights
if ($d) {write-host "Permissions: $Perms"}
	$ADE=Get-JADEntry -id $User -exact -pso -properties Name,DisplayName,member
	if ($ADE.DisplayName -eq $null)
		{
		$Name=$ADE.Name
		} else
		{
		$Name=$ADE.DisplayName
		}
if ($d) {write-host "Name: $Name"}
	$CSV+="`"$Name`",`"$Account`",`"$Perms`"`n"
if ($d) {write-host "CSV: `n$CSV"}
	if ($ADE.member -ne $null)
		{
if ($d) {write-host "PRocessing: $($ADE.Name)"}
		$members=Get-JGroupMembers $($ADE.Name)
		$members= $members | sort SAMAccountName -unique
		foreach ($Member IN $members)
			{
			$Name=$Member.displayname
if ($d) {write-host "Name: $Name"}
			$Account=$Member.samaccountname
if ($d) {write-host "ID: $Account"}
			$CSV+="`"$Name`",`"$Account`",`"$Perms`"`n"
if ($d) {write-host "CSV: `n$CSV"}
			}
		}
	}
$CSV+="`n"
$CSV+="Delegate Permissions`n"
$CSV+="`"Name`",`"Address`",`"Calendar Permissions`",`"Copy Delegate On Invite`",`"Who Receives Invite Copies`",`"Inbox Permissions`",`"Contacts Permissions`",`"Tasks Permissions`",`"Notes Permissions`",`"Private Items Visible`"`n"
$Delegates=Get-JEWSDelegates -ID $ID
foreach ($del IN $Delegates)
	{
if ($d) {write-host "Checking: $del`n"}
	if ($del.DelegateDisplayName -ne $null)
		{
if ($d) {write-host "Processing: $($del.DelegatePrimarySmtpAddress)`n"}
		$Name=$del.DelegateDisplayName
		$Address=$del.DelegatePrimarySmtpAddress
		$Cal=$del.CalendarFolderPermissionLevel
		$Copy=$del.ReceiveCopiesOfMeetingMessages
		$MailOpt=$del.MeetingRequestsDeliveryScope
		$Inbox=$del.InboxFolderPermissionLevel
		$Contacts=$del.ContactsFolderPermissionLevel
		$Taaks=$del.TasksFolderPermissionLevel
		$Notes=$del.NotesFolderPermissionLevel
		$Priv=$del.ViewPrivateItems
		$CSV+="`"$Name`",`"$Address`",`"$Cal`",`"$Copy`",`"$MailOpt`",`"$Inbox`",`"$Contacts`",`"$Tasks`",`"$Notes`",`"$Priv`"`n"
if ($d) {write-host "CSV: `n$CSV"}
		}
	}
$CSV+="`n"
$CSV+="Folder Permissions`n"
$CSV+="`"Folder`",`"Name`",`"Permission`"`n"

if ($d) {write-host "Enumerating Folders..."}
$folders=Get-MailboxFolderStatistics -id $ID
if ($d) {write-host "Folders found: $($folders.count)"}

ForEach ($folder in $Folders)
	{
	$path=$folder.Identity
if ($d) {write-host "Examining $path ..."}
	$length=$path.length
	$path=$path.replace("$ID","$ID`:")
	$path=$path.replace("Top of Information Store","")
	$Perms=Get-MailboxFolderPermission -id $path -erroraction silentlycontinue
	foreach ($perm IN $Perms)
		{
		$path=$path.replace("$ID`:","")
		$name=$Perm.User
		$Access=$Perm.AccessRights
		$CSV+="`"$path`",`"$Name`",`"$Access`"`n"
if ($d) {write-host "CSV: `n$CSV"}
		}
	}
$errorstring=$error | out-string
$CSV+="`n"
$CSV+="`"Error Log`:`"`n"
$CSV+="$errorstring"
}

# The End section executes once regardless of how many objects are passed through the pipeline
end
{
if ($report) {$CSV | out-file -file $file -encoding ASCII}
	else
	{
	$strout=$CSV
	$strout=$strout.replace(", ","`*`&`*")
	$strout=$strout.replace(","," - ")
	$strout=$strout.replace("`*`&`*",", ")
	$strout=$strout.replace("`"","")
	$strout | out-string
	}
if ($d) {write-host "CSV: `n$CSV"}

remove-variable errorstring
remove-variable Access
remove-variable Account
remove-variable Address
remove-variable ADE
remove-variable Cal
remove-variable Contacts
remove-variable Copy
remove-variable CSV
remove-variable datecode
remove-variable Delegates
remove-variable Displayname
remove-variable errorstring
remove-variable file
remove-variable folders
remove-variable Inbox
remove-variable length
remove-variable MailOpt
remove-variable MBPs
remove-variable members
remove-variable MsgFolderRoot
remove-variable Name
remove-variable Notes
remove-variable path
remove-variable Perms
remove-variable Priv
remove-variable SAMAccountName
remove-variable success
remove-variable successful
remove-variable Taaks
remove-variable test
remove-variable User
remove-variable strout
}
} # Function Show-JMailboxPermissions

Function Set-JGroupMaintenanceMode
{
<#
	.SYNOPSIS
		Puts a group of systems in SCOM in Maintenance Mode

	.DESCRIPTION
		Puts a group of systems in SCOM in Maintenance Mode

	.PARAMETER  ID
		Enter the name of the server group to put in Maintenance mode.

	.PARAMETER  End
		Enter the Date and time that maintenance mode should end.
		NOTE: This will default to 30 minutes after the cmdlet is executed.

	.PARAMETER  Reason
		Enter the Reason for the Maintenance Mode that Administrators will see in SCOM.
		NOTE: Valid Values are: PlannedOther, UnplannedOther, PlannedHardwareMaintenance, UnplannedHardwareMaintenance, PlannedHardwareInstallation, UnplannedHardwareInstallation, PlannedOperatingSystemReconfiguration, UnplannedOperatingSystemReconfiguration, PlannedApplicationMaintenance, ApplicationInstallation, ApplicationUnresponsive, ApplicationUnstable, SecurityIssue, LossOfNetworkConnectivity. 
		NOTE: Reason will default to PlannedApplicationMaintenance

	.PARAMETER  Comment
		Enter the message Administrators will see in SCOM for this maintenance mode.

	.PARAMETER  d
		Specify that Debug information should be displayed.

	.EXAMPLE
		Set-JGroupMaintenanceMode -id "Janus zzz Test Group"  -end "12/21/2012 1:00 am" -Comment "Testing"
		
		Description
		-----------
		Sets "Janus zzz Test Group" in Maintenance Mode in SCOM through 12-21-2012 1:00 am.
		
	.NOTES
		Requires the Loading of the Janus Module.
		Requires the Loading of the SCOM Module.

#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the name of the comuter to put in maintenence mode")]
		[String]
        [Alias("Identity")]
        [Alias("DisplayName")]
$ID=$(Throw "You must specify an AD ID (e.g., P-UCADM01)."),
        [Parameter(Position = 1, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the date and time that the maintenance mode is supposed to end")]
        [Alias("expires")]
		[system.datetime]
$End=$((get-date).addminutes(30)),
        [Parameter(Position = 3, Mandatory = $false, HelpMessage="Enter the reason for the maintenence mode")]
		[String]
$Reason="PlannedApplicationMaintenance",
        [Parameter(Position = 2, Mandatory = $false, HelpMessage="Enter the maintenence mode comment")]
		[String]
        [Alias("Explanation")]
        [Alias("Note")]
$Comment="IT Maintenance",
		[switch]
$d
)

begin
{
if (test-connection p-scom02.janus.cap) {$server="p-scom02.janus.cap"}
if (test-connection p-scom03.janus.cap) {$server="p-scom03.janus.cap"}
if (test-connection p-scom04.janus.cap) {$server="p-scom04.janus.cap"}
if (test-connection b-scom01.janus.cap) {$server="b-scom01.janus.cap"}
if (test-connection b-scom02.janus.cap) {$server="b-scom02.janus.cap"}
if ($d) {write-host("Entering Function...")}
[system.datetime]$now=get-date
if ($d) {write-host("Checking Date...")}
if ($end -le $now)
	{
	$errorsting="$end is an invalid date/time, Exiting..."
	Write-output("$errorsting")
	$error | out-string
	Write-Error("$errorsting") -erroraction stop
	}
if ($d) {write-host("Confirmed Valid End Date.")}
	
if ($d) {write-host("Testing to see if SCOM Module is loaded.")}
$test=Get-SCOMCommand -erroraction silentlycontinue
if ($test -eq $null)
	{
	if (test-path "C:\Program Files\System Center Operations Manager 2012\Powershell\OperationsManager")
		{
		Add-PSSnapin Microsoft.EnterpriseManagement.OperationsManager.Client -erroraction silentlycontinue
		. "C:\Program Files\System Center Operations Manager 2012\Powershell\OperationsManager\Functions.ps1"
		Start-OperationsManagerClientShell -ManagementServerName: $server -PersistConnection: $true -Interactive: $true
		if ($? -ne $true) {write-error ("Error Loading SCOM Module, exiting...") -erroraction stop}
		}
	}
if ($d) {write-host("Confirmed that SCOM Module is loaded.")}
} # begin

Process
{
if ($d) {write-host("Processing Group $ID ...")}
$test=Get-ScomGroup -DisplayName $ID -erroraction silentlycontinue
if ($test -ne $null)
	{
if ($d) {write-host("Group $ID confirmed.")}
if ($d) {write-host("Creating Geroup Connection...")}
	New-SCOMManagementGroupConnection -ComputerName $server
	$Group=Get-ScomGroup -DisplayName $ID
	$success=$?
if ($d) {write-host("Instance created for $ID - success: $success")}
	if ($success -ne $true)
		{
		$errorsting="Could not get Group $ID in SCOM, Exiting..."
		Write-output("$errorstring")
		$error | out-string
		Write-Error("$errorstring") -erroraction stop
		}
if ($d) {write-host("Setting Maintenance mode for $ID.")}
	$Group.ScheduleMaintenanceMode($now.touniversaltime(), $end.touniversaltime(), "$Reason", "$Comment", "Recursive")
if ($d) {write-host("Maintenance mode for $ID set.")}
	$success=$?

	if ($success -ne $true)
		{
		$errorsting="Failed to Start Maintenence Mode  for $ID in SCOM, Exiting..."
		Write-output("$errorstring")
		$error | out-string
		Write-Error("$errorstring") -erroraction stop
		}
}
write-host("Maintenance Mode Started.")
} # Process

end
{
remove-variable test
remove-variable Now
remove-variable Group
remove-variable Success
} # end
} # Function Set-JGroupMaintenanceMode

Function Get-JUserExchangeRoles
{
<#
	.SYNOPSIS
		List Exchange Roles assigned to a user

	.DESCRIPTION
		List Exchange Roles assigned to a user

	.PARAMETER  ID
		Enter the name of the server group to put in Maintenance mode.

	.PARAMETER  d
		Specify that Debug information should be displayed.

	.EXAMPLE
		Get-JUserExchangeRoles -ID JM27253
		
		Description
		-----------
		Lists all roles assignesd to JM27253
		
	.NOTES
		Requires the Loading of the Exchange SnapIn.

	.LINK
		http://andersonpatricio.ca/how-to-check-admin-roles-for-an-user/
#>

	[CmdletBinding()]
param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the name of the comuter to put in maintenence mode")]
		[String]
        [Alias("Identity")]
        [Alias("DisplayName")]
$ID=$(Throw "You must specify an AD ID (e.g., JM27253)."),
        [Parameter(Position = 1, Mandatory = $false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the date and time that the maintenance mode is supposed to end")]
		[switch]
$d
)

if ($global:JanusPSModule -ne $true) {Import-Module Janus -erroraction stop;$global:JanusPSModule=$?}

if ($d) {Write-host(Write-host("Checking for Exchange 2010 Module..."))}
$test=Get-DatabaseAvailabilityGroup
$global:ExchangeSnapIn=$?
if ($global:ExchangeSnapIn -ne $true)
	{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
	$test=Get-DatabaseAvailabilityGroup
	$success=$?
	If ($success -ne $true)
		{
		write-error ("Unable to load Exchange 2010 Module, exiting...") -erroraction stop
		} else # If ($success -ne $true)
		{
		$global:ExchangeSnapIn = $true
		} # else
	} # if ($global:ExchangeSnapIn -ne $true)
if ($d) {Write-host("Exchange 2010 Module Loaded!")}

if ($d) {write-host "Verifying $ID is a valid user"}
$test=Get-User $ID
$success=$?
if ($d) {write-host "Can Exchange find $ID - $success"}
if ($success -eq $true)
	{
if ($d) {write-host "Processing $ID ..."}
	Get-ManagementRoleAssignment –RoleAssignee $ID
if ($d) {write-host "$ID precessed."}
	}
if ($d) {write-host "Execution complete"}
}

function Get-JScheduledTasks
{
<#
.SYNOPSIS
Returns Scheduled Tasks.

.DESCRIPTION
Returns Scheduled Tasks as PSObjects for the specified computer.

.PARAMETER ComputerName
Specify the computer or list of computers to check.
NOTE: Will run against all servers with an AD Description of "UC Server" if not specified.

.PARAMETER Subfolders
Specifies whether to support task subfolders (Windows Vista/Server 2008 or later only).

.PARAMETER Hidden
Specifies whether to output hidden tasks.

.PARAMETER Inventory
Specifies that Tasks are supposed to be exported to an inventory file (Located at D:\Scripts\Logs\ScheduledTasks-Inventory-mmddyy.CSV).

.EXAMPLE
PS C:\> Get-ScheduledTask -ComputerName SERVER1
This command outputs scheduled tasks in the root tasks folder on the computer SERVER1.

.EXAMPLE
PS C:\> Get-Content Computers.txt | Get-ScheduledTask
This command outputs all scheduled tasks for each computer listed in the file Computers.txt.

.LINK
http://windowsitpro.com/powershell/how-use-powershell-report-scheduled-tasks
#>

[CmdletBinding()]
param(
[parameter(Position=0,ValueFromPipeline=$TRUE)]
[String[]]
[Alias("Identity")]
[Alias("Computer")]
[Alias("System")]
[Alias("DNSHostNAme")]
[Alias("SAMAccountName")]
[Alias("Server")]
[Alias("ServerName")]
$ComputerName="*",
[switch]
$Subfolders,
[switch]
$Hidden,
[switch]
$Inventory
)

begin
{
$PIPELINEINPUT = (-not $PSBOUNDPARAMETERS.ContainsKey("ComputerName")) -and (-not $ComputerName)
$MIN_SCHEDULER_VERSION = "1.2"
$TASK_ENUM_HIDDEN = 1
$TASK_STATE = @{0 = "Unknown"; 1 = "Disabled"; 2 = "Queued"; 3 = "Ready"; 4 = "Running"}
$ACTION_TYPE = @{0 = "Execute"; 5 = "COMhandler"; 6 = "Email"; 7 = "ShowMessage"}
[String[]] $TaskName="*"
[String[]] $CPUs=@()
$results=@()
[string]$datecode=get-date -DisplayHint Date -Format MMddyy

# If It is *, then we need to collect all the servers with an AD Descriptionm of "UC Server"
if ($ComputerName -eq "*")
	{
	$ComputerName=Get-JADEntry "UC Server" -description -pso -properties dnshostname | select dnshostname
	} else
	{
	foreach ($ID in $ComputerName)
		{
		if (Test-Connection -computername $ID -count 1 -quiet) {$CPUs+=(Get-JADEntry -id $ID -exact -pso -properties dnshostname).dnshostname}
		}
	$ComputerName=$CPUs
	}

# Try to create the TaskService object on the local computer; throw an error on failure
try
	{
	$TaskService = new-object -comobject "Schedule.Service"
	}
catch [System.Management.Automation.PSArgumentException]
	{
	throw $_
	}

# Returns a version number as a string (x.y); e.g. 65537 (10001 hex) returns "1.1"
function convertto-versionstr([Int] $version)
{
$major = [Math]::Truncate($version / [Math]::Pow(2, 0x10)) -band 0xFFFF
$minor = $version -band 0xFFFF
"$($major).$($minor)"
}

# Returns a string "x.y" as a version number; e.g., "1.3" returns 65539 (10003 hex)
function convertto-versionint([String] $version)
{
$parts = $version.Split(".")
$major = [Int] $parts[0] * [Math]::Pow(2, 0x10)
$major -bor [Int] $parts[1]
}

# Returns a list of all tasks starting at the specified task folder
function get-task($taskFolder)
{
# Gather the tasks
$tasks = $taskFolder.GetTasks($Hidden.IsPresent -as [Int])
# Return the Tasks
$tasks | foreach-object { $_ }
if ($SubFolders)
	{
	try
		{
		$taskFolders = $taskFolder.GetFolders(0)
		$taskFolders | foreach-object { get-task $_ $TRUE }
		}
	catch [System.Management.Automation.MethodInvocationException]
		{
		}
	}
Remove-Variable taskFolders -erroraction silentlycontinue
Remove-Variable tasks
}

# Returns a date if greater than 12/30/1899 00:00; otherwise, returns nothing
function get-OLEdate($date)
{
if ($date -gt [DateTime] "12/30/1899") { $date }
}

function get-scheduledtask2($computerName)
{
# Assume $NULL for the schedule service connection parameters unless -ConnectionCredential used
$userName = $domainName = $connectPwd = $NULL
try
	{
# You need to connect to the computer in order to retreive the Tasks. With User, pass and domain set to Null, it uses pass through cred.
	$TaskService.Connect($ComputerName, $userName, $domainName, $connectPwd)
	}
catch [System.Management.Automation.MethodInvocationException]
	{
	write-warning "$computerName - $_"
	return
	}
$serviceVersion = convertto-versionstr $TaskService.HighestVersion
$vistaOrNewer = (convertto-versionint $serviceVersion) -ge (convertto-versionint $MIN_SCHEDULER_VERSION)
$rootFolder = $TaskService.GetFolder("\")
# Call the function to return the tasks on the computer
$taskList = get-task $rootFolder
# If Tasklist is null or 0, there are no tasks on that computer
if (-not $taskList) { return }
#Otherwise, we need to retreive the properties of each Task
foreach ($task in $taskList)
	{
	foreach ($name in $TaskName)
		{
# Assume root tasks folder (\) if task folders supported
		if ($vistaOrNewer)
			{
			if (-not $name.Contains("\")) { $name = "\$name" }
			}
			if ($task.Path -notlike $name) { continue }
# Retreive the Task paramters
			$taskDefinition = $task.Definition
			$actionCount = 0
			foreach ($action in $taskDefinition.Actions)
				{
				$actionCount += 1
				$output = new-object PSObject
# PROPERTY: ComputerName
				$output | add-member NoteProperty ComputerName $computerName
# PROPERTY: ServiceVersion
				$output | add-member NoteProperty ServiceVersion $serviceVersion
# PROPERTY: TaskName
				if ($vistaOrNewer)
					{
					$output | add-member NoteProperty TaskName $task.Path
					} else
					{
					$output | add-member NoteProperty TaskName $task.Name
					}
#PROPERTY: Enabled
				$output | add-member NoteProperty Enabled ([Boolean] $task.Enabled)
# PROPERTY: ActionNumber
				$output | add-member NoteProperty ActionNumber $actionCount
# PROPERTIES: ActionType and Action
# Old platforms return null for the Type property
				if ((-not $action.Type) -or ($action.Type -eq 0))
					{
					$output | add-member NoteProperty ActionType $ACTION_TYPE[0]
					$output | add-member NoteProperty Action "$($action.Path) $($action.Arguments)"
					} else
					{
					$output | add-member NoteProperty ActionType $ACTION_TYPE[$action.Type]
					$output | add-member NoteProperty Action $NULL
					}
# PROPERTY: LastRunTime
				$output | add-member NoteProperty LastRunTime (get-OLEdate $task.LastRunTime)
# PROPERTY: LastResult
				if ($task.LastTaskResult)
					{
# If negative, convert to DWORD (UInt32)
					if ($task.LastTaskResult -lt 0)
						{
						$lastTaskResult = "0x{0:X}" -f [UInt32] ($task.LastTaskResult + [Math]::Pow(2, 32))
						} else
						{
						$lastTaskResult = "0x{0:X}" -f $task.LastTaskResult
						}
					} else
					{
					$lastTaskResult = $NULL# fix bug in v1.0-1.1 (should output $NULL)
					}
				$output | add-member NoteProperty LastResult $lastTaskResult
# PROPERTY: NextRunTime
				$output | add-member NoteProperty NextRunTime (get-OLEdate $task.NextRunTime)
# PROPERTY: State
				if ($task.State)
					{
					$taskState = $TASK_STATE[$task.State]
					}
				$output | add-member NoteProperty State $taskState
				$regInfo = $taskDefinition.RegistrationInfo
# PROPERTY: Author
				$output | add-member NoteProperty Author $regInfo.Author
# The RegistrationInfo object's Date property, if set, is a string
				if ($regInfo.Date)
					{
					$creationDate = [DateTime]::Parse($regInfo.Date)
					}
				$output | add-member NoteProperty Created $creationDate
# PROPERTY: RunAs
				$principal = $taskDefinition.Principal
				$output | add-member NoteProperty RunAs $principal.UserId
# PROPERTY: Elevated
				if ($vistaOrNewer)
					{
					if ($principal.RunLevel -eq 1) { $elevated = $TRUE } else { $elevated = $FALSE }
					}
				$output | add-member NoteProperty Elevated $elevated
# Output the object
				$output
				}
			}
		}
Remove-Variable actionCount -erroraction silentlycontinue
Remove-Variable creationDate -erroraction silentlycontinue
Remove-Variable lastTaskResult -erroraction silentlycontinue
Remove-Variable major -erroraction silentlycontinue
Remove-Variable minor -erroraction silentlycontinue
Remove-Variable output -erroraction silentlycontinue
Remove-Variable parts -erroraction silentlycontinue
Remove-Variable principal -erroraction silentlycontinue
Remove-Variable regInfo -erroraction silentlycontinue
Remove-Variable rootFolder -erroraction silentlycontinue
Remove-Variable serviceVersion -erroraction silentlycontinue
Remove-Variable taskDefinition -erroraction silentlycontinue
Remove-Variable taskFolders -erroraction silentlycontinue
Remove-Variable taskList -erroraction silentlycontinue
Remove-Variable TaskName -erroraction silentlycontinue
Remove-Variable tasks -erroraction silentlycontinue
Remove-Variable TaskService -erroraction silentlycontinue
Remove-Variable taskState -erroraction silentlycontinue
Remove-Variable userName -erroraction silentlycontinue
Remove-Variable vistaOrNewer -erroraction silentlycontinue
	}
}

process {
if ($PIPELINEINPUT)
	{
	$results=get-scheduledtask2 $_
	} else
	{
	$results=$ComputerName | foreach { get-scheduledtask2 $_ }
	}
}

end
{
if ($Inventory)
	{
	$results | export-csv D:\Scripts\Logs\ScheduledTasks-Inventory-$datecode.CSV
	} else
	{
	$results
	}

Remove-Variable results
Remove-Variable ACTION_TYPE -erroraction silentlycontinue
Remove-Variable actionCount -erroraction silentlycontinue
Remove-Variable CPUs -erroraction silentlycontinue
Remove-Variable creationDate -erroraction silentlycontinue
Remove-Variable lastTaskResult -erroraction silentlycontinue
Remove-Variable major -erroraction silentlycontinue
Remove-Variable MIN_SCHEDULER_VERSION -erroraction silentlycontinue
Remove-Variable minor -erroraction silentlycontinue
Remove-Variable output -erroraction silentlycontinue
Remove-Variable parts -erroraction silentlycontinue
Remove-Variable PIPELINEINPUT -erroraction silentlycontinue
Remove-Variable principal -erroraction silentlycontinue
Remove-Variable regInfo -erroraction silentlycontinue
Remove-Variable rootFolder -erroraction silentlycontinue
Remove-Variable serviceVersion -erroraction silentlycontinue
Remove-Variable TASK_ENUM_HIDDEN -erroraction silentlycontinue
Remove-Variable TASK_STATE -erroraction silentlycontinue
Remove-Variable taskDefinition -erroraction silentlycontinue
Remove-Variable taskFolders -erroraction silentlycontinue
Remove-Variable taskList -erroraction silentlycontinue
Remove-Variable TaskName -erroraction silentlycontinue
Remove-Variable tasks -erroraction silentlycontinue
Remove-Variable TaskService -erroraction silentlycontinue
Remove-Variable taskState -erroraction silentlycontinue
Remove-Variable userName -erroraction silentlycontinue
Remove-Variable vistaOrNewer -erroraction silentlycontinue
}

} # Get-JScheduledTasks

Function Set-JScheduledTaskPassword
{
<#
.SYNOPSIS
Sets the password for one or more scheduled tasks on a computer.

.DESCRIPTION
Sets the password for one or more scheduled tasks on a computer.

.PARAMETER TaskName
One or more scheduled task names. Wildcard values are not accepted. This parameter accepts pipeline input.

.PARAMETER TaskCredential
The password for the scheduled task. If you don't specify this parameter, you will be prompted for credentials.

.PARAMETER ComputerName
The computer name where the scheduled task(s) reside.
NOTE: If not specified, uses the computer that it is executed from.

.EXAMPLE
Set-JScheduledTaskPassword "My Scheduled Task"
This command will prompt for credentials and configure the specified task using those credentials.

.LINK
http://windowsitpro.com/scripting/updating-scheduled-tasks-credentials
#>

[CmdletBinding(SupportsShouldProcess=$TRUE)]
param(
[parameter(Mandatory=$TRUE,ValueFromPipeline=$TRUE)]
[String[]]
$TaskName,
[String]
$TaskCredential,
[String]
$ComputerName=$ENV:COMPUTERNAME
)

begin
{
$PIPELINEINPUT = (-not $PSBOUNDPARAMETERS.ContainsKey("TaskName")) -and (-not $TaskName)
$TASK_LOGON_PASSWORD = 1
$TASK_LOGON_S4U = 2
$TASK_UPDATE = 4
$MIN_SCHEDULER_VERSION = 0x00010002

# Make sure the COmputer exists and is online.
if (-not(test-connection -computername $ComputerName -count 1 -quiet)) {write-error "Cannot locate system named $ComputerName, exiting..." -errroaction stop}

# Try to create the TaskService object on the local computer; throw an error on failure
try
	{
	$TaskService = new-object -comobject "Schedule.Service"
	}
catch [System.Management.Automation.PSArgumentException]
	{
	throw $_
	}

# Assume $NULL for the schedule service connection parameters unless -ConnectionCredential used
$userName = $domainName = $connectPwd = $NULL
try
	{
# Connects us to the COnputer using the current credentials (pass through creds).
	$TaskService.Connect($ComputerName, $userName, $domainName, $connectPwd)
	}
catch [System.Management.Automation.MethodInvocationException]
	{
	write-error "Error connecting to '$ComputerName' - '$_'" -erroraction stop
	}

# Returns a 32-bit unsigned value as a version number (x.y, where x is the
# most-significant 16 bits and y is the least-significant 16 bits).
function convertto-versionstr([UInt32] $version)
{
$major = [Math]::Truncate($version / [Math]::Pow(2, 0x10)) -band 0xFFFF
$minor = $version -band 0xFFFF
"$($major).$($minor)"
}

if ($TaskService.HighestVersion -lt $MIN_SCHEDULER_VERSION)
	{
	write-error ("Schedule service on '$ComputerName' is version $($TaskService.HighestVersion) " +
	"($(convertto-versionstr($TaskService.HighestVersion))). The Schedule service must " +
	"be version $MIN_SCHEDULER_VERSION ($(convertto-versionstr $MIN_SCHEDULER_VERSION)) " +
	"or higher.") -erroraction stop
	}

# This prevents a scoping problem--if the $TaskCredential variable
# doesn't exist, it won't get created in the correct scope--create
# new variable as a workaround
$NewTaskCredential = $TaskCredential
if (-not $NewTaskCredential)
	{
# Retreive passowrd if it is not specified
	$NewTaskCredential = Read-Host "Please specify password for the scheduled task."
	if (-not $NewTaskCredential)
		{
		write-error "You must specify credentials." -errroaction stop
		}
	}

function set-scheduledtaskcredential2($taskName)
{
$rootFolder = $TaskService.GetFolder("\")
try
	{
# Retrieves all of the task data using the existing connection
	$taskDefinition = $rootFolder.GetTask($taskName).Definition
	}
catch [System.Management.Automation.MethodInvocationException]
	{
	write-error "Scheduled task '$taskName' not found on '$computerName'." -erroraction stop
	}

# Needed to set the creds on the tasks. We use the existing user rather than change it because we are only updating the password.
$TaskUser=$taskDefinition.Principal.UserId

$logonType = $taskDefinition.Principal.LogonType
# No need to set credentials for tasks that don't have stored credentials.
if (-not (($logonType -eq $TASK_LOGON_PASSWORD) -or ($logonType -eq $TASK_LOGON_S4U)))
	{
	write-error "Scheduled task '$taskName' on '$ComputerName' doesn't have stored credentials." -erroraction stop
	}
if (-not $PSCMDLET.ShouldProcess("Task '$taskName' on computer '$ComputerName'", "Set scheduled task credentials")) { return }
try
	{
# This is the cmdlet that sets the password
	[Void] $rootFolder.RegisterTaskDefinition($taskName, $taskDefinition, $TASK_UPDATE, $TaskUser, $NewTaskCredential, $logonType)
	}
catch [System.Management.Automation.MethodInvocationException]
	{
	write-error "Error updating scheduled task '$taskName' on '$computerName' - '$_'"
	}
} # function set-scheduledtaskcredential2($taskName)

}

process
{
if ($PIPELINEINPUT)
	{
	set-scheduledtaskcredential2 $_
	}
	else
	{
	$TaskName | foreach-object {set-scheduledtaskcredential2 $_}
	}
}
} # Function Set-JScheduledTaskPassword


###########################################################################
######################### End of Functions ################################
###########################################################################

write-host("`nThe following cmdlets are loaded:`nAdd-JADGroup`nAdd-JContact`nAdd-JCRDelegate`nAdd-JEWSDelegate`nAdd-JHolidays`nAdd-JMailboxAccess`nAdd-JSerenaTicket`nCompare-JFiles`nConvertFrom-JOctetToGUID`nConvertTo-JHereString`nConvertTo-JZip`nExport-JPST`nExport-PSOToCSV`nGet-JADEntry`nGet-JADIDfromSID`nGet-Janus`nGet-JConstructors`nGet-JDCs`nGet-JDirInfo`nGet-JDSACLs`nGet-JEWSDelegates`nGet-JFileEncoding`nGet-JGroupMembers`nGet-JHotFixbyDate`nGet-JInheritance`nGet-JLogonSessions`nGet-JMailboxCalendarDelegate`nGet-JMessageTracking`nGet-JModule`nGet-JPassword`nGet-JRDPSession`nGet-JScheduledTasks`nGet-JStringMatches`nGet-JWMILogonSessions`nImport-JPST`nImport-JSCOM`nMove-JADEntry`nMove-JMailbox`nMove-JMailboxesInBulk`nNew-JSyslogger`nOut-JExcel`nProtect-JParameter`nRemove-JADGroupMembership`nRemove-JCRDelegate`nRemove-JCRMeetingbyOrganizer`nRemove-JRDPSession`nRepair-JService`nRevoke-JMailboxAccess`nSend-JRemoteCMD`nSend-JTCPRequest`nSet-JADEntry`nSet-JDSACLs`nSet-JGroupMaintenanceMode`nSet-JIISPass`nSet-JInheritance`nSet-JMailboxForwarding`nSet-JMaintenanceMode`nSet-JOOFMessage`nSet-JPassword`nSet-JRandomADPassword`nShow-JActiveSyncUserInfo`nShow-JADLockoutStatus`nShow-JASDevices`nShow-JCRPermissions`nShow-JDiskInfo`nShow-JDistributionListMembers`nShow-JETPMessageStatus`nShow-JFailureAuditEvents`nShow-JIISSettings`nShow-JLyncUserInfo`nShow-JMailboxPermissions`nShow-JMessageRecipients`nShow-JOOFSettings`nShow-JPublicFolderInfo`nShow-JRDPSessions`nShow-JSearchEstimate`nShow-JServices`nShow-JStatus`nShow-KVSTemp`nStart-SOAPRequest`nUpdate-JCalendarTentativeSettings`nUpdate-JDisabledUser`nUpdate-JEWSDelegate`nUpdate-JPrimarySMTPAddress`nUpdate-KVSTempUser`n")

# Test if bpp_exchange or bps_exchange is available
write-host ("Checking Domain...")
$admin=get-jadentry -id "_Exchange" -pso -person -properties mail,distinguishedname -ErrorAction SilentlyContinue

# Test to see if the EWS Mail Module is loaded and working
write-host ("Checking for EWS Module...")
$ewstest=Get-EWSMailMessage -ResultSize 1 -Mailbox $admin.mail -ErrorAction SilentlyContinue
$global:EWSModule=$?
if (-not ($global:EWSModule))
	{
	write-output("WARNING: EWS module is not loaded, come cmdlets will not work correctly...`n")
	write-output $error
	}

$global:JanusPSModule=$true

remove-variable ewstest -ErrorAction SilentlyContinue
remove-variable admin -ErrorAction SilentlyContinue

# SIG # Begin signature block
# MIIX9AYJKoZIhvcNAQcCoIIX5TCCF+ECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU49n9f5ME9279tPYhYAj9cpkX
# THegghWFMIIF/TCCBOWgAwIBAgIKYSlt5QAAAAAAAjANBgkqhkiG9w0BAQUFADCB
# hzELMAkGA1UEBhMCVVMxJTAjBgNVBAoTHEphbnVzIENhcGl0YWwgTWFuYWdlbWVu
# dCBMTEMxGzAZBgNVBAsTElByb2R1Y3Rpb24gTmV0d29yazE0MDIGA1UEAxMrSmFu
# dXMgUHJvZHVjdGlvbiBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eTAeFw0wOTA3
# MTMyMTE1MTBaFw0xOTA3MTMyMTI1MTBaMIGoMQswCQYDVQQGEwJVUzElMCMGA1UE
# ChMcSmFudXMgQ2FwaXRhbCBNYW5hZ2VtZW50IExMQzEbMBkGA1UECxMSUHJvZHVj
# dGlvbiBOZXR3b3JrMRIwEAYDVQQLEwlTZXJ2ZXIgMDExQTA/BgNVBAMTOEphbnVz
# IFByb2R1Y3Rpb24gRW50ZXJwcmlzZSBQb2xpY3kgQ2VydGlmaWNhdGUgQXV0aG9y
# aXR5MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA0YtE/HXJpAGxgBnL
# Me6+DRTE7EAGXTtLpJHfFvCQaDv8E25gtUU97+sAxtuRhFg3NIPCP5fhSZltMZ71
# T0ltniVwpFUyxNW4oyiGejJjaKiaANsz0AWOZBEkZNeH021skKTVNgOJ6aEZpJ8O
# zNlwN4mBdazKnrVazWDQ7nDybZlDbWKDHxtXzbVbmfdTNeiyb2OwHWDZThehdj0m
# 9A+AE9DpZz0YtLaIW9MRs0pdQ3vxsVFBsTwExmgkOLCGzL+wjz6Zqt1k+h2aY93E
# WWBBeZxuOaQmtMH5Qqy7CbMm5gePbSsv/6jmlJAxf/4jOsj7Z2oUYaKN9lfM/CJM
# rCICHQIDAQABo4ICRjCCAkIwDwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4EFgQUdXfU
# lawYust3PC0OwjkUKJeFAgIwCwYDVR0PBAQDAgGGMBAGCSsGAQQBgjcVAQQDAgEA
# MBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMB8GA1UdIwQYMBaAFBXgeYAZ5Axx
# Xn/KWVJYZVMrsGz/MIHXBgNVHR8Egc8wgcwwgcmggcaggcOGUmh0dHA6Ly9wLWNh
# cm9vdC9DZXJ0RW5yb2xsL0phbnVzJTIwUHJvZHVjdGlvbiUyMFJvb3QlMjBDZXJ0
# aWZpY2F0ZSUyMEF1dGhvcml0eS5jcmyGSmZpbGU6Ly9QLUNBUk9PVC9DZXJ0RW5y
# b2xsL0phbnVzIFByb2R1Y3Rpb24gUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHku
# Y3JshiFodHRwOi8vY3JsLm15am9ubGluZS5jb20vcm9vdC5jcmwwgdoGCCsGAQUF
# BwEBBIHNMIHKMGcGCCsGAQUFBzAChltodHRwOi8vcC1jYXJvb3QvQ2VydEVucm9s
# bC9QLUNBUk9PVF9KYW51cyUyMFByb2R1Y3Rpb24lMjBSb290JTIwQ2VydGlmaWNh
# dGUlMjBBdXRob3JpdHkuY3J0MF8GCCsGAQUFBzAChlNmaWxlOi8vUC1DQVJPT1Qv
# Q2VydEVucm9sbC9QLUNBUk9PVF9KYW51cyBQcm9kdWN0aW9uIFJvb3QgQ2VydGlm
# aWNhdGUgQXV0aG9yaXR5LmNydDANBgkqhkiG9w0BAQUFAAOCAQEAY9rY8GMA0BPP
# CrlEPZbw/Ib1sBAUx8TG305X3yIcbaR6gBT75PdlNMYSA7x33haArxe6SRnEF0ne
# TXdKS9oJCj8iJU/alk0npvEhendLLQB4P0QCqZoYZMQ5aBPX0YYCBEe/dOEyvUmD
# 3KzyP7sCq1a0W6W3DaqXbjtfEz4mR8YslkQ8+5qXd5T8kSYDw5TotZvOcUrxi/iB
# iNdhxJiUvJ+SK6jrMCzVnyEUtd9MCUMHt2yT+l/zkHbxdmEVNFSJFsrhVDfDPB8L
# Kupihij4HoJIZr16PuagwD0HqA7yxDGNZvJUpLQaxt/mAKHfdW3CC+qLCsboADQr
# kBMdo2rKnDCCB68wggaXoAMCAQICClui1bwAAQAAFm0wDQYJKoZIhvcNAQEFBQAw
# gakxCzAJBgNVBAYTAlVTMSUwIwYDVQQKExxKYW51cyBDYXBpdGFsIE1hbmFnZW1l
# bnQgTExDMRswGQYDVQQLExJQcm9kdWN0aW9uIE5ldHdvcmsxEjAQBgNVBAsTCVNl
# cnZlciAwMTFCMEAGA1UEAxM5SmFudXMgUHJvZHVjdGlvbiBFbnRlcnByaXNlIElz
# c3VpbmcgQ2VydGlmaWNhdGUgQXV0aG9yaXR5MB4XDTExMDkyODIyMjAwNFoXDTEy
# MDkyNzIyMjAwNFowgYUxEzARBgoJkiaJk/IsZAEZFgNjYXAxFTATBgoJkiaJk/Is
# ZAEZFgVqYW51czEOMAwGA1UECxMFSmFudXMxDjAMBgNVBAsTBVVzZXJzMRAwDgYD
# VQQDEwdKTTI4MDA0MSUwIwYJKoZIhvcNAQkBFhZEYXZlLk1pY2hhZWxAamFudXMu
# Y29tMIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQDNSFfWigk5LrVx7hr8Adl6
# rRzC9vUcqD4l3++fNOMbFF2RIMiY7zPNDCe5QwZU0krzRyEX5IiLerkNDICHM2zo
# ks0BNZskODENtg4EvEjBysjUY2aU/7bCnIx3eUTTksCNB2BFhq38xnnxBbWaXu9r
# d753fWIUqxHKfKMcBR89AQIDAQABo4IEfTCCBHkwCwYDVR0PBAQDAgbAMD4GCSsG
# AQQBgjcVBwQxMC8GJysGAQQBgjcVCIa1k3SCoJlLg7WdHYLQiT+D29tqgSuF46sU
# gqfOXQIBZAIBAjAdBgNVHQ4EFgQUkRi09hlaAsnb3oYADGefGQtymjUwHwYDVR0j
# BBgwFoAUsvXjUZ+k0s9omnq35iFkv/Ldje0wggGqBgNVHR8EggGhMIIBnTCCAZmg
# ggGVoIIBkYaB9GxkYXA6Ly8vQ049SmFudXMlMjBQcm9kdWN0aW9uJTIwRW50ZXJw
# cmlzZSUyMElzc3VpbmclMjBDZXJ0aWZpY2F0ZSUyMEF1dC0wNjc4OSgxKSxDTj1Q
# LUNBSVNTVUUwMSxDTj1DRFAsQ049UHVibGljJTIwS2V5JTIwU2VydmljZXMsQ049
# U2VydmljZXMsQ049Q29uZmlndXJhdGlvbixEQz1qYW51c2FkbWluLERDPW5ldD9j
# ZXJ0aWZpY2F0ZVJldm9jYXRpb25MaXN0P2Jhc2U/b2JqZWN0Q2xhc3M9Y1JMRGlz
# dHJpYnV0aW9uUG9pbnSGcmh0dHA6Ly9wLWNhaXNzdWUwMS5qYW51cy5jYXAvQ2Vy
# dEVucm9sbC9KYW51cyUyMFByb2R1Y3Rpb24lMjBFbnRlcnByaXNlJTIwSXNzdWlu
# ZyUyMENlcnRpZmljYXRlJTIwQXV0aG9yaXR5KDEpLmNybIYkaHR0cDovL2NybC5t
# eWpvbmxpbmUuY29tL2lzc3VlMDEuY3JsMIIBkgYIKwYBBQUHAQEEggGEMIIBgDCB
# 5QYIKwYBBQUHMAKGgdhsZGFwOi8vL0NOPUphbnVzJTIwUHJvZHVjdGlvbiUyMEVu
# dGVycHJpc2UlMjBJc3N1aW5nJTIwQ2VydGlmaWNhdGUlMjBBdXQtMDY3ODksQ049
# QUlBLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENOPVNlcnZpY2VzLENOPUNv
# bmZpZ3VyYXRpb24sREM9amFudXNhZG1pbixEQz1uZXQ/Y0FDZXJ0aWZpY2F0ZT9i
# YXNlP29iamVjdENsYXNzPWNlcnRpZmljYXRpb25BdXRob3JpdHkwgZUGCCsGAQUF
# BzAChoGIaHR0cDovL3AtY2Fpc3N1ZTAxLmphbnVzLmNhcC9DZXJ0RW5yb2xsL1At
# Q0FJU1NVRTAxLmphbnVzLmNhcF9KYW51cyUyMFByb2R1Y3Rpb24lMjBFbnRlcnBy
# aXNlJTIwSXNzdWluZyUyMENlcnRpZmljYXRlJTIwQXV0aG9yaXR5KDEpLmNydDAp
# BgNVHSUEIjAgBgorBgEEAYI3CgUBBggrBgEFBQcDAwYIKwYBBQUHAwIwNQYJKwYB
# BAGCNxUKBCgwJjAMBgorBgEEAYI3CgUBMAoGCCsGAQUFBwMDMAoGCCsGAQUFBwMC
# MEQGA1UdEQQ9MDugIQYKKwYBBAGCNxQCA6ATDBFKTTI4MDA0QGphbnVzLmNhcIEW
# RGF2ZS5NaWNoYWVsQGphbnVzLmNvbTANBgkqhkiG9w0BAQUFAAOCAQEAVnzrgyqW
# O8Ictnp90c2OkxDQle7dUtP89E1HhHiPIMVJWSR+8ZsqF0q9021pYSlD6D5yDkyP
# XzNHULSKsPoDuQP4cB8RzEYBDLszM+QNKeTZuEvtGU1sX/hsi8FW8hr4FtHHq5mQ
# dohLjmRkSBhUjYWo8PWe/vyEKlT4sSANzObPdeGb8O+uAnVwxNimELEsSjVWK5aS
# Rs7UQR0hsDgNCzUV9xJ63TOeyNbQYIppAzfYVg5A0JpTprzt7tMfosRKQOgVf/3C
# oYC/BqzzX1axX8IUrfNhjxCmB7aFIyApIHeYSVUOlSe34ksY0W5KQe0BMDNaBcKM
# HU+2P6wmnh+32jCCB80wgga1oAMCAQICClIv1cEAAAAAAA0wDQYJKoZIhvcNAQEF
# BQAwgagxCzAJBgNVBAYTAlVTMSUwIwYDVQQKExxKYW51cyBDYXBpdGFsIE1hbmFn
# ZW1lbnQgTExDMRswGQYDVQQLExJQcm9kdWN0aW9uIE5ldHdvcmsxEjAQBgNVBAsT
# CVNlcnZlciAwMTFBMD8GA1UEAxM4SmFudXMgUHJvZHVjdGlvbiBFbnRlcnByaXNl
# IFBvbGljeSBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMTEwNzEyMTU1MzQyWhcN
# MTYwNzEwMTU1MzQyWjCBqTELMAkGA1UEBhMCVVMxJTAjBgNVBAoTHEphbnVzIENh
# cGl0YWwgTWFuYWdlbWVudCBMTEMxGzAZBgNVBAsTElByb2R1Y3Rpb24gTmV0d29y
# azESMBAGA1UECxMJU2VydmVyIDAxMUIwQAYDVQQDEzlKYW51cyBQcm9kdWN0aW9u
# IEVudGVycHJpc2UgSXNzdWluZyBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwggEiMA0G
# CSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDmqlohcIdaoZaHO2K+FoBj2HCVdJAJ
# jThY/i7VFUkrs+CezVFLUVf99DU3l2qSkUWkzfFkAF9WA2nPRsuz+Ocb2w8uWXO/
# Gl3koF4NEp1u7hpz4LRECbDGXDu3UWtIAm2BzrlBu4Dp2jhCVVEPR6sKWl3dBSIR
# zT6CxK2aLE9z6BdkKZDEDk3HzjZ7GbLlWaYSOkzkimvO3r+GN4LLOaqjImfTM6iN
# rGy9tdpuyhOQMcYqjhsAH1uGW9VIo5ou4lJx2Of+e+1rqjgHT4LszzTi63qa8Z47
# ZJGuEzoqYMcPNCHWvzqckRlLc2cxgwpveDRWa6fBLncz3lGmCVSm3DXNAgMBAAGj
# ggP0MIID8DAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBSy9eNRn6TSz2iaerfm
# IWS/8t2N7TALBgNVHQ8EBAMCAYYwEgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEE
# AYI3FQIEFgQUSLILVoUKJjO+JQl5clcretA2IPAwGQYJKwYBBAGCNxQCBAweCgBT
# AHUAYgBDAEEwHwYDVR0jBBgwFoAUdXfUlawYust3PC0OwjkUKJeFAgIwggGmBgNV
# HR8EggGdMIIBmTCCAZWgggGRoIIBjYaB8mxkYXA6Ly8vQ049SmFudXMlMjBQcm9k
# dWN0aW9uJTIwRW50ZXJwcmlzZSUyMFBvbGljeSUyMENlcnRpZmljYXRlJTIwQXV0
# aC0wMzQ2MSxDTj1QLUNBUE9MSUNZMDEsQ049Q0RQLENOPVB1YmxpYyUyMEtleSUy
# MFNlcnZpY2VzLENOPVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9amFudXNh
# ZG1pbixEQz1uZXQ/Y2VydGlmaWNhdGVSZXZvY2F0aW9uTGlzdD9iYXNlP29iamVj
# dENsYXNzPWNSTERpc3RyaWJ1dGlvblBvaW50hm9odHRwOi8vcC1jYXBvbGljeTAx
# LmphbnVzLmNhcC9DZXJ0RW5yb2xsL0phbnVzJTIwUHJvZHVjdGlvbiUyMEVudGVy
# cHJpc2UlMjBQb2xpY3klMjBDZXJ0aWZpY2F0ZSUyMEF1dGhvcml0eS5jcmyGJWh0
# dHA6Ly9jcmwubXlqb25saW5lLmNvbS9wb2xpY3kwMS5jcmwwggGQBggrBgEFBQcB
# AQSCAYIwggF+MIHlBggrBgEFBQcwAoaB2GxkYXA6Ly8vQ049SmFudXMlMjBQcm9k
# dWN0aW9uJTIwRW50ZXJwcmlzZSUyMFBvbGljeSUyMENlcnRpZmljYXRlJTIwQXV0
# aC0wMzQ2MSxDTj1BSUEsQ049UHVibGljJTIwS2V5JTIwU2VydmljZXMsQ049U2Vy
# dmljZXMsQ049Q29uZmlndXJhdGlvbixEQz1qYW51c2FkbWluLERDPW5ldD9jQUNl
# cnRpZmljYXRlP2Jhc2U/b2JqZWN0Q2xhc3M9Y2VydGlmaWNhdGlvbkF1dGhvcml0
# eTCBkwYIKwYBBQUHMAKGgYZodHRwOi8vcC1jYXBvbGljeTAxLmphbnVzLmNhcC9D
# ZXJ0RW5yb2xsL1AtQ0FQT0xJQ1kwMS5qYW51cy5jYXBfSmFudXMlMjBQcm9kdWN0
# aW9uJTIwRW50ZXJwcmlzZSUyMFBvbGljeSUyMENlcnRpZmljYXRlJTIwQXV0aG9y
# aXR5LmNydDANBgkqhkiG9w0BAQUFAAOCAQEAMfswWEPaj7SwVJNZie3pvcnPPlKy
# Jgxe/0+8KON0Qk0IZ/lVhPDI4+P4knuNxt4k0oa4ff3rSb2KnMDa91h0ulI6O3mr
# brhOuC4w8v3uH+4BUTriiXvB23NZd9OrulnMCGfFzg4/SwkFbVo4A87hXqzGDJHq
# A+z4Km63S8WQRskI2Yf9kzjDQXNJnUg9NhrD0kUAPfp2xi+AoHtEnzY6Hb5SGhyG
# m1e1EaZpKIoNFHSE7lxTnT6V6Rr2GkzqLkilJx91cW8wXb/2TKsR/hDn+4MQHZYp
# vlW5MxYXXQmn4Aa5mKj77FQgO8i8FYhpW8hvk2T28KcWd5o5WHnQzgsmFDGCAdkw
# ggHVAgEBMIG4MIGpMQswCQYDVQQGEwJVUzElMCMGA1UEChMcSmFudXMgQ2FwaXRh
# bCBNYW5hZ2VtZW50IExMQzEbMBkGA1UECxMSUHJvZHVjdGlvbiBOZXR3b3JrMRIw
# EAYDVQQLEwlTZXJ2ZXIgMDExQjBABgNVBAMTOUphbnVzIFByb2R1Y3Rpb24gRW50
# ZXJwcmlzZSBJc3N1aW5nIENlcnRpZmljYXRlIEF1dGhvcml0eQIKW6LVvAABAAAW
# bTAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG
# 9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIB
# FTAjBgkqhkiG9w0BCQQxFgQUFj14h4bRVNqlGE/GpZdmZcv6jdYwDQYJKoZIhvcN
# AQEBBQAEgYByhNoOpGXgJGyhJmG87cFnLraNTpzJziEFtxhzVlZXuCakH2pwosGZ
# cYm3cuKJUnwQCsCY7NV3YDAsbG7pdTH4iagrAEIpOgzf/LsYOb2uABlARmli4Bbd
# 7slfv053oUN4bzaM1kmni70hgTDIN0mgRX4dhw2s5Xz6AZkshFQIrQ==
# SIG # End signature block

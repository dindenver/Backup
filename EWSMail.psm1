function Get-EWSMailMessage {
<#
	.SYNOPSIS
		This function retrieves messages from an Exchange Mailbox

	.DESCRIPTION
		This function retrieves messages from an Exchange mailbox using the
		EWS Managed API.

	.PARAMETER  Mailbox
		Specifies the email address of the mailbox to search. If no value is provided, 
		the mailbox of the user running the function will be targeted. When specifying an 
		alternate mailbox you'll need to be assigned the ApplicationImpersonation 
		RBAC role.

	.PARAMETER  SearchQuery
		This parameter allows you to specify a query string based on Advanced
		Query Syntax (AQS). You can use an AQS query to specific properties
		of a message using word phrase restriction, date range restriction,
		and message type restriction. See the following article for details:
		http://msdn.microsoft.com/en-us/library/ee693615.aspx
		NOTE: This feature does not work in Exchange 2007

	.PARAMETER  ResultSize
		Specifies the number of messages that should be returned by your search.
		This values is set to 1000 by default.
		
	.PARAMETER  Folder
		Allows you to specify which well knwon mailbox folder should be searched 
		in your command. If you do not specify a value the Inbox folder will be 
		used.
		
	.EXAMPLE
		Get-EWSMailMessage -ResultSize 10
		
		Description
		-----------
		Retrieves the first 10 messages in the callers Inbox.	

	.EXAMPLE
		Get-EWSMailMessage -ResultSize 1 -Mailbox sysadmin@contoso.com
		
		Description
		-----------
		Returns the newest message in the sysadmin Inbox.		
		
	.NOTES
		Reference: http://msdn.microsoft.com/en-us/library/dd633696%28v=EXCHG.80%29.aspx		

#>
	[CmdletBinding()]
    param(
        [Parameter(Position = 0, Mandatory = $false)]
		[String]
        $Mailbox,
        [Parameter(Position = 1, Mandatory = $false)]
		[String]
        $SearchQuery = "*",
        [Parameter(Position = 2, Mandatory = $false)]
		[int]
        $ResultSize = 10000,        
        [Parameter(Position = 3, Mandatory = $false)]
		[string]
        $Folder = "Inbox",
        [Parameter(Position = 4, Mandatory = $false)]
		[switch]
	$d
    )
	
	#Get the email address of the user running the function
	begin {
		$sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
	    $smtpAddress = ([ADSI]"LDAP://<SID=$sid>").properties.Mail

if ($d)
{
	write-host ("SID: $sid")
	write-host ("SMTP: $smtpAddress")
}

	} # Begin Section

    process {	
$ofolder = $null
		#Do we need to use impersonation?
		switch($Mailbox) {
			$null {$Mailbox = $smtpAddress ; $impersonate = $false ; break}
			$smtpAddress {$Mailbox = $smtpAddress ; $impersonate = $false ; break}
			default {$impersonate = $true}
		}
if ($d)
{
	write-host("Impersonation: $impersonate")
}
		
		#Instatiate the EWS service object
        $service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)
#       $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
		
		#Set the impersonated user id on the service object if required
        if($impersonate) {
            $ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress),$Mailbox
            $service.ImpersonatedUserId = $ImpersonatedUserId
		}
if ($d)
{
	write-host("Impersonation ID: $ImpersonatedUserId")
}
        
		#Determine the EWS end-point using Autodiscover
try
	{
        $service.AutodiscoverUrl($Mailbox)
	}


# Manually assigns the url if there is an error
Catch [system.exception]
 {
	$URI='https://p-uccas03.janus.cap/EWS/Exchange.asmx'
	$service.URL = New-Object Uri($URI)
	write-output ("Caught an Autodiscover URL exception, recovering...")
 }

if ($d)
{
try
	{
	[string]$url=$service.AutodiscoverUrl($Mailbox)
	}


# Manually assigns the url if there is an error
Catch [system.exception]
 {
	$URI='https://p-uccas03.janus.cap/EWS/Exchange.asmx'
	$service.URL = New-Object Uri($URI)
	write-output ("Caught an Autodiscover URL exception, recovering...")
 }
	write-host("CAS URL: $url")
}

switch ($folder.tolower())
	{
"Calendar" {break}
"Contacts" {break}
"DeletedItems" {break}
"Drafts" {break}
"Inbox" {break}
"Journal" {break}
"Notes" {break}
"Outbox" {break}
"SentItems" {break}
"Tasks" {break}
"MsgFolderRoot" {break}
"PublicFoldersRoot" {break}
"Root" {break}
"JunkEmail" {break}
"SearchFolders" {break}
"VoiceMail" {break}
"RecoverableItemsRoot" {break}
"RecoverableItemsDeletions" {break}
"RecoverableItemsVersions" {break}
"RecoverableItemsPurges" {break}
default
	{
	$rfRootFolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$smtpAddress)
	$rfRootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$rfRootFolderID)
	$fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(10000);
	$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
	$fvFolderView.PropertySet = $Propset
	$ffResponse = $rfRootFolder.FindFolders($fvFolderView);
	foreach ($ffolder in $ffResponse.Folders)
	{
		if ($ffolder.Displayname -like $folder)
		{
		$ofolder=$ffolder
		}
	}
	}# default
	} # switch
[Microsoft.Exchange.WebServices.Data.FolderId]$folderID=$ofolder.Id
if($folderID -eq $null) {[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]$folderID=[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$folder}
write-host("Folder ID: $folderID")
		#Create a view based on the $ResultSize parameter value
        $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList $ResultSize
		
		#Define which properties we want to retrieve from each message
        $propertyset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
        
        $view.PropertySet = $propertyset
		
		#Use FindItems method for the specified folder, AQS query and number of messages
        $items = $service.FindItems($folderID,$view)
if ($d)
{
	$count=($items | measure-object).count
	write-host("Items: $count")
}

#       $items = $service.FindItems($folderID,$SearchQuery,$view)
		
		#Loop through each message returned by FindItems
        $items | %{
			#The FindItem method does not return the message body so we need to bind to 
			#the message using the Bind method of the EmailMessage class
			$emailProps = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
			$emailProps.RequestedBodyType = "Text"

			if ($folder -eq "Calendar")
				{
				$email = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind($service, $_.Id, $emailProps)
				}
				else
				{
				$email = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service, $_.Id, $emailProps)
				}
if ($d)
{
$email | fl
$email | get-member
}

			#Create a custom object that returns the desired message properties
New-Object PSObject -Property @{
AdjacentMeetingCount = $email.AdjacentMeetingCount
AdjacentMeetings = $email.AdjacentMeetings
AllowedResponseActions = $email.AllowedResponseActions
AllowNewTimeProposal = $email.AllowNewTimeProposal
AppointmentReplyTime = $email.AppointmentReplyTime
AppointmentSequenceNumber = $email.AppointmentSequenceNumber
AppointmentState = $email.AppointmentState
AppointmentType = $email.AppointmentType
Attachments = $email.Attachments
BccRecipients = $email.BccRecipients
Body = $email.Body
Categories = $email.Categories
CcRecipients = $email.CcRecipients
ConferenceType = $email.ConferenceType
ConflictingMeetingCount = $email.ConflictingMeetingCount
ConflictingMeetings = $email.ConflictingMeetings
ConversationId = $email.ConversationId
ConversationIndex = $email.ConversationIndex
ConversationTopic = $email.ConversationTopic
Culture = $email.Culture
DateTimeCreated = $email.DateTimeCreated
DateTimeReceived = $email.DateTimeReceived
DateTimeSent = $email.DateTimeSent
DeletedOccurrences = $email.DeletedOccurrences
DisplayCc = $email.DisplayCc
DisplayTo = $email.DisplayTo
CC = $email.DisplayCc
To = $email.DisplayTo
Duration = $email.Duration
EffectiveRights = $email.EffectiveRights
End = $email.End
EndTimeZone = $email.EndTimeZone
ExtendedProperties = $email.ExtendedProperties
FirstOccurrence = $email.FirstOccurrence
From = $email.Sender.Name
HasAttachments = [bool]$email.HasAttachments
ICalDateTimeStamp = $email.ICalDateTimeStamp
ICalRecurrenceId = $email.ICalRecurrenceId
ICalUid = {if($email.ICalUid -ne $null) {$email.ICalUid.ToString()} else {""}}
Id = $email.Id.ToString()
Importance = $email.Importance
InReplyTo = $email.InReplyTo
InternetMessageHeaders = $email.InternetMessageHeaders
InternetMessageId = $email.InternetMessageId
IsAllDayEvent = $email.IsAllDayEvent
IsAssociated = $email.IsAssociated
IsAttachment = $email.IsAttachment
IsCancelled = $email.IsCancelled
IsDeliveryReceiptRequested = $email.IsDeliveryReceiptRequested
IsDirty = $email.IsDirty
IsDraft = $email.IsDraft
IsFromMe = $email.IsFromMe
IsMeeting = $email.IsMeeting
IsNew = $email.IsNew
IsOnlineMeeting = $email.IsOnlineMeeting
IsRead = $email.IsRead
IsReadReceiptRequested = $email.IsReadReceiptRequested
IsRecurring = $email.IsRecurring
IsReminderSet = $email.IsReminderSet
IsResend = $email.IsResend
IsResponseRequested = $email.IsResponseRequested
IsSubmitted = $email.IsSubmitted
IsUnmodified = $email.IsUnmodified
LastModifiedName = $email.LastModifiedName
LastModifiedTime = $email.LastModifiedTime
LastOccurrence = $email.LastOccurrence.End
LegacyFreeBusyStatus = $email.LegacyFreeBusyStatus
Location = $email.Location
Mailbox = $Mailbox
MeetingRequestWasSent = $email.MeetingRequestWasSent
MeetingWorkspaceUrl = $email.MeetingWorkspaceUrl
MessageClass = $email.ItemClass
MimeContent = $email.MimeContent
ModifiedOccurrences = $email.ModifiedOccurrences
MyResponseType = $email.MyResponseType
NetShowUrl = $email.NetShowUrl
OptionalAttendees = $email.OptionalAttendees
Organizer = $email.Organizer
OriginalStart = $email.OriginalStart
ParentFolderId = $email.ParentFolderId
Received = $email.DateTimeReceived
ReceivedBy = $email.ReceivedBy
ReceivedRepresenting = $email.ReceivedRepresenting
Recurrence = $email.Recurrence
References = $email.References
ReminderDueBy = $email.ReminderDueBy
ReminderMinutesBeforeStart = $email.ReminderMinutesBeforeStart
ReplyTo = $email.ReplyTo
RequiredAttendees = $email.RequiredAttendees
Resources = $email.Resources
Schema = $email.Schema
Sender = $email.Sender
Sensitivity = $email.Sensitivity
Sent = $email.DateTimeSent
Service = $email.Service
Size = [int]$email.Size
Start = $email.Start
StartTimeZone = $email.StartTimeZone
Subject = $email.Subject
TimeZone = $email.TimeZone
ToRecipients = $email.ToRecipients
UniqueBody = $email.UniqueBody
WebClientEditFormQueryString = $email.WebClientEditFormQueryString
WebClientReadFormQueryString = $email.WebClientReadFormQueryString
When = $email.When
} # End of Object Creation
        }
    }
}

function Move-EWSMailMessage {
<#
	.SYNOPSIS
		This function allows you to move an email message.

	.DESCRIPTION
		This function uses the EWS Managed API to move an email message 
		from a source folder to a target folder in the same mailbox.

	.PARAMETER  Id
		Specifies the message id of the message that should be moved.

	.PARAMETER  Mailbox
		Specifies the email address of the mailbox to search. If no value is provided, 
		the mailbox of the user running the function will be targeted. When specifying an 
		alternate mailbox you'll need to be assigned the ApplicationImpersonation 
		RBAC role.

	.PARAMETER  TargetFolder
		Allows you to specify which well knwon mailbox folder should be the destination  
		of the message.
		
		The following values are valid for this parameter:
		
			Calendar
			Contacts
			DeletedItems
			Drafts
			Inbox
			Journal
			Notes
			Outbox
			SentItems
			Tasks
			MsgFolderRoot
			PublicFoldersRoot
			Root
			JunkEmail
			SearchFolders
			VoiceMail
			RecoverableItemsRoot
			RecoverableItemsDeletions
			RecoverableItemsVersions
			RecoverableItemsPurges
			ArchiveRoot
			ArchiveMsgFolderRoot
			ArchiveDeletedItems
			ArchiveRecoverableItemsRoot
			ArchiveRecoverableItemsDeletions
			ArchiveRecoverableItemsVersions
			ArchiveRecoverableItemsPurges		

	.EXAMPLE
		Get-EWSMailMessage -SearchQuery "Subject:'Hello World'" | Move-EWSMailMessage -TargetFolder Drafts
		
		Description
		-----------
		Moves all messages with the subject "Hello World" to the drafts 
		folder.

	.EXAMPLE
		Get-EWSMailMessage -Folder DeletedItems | Move-EWSMailMessage -TargetFolder Inbox -Confirm:$false
		
		Description
		-----------
		Moves the first 1000 messages in the Deleted Items folder to the Inbox 
		without confirmation.
		
	.NOTES
		Reference: http://msdn.microsoft.com/en-us/library/dd633696%28v=EXCHG.80%29.aspx		

#>
	[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "High")]
    param(
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
		[string]
        $Id,
        [Parameter(Position = 1, Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
		[string]
        $Mailbox,        
        [Parameter(Position = 2, Mandatory = $true)]
		[string]
        $TargetFolder       
    )
    
	#Get the email address of the user running the function
	begin {
		$sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
	    $smtpAddress = ([ADSI]"LDAP://<SID=$sid>").properties.Mail
	}
    
    process {	
	
		#Do we need to use impersonation?
		switch($Mailbox) {
			$null {$Mailbox = $smtpAddress ; $impersonate = $false ; break}
			$smtpAddress {$Mailbox = $smtpAddress ; $impersonate = $false ; break}
			default {$impersonate = $true}
		}
		
		#Instatiate the EWS service object
        $service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)
#       $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
		
		#Set the impersonated user id on the service object if required
        if($impersonate) {
            $ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress),$Mailbox
            $service.ImpersonatedUserId = $ImpersonatedUserId
		}
        
		#Determine the EWS end-point using Autodiscover
try
	{
        $service.AutodiscoverUrl($Mailbox)
	}


# Manually assigns the url if there is an error
Catch [system.exception]
 {
	$URI='https://p-uccas03.janus.cap/EWS/Exchange.asmx'
	$service.URL = New-Object Uri($URI)
	write-output ("Caught an Autodiscover URL exception, recovering...")
 }
		
		#Create a view for a single item
        $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList 1
		
		#Create a propertyset specifying only the message id
		$propertyset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
        $view.PropertySet = $propertyset
		
		#Use the Bind method to create an instance of the message based off the message id
        $item = [Microsoft.Exchange.WebServices.Data.Item]::Bind($service, $Id)
        
		#Use the Move method to move the message to the target folder
		#Return the message subject for confirmation and -whatif parameter
		if ($pscmdlet.ShouldProcess($item.Subject)) {
			$item.Move($TargetFolder)
		}
    }
}

function Get-EWSMailMessageHeader {
<#
	.SYNOPSIS
		This function retrieves the message headers for a single email message.

	.DESCRIPTION
		This function retrieves the message headers for a single email message.

	.PARAMETER  Id
		Specifies the message id of the message that should be moved.

	.PARAMETER  Mailbox
		Specifies the email address of the mailbox to search. If no value is provided, 
		the mailbox of the user running the function will be targeted. When specifying an 
		alternate mailbox you'll need to be assigned the ApplicationImpersonation 
		RBAC role.

	.EXAMPLE
		Get-EWSMailMessage -ResultSize 1 | Get-EWSMessageHeader
		
		Description
		-----------
		Retrieves the message headers for the first item in the callers Inbox.	

	.EXAMPLE
		Get-EWSMailMessage -SearchQuery "Subject:'Sales meeting on 4/12'" | Get-EWSMessageHeader
		
		Description
		-----------
		Retrieves the message headers for an item with a specific subject in the callers Inbox.		
		
	.NOTES
		Reference: http://msdn.microsoft.com/en-us/library/dd633696%28v=EXCHG.80%29.aspx		

#>
	[CmdletBinding()]
    param(
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
		[string]
        $Id,
        [Parameter(Position = 1, Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
		[string]
        $Mailbox
    )
    
	#Get the email address of the user running the function
	begin {
		$sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
	    $smtpAddress = ([ADSI]"LDAP://<SID=$sid>").properties.Mail
	}
    
    process {	
	
		#Do we need to use impersonation?
		switch($Mailbox) {
			$null {$Mailbox = $smtpAddress ; $impersonate = $false ; break}
			$smtpAddress {$Mailbox = $smtpAddress ; $impersonate = $false ; break}
			default {$impersonate = $true}
		}
		
		#Instatiate the EWS service object
        $service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)
#       $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
		
		#Set the impersonated user id on the service object if required
        if($impersonate) {
            $ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress),$Mailbox
            $service.ImpersonatedUserId = $ImpersonatedUserId
		}
		
        #Determine the EWS end-point using Autodiscover
try
	{
        $service.AutodiscoverUrl($Mailbox)
	}


# Manually assigns the url if there is an error
Catch [system.exception]
 {
	$URI='https://p-uccas03.janus.cap/EWS/Exchange.asmx'
	$service.URL = New-Object Uri($URI)
	write-output ("Caught an Autodiscover URL exception, recovering...")
 }
		
		#Create a view for a single item
        $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList 1
		
		#Create a propertyset specifying only the message headers
        $propertyset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.ItemSchema]::InternetMessageHeaders)
        $view.PropertySet = $propertyset
		
		#Use the Bind method to create an instance of the message based off the message id
        $item = [Microsoft.Exchange.WebServices.Data.Item]::Bind($service, $Id, $view.PropertySet)
		
		#Return the message headers
        $item.InternetMessageHeaders
    }
}

function Remove-EWSMailMessage {
<#
	.SYNOPSIS
		This function deletes messages from a mailbox.

	.DESCRIPTION
		This function uses the EWS Managed API to delete messages from 
		an Exchange mailbox.

	.PARAMETER  Id
		Specifies the message id of the message that should be moved.

	.PARAMETER  Mailbox
		Specifies the email address of the mailbox to search. If no value is provided, 
		the mailbox of the user running the function will be targeted. When specifying an 
		alternate mailbox you'll need to be assigned the ApplicationImpersonation 
		RBAC role.
		
	.PARAMETER  DeleteMode
		Specifies the delete operation that should be performed. The following 
		values are valid for this parameter:
		
			HardDelete
			SoftDelete
			MoveToDeletedItems

	.EXAMPLE
		Get-EWSMailMessage -SearchQuery "Subject:'Your Mailbox is Full'" | Remove-EWSMailMessage -DeleteMode HardDelete
		
		Description
		-----------
		Removes messages with the specified message subject permanently.		

	.EXAMPLE
		Get-EWSMailMessage -SearchQuery "Subject:'Your Mailbox is Full'" | Remove-EWSMailMessage -DeleteMode SoftDelete
		
		Description
		-----------
		Removes messages with the specified message subject from the mailbox, but the mailbox 
		owner can restore the message from Recoverable Items.
		
	.EXAMPLE
		Get-EWSMailMessage -SearchQuery "Subject:'Your Mailbox is Full'" | Remove-EWSMailMessage -DeleteMode MoveToDeletedItems
		
		Description
		-----------
		Moves messages with the specified message subject to the Deleted Items folder.
		
	.NOTES
		Reference: http://msdn.microsoft.com/en-us/library/dd633696%28v=EXCHG.80%29.aspx		

#>
	[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "High")]
    param(
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
		[string]
        $Id,
        [Parameter(Position = 1, Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
		[string]
        $Mailbox,
        [Parameter(Position = 2, Mandatory = $false)]
		[ValidateSet(
			'HardDelete',
			'SoftDelete',
			'MoveToDeletedItems'
		)]
        $DeleteMode = "MoveToDeletedItems"
    )
    
	#Get the email address of the user running the function
	begin {
		$sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
	    $smtpAddress = ([ADSI]"LDAP://<SID=$sid>").properties.Mail
	}
    
    process {	
	
		#Do we need to use impersonation?
		switch($Mailbox) {
			$null {$Mailbox = $smtpAddress ; $impersonate = $false ; break}
			$smtpAddress {$Mailbox = $smtpAddress ; $impersonate = $false ; break}
			default {$impersonate = $true}
		}
		
		#Instatiate the EWS service object
        $service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)
#       $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
		
		#Set the impersonated user id on the service object if required
        if($impersonate) {
            $ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress),$Mailbox
            $service.ImpersonatedUserId = $ImpersonatedUserId
		}
        
		#Determine the EWS end-point using Autodiscover
try
	{
        $service.AutodiscoverUrl($Mailbox)
	}


# Manually assigns the url if there is an error
Catch [system.exception]
 {
	$URI='https://p-uccas03.janus.cap/EWS/Exchange.asmx'
	$service.URL = New-Object Uri($URI)
	write-output ("Caught an Autodiscover URL exception, recovering...")
 }
		
		#Create a view for a single item
        $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList 1
		
		#Create a propertyset specifying only the message id
        $propertyset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
        $view.PropertySet = $propertyset
		
		#Use the Bind method to create an instance of the message based off the message id
        $item = [Microsoft.Exchange.WebServices.Data.Item]::Bind($service, $Id)
		
		#Use the Move method to move the message to the target folder
		#Return the message subject for confirmation and -whatif parameter		
		if ($pscmdlet.ShouldProcess($item.Subject)) {
        	$item.Delete($DeleteMode)
		}
    }
}

function Send-EWSMailMessage {
<#
	.SYNOPSIS
		The function allows you to send an email from an Exchange mailbox.

	.DESCRIPTION
		This function uses the EWS Managed API to send an email message from 
		an Exchange mail.box.
		
	.PARAMETER  To
		Specifies one or more recipient email address.
		
	.PARAMETER  CcRecipients
		Specifies one or more carbon copy recipient email address.

	.PARAMETER  BccRecipients
		Specifies one or more blind copy recipient email address.
		
	.PARAMETER  From
		Specifies the sender email address. If no value is provided, 
		the message will be sent from the callers mailbox. When specifying an 
		alternate email address you'll need to be assigned the 
		ApplicationImpersonation RBAC role.
		
	.PARAMETER  Subject
		Specifies the subject of the email message.
		
	.PARAMETER  Body
		Specifies the body of the email message.	

	.EXAMPLE
		Send-EWSMailMessage -To sysadmin@contoso.com -Subject 'Hello World' -Body 'This is a test'
		
		Description
		-----------
		Sends an email message to a single recipient.	

	.EXAMPLE
		$subject = 'Hello World'
		$body = 'This is a test'
		Send-EWSMailMessage -To sysadmin@contoso.com,support@contoso.com -Subject $subject -Body $body
		
		Description
		-----------
		Sends an email message to multiple recipients.
		
	.NOTES
		Reference: http://msdn.microsoft.com/en-us/library/dd633696%28v=EXCHG.80%29.aspx		
#>
    [CmdletBinding()]
    param(
        [Parameter(Position=0, Mandatory=$true)]
		[String[]]
        $To,
        [Parameter(Position=1, Mandatory=$false)]
		[String[]]
        $CcRecipients,	
        [Parameter(Position=2, Mandatory=$false)]
		[String[]]
        $BccRecipients,		
        [Parameter(Position=3, Mandatory=$false)]
		[String]
        $From,		
        [Parameter(Position=4, Mandatory=$true)]
        [String]
        $Subject,
        [Parameter(Position=5, Mandatory=$true, ValueFromPipeline = $true)]
        [String]
        $Body
        )
	
	#Get the email address of the user running the function
	begin {
		$sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
	    $smtpAddress = ([ADSI]"LDAP://<SID=$sid>").properties.Mail
	}
    
    process {	
	
		#Do we need to use impersonation?
		switch($From) {
			$null {$From = $smtpAddress ; $impersonate = $false ; break}
			$smtpAddress {$From = $smtpAddress ; $impersonate = $false ; break}
			default {$impersonate = $true}
		}
		
		#Instatiate the EWS service object
        $service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)
#       $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
		
		#Set the impersonated user id on the service object if required
        if($impersonate) {
            $ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress),$From
            $service.ImpersonatedUserId = $ImpersonatedUserId
		}
        
		#Determine the EWS end-point using Autodiscover
try
	{
        $service.AutodiscoverUrl($From)
	}


# Manually assigns the url if there is an error
Catch [system.exception]
 {
	$URI='https://p-uccas03.janus.cap/EWS/Exchange.asmx'
	$service.URL = New-Object Uri($URI)
	write-output ("Caught an Autodiscover URL exception, recovering...")
 }
		
		#Create a new email message object
		$mail = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($service)
		
		#Set the subject and body based on function parameters
		$mail.Subject = $Subject
		$mail.Body = $Body
		
		#Loop through each recipient based on function parameters
		$To | %{ [Void]$mail.ToRecipients.Add($_) }
		if($CcRecipients) {$CcRecipients | %{ [Void]$mail.CcRecipients.Add($_) }} 
		if($BccRecipients) {$BccRecipients | %{ [Void]$mail.BccRecipients.Add($_) }}
		
		#Send the message and save a copy in the sent items folder
		$mail.SendAndSaveCopy()
	}
}

#requires –Version 2 
 
<# 
.SYNOPSIS 
    Gets items from an exchange mailbox. 
 
.DESCRIPTION 
    Returns items form an exchange mailbox folder. 
    As the default settings it means downloading Mail items from Inbox , but Calendar items, Tasks, Notes, Contacts can be retrieved also. 
    Can be used as replacement of the xp_readmail SQL stored procedure. (xp_readmail does not work on 64-bit SQL server) 
    It is independent from outlook. Does not use outlook MAPI but use EWS. 
    Client side requirements: Microsoft .NET Framework 3.5, Microsoft Exchange Web Services (EWS) Managed API 1.1 installed or at least a copy of Microsoft.Exchange.WebServices.dll. 
    Server side requirements: Exchange 2007 SP1 or newer. 
     
.INPUTS 
    You can pipe UserMailAddress as string input to this script. 
 
.OUTPUTS 
    EWS Exchange item object. Example:[Microsoft.Exchange.WebServices.Data.EmailMessage] 
    or [Microsoft.Exchange.WebServices.Data.Folder] in case of -ListFolders specified. 
     
.EXAMPLE 
    $mails = Get-MailboxItem -Verbose 
    $mails | Select-Object -Property:From,DateTimeReceived,Subject 
     
    From                                           DateTimeReceived             Subject 
    ----                                           ----------------             ------- 
    Mail Delivery Subsystem <SMTP:MAILER-DAEMON... 1/17/2011 10:11:55 AM        Delivery Delayed: Test1 
    Mail Delivery Subsystem <SMTP:MAILER-DAEMON... 1/16/2011 10:29:21 AM        Delivery Delayed: Test2 
 
    This is the most simple example using default parameter values. ( Retrieving mail items from Inbox. EWS is installed. ) 
    .\ prefix must be used when the script file is in the current directory, except the current directory is listed in PATH environment variable. 
 
.EXAMPLE 
    Get-MailboxItem -ModuleDllPath:Modules\EWS\Microsoft.Exchange.WebServices.dll -Verbose | Select-Object -Property:From,DateTimeReceived,Subject 
     
    You can specify the path of the Microsoft.Exchange.WebServices.dll if EWS is not installed or non-standard installation. 
 
.EXAMPLE 
    Get-MailboxItem -FolderName:Calendar -ItemLimit:2 | Select Start,End,Subject 
     
    Start                   End                     Subject 
    -----                   ---                     ------- 
    1/24/2011 2:00:00 PM    1/24/2011 3:00:00 PM    Test1 
    1/24/2011 12:00:00 AM   1/25/2011 12:00:00 AM   Test2 
 
    Retrieve some Calendar items.  
 
.EXAMPLE 
    "TestAccount@yourdomain.com" | Get-MailboxItem 
    Get-MailboxItem -UserMailAddress:"TestAccount@yourdomain.com" 
     
    Exchange account can be specified. If no password specified for the account, then windows authentication will be used. 
    If the account is not specified, then it will be autodetected based on currently logged on user. 
 
.EXAMPLE 
    Get-MailboxItem -ListFolder * -Verbose | Format-Table *Name,*Count -AutoSize 
    WellKnownFolderName DisplayName              ChildFolderCount TotalCount UnreadCount 
    ------------------- -----------              ---------------- ---------- ----------- 
    Calendar            Calendar                                0         16 
    Contacts            Contacts                                0          2 
    DeletedItems        Deleted Items                           4          9 2 
    Drafts              Drafts                                  0          8 5 
    Inbox               Inbox                                   5          7 2 
    Journal             Journal                                 0          0 0 
    JunkEmail           Junk E-mail                             0          0 0 
    MsgFolderRoot       Top of Information Store               15          0 0 
    Notes               Notes                                   0          0 0 
    Outbox              Outbox                                  0          0 0 
    PublicFoldersRoot   IPM_SUBTREE                            11          0 0 
    Root                                                       19          3 0 
    SearchFolders       Finder                                 10          0 0 
    SentItems           Sent Items                              0         40 0 
    Tasks               Tasks                                   0          2 2 
     
    Listing all available folders of the mailbox of the current user. 
 
.EXAMPLE     
    Get-MailboxItem -FolderName:Calendar -ItemLimit:1000 | Where-Object { $_.End -le (Get-Date).AddMonths(-1) } | Foreach-Object { 
        Write-Host "Deleting old Calendar item: $($_.Subject)" 
        $_.Delete("SoftDelete") 
    } 
     
    Methods of retrieved items can be used in several case. 
    Example: Deletion of Calendar items older than 1 month. 
 
.EXAMPLE 
    Get-MailboxItem -FolderName Contacts -ItemLimit 1 -Verbose | Format-Table -AutoSize -Property DisplayName, 
      @{n="email";e={$_.EmailAddresses.Item(0)}}, 
      @{n="BusinessAddress";e={$_.PhysicalAddresses.Item(0).CountryOrRegion}}, 
      @{n="BusinessPhoneNr";e={$_.PhoneNumbers.Item(2)}} 
 
    DisplayName  email              BusinessAddress           BusinessPhoneNr 
    -----------  -----              ---------------           --------------- 
    TestContact  TestCo@sample.com  United States of America  +1 234567890 
     
    Retrieving conatact details. 
    MSDN articles can explain why do we need indexing some properties: 
    PhoneNumberKey Enumeration        : http://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.phonenumberkey(v=EXCHG.80).aspx 
    PhysicalAddressKey Enumeration    : http://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.physicaladdresskey(v=EXCHG.80).aspx 
     
 
.LINK 
    http://povvershell.blogspot.com 
 
.LINK 
    http://msdn.microsoft.com/en-us/library/gg248112(v=EXCHG.80).aspx 
 
.NOTES 
    author:        karaszmiklos@gmail.com 
    version:    20110924 
#> 
 
function Get-MailboxItem 
{ 
    [CmdletBinding(DefaultParameterSetName='Item', 
      SupportsShouldProcess=$true, 
      ConfirmImpact='Medium')] 
    param ( 
            # Path of Microsoft.Exchange.WebServices.dll 
            # http://www.microsoft.com/downloads/en/details.aspx?FamilyID=c3342fb3-fbcc-4127-becf-872c746840e1 
            # Default: "$env:SystemDrive\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll" 
            [Alias('DllPath')] 
            [string] $ModuleDllPath = "$env:SystemDrive\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll", 
             
            # Force using specified RequestedServerVersion. Examples: 
            # Exchange2007_SP1, Exchange2010, Exchange2010_SP1, AutoDetect 
            # Default: Exchange2007_SP1 
            [ValidatePattern('^Exchange[0-9]{4}.{0,4}$|^AutoDetect$')] 
            [string] $ServerVersion = "Exchange2007_SP1", 
             
            # youraccount@yourdomain.com | AutoDetect 
            # Default: AutoDetect 
            [parameter(ValueFromPipeline=$true)] 
            [ValidatePattern('^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}$|^AutoDetect$')] 
            [Alias('MailAddress')] 
            [string] $UserMailAddress = "AutoDetect", 
             
            # Windows Authentication will be used if not specified. 
            [string] $Password = $null, 
             
            # Service Timeout in MilliSeconds 
            # Default: 30000 
            [ValidateRange(0,86400000)] 
            [int] $Timeout, 
             
            # Url of Exchange Web Services. Typically: https://[web]mail.yourdomain.com/ews/exchange.asmx 
            # Default: "AutoDetect" 
            [ValidatePattern('^https?://[^/]*/ews/exchange.asmx$|^AutoDetect$')] 
            [Alias('Url')] 
            [string] $EwsUrl = "AutoDetect", 
             
            # The top level folder that you want to items be retrieved from. 
            # http://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.wellknownfoldername(v=exchg.80).aspx 
            # Default: "Inbox" 
            [Parameter(ParameterSetName='Item')] 
            [ValidateSet( 
                "ArchiveDeletedItems", 
                "ArchiveMsgFolderRoot", 
                "ArchiveRecoverableItemsDeletions", 
                "ArchiveRecoverableItemsPurges", 
                "ArchiveRecoverableItemsRoot", 
                "ArchiveRecoverableItemsVersions", 
                "ArchiveRoot", 
                "Calendar", 
                "Contacts", 
                "DeletedItems", 
                "Drafts", 
                "Inbox", 
                "Journal", 
                "JunkEmail", 
                "MsgFolderRoot", 
                "Notes", 
                "Outbox", 
                "PublicFoldersRoot", 
                "RecoverableItemsDeletions", 
                "RecoverableItemsPurges", 
                "RecoverableItemsRoot", 
                "RecoverableItemsVersions", 
                "Root", 
                "SearchFolders", 
                "SentItems", 
                "Tasks", 
                "VoiceMail" 
            )] 
            [Alias('TopLevelFolderName','WellKnownFolderName')] 
            [string] $FolderName = "Inbox", 
             
            # Maximum number of the most recent Exchange Items to be retrieved. 
            # Default: 10 
            [Parameter(ParameterSetName='Item')] 
            [Alias('Limit','MaxItems')] 
            [ValidateRange(1,32768)] 
            [Int] $ItemLimit = 10, 
             
            # Defines the type of body of an item. 
            # Default: Text 
            [Parameter(ParameterSetName='Item')] 
            [ValidateSet("HTML","Text")] 
            [String] $BodyType = "Text", 
             
            # Gets the specified WellKnownFolderNames. Enter the folder names in a comma-separated list. Wildcards are permitted. To get all, enter a value of *. 
            [Parameter(ParameterSetName='Folder')] 
            [String[]]$ListFolder 
    ) 
 
    begin 
    { 
        Write-Verbose "Performing operation: $($MyInvocation.Line)" 
        # Loading Module: Microsoft.Exchange.WebService ( if not loaded yet ) 
        if ( -not (Get-Module -Name:Microsoft.Exchange.WebServices) ) 
        { 
            try 
            { 
                Import-Module -Name:$ModuleDllPath -ErrorAction:Stop 
            } 
            catch 
            { 
                throw "$_`nPlease install EWS: http://www.microsoft.com/download/en/details.aspx?id=13480" 
            } 
        } 
 
        # initializing EWS ExchangeService 
        if ( $ServerVersion -eq "AutoDetect" ) { 
            $ExchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService 
        } else { 
            $ExchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$ServerVersion) 
        } 
        Write-Verbose "RequestedServerVersion: $($ExchangeService.RequestedServerVersion)" 
        Write-Verbose "UserAgent: $($ExchangeService.UserAgent)" 
    } # begin 
 
    process { 
        if ( $Timeout ) { 
            if ( $Timeout -lt $ExchangeService.Timeout ) { 
                Write-Warning "Timeout < builtin default value. ($Timeout<$($ExchangeService.Timeout)) Define grater in case the function is not responding." 
            } 
            $ExchangeService.Timeout = $Timeout 
        } 
         
        if ( ($UserMailAddress -ne "AutoDetect") -and ($Password -ne [string]$null) ) { 
            $ExchangeService.Credentials = New-Object Net.NetworkCredential($($UserMailAddress -split "\@" | Select-Object -First:1), $Password) 
            Write-Verbose "Credentials: $($ExchangeService.Credentials.Credentials.UserName)" 
        } else { 
            if ( $UserMailAddress -eq "AutoDetect" ) { 
                Write-Verbose "Detecting mail address of current user from AD..." 
                try 
                { 
                    $UserMailAddress = Get-ADMailAddress -samAccountName:$env:UserName -ErrorAction:Stop 
                } 
                catch 
                { 
                    throw "mailaddress of current user($($env:UserName)) not found. Error: $_" 
                } 
                Write-Verbose "detected: $UserMailAddress" 
            } 
            $ExchangeService.UseDefaultCredentials = $true 
            Write-Verbose "UseDefaultCredentials: $($ExchangeService.UseDefaultCredentials)" 
        } 
 
        if ( $EwsUrl -eq "AutoDetect" ) { 
            Write-Verbose "AutodiscoverUrl($UserMailAddress) ..." 
            $ExchangeService.AutodiscoverUrl($UserMailAddress) 
        } else { 
            $ExchangeService.Url = new-object Uri($EwsUrl) 
        } 
        Write-Verbose "EwsUrl: $($ExchangeService.Url)" 
    } # process 
 
    end 
    { 
        if ( $MyInvocation.BoundParameters.Debug ) { $DebugPreference = "Continue" } 
        switch ($PsCmdlet.ParameterSetName)  
        {  
            "Item" 
            { 
                try 
                { 
                    $ExchangeFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind( 
                        $ExchangeService, 
                        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$FolderName 
                    ) 
                    Write-Verbose "Retriveing last $ItemLimit items from $FolderName ..." 
                    [array] $ExchangeItems = $ExchangeFolder.FindItems($ItemLimit) 
                    Write-Verbose "received $($ExchangeItems.Count)." 
                    $PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
                    $PropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::$BodyType 
                    if ( $ExchangeItems.Count -gt 0 ) 
                    { 
                        Write-Verbose "Loading items, applying $BodyType BodyType ..." 
                    } 
                    foreach ( $ExchangeItem in $ExchangeItems ) { 
                        try 
                        { 
                            $ExchangeItem.Load($PropertySet) 
                        } 
                        catch 
                        { 
                            Write-Warning "$_ : $($ExchangeItem | Format-List ConversationTopic,*Date* | Out-String)".Trim() 
                        } 
                        # return Item: 
                        Write-Output $ExchangeItem 
                    } 
                } 
                catch 
                { 
                    Write-Error "$_" 
                } 
                break #switch 
            }  
            "Folder" 
            { 
                [array]$AllFolders = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName] | Get-Member -Static -MemberType:Property 
                [array]$MatchFolders = $AllFolders | Select-Object -ExpandProperty:Name | Where-Object { $_ -like $ListFolder } 
                Write-Verbose "Listing $($MatchFolders.Count) matching of $($AllFolders.Count) foldernames ..." 
                $AvailableFolderCount = $null 
                foreach ( $WkFN in $MatchFolders ) { 
                    try 
                    { 
                        $ReturnFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind( 
                            $ExchangeService, 
                            [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$WkFN 
                        ) | Add-Member -MemberType:NoteProperty -Name:WellKnownFolderName -Value:$WkFN -PassThru -Force 
                        if ( $ReturnFolder ) 
                        { 
                            # return valid folders of the current mailbox: 
                            Write-Output $ReturnFolder 
                            $AvailableFolderCount ++ 
                        } 
                    } 
                    catch 
                    { 
                        Write-Debug "$_" 
                    } 
                } 
                Write-Verbose "$AvailableFolderCount is available in the current mailbox." 
                break #switch 
            }  
        } 
    } # end 
} #function 
 
function Get-ADMailAddress 
{ 
    param( 
        [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Medium')] 
         
        # Active Directory User Account (samAccountName) 
        [parameter(ValueFromPipeline=$true)] [string] $samAccountName = $env:UserName 
    ) 
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher 
    $objSearcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry 
    $objSearcher.Filter = "(&(objectCategory=Person)(objectClass=user)(samAccountName=$samAccountName))" 
    $objSearcher.PageSize = 1000 
    $objSearcher.findone().Properties.mail 
} 
         
Set-Alias -Name gmi -Value Get-MailboxItem

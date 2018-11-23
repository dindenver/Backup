<#-----------------------------------------------------------------------------
Russell Tomkins
Microsoft Premier Field Engineer

Name:           CRLFreshChecks.psm1
Description:    A Powershell module that contains a public and an internal
				function to perform CRL freshness checks on HTTP servers.
				
				Will send an email warning if a CRL is expired or expiring in
				near future. Will use the NextPublish extension by default, 
				4,3,2 and 1 day warnings if missing or can use a manual 
				override value.
				
				Can be use individually or in batch mode for bulk checks. 
				
Usage:          Import-Module .\CRLFreshCheck.psm1 -Function Get-CRLFreshness 
Date:           1.0 - 29-04-2016 - RT - Initial Release
-------------------------------------------------------------------------------
Disclaimer
The sample scripts are not supported under any Microsoft standard support 
program or service. 
The sample scripts are provided AS IS without warranty of any kind. Microsoft
further disclaims all implied warranties including, without limitation, any 
implied warranties of merchantability or of fitness for a particular purpose.
The entire risk arising out of the use or performance of the sample scripts and 
documentation remains with you. In no event shall Microsoft, its authors, or 
anyone else involved in the creation, production, or delivery of the scripts be
liable for any damages whatsoever (including, without limitation, damages for 
loss of business profits, business interruption, loss of business information, 
or other pecuniary loss) arising out of the use of or inability to use the 
sample scripts or documentation, even if Microsoft has been advised of the 
possibility of such damages.
-----------------------------------------------------------------------------#>

Function Get-CRLFreshness {
  <#
  .SYNOPSIS
  Checks the Validity/Freshness of a CRL from the provided CRL Distribution Point (CDP) provided.
  .DESCRIPTION
  Downloads and examines the CRL from the provided CDP. Sends e-mail alert if CRL is expired or expiring (past next publish or less than an hour value provided)
  .EXAMPLE
  Get-CRLFreshness -CDP http://crl.domain.com/IssuingCA.crl
  Performs a CRL Freshness check against the CDP URI provided. 
  .EXAMPLE
  Get-CRLFreshness -CDP http://crl.domain.com/IssuingCA.crl -ServerIP 123.123.123.123
  Performs a CRL Freshness check against the CDP URI provided with an override server IP. Direct IP connection, will not use DNS for server lookup. Useful to directly check a web server behind a load balancer.
  .EXAMPLE
  Get-CRLFreshness -CDP http://crl.domain.com/IssuingCA.crl -WarningHours 1
  Performs a CRL Freshness check against the CDP URI provided and an override warning period. Overrides the NextCRLPublish extension if present. 
  .EXAMPLE
  Get-CRLFreshness -CDP http://crl.domain.com/IssuingCA.crl -ServerIP 123.123.123.123 -WarningHours 5
  Performs a CRL Freshness check against the CDP URI provided with an override server IP and override warning period.
  .EXAMPLE
  Get-CRLFreshness -Batch -InputFile C:\Admin\CRLChecks.csv
  Loads an input CSV file and performs a batch CRL freshness check. All options above available in batch mode. Warning, No input error checking performed.
  .EXAMPLE
   ---- Batch CSV Inputfile Example ------
  CASubjectCN,CDP,ServerIP,WarningHours
  Baltimore CyberTrust Root,http://cdp1.public-trust.com/CRL/Omniroot2025.crl,117.18.237.191,120
  
  .PARAMETER ServerIP
  The IP of the Server to be checked. Allows the bypass of load balanced virtual IP's to directly query the Web Server
  .PARAMETER WarningHours
  The number of hours to give a warning if "Next CRL Publish" is not available or you wish to override it.
  .PARAMETER LogFile
  Name of the log file to log to. Defaults to .\CRLFreshness.log.
  .PARAMETER TempFile
  Name of the temporary CRL file that is downloaded.
  .PARAMETER MailFrom
  Sender Email Address for Mail Notifications
  .PARAMETER MailTo
  Receipients Email Address for Mail Notificaitons
  .PARAMETER MailServer
  Mail Server to submit mail to (assumes no relay required)

  #>
  # Configure the Commandlet
  [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True,Position=0,ParameterSetName = "Single")][ValidateNotNullOrEmpty()][String]$CDP,
        [Parameter(Mandatory=$False,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True,Position=1,ParameterSetName = "Single")][ValidatePattern("\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b")][string]$ServerIP = $Null,        
        [parameter(Mandatory=$false,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True,Position=2,ParameterSetName = "Single")][Int]$WarningHours = 0,
        [parameter(Mandatory=$true,ParameterSetName = "Batch")][Switch]$Batch,
        [parameter(Mandatory=$true,ParameterSetName = "Batch")][String]$InputFile,
        [string]$Logfile = ".\CRLFreshCheck.Log",
        [string]$TempFile = ".\TempCRL.crl",
        [string]$MailFrom = "CRL Freshness Check Script <crlfreshcheck@yourdomain.com>",
		[string]$MailTo = "Receiver Address <receipient@youromdain.com>",
		[string]$MailServer = "your-smtp.server.com")

# Begin the Commandlet
Begin {
  	$ReturnObject = @()
	$LogTime = Get-Date -Format "yyyy.MM.dd hh:mm"
	$Now = (Get-Date).ToLocalTime()
	#$Now = Get-Date "5/4/2016 9:00:00 PM"		# Uncomment to manually set the date and time to confirm email and reporting behaviour 

    Add-Content $LogFile "$LogTime`tInfo:   `t----- Begin Processing -----"
    If ($Batch) {
            $CDPPaths = Import-CSV $InputFile
            $Count = $CDPPaths.Count
			Add-Content $LogFile "$LogTime`tInfo:   `tInput Mode: Batch $Count Paths"}
	Else {	Add-Content $LogFile "$LogTime`tInfo:   `tInput Mode: Single"}
    }

# Main Program
  Process {
	   
    # If we are in Batch mode, loop through each CDP entry otherwise just check the passed CRL.
    If ($Batch) {
            $Counter=-1
            ForEach ($Entry in $CDPPaths) {
                    $Counter++
                    Add-Content $LogFile "$LogTime`tInfo:   `t--- CDP $Counter ---"
					$CDP = $Entry.CDP
					Write-Progress -activity "Processing CDP Paths" -status "Processing $CDP`: " -percentComplete (($Counter / $Count)  * 100)
					$ReturnObject += Get-CRLFile $Entry.CDP $Entry.ServerIP $Entry.WarningHours $Counter}
	} #End If
    Else {$ReturnObject = Get-CRLFile $CDP $ServerIP $WarningHours 0}
  	
	} # End Process

 # Tidy Up
 End {
	Add-Content $LogFile "$LogTime`tInfo:   `t--- Finished Processing ---"
    Return $ReturnObject
	} # End End
} # End of Get-CRLFreshness Function

# ---- The Primary Get-CRLFile Function ----

Function Get-CRLFile ($CDP,$ServerIP,$WarningHours,$Counter){

    # Start with some logging    
    Add-Content $LogFile "$LogTime`tInfo:   `t[$Counter] Begin CRL Check of $CDP"
    
    # Build our Custom Status Object to Return
	$Result = "" | Select-Object "CDP","Status","Description","HoursTilExpiry","Issuer","AKI","ServerIP","DownloadOK","ValidFrom","ValidTo","NextCRLPublish","CurrentDate","BaseCRL","HashAlgorithm","CRLNumber"	
	$Result.CDP = $CDP
	 
    # Grab the CDP Host Header from the CDP Variable
	$HostHeader = ([System.Uri]$CDP).Host
    
	# If we recieved a ServerIP Address, Extract the Host header and update the CDP Path. This will override DNS lookup so we can query load balanced web servers. Log how we are performing the server lookup
	If($ServerIP){
            $CDP = $CDP -Replace ($HostHeader,$ServerIP)
            $Result.ServerIP = $ServerIP
            Add-Content $LogFile "$LogTime`tInfo:   `t[$Counter] CRL Download Direct from IP $ServerIP"}
    Else {
           $Result.ServerIP = "Via DNS Lookup"
           Add-Content $LogFile "$LogTime`tInfo:   `t[$Counter] CRL Download via DNS Lookup"
    	   }
   	
    # Attempt to download the file from the server
	Try { Invoke-WebRequest $CDP -Headers @{Host = $HostHeader} -OutFile $TempFile}	
	Catch { 
		Add-Content $LogFile "$LogTime`tError:    `t[$Counter] CRL Download Failed from $CDP. Unable to check freshness"
		$Result.Status = "NoDownload"
		$Result.Description = "NoDownload - Failed to download CRL"
		$Result.DownloadOK = $False
		}
	If ($Result.Status -ne "NoDownload"){

		Add-Content $LogFile "$LogTime`tInfo:   `t[$Counter] CRL Downloaded Complete"
		$Result.DownloadOK = $True

		# Open the CRL as A Byte file and then convert to Base64
		$CRLContents = [System.Convert]::ToBase64String((Get-Content $TempFile -Encoding Byte))
			
		# Create a X509 CRL Object and Intiliaze all of the CRL data into It
		$CRL = New-Object -ComObject "X509Enrollment.CX509CertificateRevocationList"
		$CRL.InitializeDecode($CRLContents,1) 									# 1 = XCN_CRYPT_STRING_BASE64
			
		# Grab the current ValidFrom/ValidTo Extensions
		$ThisUpdate = ($CRL.ThisUpdate).ToLocalTime()
		$NextUpdate = ($CRL.NextUpdate).ToLocalTime()

		# Attempt to grab the Next CRL Publish Date Extension
		Try {
			$NextPublishExtension = ($CRL.X509Extensions | Where-Object {$_.ObjectID.Value -eq '1.3.6.1.4.1.311.21.4'})
			$NextPublishData = [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($NextPublishExtension.RawData("1")))
			$NextPublishData  = $NextPublishData.Remove(0,2) 					# Strip the first 2 bytes (type and length)
			$NextCRLPublish = [DateTime]::ParseExact($NextPublishData,"yyMMddHHmmss\Z",$null)
			$NextCRLPublish = $NextCRLPublish.ToLocalTime()
			}
		Catch { $NextCRLPublish = "Not Available"}
			
		# Attempt to grab the CRL Number Extension
		Try {
			$CRLNumberExtension = ($CRL.X509Extensions | Where-Object {$_.ObjectID.Value -eq '2.5.29.20'})
			$CRLNumberExtensionData = $CRLNumberExtension.RawData("4")
			$CRLNumberExtensionData = $CRLNumberExtensionData -Replace(" ","")
			$CRLNumber = $CRLNumberExtensionData.Remove(0,4) 					# Strip the first 2 bytes (type and length)
			}
		Catch { $CRLNumber = "Not Available"}

		# Attempt to grab the AKI Extension
		Try {
			$AKIExtension = ($CRL.X509Extensions | Where-Object {$_.ObjectID.Value -eq '2.5.29.35'})
			$AKIExtensionData = $AKIExtension.RawData("4")
			$AKIExtensionData = $AKIExtensionData -Replace(" ","")
			$AKIExtensionData = $AKIExtensionData -Replace("`n","")
			$AKI = $AKIExtensionData.Remove(0,8) 					# Strip the first 4 bytes
			}
		Catch {$AKI = "Not Available"}

		# Update our Log
		Add-Content $LogFile "$LogTime`tInfo:   `t[$Counter] Beginning Freshness Checks"
		If ($WarningHours -ne 0) {Add-Content $LogFile "$LogTime`tInfo:   `t[$Counter] Warning Hours Override = $WarningHours"}

		# Compare against Next CRL Publish if it's present or generic 3/2/1 days unless WarningHours Override specified.
		$HoursTilExpiry = [Decimal]::Round((New-TimeSpan $Now $NextUpdate).TotalHours)
			
		# First check if We have simply expired
		If ($Now -gt $NextUpdate) {
			$Status = "Expired"
			$Description = "Expired - CRL Expired by $HoursTilExpiry hours"
			}
		ElseIf ($WarningHours -eq 0) {				# If there is No Override, use NextCRLPublish
			If ($NextCRLPublish -ne "Not Available") {			# If there is No Override and we have a NextCRL value, compare it with a 30 minute skew
				If($Now -gt $NextCRLPublish.AddMinutes(30)) {$Status = "Expiring";$Description = "Expiring - Expires in $HoursTilExpiry hours"}
				Else {$Status = "Fresh";$Description = "Fresh - Expires in $HoursTilExpiry hours"}
			}
			Else{							# If there is No Override and we don't have a NextCRL, compare it against 4,3,2 and 1 Day.
				Add-Content $LogFile "$LogTime`tWarning:   `t[$Counter] NextCRLPublish and WarningHours both missing. Performing 4,3,2 and 1 day comparison"
				If($Now -gt $NextUpdate.AddHours(-24)) {$Status = "Expiring";$Description = "Expiring - Less than 1 day validity"}
					ElseIf($Now -gt $NextUpdate.AddHours(-48)) {$Status = "Expiring";$Description = "Expiring - Less than 2 days validity"}
					ElseIf($Now -gt $NextUpdate.AddHours(-72)) {$Status = "Expiring";$Description = "Expiring - Less than 3 days validity"}
					ElseIf($Now -gt $NextUpdate.AddHours(-96)) {$Status = "Expiring";$Description = "Expiring - Less than 4 days validity"}
					Else {$Status = "Fresh";$Description = "Fresh - Expires in $HoursTilExpiry hours"
				}
			}	
		}
		Else{								# If we have a Override Warning hours, use it
			If($HoursTilExpiry -lt $WarningHours){$Status = "Expiring";$Description = "Expiring - Expires in $HoursTilExpiry hours (override value is $WarningHours hours)"}
			Else {$Status = "Fresh";$Description = "Fresh - Expires in $HoursTilExpiry hours (Overridden to alert at less than $WarningHours hours remaining)"}
		}	
		
		# Populate the Return Object with more details
		$Result.Status = $Status
		$Result.Description = $Description
		$Result.HoursTilExpiry = $HoursTilExpiry
		$Result.ValidFrom = $ThisUpdate
		$Result.ValidTo = $NextUpdate
		$Result.NextCRLPublish = $NextCRLPublish
		$Result.CurrentDate = $Now
		$Result.Issuer = (($CRL.Issuer.Name).Split(",")[0]).Remove(0,3)
		$Result.BaseCRL = $CRL.BaseCRL
		$Result.HashAlgorithm = ($CRL.HashAlgorithm.FriendlyName).ToUpper()
		$Result.CRLNumber = $CRLNumber
		$Result.AKI = $AKI
			
		# Update our Log
		If ($Status -ne "Fresh") {
			Write-Host "Attempting to send alert e-mail to $MailTo"
			$MailSubject = "$Status`: CRL Freshness Check Failed for the CRL at $CDP"
			$MailBody = ($Result | Format-List | out-string)
			Try{Send-mailmessage -from $MailFrom -to $MailTo -subject $MailSubject -body $MailBody -priority High -dno never -smtpServer $MailServer -ErrorAction Stop}
			Catch{Add-Content $LogFile "$LogTime`tError:    `t[$Counter]$Error Failed to to send alert e-mail to $MailTo"}
			Add-Content $LogFile "$LogTime`tError:    `t[$Counter] $Description"
		}
		Else {Add-Content $LogFile "$LogTime`tSuccess:`t[$Counter] $Description"}
	} # End of CRL Download If

	# Close the log entry off
	Add-Content $LogFile "$LogTime`tInfo:   `t[$Counter] Finish CRL Check of $CDP"

	# Return the Results
	Return $Result
}
# --- End of Get-CRL Function ---
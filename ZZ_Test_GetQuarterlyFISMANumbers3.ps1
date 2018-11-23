<#
Replaced $NewUsers = $i | where {$_.whenCreated -gt $QSDate -and $_.whenCreated -lt $QEDate}
With $NewUsers = $1Users | where {$_.whenCreated -gt $QSDate -and $_.whenCreated -lt $QEDate}

To get privleged users that were disabled during the period run the following:
$QSDate = '04/01/2017 00:00:01 AM'
$QEDate = '06/30/2017 11:59:59 PM'
$DisabledUser1 = get-aduser -filter {Enabled -eq $False} -properties msDS-cloudExtensionAttribute20
$DisabledUser2 = $DisabledUser1 | Where-Object {($_.DistinguishedName -notlike "*OU=Messaging*") -And ($_.DistinguishedName -notlike "*OU=DOIAccess*")}
$DisabledUser = $DisabledUser2 | Where {($_.'msDS-cloudExtensionAttribute20') -And ($_.samaccountname -like "*-*a")}
ForEach ($Disabled in $DisabledUser){
$DisabledDate = get-date $Disabled.'msDS-cloudExtensionAttribute20'
If ($DisabledDate -gt $QSDate -and $DisabledDate -lt $QEDate){
write-host $Disabled.samaccountname, $Disabled.'msDS-cloudExtensionAttribute20'
}
}
#>

remove-item C:\Temp\FismaQuarterlyOutput\*  -Recurse -Force
remove-item c:\temp\fismaquarterly1.txt
remove-item c:\temp\fismaquarterly1.csv
remove-item c:\temp\fismaquarterly11.csv
#$QSDate1 = (Get-Date).AddDays(-91).ToString('MM/dd/yyyy')
#$QSDate = $QSDate1 +" 00:00:01 AM"
$QSDate = '04/01/2017 00:00:01 AM'
$QEDate = '06/30/2017 11:59:59 PM'
#$QEDate1 = (Get-Date).AddDays(-1).ToString('MM/dd/yyyy')
#$QEDate = $QEDate1 +" 11:59:59 PM"


$a = get-aduser -filter {Enabled -eq $true} -server iosdendc01.os.doi.net -properties memberOf,description,distinguishedName,SmartcardLogonRequired,altSecurityIdentities,displayname,extensionAttribute5,sn,GivenName,mail,UserPrincipalName,CanonicalName,whenCreated
$b = $a | where { $_.Enabled -eq $True }
$c = $b | where {$_.distinguishedName -notlike "*Messaging*"}
$s = $c | where {($_.distinguishedName -like "*service*") -or ($_.displayname -like "*Service*") -or ($_.description -like "*Service*") -or ($_.samaccountname -like "svc*")}
$d1 = $c | where {($_.DistinguishedName -notlike "*service*")}
$d2 = $d1 | where {($_.samaccountname -notlike "svc*")}
$d = $d2 | where {($_.description -notlike "*Service*")}
$e = $d | where {$_.name -ne "osadmin"}
$f = $e | where {$_.name -notlike "*$"}
$g = $f | where {$_.name -notlike "*TRAIN*"}
$h = $g | where {$_.name -notlike "*Student*"}
$i = $h | where {$_.name -notlike "*Instructor*"}
$1Users = $i | where {$_.extensionAttribute5 -like "*-*"}
$NewUsers = $1Users | where {$_.whenCreated -gt $QSDate -and $_.whenCreated -lt $QEDate}
$2DA = foreach ($z in (get-adgroup 'Domain Admins' -server iosdendc01.os.doi.net -properties member).member){get-aduser -server iosdendc01.os.doi.net $z -properties description,distinguishedName,SmartcardLogonRequired,altSecurityIdentities,displayname,extensionAttribute5 }
$1DA = $2DA | where {$_.Enabled -eq "True"}
$2BA = foreach ($z in (get-adgroup 'IOS_BureauAdmins' -server iosdendc01.os.doi.net -properties member).member){get-aduser -server iosdendc01.os.doi.net $z -properties description,distinguishedName,SmartcardLogonRequired,altSecurityIdentities,displayname,extensionAttribute5 }
$1BA = $2BA | where {$_.Enabled -eq "True"}
$2SA = $i | where {($_.samaccountname -like "*-sa") -or ($_.sn -like "*-sa")}
$1SA = $2SA | where {!($_.memberOf -like "CN=IOS_BureauAdmins*")}
$1WA = $i | where {$_.samaccountname -like "*-wa"}
$1DAP = $1DA | Where { $_.SmartcardLogonRequired -eq $True }
$1SAP = $1SA | Where { $_.SmartcardLogonRequired -eq $True }
$1WAP = $1WA | Where { $_.SmartcardLogonRequired -eq $True }
$1BAP = $1BA | Where { $_.SmartcardLogonRequired -eq $True }
$Users = $1Users | measure
$Users = $Users.count
$DA = $1DA | measure
$DA = $DA.count
$BA = $1BA | measure
$BA = $BA.count
$SA = $1SA | measure
$SA = $SA.count
$WA = $1WA | measure
$WA = $WA.count
$DAP = $1DAP | measure
$DAP = $DAP.count
$BAP = $1BAP | measure
$BAP = $BAP.count
$SAP = $1SAP | measure
$SAP = $SAP.count
$WAP = $1WAP | measure
$WAP = $WAP.count
$ser = $s | measure
$ser = $ser.count
$NU = $NewUsers | measure
$NU = $NU.count
$1DA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\OSDA.csv -Delimiter ";" -nti
$1SA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\OSSA.csv -Delimiter ";" -nti
$1WA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\OSWA.csv -Delimiter ";" -nti
$1BA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\OSBA.csv -Delimiter ";" -nti
$s | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\OSService.csv -Delimiter ";" -nti
$NewUsers | select surname,givenname,mail,whenCreated | export-csv C:\Temp\FismaQuarterlyOutput\OSNewUsers.csv -Delimiter ";" -nti
$header =  "Domain,Users,EnterpriseAdmins,EnterpriseAdmins_PIV_Enforced,DomainAdmins,DomainAdmins_PIV_Enforced,BureauAdmins,BureauAdmins_PIV_Enforced,ServerAdmins,ServerAdmins_PIV_Enforced,WorkstationAdmins,WorkstationAdmins_PIV_Enforced,Service_Accounts,New_UserAccounts"
$header | out-file c:\temp\fismaquarterly1.txt
write-output "OS,$Users,0,0,$DA,$DAP,$BA,$BAP,$SA,$SAP,$WA,$WAP,$SER,$NU" | out-file c:\temp\fismaquarterly1.txt -append

$a = get-aduser -filter {Enabled -eq $true} -server iesdendc01.eis.doi.net -properties description,distinguishedName,SmartcardLogonRequired,altSecurityIdentities,displayname,extensionAttribute5,sn,GivenName,mail,UserPrincipalName,CanonicalName,whenCreated
$b = $a | where { $_.Enabled -eq $True }
$c = $b | where {$_.distinguishedName -notlike "*Messaging*"}
$s = $c | where {($_.distinguishedName -like "*service*") -or ($_.displayname -like "*Service*") -or ($_.description -like "*Service*") -or ($_.samaccountname -like "svc*")}
$d1 = $c | where {($_.DistinguishedName -notlike "*service*")}
$d2 = $d1 | where {($_.samaccountname -notlike "svc*")}
$d = $d2 | where {($_.description -notlike "*Service*")}
$e = $d | where {$_.name -ne "osadmin"}
$f = $e | where {$_.name -notlike "*$"}
$g = $f | where {$_.name -notlike "*TRAIN*"}
$h = $g | where {$_.name -notlike "*Student*"}
$i = $h | where {$_.name -notlike "*Instructor*"}
$1Users = $i | where {$_.extensionAttribute5 -like "*-*"}
$NewUsers = $1Users | where {$_.whenCreated -gt $QSDate -and $_.whenCreated -lt $QEDate}
$2DA = foreach ($z in (get-adgroup 'Domain Admins' -server iesdendc01.eis.doi.net -properties member).member){get-aduser -server iesdendc01.eis.doi.net $z -properties description,distinguishedName,SmartcardLogonRequired,altSecurityIdentities,displayname,extensionAttribute5 }
$1DA = $2DA | where {$_.Enabled -eq "True"}
$1SA = $i | where {$_.samaccountname -like "*-sa"}
$1WA = $i | where {$_.samaccountname -like "*-wa"}
$1DAP = $1DA | Where { $_.SmartcardLogonRequired -eq $True }
$1SAP = $1SA | Where { $_.SmartcardLogonRequired -eq $True }
$1WAP = $1WA | Where { $_.SmartcardLogonRequired -eq $True }
$Users = $1Users | measure
$Users = $Users.count
$DA = $1DA | measure
$DA = $DA.count
$SA = $1SA | measure
$SA = $SA.count
$WA = $1WA | measure
$WA = $WA.count
$DAP = $1DAP | measure
$DAP = $DAP.count
$SAP = $1SAP | measure
$SAP = $SAP.count
$WAP = $1WAP | measure
$WAP = $WAP.count
$ser = $s | measure
$ser = $ser.count
$NU = $NewUsers | measure
$NU = $NU.count
$1DA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\EISDA.csv -Delimiter ";" -nti
$1SA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\EISSA.csv -Delimiter ";" -nti
$1WA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\EISWA.csv -Delimiter ";" -nti
$s | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\EISService.csv -Delimiter ";" -nti
$NewUsers | select surname,givenname,mail,whenCreated | export-csv C:\Temp\FismaQuarterlyOutput\EISNewUsers.csv -Delimiter ";" -nti
write-output "EIS,$Users,0,0,$DA,$DAP,$BA,$BAP,$SA,$SAP,$WA,$WAP,$SER,$NU" | out-file c:\temp\fismaquarterly1.txt -append

$a = get-aduser -filter {Enabled -eq $true} -properties memberOf,description,distinguishedName,SmartcardLogonRequired,altSecurityIdentities,displayname,extensionAttribute5,sn,GivenName,mail,UserPrincipalName,CanonicalName,whenCreated -server ibcdendc01.bc.doi.net
$b = $a | where { $_.Enabled -eq $True }
$c = $b | where {$_.distinguishedName -notlike "*Messaging*"}
$s = $c | where {($_.distinguishedName -like "*service*") -or ($_.displayname -like "*Service*") -or ($_.description -like "*Service*") -or ($_.samaccountname -like "svc*")}
$d1 = $c | where {($_.DistinguishedName -notlike "*service*")}
$d2 = $d1 | where {($_.samaccountname -notlike "svc*")}
$d = $d2 | where {($_.description -notlike "*Service*")}
$e = $d | where {$_.name -ne "bcadmin"}
$f = $e | where {$_.name -notlike "*$"}
$g = $f | where {$_.name -notlike "*TRAIN*"}
$h = $g | where {$_.name -notlike "*Student*"}
$i = $h | where {$_.name -notlike "*Instructor*"}
$1Users = $i | where {$_.extensionAttribute5 -like "*-*"}
$NewUsers = $1Users | where {$_.whenCreated -gt $QSDate -and $_.whenCreated -lt $QEDate}
$2DA = foreach ($z in (get-adgroup 'Domain Admins' -server ibcdendc01.bc.doi.net -properties member).member){get-aduser $z -server ibcdendc01.bc.doi.net -properties SmartcardLogonRequired,altSecurityIdentities }
$1DA = $2DA | where {$_.Enabled -eq "True"}
$2BA = foreach ($z in (get-adgroup 'IBC_BureauAdmins' -server ibcdendc01.bc.doi.net -properties member).member){get-aduser $z -server ibcdendc01.bc.doi.net -properties SmartcardLogonRequired,altSecurityIdentities }
$1BA = $2BA | where {$_.Enabled -eq "True"}
$2SA = $i | where {($_.samaccountname -like "*-sa") -or ($_.sn -like "*-sa")}
$1SA = $2SA | where {!($_.memberOf -like "CN=IBC_BureauAdmins*")}
$1WA = $i | where {$_.samaccountname -like "*-wa"}
$1DAP = $1DA | Where { $_.SmartcardLogonRequired -eq $True }
$1SAP = $1SA | Where { $_.SmartcardLogonRequired -eq $True }
$1WAP = $1WA | Where { $_.SmartcardLogonRequired -eq $True }
$1DAP = $1DA | Where { $_.SmartcardLogonRequired -eq $True }
$1SAP = $1SA | Where { $_.SmartcardLogonRequired -eq $True }
$1WAP = $1WA | Where { $_.SmartcardLogonRequired -eq $True }
$1BAP = $1BA | Where { $_.SmartcardLogonRequired -eq $True }
$Users = $1Users | measure
$Users = $Users.count
$DA = $1DA | measure
$DA = $DA.count
$BA = $1BA | measure
$BA = $BA.count
$SA = $1SA | measure
$SA = $SA.count
$WA = $1WA | measure
$WA = $WA.count
$DAP = $1DAP | measure
$DAP = $DAP.count
$BAP = $1BAP | measure
$BAP = $BAP.count
$SAP = $1SAP | measure
$SAP = $SAP.count
$WAP = $1WAP | measure
$WAP = $WAP.count
$ser = $s | measure
$ser = $ser.count
$NU = $NewUsers | measure
$NU = $NU.count
$1DA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\BCDA.csv -Delimiter ";" -nti
$1SA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\BCSA.csv -Delimiter ";" -nti
$1WA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\BCWA.csv -Delimiter ";" -nti
$1BA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\BCBA.csv -Delimiter ";" -nti
$s | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\BCService.csv -Delimiter ";" -nti
$NewUsers | select surname,givenname,mail,whenCreated | export-csv C:\Temp\FismaQuarterlyOutput\BCNewUsers.csv -Delimiter ";" -nti
write-output "BC,$Users,0,0,$DA,$DAP,$BA,$BAP,$SA,$SAP,$WA,$WAP,$SER,$NU" | out-file c:\temp\fismaquarterly1.txt -append

$a = get-aduser -filter {Enabled -eq $true} -properties description,distinguishedName,SmartcardLogonRequired,altSecurityIdentities,displayname,extensionAttribute5,sn,GivenName,mail,UserPrincipalName,CanonicalName,whenCreated
$b = $a | where { $_.Enabled -eq $True }
$c = $b | where {($_.distinguishedName -notlike "*Messaging*") -And ($_.distinguishedName -notlike "*DOI_OS*") -And ($_.distinguishedName -notlike "*DOI_BC*") -And ($_.distinguishedName -notlike "*DOI_OHA*")}
$s = $c | where {($_.distinguishedName -like "*service*") -or ($_.displayname -like "*Service*") -or ($_.description -like "*Service*") -or ($_.samaccountname -like "svc*")}
$d1 = $c | where {($_.DistinguishedName -notlike "*service*")}
$d2 = $d1 | where {($_.samaccountname -notlike "svc*")}
$d = $d2 | where {($_.description -notlike "*Service*")}
$e = $d | where {$_.name -ne "doiadmin"}
$f = $e | where {$_.name -notlike "*$"}
$g = $f | where {$_.name -notlike "*TRAIN*"}
$h = $g | where {$_.name -notlike "*Student*"}
$i = $h | where {$_.name -notlike "*Instructor*"}
$1Users = $i | where {$_.extensionAttribute5 -like "*-*"}
$NewUsers = $1Users | where {$_.whenCreated -gt $QSDate -and $_.whenCreated -lt $QEDate}
$2DA = foreach ($z in (get-adgroup 'Domain Admins' -properties member).member){get-aduser $z -properties SmartcardLogonRequired,altSecurityIdentities }
$2EA = foreach ($z in (get-adgroup 'Enterprise Admins' -properties member).member){get-aduser $z -properties SmartcardLogonRequired,altSecurityIdentities }
$1DA = $2DA | where {$_.Enabled -eq "True"}
$1EA = $2EA | where {$_.Enabled -eq "True"}
$1BA = $i | where {$_.samaccountname -like "*-ba"}
$1SA = $i | where {$_.samaccountname -like "*-sa"}
$1WA = $i | where {$_.samaccountname -like "*-wa"}
$1DAP = $1DA | Where { $_.SmartcardLogonRequired -eq $True }
$1EAP = $1EA | Where { $_.SmartcardLogonRequired -eq $True }
$1BAP = $1BA | Where { $_.SmartcardLogonRequired -eq $True }
$1SAP = $1SA | Where { $_.SmartcardLogonRequired -eq $True }
$1WAP = $1WA | Where { $_.SmartcardLogonRequired -eq $True }
$Users = $1Users | measure
$Users = $Users.count
$DA = $1DA | measure
$DA = $DA.count
$DAP = $1DAP | measure
$DAP = $DAP.count
$EA = $1EA | measure
$EA = $EA.count
$EAP = $1EAP | measure
$EAP = $EAP.count
$SA = $1SA | measure
$SA = $SA.count
$SAP = $1SAP | measure
$SAP = $SAP.count
$WA = $1WA | measure
$WA = $WA.count
$WAP = $1WAP | measure
$WAP = $WAP.count
$BA = $1BA | measure
$BA = $BA.count
$BAP = $1BAP | measure
$BAP = $BAP.count
$ser = $s | measure
$ser = $ser.count
$NU = $NewUsers | measure
$NU = $NU.count
$1DA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\DDOIDA.csv -Delimiter ";" -nti
$1SA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\DDOISA.csv -Delimiter ";" -nti
$1WA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\DOIWA.csv -Delimiter ";" -nti
$1EA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\DDOIEA.csv -Delimiter ";" -nti
$s | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\DOIService.csv -Delimiter ";" -nti
$NewUsers | select surname,givenname,mail,whenCreated | export-csv C:\Temp\FismaQuarterlyOutput\DOINewUsers.csv -Delimiter ";" -nti
write-output "DOI,$Users,$EA,$EAP,$DA,$DAP,$BA,$BAP,$SA,$SAP,$WA,$WAP,$SER,$NU" | out-file c:\temp\fismaquarterly1.txt -append

$c = $b | where {($_.distinguishedName -notlike "*Messaging*") -And ($_.distinguishedName -Like "*OU=DOI_BC,DC=DOI,DC=NET")}
$s = $c | where {($_.distinguishedName -like "*service*") -or ($_.displayname -like "*Service*") -or ($_.description -like "*Service*") -or ($_.samaccountname -like "svc*")}
$d1 = $c | where {($_.DistinguishedName -notlike "*service*")}
$d2 = $d1 | where {($_.samaccountname -notlike "svc*")}
$d = $d2 | where {($_.description -notlike "*Service*")}
$e = $d | where {$_.name -ne "bcadmin"}
$f = $e | where {$_.name -notlike "*$"}
$g = $f | where {$_.name -notlike "*TRAIN*"}
$h = $g | where {$_.name -notlike "*Student*"}
$i = $h | where {$_.name -notlike "*Instructor*"}
$1Users = $i | where {$_.extensionAttribute5 -like "*-*"}
$NewUsers = $1Users | where {$_.whenCreated -gt $QSDate -and $_.whenCreated -lt $QEDate}
$1BA = $i | where {$_.name -like "*-ba"}
$1SA = $i | where {$_.name -like "*-sa"}
$1WA = $i | where {$_.name -like "*-wa"}
$1BAP = $1BA | Where { $_.SmartcardLogonRequired -eq $True }
$1SAP = $1SA | Where { $_.SmartcardLogonRequired -eq $True }
$1WAP = $1WA | Where { $_.SmartcardLogonRequired -eq $True }
$Users = $1Users | measure
$Users = $Users.count
$SA = $1SA | measure
$SA = $SA.count
$SAP = $1SAP | measure
$SAP = $SAP.count
$WA = $1WA | measure
$WA = $WA.count
$WAP = $1WAP | measure
$WAP = $WAP.count
$BA = $1BA | measure
$BA = $BA.count
$BAP = $1BAP | measure
$BAP = $BAP.count
$ser = $s | measure
$ser = $ser.count
$NU = $NewUsers | measure
$NU = $NU.count
$1ba | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\DOI_BCBA.csv -Delimiter ";" -nti
$1SA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\DOI_BCSA.csv -Delimiter ";" -nti
$1WA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\DOI_BCWA.csv -Delimiter ";" -nti
$s | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\DOI_BCService.csv -Delimiter ";" -nti
$NewUsers | select surname,givenname,mail,whenCreated | export-csv C:\Temp\FismaQuarterlyOutput\DOI_BCNewUsers.csv -Delimiter ";" -nti
write-output "DOI_BC,$Users,0,0,0,0,$BA,$BAP,$SA,$SAP,$WA,$WAP,$SER,$NU" | out-file c:\temp\fismaquarterly1.txt -append

$c = $b | where {($_.distinguishedName -notlike "*Messaging*") -And ($_.distinguishedName -Like "*OU=DOI_OS,DC=DOI,DC=NET")}
$s = $c | where {($_.distinguishedName -like "*service*") -or ($_.displayname -like "*Service*") -or ($_.description -like "*Service*") -or ($_.samaccountname -like "svc*")}
$d1 = $c | where {($_.DistinguishedName -notlike "*service*")}
$d2 = $d1 | where {($_.samaccountname -notlike "svc*")}
$d = $d2 | where {($_.description -notlike "*Service*")}
$e = $d | where {$_.name -ne "osadmin"}
$f = $e | where {$_.name -notlike "*$"}
$g = $f | where {$_.name -notlike "*TRAIN*"}
$h = $g | where {$_.name -notlike "*Student*"}
$i = $h | where {$_.name -notlike "*Instructor*"}
$1Users = $i | where {$_.extensionAttribute5 -like "*-*"}
$NewUsers = $1Users | where {$_.whenCreated -gt $QSDate -and $_.whenCreated -lt $QEDate}
$1BA = $i | where {$_.name -like "*-ba"}
$1SA = $i | where {$_.name -like "*-sa"}
$1WA = $i | where {$_.name -like "*-wa"}
$1BAP = $1BA | Where { $_.SmartcardLogonRequired -eq $True }
$1SAP = $1SA | Where { $_.SmartcardLogonRequired -eq $True }
$1WAP = $1WA | Where { $_.SmartcardLogonRequired -eq $True }
$Users = $1Users | measure
$Users = $Users.count
$SA = $1SA | measure
$SA = $SA.count
$SAP = $1SAP | measure
$SAP = $SAP.count
$WA = $1WA | measure
$WA = $WA.count
$WAP = $1WAP | measure
$WAP = $WAP.count
$BA = $1BA | measure
$BA = $BA.count
$BAP = $1BAP | measure
$BAP = $BAP.count
$ser = $s | measure
$ser = $ser.count
$NU = $NewUsers | measure
$NU = $NU.count
$1ba | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\DOI_OSBA.csv -Delimiter ";" -nti
$1SA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\DOI_OSSA.csv -Delimiter ";" -nti
$1WA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\DOI_OSWA.csv -Delimiter ";" -nti
$s | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\DOI_OSService.csv -Delimiter ";" -nti
$NewUsers | select surname,givenname,mail,whenCreated | export-csv C:\Temp\FismaQuarterlyOutput\DOI_OSNewUsers.csv -Delimiter ";" -nti
write-output "DOI_OS,$Users,0,0,0,0,$BA,$BAP,$SA,$SAP,$WA,$WAP,$SER,$NU" | out-file c:\temp\fismaquarterly1.txt -append

$c = $b | where {($_.distinguishedName -notlike "*Messaging*") -And ($_.distinguishedName -Like "*OU=DOI_OHA,DC=DOI,DC=NET")}
$s = $c | where {($_.distinguishedName -like "*service*") -or ($_.displayname -like "*Service*") -or ($_.description -like "*Service*") -or ($_.samaccountname -like "svc*")}
$d1 = $c | where {($_.DistinguishedName -notlike "*service*")}
$d2 = $d1 | where {($_.samaccountname -notlike "svc*")}
$d = $d2 | where {($_.description -notlike "*Service*")}
$e = $d | where {$_.name -ne "osadmin"}
$f = $e | where {$_.name -notlike "*$"}
$g = $f | where {$_.name -notlike "*TRAIN*"}
$h = $g | where {$_.name -notlike "*Student*"}
$i = $h | where {$_.name -notlike "*Instructor*"}
$1Users = $i | where {$_.extensionAttribute5 -like "*-*"}
$NewUsers = $1Users | where {$_.whenCreated -gt $QSDate -and $_.whenCreated -lt $QEDate}
$1BA = $i | where {$_.name -like "*-ba"}
$1SA = $i | where {$_.name -like "*-sa"}
$1WA = $i | where {$_.name -like "*-wa"}
$1BAP = $1BA | Where { $_.SmartcardLogonRequired -eq $True }
$1SAP = $1SA | Where { $_.SmartcardLogonRequired -eq $True }
$1WAP = $1WA | Where { $_.SmartcardLogonRequired -eq $True }
$Users = $1Users | measure
$Users = $Users.count
$SA = $1SA | measure
$SA = $SA.count
$SAP = $1SAP | measure
$SAP = $SAP.count
$WA = $1WA | measure
$WA = $WA.count
$WAP = $1WAP | measure
$WAP = $WAP.count
$BA = $1BA | measure
$BA = $BA.count
$BAP = $1BAP | measure
$BAP = $BAP.count
$ser = $s | measure
$ser = $ser.count
$NU = $NewUsers | measure
$NU = $NU.count
$1ba | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\DOI_OHABA.csv -Delimiter ";" -nti
$1SA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\DOI_OHASA.csv -Delimiter ";" -nti
$1WA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\DOI_OHAWA.csv -Delimiter ";" -nti
$s | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired,@{name="altSecurityIdentities";expression={$_.altSecurityIdentities -join ";"}}  | export-csv C:\Temp\FismaQuarterlyOutput\DOI_OHAService.csv -Delimiter ";" -nti
$NewUsers | select surname,givenname,mail,whenCreated | export-csv C:\Temp\FismaQuarterlyOutput\DOI_OHANewUsers.csv -Delimiter ";" -nti
write-output "DOI_OHA,$Users,0,0,0,0,$BA,$BAP,$SA,$SAP,$WA,$WAP,$SER,$NU" | out-file c:\temp\fismaquarterly1.txt -append

$date = get-date -f D
$title = 'FISMA Quarterly Numbers'
$Header = $title + ' Generated on ' + $date

copy-item c:\temp\fismaquarterly1.txt c:\temp\fismaquarterly1.csv
$zz = import-csv c:\temp\fismaquarterly1.csv
$zz | sort Domain | export-csv c:\temp\fismaquarterly1.csv -nti
c:\Scripts\Transpose.ps1 -inputfile 'c:\temp\fismaquarterly1.csv' | export-csv 'c:\temp\fismaquarterly11.csv' -nti
copy-item 'c:\temp\fismaquarterly11.csv' 'c:\temp\fismaquarterly1.csv'
$c=import-csv "c:\temp\fismaquarterly1.csv" | convertto-html -as table -title "FISMA Quarterly Numbers" -body "<H2>$Header</H2>" -cssuri "HtmlReport.css" | Out-File "C:\inetpub\wwwroot\EDS\fismaquarterly.htm"

new-item -type Directory -path C:\Temp\FismaQuarterlyOutput\Working
new-item -type Directory -path C:\Temp\FismaQuarterlyOutput\Working2

import-module c:\Scripts\CombineCSV.psm1
cd C:\Temp\FismaQuarterlyOutput\
$a = gci | where {($_.name -like "*.csv") -and ($_.Length -gt "0")}
$b = $a | where {($_.name -notlike "*Service*") -and ($_.name -notlike "*WA*") -and ($_.name -notlike "*NewUsers*")}
$c = $a | where {($_.name -notlike "*Service*") -and ($_.name -notlike "*NewUsers*")}
foreach ($zz in $b.name){
copy-item $zz working
}
foreach ($zz in $c.name){
copy-item $zz working2
}

Combine-CSV -SourceFolder "C:\Temp\FismaQuarterlyOutput\Working\" -Filter "BC*.csv" -ExportFileName "BC_NoWA.csv"
move-item C:\Temp\FismaQuarterlyOutput\Working\BC_NoWA.csv C:\Temp\FismaQuarterlyOutput\BC_NoWA.csv

Combine-CSV -SourceFolder "C:\Temp\FismaQuarterlyOutput\Working\" -Filter "DOI_BC*.csv" -ExportFileName "DOI_BC_NoWA.csv"
move-item C:\Temp\FismaQuarterlyOutput\Working\DOI_BC_NoWA.csv C:\Temp\FismaQuarterlyOutput\DOI_BC_NoWA.csv

Combine-CSV -SourceFolder "C:\Temp\FismaQuarterlyOutput\Working\" -Filter "DOI_OHA*.csv" -ExportFileName "DOI_OHA_NoWA.csv"
move-item C:\Temp\FismaQuarterlyOutput\Working\DOI_OHA_NoWA.csv C:\Temp\FismaQuarterlyOutput\DOI_OHA_NoWA.csv

Combine-CSV -SourceFolder "C:\Temp\FismaQuarterlyOutput\Working\" -Filter "DOI_OS*.csv" -ExportFileName "DOI_OS_NoWA.csv"
move-item C:\Temp\FismaQuarterlyOutput\Working\DOI_OS_NoWA.csv C:\Temp\FismaQuarterlyOutput\DOI_OS_NoWA.csv

Combine-CSV -SourceFolder "C:\Temp\FismaQuarterlyOutput\Working\" -Filter "EIS*.csv" -ExportFileName "EIS_NoWA.csv"
move-item C:\Temp\FismaQuarterlyOutput\Working\EIS_NoWA.csv C:\Temp\FismaQuarterlyOutput\EIS_NoWA.csv

Combine-CSV -SourceFolder "C:\Temp\FismaQuarterlyOutput\Working\" -Filter "OS*.csv" -ExportFileName "OS_NoWA.csv"
move-item C:\Temp\FismaQuarterlyOutput\Working\OS_NoWA.csv C:\Temp\FismaQuarterlyOutput\OS_NoWA.csv

#Combine-CSV -SourceFolder "C:\Temp\FismaQuarterlyOutput\Working\" -Filter "SRV*.csv" -ExportFileName "SRV_NoWA.csv"
#move-item C:\Temp\FismaQuarterlyOutput\Working\SRV_NoWA.csv C:\Temp\FismaQuarterlyOutput\SRV_NoWA.csv

Combine-CSV -SourceFolder "C:\Temp\FismaQuarterlyOutput\Working\" -Filter "DDOI*.csv" -ExportFileName "DOI_NoWA.csv"
move-item C:\Temp\FismaQuarterlyOutput\Working\DOI_NoWA.csv C:\Temp\FismaQuarterlyOutput\DOI_NoWA.csv

Combine-CSV -SourceFolder "C:\Temp\FismaQuarterlyOutput\Working2\" -ExportFileName "ALL_OCIO_Elevated.csv"
move-item C:\Temp\FismaQuarterlyOutput\Working2\ALL_OCIO_Elevated.csv C:\Temp\FismaQuarterlyOutput\ALL_OCIO_Elevated.csv


$All = import-csv C:\Temp\FismaQuarterlyOutput\ALL_OCIO_Elevated.csv -delimiter ";"
$All2 = $All | sort-object -property altSecurityIdentities -Unique
$All2 | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | Sort SmartcardLogonRequired,samaccountname | Export-CSV 'C:\Temp\FismaQuarterlyOutput\FISMA 2.5_2.5.1.csv' -nti
Remove-Item C:\Temp\FismaQuarterlyOutput\ALL_OCIO_Elevated.csv -force

$aa = gci | where {$_.name -like "*.csv"}
$bb = $aa | where {($_.name -notlike "FISMA 2.5_2.5.1.csv") -AND ($_.name -notlike "*NewUsers.csv")}
$cc = $aa | where {$_.name -like "*NewUsers.csv"}
ForEach ($File in $bb.name){
$File1 = $File +"1"
$TMP = import-csv $File -delimiter ";"
$TMP | Select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | Export-CSV $File1 -nti
Start-Sleep -s 3
Move-Item $File1 $File -FORCE
##(Get-Content $File) | % {$_ -replace '"', ""} | out-file -FilePath $File -Force -Encoding ascii
}
ForEach ($File in $cc.name){
$File1 = $File +"1"
$TMP = import-csv $File -delimiter ";"
$TMP | Export-CSV $File1 -nti
Start-Sleep -s 3
Move-Item $File1 $File -FORCE
}
remove-item C:\Temp\FismaQuarterlyOutput\*
remove-item c:\temp\fismaquarterly1.txt
remove-item c:\temp\fismaquarterly1.csv
remove-item c:\temp\fismaquarterly11.csv
$QSDate = '10/01/2016 00:00:01 AM'
$QEDate = '12/31/2016 11:59:59 PM'

$a = get-aduser -filter {Enabled -eq $true} -server iosdendc01.os.doi.net -properties memberOf,description,distinguishedName,SmartcardLogonRequired,displayname,extensionAttribute5,sn,GivenName,mail,UserPrincipalName,CanonicalName,whenCreated
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
$NewUsers = $i | where {$_.whenCreated -gt $QSDate -and $_.whenCreated -lt $QEDate}
$2DA = foreach ($z in (get-adgroup 'Domain Admins' -server iosdendc01.os.doi.net -properties member).member){get-aduser -server iosdendc01.os.doi.net $z -properties description,distinguishedName,SmartcardLogonRequired,displayname,extensionAttribute5 }
$1DA = $2DA | where {$_.Enabled -eq "True"}
$2BA = foreach ($z in (get-adgroup 'IOS_BureauAdmins' -server iosdendc01.os.doi.net -properties member).member){get-aduser -server iosdendc01.os.doi.net $z -properties description,distinguishedName,SmartcardLogonRequired,displayname,extensionAttribute5 }
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
$1DA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\OSDA.csv -Delimiter ";" -nti
$1SA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\OSSA.csv -Delimiter ";" -nti
$1WA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\OSWA.csv -Delimiter ";" -nti
$1BA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\OSBA.csv -Delimiter ";" -nti
$s | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\OSService.csv -Delimiter ";" -nti

$a = get-aduser -filter {Enabled -eq $true} -server iesdendc01.eis.doi.net -properties description,distinguishedName,SmartcardLogonRequired,displayname,extensionAttribute5,sn,GivenName,mail,UserPrincipalName,CanonicalName,whenCreated
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
$NewUsers = $i | where {$_.whenCreated -gt $QSDate -and $_.whenCreated -lt $QEDate}
$2DA = foreach ($z in (get-adgroup 'Domain Admins' -server iesdendc01.eis.doi.net -properties member).member){get-aduser -server iesdendc01.eis.doi.net $z -properties description,distinguishedName,SmartcardLogonRequired,displayname,extensionAttribute5 }
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
$1DA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\EISDA.csv -Delimiter ";" -nti
$1SA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\EISSA.csv -Delimiter ";" -nti
$1WA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\EISWA.csv -Delimiter ";" -nti
$s | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\EISService.csv -Delimiter ";" -nti

$a = get-aduser -filter {Enabled -eq $true} -server isvdendc01.srv.doi.net -properties description,distinguishedName,SmartcardLogonRequired,displayname,extensionAttribute5,sn,GivenName,mail,UserPrincipalName,CanonicalName,whenCreated
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
$NewUsers = $i | where {$_.whenCreated -gt $QSDate -and $_.whenCreated -lt $QEDate}
$2DA = foreach ($z in (get-adgroup 'Domain Admins' -server isvdendc01.srv.doi.net -properties member).member){get-aduser -server isvdendc01.srv.doi.net $z -properties description,distinguishedName,SmartcardLogonRequired,displayname,extensionAttribute5 }
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
$1DA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\SRVDA.csv -Delimiter ";" -nti
$1SA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\SRVSA.csv -Delimiter ";" -nti
$1WA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\SRVWA.csv -Delimiter ";" -nti
$s | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\SRVService.csv -Delimiter ";" -nti

$a = get-aduser -filter {Enabled -eq $true} -properties memberOf,description,distinguishedName,SmartcardLogonRequired,displayname,extensionAttribute5,sn,GivenName,mail,UserPrincipalName,CanonicalName,whenCreated -server ibcdendc01.bc.doi.net
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
$NewUsers = $i | where {$_.whenCreated -gt $QSDate -and $_.whenCreated -lt $QEDate}
$2DA = foreach ($z in (get-adgroup 'Domain Admins' -server ibcdendc01.bc.doi.net -properties member).member){get-aduser $z -server ibcdendc01.bc.doi.net -properties SmartcardLogonRequired }
$1DA = $2DA | where {$_.Enabled -eq "True"}
$2BA = foreach ($z in (get-adgroup 'IBC_BureauAdmins' -server ibcdendc01.bc.doi.net -properties member).member){get-aduser $z -server ibcdendc01.bc.doi.net -properties SmartcardLogonRequired }
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
$1DA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\BCDA.csv -Delimiter ";" -nti
$1SA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\BCSA.csv -Delimiter ";" -nti
$1WA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\BCWA.csv -Delimiter ";" -nti
$1BA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\BCBA.csv -Delimiter ";" -nti
$s | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\BCService.csv -Delimiter ";" -nti

$a = get-aduser -filter {Enabled -eq $true} -properties description,distinguishedName,SmartcardLogonRequired,displayname,extensionAttribute5,sn,GivenName,mail,UserPrincipalName,CanonicalName,whenCreated
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
$NewUsers = $i | where {$_.whenCreated -gt $QSDate -and $_.whenCreated -lt $QEDate}
$2DA = foreach ($z in (get-adgroup 'Domain Admins' -properties member).member){get-aduser $z -properties SmartcardLogonRequired }
$2EA = foreach ($z in (get-adgroup 'Enterprise Admins' -properties member).member){get-aduser $z -properties SmartcardLogonRequired }
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
$1DA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\DDOIDA.csv -Delimiter ";" -nti
$1SA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\DDOISA.csv -Delimiter ";" -nti
$1WA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\DOIWA.csv -Delimiter ";" -nti
$1EA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\DDOIEA.csv -Delimiter ";" -nti
$s | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\DOIService.csv -Delimiter ";" -nti

#$a = get-aduser -filter {Enabled -eq $true} -properties description,distinguishedName,SmartcardLogonRequired,sn,GivenName,mail,UserPrincipalName,CanonicalName,whenCreated -SearchBase "OU=DOI_BC,DC=DOI,DC=NET"
#$b = $a | where { $_.Enabled -eq $True }
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
$NewUsers = $i | where {$_.whenCreated -gt $QSDate -and $_.whenCreated -lt $QEDate}
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
$1ba | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\DOI_BCBA.csv -Delimiter ";" -nti
$1SA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\DOI_BCSA.csv -Delimiter ";" -nti
$1WA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\DOI_BCWA.csv -Delimiter ";" -nti
$s | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\DOI_BCService.csv -Delimiter ";" -nti

#$a = get-aduser -filter {Enabled -eq $true} -properties description,distinguishedName,SmartcardLogonRequired,sn,GivenName,mail,UserPrincipalName,CanonicalName,whenCreated -SearchBase "OU=DOI_OS,DC=DOI,DC=NET"
#$b = $a | where { $_.Enabled -eq $True }
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
$NewUsers = $i | where {$_.whenCreated -gt $QSDate -and $_.whenCreated -lt $QEDate}
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
$1ba | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\DOI_OSBA.csv -Delimiter ";" -nti
$1SA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\DOI_OSSA.csv -Delimiter ";" -nti
$1WA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\DOI_OSWA.csv -Delimiter ";" -nti
$s | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\DOI_OSService.csv -Delimiter ";" -nti

#$a = get-aduser -filter {Enabled -eq $true} -properties description,distinguishedName,SmartcardLogonRequired,sn,GivenName,mail,UserPrincipalName,CanonicalName,whenCreated -SearchBase "OU=DOI_OHA,DC=DOI,DC=NET"
#$b = $a | where { $_.Enabled -eq $True }
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
$NewUsers = $i | where {$_.whenCreated -gt $QSDate -and $_.whenCreated -lt $QEDate}
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
$1ba | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\DOI_OHABA.csv -Delimiter ";" -nti
$1SA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\DOI_OHASA.csv -Delimiter ";" -nti
$1WA | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\DOI_OHAWA.csv -Delimiter ";" -nti
$s | select surname,givenname,displayname,samaccountname,Enabled,Description,distinguishedName,SmartcardLogonRequired | export-csv C:\Temp\FismaQuarterlyOutput\DOI_OHAService.csv -Delimiter ";" -nti

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

import-module c:\Scripts\CombineCSV.psm1
cd C:\Temp\FismaQuarterlyOutput\
$a = gci | where {$_.name -like "*.csv"}
$b = $a | where {($_.name -notlike "*Service*") -and ($_.name -notlike "*WA*")}
foreach ($zz in $b.name){
move-item $zz working
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

Combine-CSV -SourceFolder "C:\Temp\FismaQuarterlyOutput\Working\" -Filter "SRV*.csv" -ExportFileName "SRV_NoWA.csv"
move-item C:\Temp\FismaQuarterlyOutput\Working\SRV_NoWA.csv C:\Temp\FismaQuarterlyOutput\SRV_NoWA.csv

Combine-CSV -SourceFolder "C:\Temp\FismaQuarterlyOutput\Working\" -Filter "DDOI*.csv" -ExportFileName "DOI_NoWA.csv"
move-item C:\Temp\FismaQuarterlyOutput\Working\DOI_NoWA.csv C:\Temp\FismaQuarterlyOutput\DOI_NoWA.csv

<#
$csvs = Get-ChildItem C:\Temp\FismaQuarterlyOutput\* -Include *.csv
# $y=$csvs.Count
# Write-Host "Detected the following CSV files: ($y)"
#foreach ($csv in $csvs)
# {
# Write-Host " "$csv.Name
# }
# $outputfilename = read-host "Please enter the output file name: "
$outputfilename = 'C:\Temp\FismaQuarterlyOutput\1234.xslx'
#Write-Host Creating: $outputfilename
 $excelapp = new-object -comobject Excel.Application
 $excelapp.sheetsInNewWorkbook = $csvs.Count
 $xlsx = $excelapp.Workbooks.Add()
 $sheet=1

foreach ($csv in $csvs)
 {
 $row=1
 $column=1
 $worksheet = $xlsx.Worksheets.Item($sheet)
 $worksheet.Name = $csv.Name
 $file = (Get-Content $csv)
 foreach($line in $file)
 {
 $linecontents=$line -split ';(?!\s*\w+")'
foreach($cell in $linecontents)
 {
 $cell = $cell.TrimStart('"')
 $cell = $cell.TrimEnd('"')
 $worksheet.Cells.Item($row,$column) = $cell
 $column++
 }
 $column=1
 $row++
 }
 $sheet++
 }

$xlsx.SaveAs($outputfilename)
 $excelapp.quit()
#>

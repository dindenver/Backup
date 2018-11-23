################ALL IN ONE?? Need to test.
$Today = get-date -f d
$msg = "Disabled on $Today Per Script"
$d2 = get-date -f MMddyyyy
$log = 'c:\scripttemp\SCRIPT_Disabled'+ $d2 + '.txt'
$HLOG = 'c:\scripttemp\StaleUsers45'+ $d2 + '.txt'
$Whatif = 'c:\scripttemp\SCRIPT_DisabledOutput.txt'
$21 = (Get-Date).AddDays(-74)
$45 = (Get-Date).AddDays(-74)
$AllUsers3 = Get-ADUser -ResultPageSize 900 -filter * -properties info,description,businessCategory,SmartcardLogonRequired,description,altSecurityIdentities,Name,DistinguishedName,Enabled,LastLogonDate,whenCreated,whenChanged,PasswordLastSet,SAMAccountName,LockedOut,PasswordNeverExpires,PasswordNotRequired
$AllUsers2 = $AllUsers3 | Where-Object { $_.userPrincipalName }
$AllUsers1 = $AllUsers2 | Where-Object {$_.DistinguishedName -notlike "*OU=Messaging*"}
$AllUsers = $AllUsers1 | Where-Object {$_.DistinguishedName -like "*OU=DOI_*"}
$EnabledUsers = $AllUsers | Where-Object { $_.Enabled -eq $True }
$Enabled1 = $EnabledUsers | Where-Object { $_.whenCreated -lt $21 }
$StaleUsers45 = $Enabled1 | Where-Object { ($_.LastLogonDate -le $45) -and ($_.name -notlike "*svc*") -and ($_.description -notlike "*Service*") }
$DOIATerm = $Enabled | Where-Object { $_.altSecurityIdentities -like "*Terminated*"}
$StaleUsers45 | out-file $HLOG
Write-Output "name,lastlogon,created,buisnesscategory" | out-file $log -append
foreach ($StaleUser in $StaleUsers45){
Write-Output "Set-ADuser $StaleUser -add @{info = $StaleUser.distinguishedName}" | out-file $Whatif
Write-Output "Set-ADuser $StaleUser -description $msg" | out-file $Whatif -append
Write-Output "Set-ADUser $StaleUser -Enabled $false | out-file $Whatif -append

if ($StaleUser.businessCategory -like "*") {
Write-Output "Set-ADuser $StaleUser -replace @{businessCategory="21"}; | out-file $Whatif -append
$CVV = $staleuser.SAMAccountName +','+ $staleuser.LastLogonDate +','+ $staleuser.whenCreated +','+ $staleuser.businessCategory
Write-Output $CVV | out-file $log -append
if ($StaleUser.distinguishedName -like "*DOI_BC*"){
Write-Output "Move-ADObject $StaleUser -TargetPath "OU=30 Day Retired Accounts,OU=Retired Accounts All BC,OU=DOI_BC,DC=doi,DC=net" | out-file $Whatif -append
}
elseif ($StaleUser.distinguishedName -like "*DOI_OHA*"){
Write-Output "Move-ADObject $StaleUser -TargetPath "OU=30 Day Retired Accounts,OU=Retired Accounts All OHA,OU=DOI_OHA,DC=doi,DC=net" | out-file $Whatif -append
}
else {
Write-Output "Move-ADObject $StaleUser -TargetPath "OU=30 Day Retired Accounts,OU=Retired Accounts All OS,OU=DOI_OS,DC=doi,DC=net" | out-file $Whatif -append
}
}

else {
$CVV = $staleuser.SAMAccountName +','+ $staleuser.LastLogonDate +','+ $staleuser.whenCreated +','+ $staleuser.businessCategory
Write-Output $CVV | out-file $log -append
if ($StaleUser.distinguishedName -like "*DOI_BC*"){
Write-Output "Move-ADObject $StaleUser -TargetPath "OU=30 Day Retired Accounts,OU=Retired Accounts All BC,OU=DOI_BC,DC=doi,DC=net" | out-file $Whatif -append
}
elseif ($StaleUser.distinguishedName -like "*DOI_OHA*"){
Write-Output "Move-ADObject $StaleUser -TargetPath "OU=30 Day Retired Accounts,OU=Retired Accounts All OHA,OU=DOI_OHA,DC=doi,DC=net" | out-file $Whatif -append
}
else {
Write-Output "Move-ADObject $StaleUser -TargetPath "OU=30 Day Retired Accounts,OU=Retired Accounts All OS,OU=DOI_OS,DC=doi,DC=net" | out-file $Whatif -append
}
}
}

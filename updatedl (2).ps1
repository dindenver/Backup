$dllist = "d:\updatedl.csv"
$dlDetails=Import-CSV $dllist

foreach($UD in $dlDetails) {
	$dl = $UD.oldName
	$DLNew = $UD.NewName
        $DLAlias = $UD.alias
	$DLDesc = $UD.Desc
	$DLemail = $UD.newEmail
        $DLN=[ADSI]::Exists("LDAP://cn=$dl,ou=distribution groups,ou=janus,dc=janus,dc=cap")
        if($DLN -ne $FALSE){
           $DLObj = get-DistributionGroup $dl
           $DLobj.EmailAddresses += $DLemail
           set-DistributionGroup -identity $dl -DisplayName $DLNew -Name $DLNew -Alias $DLAlias -EmailAddresses $DLobj.EmailAddresses
           set-DistributionGroup -identity $DLNew -PrimarySMTPAddress $DLemail
           Set-Group -identity $DLNew -Universal

        }
        else{
           write-host $dl "object does not exist" -foregroundcolor red -backgroundcolor yellow
        }
}

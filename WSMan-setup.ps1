# Server-side
new-itemproperty -path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System -name LocalAccountTokenFilterPolicy -propertyType DWord -value 1
Enable-PSRemoting -Force
[string]$IP=$((get-WmiObject Win32_NetworkAdapterConfiguration | Where {($_.IPAddress -ne $null) -and($_.DefaultIPGateway -eq $null)} | get-random).IPAddress)
netsh advfirewall firewall add rule name="Windows Remote Management (HTTPS-In)" dir=in action=allow program="system" enable=yes remoteip=10.0.0.0/8 localip=$IP profile=public,private protocol=TCP localport=5986
set-item wsman:localhost\client\trustedhosts -value * -Force
winrm set winrm/config/listener?Address=*+Transport=HTTP `@`{Enabled=`"false`"`}
[string]$thumbprint=(Get-ExchangeCertificate | where {$_.subject -like "*$(gc env:computername)*"}).thumbprint
winrm create winrm/config/Listener?Address=*+Transport=HTTPS `@`{Hostname=`"$([string](gc env:computername))`"`;CertificateThumbprint=`"$thumbprint`"`}
winrm enumerate winrm/config/listener

# Client-side
new-itemproperty -path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System -name LocalAccountTokenFilterPolicy -propertyType DWord -value 1
Enable-PSRemoting -Force
netsh advfirewall firewall add rule name="Windows Remote Management (HTTPS-In)" dir=in action=allow program="system" enable=yes remoteip=10.0.0.0/8 profile=public,private protocol=TCP localport=5986
set-item wsman:localhost\client\trustedhosts -value *
[string]$pass=Find-PasswordInKeePassDB -PathToDB 'C:\Windows\system32\WindowsPowerShell\v1.0\Modules\KeePass\Dave''s.kdbx' -entryToFind gsmtptest1 -PasswordToDB '1_f0rget'
$SString=ConvertTo-SecureString -String $pass -AsPlainText -Force
$credential=New-Object System.Management.Automation.PSCredential("dmichael",$sstring)
$credential=Get-Credential
$PSSO=New-PSSessionOption -SkipCACheck -SkipCNCheck -OpenTimeout 60000 -OperationTimeout 0
Enter-PSSession -ComputerName 10.10.82.31 -Credential $credential -UseSSL -SessionOption $PSSO

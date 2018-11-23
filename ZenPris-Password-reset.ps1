param ([string]$pass)

Set-Pass bpp_zenprise -pass $pass

write-host ("Update the password in the ZenPrise App!")
read-host ("Press Enter to continue...")

control-service -sys p-uczenm01 -svc Zenprise -pass $pass -start
set-iispass -sys p-uczenm01 -svcact bpp_zenprise -pass $pass

write-host ("Reboot the server and test ZenPrise...")

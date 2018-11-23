Set-Pass bpp_zenprise -pass M3ssagingG0ds!

write-host ("Update the password in the ZenPrise App!")
read-host ("Press Enter to continue...")

control-service -sys p-uczenm01 -svc Zenprise -pass M3ssagingG0ds! -start
set-iispass -sys p-uczenm01 -svcact bpp_zenprise -pass M3ssagingG0ds!

Set-Pass bpp_zenprise -pass M3ssagingG0ds!

write-host ("Reboot the server and test ZenPrise...")

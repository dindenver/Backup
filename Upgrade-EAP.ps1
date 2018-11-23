# **********************************************************************************
#
# Script Name: Script_Name_Here.ps1
# Version: 1.0
# Author: You_name_here
# Date Created: Enter Date here
# _______________________________________
#
# MODIFICATIONS:
# Date Modified: N/A
# Modified By: N/A
# Reason for modification: N/A
# What was modified: N/A
# Description: The description of what the script does goes here.
#
# Usage:
# ./Script_Name_Here.ps1 -paramater "Parameters go here"
#
# **********************************************************************************

# Functions and Filters


# Main Script
Set-EmailAddressPolicy -id  "Perkins mailboxes" -UseRusServer ucmailv06 -ForceUpgrade -RecipientFilter {samaccountname -like "*PE0*"} -confirm:$false
Set-EmailAddressPolicy -id  "Intech Mailboxes" -UseRusServer ucmailv06 -ForceUpgrade -RecipientFilter {samaccountname -like "*IN0*"} -confirm:$false
Set-EmailAddressPolicy -id  "Default Policy" -IncludedRecipients AllRecipients -UseRusServer ucmailv06 -ForceUpgrade -confirm:$false
write-output ('Waiting...')
start-sleep -s 900
./User-Setup -first EAP -last zzTest -ID ZZTest-EAP -pass 'l2E4$67B' -location Janus
./User-Setup -first IN0EAP -last zzTest -ID IN0Test-EAP -pass 'l2E4$67B' -location Intech
./User-Setup -first PE0EAP -last zzTest -ID PE0Test-EAP -pass 'l2E4$67B' -location Perkins
write-output ('Waiting...')
start-sleep -s 900
(get-jadentry -id ZZTest-EAP -pso -properties proxyaddresses).proxyaddresses | fl
(get-jadentry -id IN0Test-EAP -pso -properties proxyaddresses).proxyaddresses | fl
(get-jadentry -id PE0Test-EAP -pso -properties proxyaddresses).proxyaddresses | fl
write-output ('Waiting...')
start-sleep -s 900
./AD-Decom.ps1 -ID 'ZZTest-EAP'
./AD-Decom.ps1 -ID 'IN0Test-EAP'
./AD-Decom.ps1 -ID 'PE0Test-EAP'

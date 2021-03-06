#
# Windows PowerShell in Action Second Edition
#
# Chapter 19 WMI and WSMan
#
# THe file contains examples showing how to use WSMan from PowerShell.
# WMI examples are included to compare them with the WSMan examples.
#

# Get BIOS information using WMI
$properties = "Caption", "OSArchitecture","Version",
  "WindowsDirectory", "SystemDirectory", "SystemDrive"
  
Get-WmiObject Win32_OperatingSystem |
      Format-List $properties
      
Get-WmiObject -Namespace root\cimv2 `
  -Class Win32_OperatingSystem |
    Format-List $properties
    
Get-WSManInstance `
  -ResourceURI wmicimv2/Win32_OperatingSystem |
    Format-List $properties

Get-WmiObject -Namespace root\cimv2 `
  -Class Win32_OperatingSystem |
    Get-Member -MemberType method

Get-WSManInstance `
  -ResourceURI wmicimv2/Win32_OperatingSystem |
    Get-Member -MemberType method

Get-WSManInstance -ResourceURI wmicimv2/Win32_Process
Get-WSManInstance -Enumerate `
  -ResourceURI wmicimv2/Win32_Process |
    select -First 5  |
      Format-Table -AutoSize Name, Handle, ParentProcessId

Get-WSManInstance -Enumerate -ResourceURI wmicimv2/* -Filter @"
  select Name,Handle, ParentProcessId from win32_process
    where name = 'powershell.exe'
"@  | Format-Table -AutoSize Name, Handle, ParentProcessId

Get-WmiObject -Query @"
  select Name,Handle, ParentProcessId from win32_process
    where name = 'powershell.exe'
"@ | Format-Table -AutoSize Name, Handle, ParentProcessId

calc
Get-WSManInstance -Enumerate -ResourceURI wmicimv2/* -Filter @"
  select Name,Handle from win32_process
    where name = 'calc.exe'
"@  | Format-Table -AutoSize Name, Handle

Get-WSManInstance -ResourceURI wmicimv2/win32_process `
  -SelectorSet @{ Handle = 7208 } |
    Format-Table -AutoSize Name, Handle

$target = @{
 ResourceURI = "wmicimv2/Win32_Process";
 SelectorSet = @{ Handle = "7208" }
}

Get-WSManInstance @target | Format-Table -AutoSize Name,Handle
$result = Invoke-WSManAction @target -Action Terminate
$result.ReturnValue
Format-XmlDocument -String $result.OuterXml

Invoke-WSManAction -ResourceURI wmicimv2/win32_process `
  -Action Create -ValueSet @{
    CommandLine = 'calc'
    CurrentDirectory = 'c:\'
  }

Get-WSManInstance -ResourceURI wmicimv2/win32_process `
  -SelectorSet @{ Handle = 8524 } |
    Format-Table Name,Handle

Invoke-WSManAction wmicimv2/win32_process Terminate `
  -SelectorSet @{Handle = 8524 }

Get-WSManInstance -ResourceURI wmicimv2/win32_process `
  -SelectorSet @{ Handle = 8524 } |
    Format-Table Name,Handle
    
Format-XmlDocument -string ($error[0].Exception.Message)
    
##############################################
       
Get-WSManInstance wmicimv2/win32_service `
  -selectorset @{name="winrm"} -computername brucepayx61

Get-WSManInstance wmicimv2/win32_service `
  -selectorset @{name="spooler"} -fragment status -computername 

get-wsmaninstance -enumerate wmicimv2/win32_process

get-wsmaninstance -enumerate wmicimv2/win32_service -returntype epr

Get-WSManInstance -Enumerate wmicimv2/* `
  -filter "select * from * where StartMode = 'Auto' and State = 'Stopped'"

get-wsmaninstance winrm/config/listener `
  -selectorset @{Address="*";Transport="http"} `
  -computername brucepayx61

Get-WSManInstance -Enumerate `
   -Dialect association `
   -filter "{Object=win32_service?name=winrm}" `
   -res wmicimv2/*

##########################
Get-WSManInstance `
  -ResourceURI wmicimv2/win32_process `
  -Enumerate | ft name, handle
  
Get-WSManInstance `
  -ResourceURI wmicimv2/win32_process `
  -Enumerate `
  -Fragment Name

# Get a specific instance identified by the key property
Get-WSManInstance `
  -ResourceURI wmicimv2/win32_process `
  -SelectorSet @{handle = "1800"}

# Return a single property or piece of information from the object
Get-WSManInstance `
  -ResourceURI wmicimv2/win32_process `
  -SelectorSet @{handle = "4920"} `
  -Fragment Name
  
Get-WSManInstance `
  -ResourceURI wmicimv2/win32_process `
  -Enumerate -ReturnType epr

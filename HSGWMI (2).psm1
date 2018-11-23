# ******************************************************************************
#
# Script Name: Script_Name_Here.ps1
# Version: 1.0
# Author: Author_Name
# Date Created: Month Day, Year
# _______________________________________
#
# MODIFICATIONS: 
# Date Modified: 
# Modified By: 
# Reason for modification: 
# What was modified: 
# Description: This script lists all Members of a DL.
#
# Usage:
# ./ Script_Name_Here.ps1 -list "1MISC-Exchange Administrators" [-id #"JMXXXX"]
#
# ******************************************************************************

# ----------------------------------------------------------------------------- 
# Script: HSGWMImoduleV6.psm1 
# Author: ed wilson, msft 
# Date: 10/27/2011 18:38:33 
# Keywords: scripting techniques, wmi 
# comments: 
# HSG: hsg-10-29-11 
# ----------------------------------------------------------------------------- 
Function Get-WMIClassesWithQualifiers 
{ 
  <# 
   .Synopsis 
    This function finds wmi classes with a specific qualifier 
   .Description 
    This function allows you to explore WMI namespaces to find classes with a 
    specific qualifier. Search for qualifiers such as dynamic, abstract,  
    supportsupdate, singleton 
   .Example 
    Get-WmiClassesWithQualifiers -q dynamic 
    finds all wmi classes in default namespace (root/cimv2) that are dynamic 
   .Example 
    Get-WmiClassesWithQualifiers -q supportsupdate -n root/wmi 
    Finds all wmi classes in root/wmi that support updating 
   .Parameter Qualifier 
    The qualifier to search for. Ex: dynamic, abstract, supportsupdate, 
    Singleton 
   .Parameter Namespace 
    The namespace to search. Default is Root/Cimv2 
   .Role 
    Meta 
   .Component 
    HSGWMIModuleV6 
   .Notes 
    NAME:  Get-WmiClassesWithQualifiers 
    AUTHOR: ed wilson, msft 
    LASTEDIT: 10/16/2011 13:49:42 
    KEYWORDS: Scripting Techniques, WMI 
    HSG: HSG-10-22-11, 10-24-11 
   .Link 
     Http://www.ScriptingGuys.com 
 #Requires -Version 2.0 
 #> 
 Param([string]$qualifier = "dynamic", 
  [string]$namespace = "root\cimv2") 
 $classes = Get-WmiObject -list -namespace $namespace 
 foreach($class in $classes) 
 { 
  $query = "Select * from meta_class where __this isa ""$($class.name)"" " 
  $a = Get-WmiObject -Query $query -Namespace $namespace |  
  Select-Object -Property __class, qualifiers 
   if($a.qualifiers | ForEach-Object { $_ | Where-Object { $_.name -match "$qualifier" }}) 
    { $a.__class } 
  } #end foreach $class 
} #end function Get-WMIClassesWithQualifiers  
 
Function Out-TempFile 
{ 
<# 
  .Role 
   Helper 
  .Component 
   HSGWMIModuleV6  
#> 
 begin { $tmpfile= [io.path]::GetTempFileName() } 
 Process { $_ >> $tmpFile } 
 End { notepad $tmpfile | out-null 
       Remove-Item $tmpFile } 
} 
 
function New-Underline 
{ 
<# 
.Synopsis 
 Creates an underline the length of the input string 
.Example 
 New-Underline -strIN "Hello world" 
.Example 
 New-Underline -strIn "Morgen welt" -char "-" -sColor "blue" -uColor "yellow" 
.Example 
 "this is a string" | New-Underline 
.Role 
 Helper 
.Component 
 HSGWMIModuleV6  
.Notes 
 NAME: 
 AUTHOR: Ed Wilson 
 LASTEDIT: 5/20/2009 
 KEYWORDS: 
.Link 
 Http://www.ScriptingGuys.com 
#> 
[CmdletBinding()] 
param( 
      [Parameter(Mandatory = $true,Position = 0,valueFromPipeline=$true)] 
      [string] 
      $strIN, 
      [string] 
      $char = "=", 
      [string] 
      $sColor = "Green", 
      [string] 
      $uColor = "darkGreen", 
      [switch] 
      $pipe 
 ) #end param 
 $strLine= $char * $strIn.length 
 if(-not $pipe) 
  { 
   Write-Host -ForegroundColor $sColor $strIN 
   Write-Host -ForegroundColor $uColor $strLine 
  } 
  Else 
  { 
  $strIn 
  $strLine 
  } 
} #end new-underline function 
 
Function Get-WmiClassMethods 
{  
 <# 
   .Synopsis 
    This function returns implemented methods for a WMI class  
   .Description 
    This function returns implemented methods for a WMI class 
   .Example 
    Get-WmiClassMethods Win32_logicaldisk 
    Returns implemented methods from the Win32_logicaldisk class 
   .Example 
    Get-WmiClassMethods -class bcdstore -namespace root\wmi 
    Returns methods of the bcdStore WMI class from the root\wmi namespace 
   .EXAMPLE 
    Get-WmiClassMethods -class Win32_networkadapter -computer DC1 
    Returns methods from the Win32_networkadapter wmi class in the root\cimv2 
    namespace from a remote server named DC1   
   .Parameter Class 
    The name of the WMI class 
   .Parameter Namespace 
    The name of the WMI namespace. Defaults to root\cimv2 
   .Parameter Computer 
    The name of the computer. Defaults to local computer 
   .Role 
    Meta 
   .Component 
    HSGWMIModuleV6  
   .Notes 
    NAME:  Get-WmiClassMethods  
    AUTHOR: ed wilson, msft 
    LASTEDIT: 10/17/2011 13:43:24 
    KEYWORDS: 
    HSG: HSG-10-24-11, based upon WES-3-12-11 
   .Link 
     Http://www.ScriptingGuys.com 
 #Requires -Version 2.0 
 #> 
 Param( 
   [Parameter(Mandatory = $true,Position = 0)] 
   [string]$class, 
   [string]$namespace = "root\cimv2", 
   [string]$computer = $env:computername 
) 
 $abstract = $false 
 $method = $null 
 [wmiclass]$class = "\\{0}\{1}:{2}" -f $computer,$namespace,$class 
  Foreach($q in $class.Qualifiers) 
   { if ($q.name -eq 'Abstract') {$abstract = $true} } 
  If(!$abstract)  
    {  
     Foreach($m in $class.methods) 
      {  
       Foreach($q in $m.qualifiers)  
        {  
         if($q.name -match "implemented")  
          {  
            $method += $m.name + "`r`n" 
          } #end if name 
        } #end foreach q 
      } #end foreach m 
      if($method)  
        { 
         New-Underline -strIN $class.name  
         New-Underline "METHODS" -char "-" 
        } 
      $method 
    } #end if not abstract 
  $abstract = $false 
  $method = $null 
# } #end foreach class 
} #end function Get-WmiClassMethods 
 
Function Get-WmiClassProperties 
{  
<# 
   .Synopsis 
    This function returns writable properties for a WMI class  
   .Description 
    This function returns writable properties for a WMI class 
   .Example 
    Get-WMIClassProperties Win32_logicaldisk 
    Returns writable properties from the Win32_logicaldisk class 
   .Example 
    Get-WMIClassProperties -class bcdstore -namespace root\wmi 
    Returns properties of the bcdStore WMI class from the root\wmi namespace 
   .EXAMPLE 
    Get-WMIClassProperties -class Win32_networkadapter -computer DC1 
    Returns properties from the Win32_networkadapter wmi class in the root\cimv2 
    namespace from a remote server named DC1   
   .Parameter Class 
    The name of the WMI class 
   .Parameter Namespace 
    The name of the WMI namespace. Defaults to root\cimv2 
   .Parameter Computer 
    The name of the computer. Defaults to local computer 
   .Role 
    Meta 
   .Component 
    HSGWMIModuleV6  
   .Notes 
    NAME:  Get-WMIClassProperties  
    AUTHOR: ed wilson, msft 
    LASTEDIT: 10/17/2011 13:43:24 
    KEYWORDS: 
    HSG: HSG-10-24-11, based upon WES-3-12-11 
   .Link 
     Http://www.ScriptingGuys.com 
 #Requires -Version 2.0 
 #> 
 Param( 
   [Parameter(Mandatory = $true,Position = 0)] 
   [string]$class, 
   [string]$namespace = "root\cimv2", 
   [string]$computer = $env:computername 
) 
 $abstract = $false 
 $property = $null 
 [wmiclass]$class = "\\{0}\{1}:{2}" -f $computer,$namespace,$class 
  Foreach($q in $class.Qualifiers) 
   { if ($q.name -eq 'Abstract') {$abstract = $true} } 
  If(!$abstract)  
    {  
     Foreach($p in $class.Properties) 
      {  
       Foreach($q in $p.qualifiers)  
        {  
         if($q.name -match "write")  
          {  
            $property += $p.name + "`r`n" 
          } #end if name 
        } #end foreach q 
      } #end foreach p 
      if($property)  
        { 
         New-Underline -strIN $class.name 
         New-Underline "PROPERTIES" -char "-" 
        } 
      $property 
    } #end if not abstract 
  $abstract = $false 
  $property = $null 
# } #end foreach class 
} #end function Get-WmiClassProperties 
 
function Get-WmiKey 
{  
  <# 
   .Synopsis 
    This function returns the key property of a WMI class 
   .Description 
    This function returns the key property of a WMI class 
   .Example 
    Get-WMIKey win32_bios 
    Returns the key properties for the Win32_bios WMI class in root\ciimv2 
   .Example 
    Get-WmiKey -class Win32_product 
    Returns the key properties for the Win32_Product WMI class in root\cimv2 
   .Example 
    Get-WmiKey -class systemrestore -namespace root\default 
    Gets the key property from the systemrestore WMI class in the root\default 
    WMI namespace.  
   .Parameter Class 
    The name of the WMI class 
   .Parameter Namespace 
    The name of the WMI namespace. Defaults to root\cimv2 
   .Parameter Computer 
    The name of the computer. Defaults to local computer 
   .Role 
    Meta 
   .Component 
    HSGWMIModuleV6  
   .Notes 
    NAME:  Get-WMIKey 
    AUTHOR: ed wilson, msft 
    LASTEDIT: 10/18/2011 17:38:20 
    KEYWORDS: Scripting Techniques, WMI 
    HSG: HSG-10-24-2011 
   .Link 
     Http://www.ScriptingGuys.com 
 #Requires -Version 2.0 
 #> 
 Param( 
   [Parameter(Mandatory = $true,Position = 0)] 
   [string]$class, 
   [string]$namespace = "root\cimv2", 
   [string]$computer = $env:computername 
 ) 
  [wmiclass]$class = "\\{0}\{1}:{2}" -f $computer,$namespace,$class 
  $class.Properties |  
      Select-object @{Name="PropertyName";Expression={$_.name}} ` 
        -ExpandProperty Qualifiers |  
      Where-object {$_.Name -eq "key"} |  
      ForEach-Object {$_.PropertyName}  
} #end GetWmiKey 
 
 
Function Get-WmiKeyvalue 
{ 
  <# 
   .Synopsis 
    This function gets the __Path values for a WMI class 
   .Description 
    This function gets the __path values which will show value of key property 
   .Example 
    Get-WmiKeyvalue win32_bios 
    Gets the path to the Win32_bios class 
   .Example 
    Get-WmiKeyvalue -class win32_process  
    Gets the paths to each process running on computer 
   .Example 
    Get-WmiKeyvalue -class SystemRestore -namespace root\default 
    Gets paths of SystemRestore (requires admin rights) 
   .Parameter Class 
    The wmi class name 
   .Parameter Computername 
    The name of the computer 
   .Parameter Namespace 
    The namespace containing the WMI class 
   .Role 
    Meta 
   .Component 
    HSGWMIModuleV6  
   .Notes 
    NAME:  Get-WmiKeyvalue 
    AUTHOR: ed wilson, msft 
    LASTEDIT: 10/19/2011 17:52:10 
    KEYWORDS: Scripting Techniques, WMI 
    HSG: HSG-10-26-11 
   .Link 
     Http://www.ScriptingGuys.com 
 #Requires -Version 2.0 
 #> 
 Param( 
    [Parameter(Mandatory=$true)] 
    [string]$class, 
    [string]$computername = $env:COMPUTERNAME, 
    [string]$namespace = "root\cimv2" 
) 
  Get-WmiObject -Class $class -ComputerName $computername -Namespace $namespace |  
  Select __PATH 
} #end function get-WmiKeyvalue 
 
Filter HasWmiValue 
{ 
  <# 
   .Synopsis 
    This is a filter that will remove empty property values 
    from the returned WMI information  
   .Description 
    This removes empty property values from returned WMI information. 
    It is useful because most WMI classes return many lines of empty 
    properties. By using this filter, the returned WMI information is 
    easier to read. It also makes it easier to find required information. 
   .Example 
    Get-WmiObject -class win32_bios | HasWMiValue 
    Returns BIOS information from the local computer. Only WMI properties 
    that contain a value are returned. 
   .Example 
    gwmi -cl win32_bios -cn remotehost | HasWMIValue 
    Returns BIOS information from a remote computer named remotehost. Only 
    WMI properties that contain a value are returned.  
   .Role 
    Helper 
   .Component 
    HSGWMIModuleV6  
   .Notes 
    NAME:  HasWmiValue 
    AUTHOR: ed wilson, msft 
    LASTEDIT: 10/20/2011 12:33:45 
    KEYWORDS: 
    HSG: HSG-10-27-2011 
   .Link 
     Http://www.ScriptingGuys.com 
 #Requires -Version 2.0 
 #> 
   $_.properties | 
   foreach-object -BEGIN { $_.path | new-underline -pipe} -Process { 
     If($_.value -AND $_.name -notmatch "__") 
      { 
        @{ $($_.name) = $($_.value) } 
      } #end if 
    } #end foreach property 
} #end filter HasWmiValue 
 
Function Get-WmiClassesAndQuery 
{ 
  <# 
   .Synopsis 
    This function searches for WMI classes based upon a wild card 
    and then it will query those WMI classes.   
   .Description 
    This function searches for WMI classes based upon a wild card 
    pattern, and then it will query those WMI classes. It can use 
    standard wild card patterns. You can specify namespace, and  
    remote computer name. This version does not accept alternate 
    credentials, but you can use runas to run PowerShell with alternate 
    credentials. It automatically filters out abstract classes.  
   .Example 
    Get-WMIClassesAndQuery -class *disk* 
    Searches for all WMI classes that contain the letters disk in the class  
    name. It then filters out only the non-abstract classes, and then queries  
    them. The WMI classes queried come from the root\cimv2 namespace and are 
    on the local computer. 
   .Example 
    Get-WmiClassesAndQuery -class *adapter* -namespace "root\wmi" -cn dc1 
    Gets all WMI classes from the root\wmi namespace that are on the remote 
    computer named dc1. It then queries those non-abstract classes and returns 
    the information.  
   .Parameter Class 
    A WMI class name, or wild card pattern that will find WMI class names 
   .Parameter Namespace 
    A valid WMI namespace. By default it is root\cimv2 
   .Parameter Computer 
    The computer from which to return information. Defaults to local computer. 
   .Role 
    Query 
   .Component 
    HSGWMIModuleV6  
   .Notes 
    NAME:  Get-WmiClassesAndQuery 
    AUTHOR: ed wilson, msft 
    LASTEDIT: 10/27/2011 13:52:46 
    KEYWORDS: Scripting Techniques, WMI 
    HSG: HSG-10-28-11 
   .Link 
     Http://www.ScriptingGuys.com 
 #Requires -Version 2.0 
 #> 
 Param( 
   [Parameter(Mandatory=$true,Position=0,valueFromPipeline=$true)] 
   [string]$class, 
   [string]$namespace = "Root\cimv2", 
   [string]$computer = $env:computername 
 ) 
  Get-WmiObject -List $class -Namespace $namespace -ComputerName $computer | 
  ForEach-Object {  
   $abstract = $false 
   [wmiclass]$class = "\\{0}\{1}:{2}" -f $computer,$namespace,$_.name 
  Foreach($q in $_.Qualifiers) 
   { if ($q.name -eq 'Abstract') {$abstract = $true} } 
  If(!$abstract)  
    { 
     Get-WmiObject -Class $_.name -Namespace $namespace -ComputerName $computer 
     } #end if !$abstract  
  } #end Foreach-object 
} #end function get-wmiclassesandquery 
 
 
 # *** Aliases and descriptions *** 
 New-Alias -Name gwq -Value Get-WmiClassesAndQuery -Description "HSG WMI module: query" -Scope "global" 
 New-Alias -Name gwcq -Value Get-WMIClassesWithQualifiers -Description "HSG WMI module: meta" -Scope "global" 
 New-Alias -Name gwcm -Value Get-WmiClassMethods -Description "HSG WMI module: meta" -Scope "global" 
 New-Alias -Name gwcp -Value Get-WmiClassProperties -Description "HSG WMI module: meta" -Scope "global" 
 New-Alias -Name gwck -Value Get-WmiKey -Description "HSG WMI module: meta" -Scope "global" 
 New-Alias -Name gwkv -Value Get-WmiKeyvalue -Description "HSG WMI module: meta" -Scope "global" 
 New-Alias -Name hwv -Value haswmivalue -Description "HSG WMI module: helper" -Scope "global" 
 New-Alias -Name nu -Value New-Underline -Description "HSG WMI module: helper" -Scope "global" 
 New-Alias -Name otf -Value Out-TempFile -Description "HSG WMI module: helper" -Scope "global" 
 
Export-ModuleMember -Function * -Alias *

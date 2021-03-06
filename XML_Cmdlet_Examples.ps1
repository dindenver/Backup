#
# Windows PowerShell in Action
#
# Chapter 16 
#
# Examples using the ConvertTo-Xml 
# and Export-CliXml cmdlets
#

# ConvertTo-XML examples
$services = (Get-Service)[0..2]
$doc = $services| ConvertTo-Xml
$doc
$doc.Objects.Object
$services[0] | ConvertTo-Xml -As string
$services[0] | ConvertTo-Xml -NoTypeInformation -As string


# Export-Clixml - save data in a recoverable format

$data = @{a=1;b=2;c=3},"Hi there", 3.5
$data | Format-List
$data | Export-Clixml -Path savedData.xml
"Recovered data has the same structure"
$recoveredData = Import-Clixml savedData.xml 
$recoveredData | Format-List


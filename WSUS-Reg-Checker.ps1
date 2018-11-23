$CPUs=gc .\CPUs.txt
foreach ($CPU IN $CPUs)
{
$sys="$CPU";$AU=Invoke-Command -ComputerName $sys -ScriptBlock {$AU=$null ; $sys = gc env:computername ; cd HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate ; $AU=(Get-ItemProperty AU).NoAutoRebootWithLoggedOnUsers ; return $AU};if ($AU -eq $null) {$AU="Default"};write-output "$sys - $AU"
}

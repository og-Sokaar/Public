<#
+---------------------------------------------------------------------------------------------+ 
| NetScan                                                                                     | 
+---------------------------------------------------------------------------------------------+ 
|                                                                                             |
|   AUTHOR      : og-sokaar                                                                   |
|   DESCRIPTION : Network scan performed slowly to to simulate a user manual scan and avoid   |
|                 IPS/IDS Systems works with /24 networks                                     |
+---------------------------------------------------------------------------------------------+ 
#>

$Network = "192.168.16"
$Range = 1..254
$ErrorActionPreference = 'silentlycontinue'
$ScanOut = @()
$csv = "C:\netscan.csv"
foreach ($Add in $Range){
$ip = "{0}.{1}" -f $Network,$Add
$test = Test-NetConnection -ComputerName $ip -ErrorAction $ErrorActionPreference -WarningAction $ErrorActionPreference | select ComputerName, PingSucceeded
$Scan = New-Object PSObject
Add-Member -InputObject $Scan -MemberType NoteProperty -Name "Computer Name" -Value $test.ComputerName
Add-Member -InputObject $Scan -MemberType NoteProperty -Name "Response" -Value $test.PingSucceeded
$ScanOut += $Scan
}

$ScanOut | Export-Csv $csv -NoClobber -NoTypeInformation
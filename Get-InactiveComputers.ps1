# Gets time stamps for all computers in the domain that have NOT logged in since after specified date

Import-Module ActiveDirectory
$domain = “asiapacific.hagemeyer.pri”
$DaysInactive = 30
$time = (Get-Date).Adddays(-($DaysInactive))

# Get all AD computers with lastLogonTimestamp less than our time and export to CSV
Get-ADComputer -Server $domain -Filter {(enabled -eq $True) -and (LastLogonTimeStamp -lt $time)} -Properties LastLogonTimeStamp | export-csv c:\temp\StalePCs.csv -notypeinformation

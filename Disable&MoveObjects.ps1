Import-Module ActiveDirectory

# Enter the domain here:
$Domain = "asiapacific.hagemeyer.pri"

#Enter the disabled objects OU here:
$DisabledOU = "OU=Clients,OU=Inactive Computer Accounts,DC=asiapacific,DC=hagemeyer,DC=pri"

#Enter the path to the CSV here:
$StalePCs = Import-CSV C:\Temp\StalePCs.CSV

#Disable objects and move them to the disabled objects OU.
foreach ($StalePC in $StalePCs) {
Set-ADComputer -Server $domain -Identity $StalePC.DistinguishedName -Enabled $false
Move-AdObject -Server $domain -Identity $StalePC.DistinguishedName -TargetPath $DisabledOU
}

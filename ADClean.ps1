#                   lllllll   iiii  kkkkkkkk           kkkkkkkk           
#                   l:::::l  i::::i k::::::k           k::::::k           
#                   l:::::l   iiii  k::::::k           k::::::k           
#                   l:::::l         k::::::k           k::::::k           
#     ooooooooooo    l::::l iiiiiii  k:::::k    kkkkkkk k:::::k    kkkkkkk   aaaaaaaaaaaaa       
#   oo:::::::::::oo  l::::l i:::::i  k:::::k   k:::::k  k:::::k   k:::::k    a::::::::::::a      
#  o:::::::::::::::o l::::l  i::::i  k:::::k  k:::::k   k:::::k  k:::::k     aaaaaaaaa:::::a    
#  o:::::ooooo:::::o l::::l  i::::i  k:::::k k:::::k    k:::::k k:::::k               a::::a     
#  o::::o     o::::o l::::l  i::::i  k::::::k:::::k     k::::::k:::::k         aaaaaaa:::::a     
#  o::::o     o::::o l::::l  i::::i  k:::::::::::k      k:::::::::::k        aa::::::::::::a     
#  o::::o     o::::o l::::l  i::::i  k:::::::::::k      k:::::::::::k       a::::aaaa::::::a     
#  o::::o     o::::o l::::l  i::::i  k::::::k:::::k     k::::::k:::::k     a::::a    a:::::a     
#  o:::::ooooo:::::ol::::::li::::::ik::::::k k:::::k   k::::::k k:::::k    a::::a    a:::::a     
#  o:::::::::::::::ol::::::li::::::ik::::::k  k:::::k  k::::::k  k:::::k   a:::::aaaa::::::a      
#   oo:::::::::::oo l::::::li::::::ik::::::k   k:::::k k::::::k   k:::::k   a::::::::::aa:::a    
#     ooooooooooo   lllllllliiiiiiiikkkkkkkk    kkkkkkkkkkkkkkk    kkkkkkk  aaaaaaaaaa  aaaa     
#                  
#         
#                 
#=============================================================================
# REVISION HISTORY
#=============================================================================
# Version : 1.0
# Date    : [December 2013]
# Author  : Timothy Mukaibo
# Notes   : Initial version
#=============================================================================
#* FileName: Get-OldADObjects.ps1
#*=============================================================================
#* Script Name: [Get-OldADObjects]
#* Created: [December 2013]
#* Author: Timothy Mukaibo
#* Company: Olikka, Specialised - Focussed - Different
#* Web: http://olikka.com.au/
#* Requirements: AD Powershell Module
#* Example 1. Print a list of all computers who haven't logged in for 45 days: 
#*    "powershell.exe -File Get-OldADObjects.ps1
#* Example 2. Disable and move all computer objects which haven't logged in for 45 days: 
#*    "powershell.exe -File Get-OldADObjects.ps1 -Cleanup
#*=============================================================================
#* Purpose: This script identifies Computers which haven't logged on in more than 90 days
#*
#* Parameters:
#*    -Cleanup        - This will disable and move all identified objects to $OU
#*    -WriteOutput    - Outputs the object and action taken to a CSV file in the current directory (default: to screen)
#*    -AsiaPacific    - Look at the old domain instead of internal.pri
#*           
#*=============================================================================
Param(
    [Parameter(Mandatory=$false)][switch]$WriteOutput=$false,
    [Parameter(Mandatory=$false)][switch]$Cleanup=$false,
    [Parameter(Mandatory=$false)][switch]$asiapacific=$false
)
# Global Variables. Change these as appropriate!
If($asiapacific) {
    $DisabledOU = "OU=Clients,OU=Inactive Computer Accounts,DC=asiapacific,DC=hagemeyer,DC=pri"
    $Domain = "asiapacific.hagemeyer.pri"
}
Else {
    $DisabledOU = "OU=Clients,OU=Inactive,DC=INTERNAL,DC=PRI"
    $Domain = "internal.pri"
}

$GeneralADDaysToCheck = 45 # Days to check for Computers

#*=============================================================================
#* FUNCTION LISTINGS
#*=============================================================================
Function Write-ADCollection() {
    Param(
        [Parameter(Mandatory=$true)]$ADCollection
    )
    $ADCollection | select Name,@{Name="lastLogonTimestamp"; Expression= { [DateTime]::FromFileTime($_.lastLogonTimestamp) }},distinguishedName
}

Function DisableADObject($thingy) {
    # Format the date with the Aussie locale
    $timestamp = get-date -uformat "%d/%m/%Y %r" 
    $description = "Disabled: $timestamp $($thingy.description)"
    $thingy | Set-ADObject -Description $description
    $thingy | Disable-ADAccount
    $thingy | Move-ADObject -TargetPath $DisabledOU
    
}

Function Check-ADObject() {
    Param(
        [Parameter(Mandatory=$true)][string]$ObjectType,
        [Parameter(Mandatory=$true)][int]$Days
    )
    $OlderThan = (Get-Date).AddDays(-($days))
    
    $Collection = Get-ADComputer -Filter { lastLogonTimestamp -lt $OlderThan -and enabled -eq $true } -properties lastLogonTimestamp,objectClass -Server $domain
    
    Write-Host "Found $($Collection.count) $ObjectType objects older than $days days"
    
    If($Cleanup) {
        $Counter = 0
        ForEach($object in $Collection) {
            Write-Progress -Activity "Disabling $objectType objects older than $days days" -Status "Disabling $($object.name)" -PercentComplete (($Counter / $collection.count) * 100)
            DisableADObject($object)
            $object | Add-Member -type NoteProperty -Name "Action Performed" -Value "Disabled" -Force
            $counter++
        }
    }
    
    Write-ADCollection $Collection

    If($WriteOutput) {
        $Timestamp = (get-date -format u).replace(":", "-")
        $filename = "$($objectType)s $timestamp.csv"
        Write-ADCollection $Collection | Export-CSV "$filename" -NoTypeInformation 
        Write-Host -ForegroundColor green "`nWrote $objectType output to `"$filename`""
    }
}

#*=============================================================================
#* SCRIPT BODY
#*=============================================================================
Import-Module ActiveDirectory

Check-ADObject -objectType "computer" -days $GeneralADDaysToCheck


#*=============================================================================
#* END OF SCRIPT
#*=============================================================================

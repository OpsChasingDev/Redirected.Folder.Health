# returns list of folders corresponding to SIDs on a machine
# Get-ChildItem 'REGISTRY::HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\' | Select Name

# returns the username (SamAccountName) and corresponding SID for each user profile on the computer
$SID = Get-ChildItem 'REGISTRY::HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\' | Select -ExpandProperty Name
foreach ($s in $SID) {
    Get-ItemProperty -Path "REGISTRY::$s" -Name "ProfileImagePath"
}


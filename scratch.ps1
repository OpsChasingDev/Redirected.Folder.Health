break
# returns list of folders corresponding to SIDs on a machine
# Get-ChildItem 'REGISTRY::HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\' | Select Name

# returns the username (SamAccountName) and corresponding SID for each user profile on the computer
$SID = Get-ChildItem 'REGISTRY::HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\' | Select-Object -ExpandProperty Name
foreach ($s in $SID) {
    $Prof = Get-ItemProperty -Path "REGISTRY::$s" -Name "ProfileImagePath"
    $User = ($Prof.ProfileImagePath.ToString()).Split('\')[-1]
    $obj = [PSCustomObject]@{
        UserSID = $Prof.PSChildName
        UserName = $User
    }
    Write-Output $obj
}
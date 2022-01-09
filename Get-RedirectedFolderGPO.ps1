Function Get-RedirectedFolderGPO {

    # grabs all GPOs, gets an XML report of their config, and returns GPOs that have folder redirection settings
    $AllGPO = Get-GPO -All
    $AllGuid = $AllGPO.Id.Guid
    ForEach ($Guid in $AllGuid) {
        [xml]$Report = Get-GPOReport -Guid $Guid -ReportType xml
        If ($Report.GPO.User.ExtensionData.Name -contains 'Folder Redirection'){

            # initial object creation and member assignments
            $RedirectionGPO = Get-GPO -Guid $Guid
            $obj = New-Object -TypeName psobject
            $obj | Add-Member -Name 'DisplayName' -MemberType NoteProperty -Value $RedirectionGPO.DisplayName
            $obj | Add-Member -Name 'GUID' -MemberType NoteProperty -Value $RedirectionGPO.Id
            $obj | Add-Member -Name 'Enabled' -MemberType NoteProperty -Value $Report.GPO.LinksTo.Enabled
            $obj | Add-Member -Name 'Enforced' -MemberType NoteProperty -Value $Report.GPO.LinksTo.NoOverride
            $obj | Add-Member -Name 'Link' -MemberType NoteProperty -Value $Report.GPO.LinksTo.SOMPath

            # empty array that will hold the letters abbreviating all the libraries found
            $AbbreviationCollection = @()

            # determine the libraries in the GPO
            $obj | Add-Member -Name 'RedirectionPath' -MemberType NoteProperty -Value $Report.GPO.User.ExtensionData.Extension.Folder.Location.DestinationPath

            # adds each library to the object and the collection of library abbreviations if they are found
            ForEach ($l in $obj.RedirectionPath) {
                If ($l -match 'Desktop') {
                    $obj | Add-Member -Name 'Desktop' -MemberType NoteProperty -Value $l
                    $AbbreviationCollection += 'D'
                }
                If ($l -match 'Documents') {
                    $obj | Add-Member -Name 'Documents' -MemberType NoteProperty -Value $l
                    $AbbreviationCollection += 'O'
                }
                If ($l -match 'Downloads') {
                    $obj | Add-Member -Name 'Downloads' -MemberType NoteProperty -Value $l
                    $AbbreviationCollection += 'W'
                }
                If ($l -match 'Music') {
                    $obj | Add-Member -Name 'Music' -MemberType NoteProperty -Value $l
                    $AbbreviationCollection += 'M'
                }
                If ($l -match 'Pictures') {
                    $obj | Add-Member -Name 'Pictures' -MemberType NoteProperty -Value $l
                    $AbbreviationCollection += 'P'
                }
                If ($l -match 'Videos') {
                    $obj | Add-Member -Name 'Videos' -MemberType NoteProperty -Value $l
                    $AbbreviationCollection += 'V'
                }
                If ($l -match 'Favorites') {
                    $obj | Add-Member -Name 'Favorites' -MemberType NoteProperty -Value $l
                    $AbbreviationCollection += 'F'
                }
                If ($l -match 'AppData') {
                    $obj | Add-Member -Name 'AppData' -MemberType NoteProperty -Value $l
                    $AbbreviationCollection += 'A'
                }
                If ($l -match 'Start Menu') {
                    $obj | Add-Member -Name 'Start Menu' -MemberType NoteProperty -Value $l
                    $AbbreviationCollection += 'S'
                }
                If ($l -match 'Contacts') {
                    $obj | Add-Member -Name 'Contacts' -MemberType NoteProperty -Value $l
                    $AbbreviationCollection += 'C'
                }
                If ($l -match 'Links') {
                    $obj | Add-Member -Name 'Links' -MemberType NoteProperty -Value $l
                    $AbbreviationCollection += 'L'
                }
                If ($l -match 'Searches') {
                    $obj | Add-Member -Name 'Searches' -MemberType NoteProperty -Value $l
                    $AbbreviationCollection += 'H'
                }
                If ($l -match 'Saved Games') {
                    $obj | Add-Member -Name 'Saved Games' -MemberType NoteProperty -Value $l
                    $AbbreviationCollection += 'G'
                }
            }

            # creates object member containing all the library abbreviations as a string array
            $obj | Add-Member -Name 'Library' -MemberType NoteProperty -Value $AbbreviationCollection

            # outputs the object
            Write-Output $obj
        }
    }
}
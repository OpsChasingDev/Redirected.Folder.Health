Function Get-RedirectedFolderGPO {
    <#
    .SYNOPSIS
        Retrieves information about a GPO in the current domain that has Folder Redirection settings.
    .DESCRIPTION
        Retrieves information about a GPO in the current domain that has Folder Redirection settings.  By default, this cmdlet will look through all GPOs in the current domain and output information about all GPOs that have settings pertaining to Folder Redirections.
The intended use case for this function is to easily get information on which policies control Folder Redirections.  The output of this cmdlet contains a property called "Library" which will hold a single letter abbreviation for each library the GPO is found to redirect.  This is meant to be used for constructing scripts where the letter abbreviations are stored in a variable and then passed to the Get-RFH function's -Library parameter to check the status of folder redirections on computers.
    .EXAMPLE
        PS C:\> <example usage>
        Explanation of what the example does
    .INPUTS
        Inputs (if any)
    .OUTPUTS
        Output (if any)
    .NOTES
        General notes
    #>
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [Alias('Get-RFGPO')]
    Param (

        [Parameter(ParameterSetName = 'Name')]
        [string]$Name,

        [Parameter(ParameterSetName = 'Id',
                    ValueFromPipeline)]
        [Alias('Guid')]
        [string]$Id

    )
    
    # determines which GPOs will be used in the function based on parameter input
    If (!$Name -and !$Id) {
        $AllGPO = Get-GPO -All
    }
    ElseIf ($Name -and !$Id) {
        $AllGPO = Get-GPO -Name $Name
    }
    ElseIf (!$Name -and $Id) {
        $AllGPO = Get-GPO -Guid $Id
    }

    # gets XML report of specified GPO(s) (does this for all GPOs by default), and returns GPOs that have folder redirection settings
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
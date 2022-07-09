Function Get-RedirectedFolderGPO {
    <#
    .SYNOPSIS
        Retrieves information about a GPO in the current domain that has Folder Redirection settings.
    .DESCRIPTION
        Retrieves information about a GPO in the current domain that has Folder Redirection settings.  By default, this cmdlet will look through all GPOs in the current domain and output information about all GPOs that have settings pertaining to Folder Redirections.
The intended use case for this function is to easily get information on which policies control Folder Redirections.  For more robust functionality, the output of this cmdlet contains a property called "Library" which will hold a single letter abbreviation for each library the GPO is found to redirect.  This is meant to be used for constructing scripts where the letter abbreviations are stored in a variable and then passed to the Get-RFH function's -Library parameter to check the status of folder redirections on computers.
    .EXAMPLE
        PS C:\> Get-RedirectedFolderGPO

        Output:
DisplayName     : RedirectedFolders
GUID            : 368bd0ae-e4ad-463a-b640-f4898f380cea
Enabled         : {false, false}
Enforced        : {false, false}
Link            : {savylabs.local/SavyLabs/SL_Users, savylabs.local/SavyLabs/SL_Groups}
RedirectionPath : {\\SL-DC-01\RedirectedFolders\%USERNAME%\Downloads, \\SL-DC-01\RedirectedFolders\%USERNAME%\Favorites, \\SL-DC-01\RedirectedFolders\%USERNAME%\Videos, \\SL-DC-01\RedirectedFolders\%USERNAME%\Music…}
Downloads       : \\SL-DC-01\RedirectedFolders\%USERNAME%\Downloads
Favorites       : \\SL-DC-01\RedirectedFolders\%USERNAME%\Favorites
Videos          : \\SL-DC-01\RedirectedFolders\%USERNAME%\Videos
Music           : \\SL-DC-01\RedirectedFolders\%USERNAME%\Music
Pictures        : \\SL-DC-01\RedirectedFolders\%USERNAME%\Pictures
Documents       : \\SL-DC-01\RedirectedFolders\%USERNAME%\Documents
Start Menu      : \\SL-DC-01\RedirectedFolders\%USERNAME%\Start Menu
Contacts        : \\SL-DC-01\RedirectedFolders\%USERNAME%\Contacts
Links           : \\SL-DC-01\RedirectedFolders\%USERNAME%\Links
Searches        : \\SL-DC-01\RedirectedFolders\%USERNAME%\Searches
Saved Games     : \\SL-DC-01\RedirectedFolders\%USERNAME%\Saved Games
Library         : {W, F, V, M…}

DisplayName     : RedirectedFolders_New02
GUID            : d8922189-cf31-4c47-bdbf-c70299a7aba4
Enabled         : true
Enforced        : true
Link            : savylabs.local
RedirectionPath : {\\SL-DC-01\RedirectedFolders\%USERNAME%\Pictures, \\SL-DC-01\RedirectedFolders\%USERNAME%\Music, \\SL-DC-01\RedirectedFolders\%USERNAME%\Videos, \\SL-DC-01\RedirectedFolders\%USERNAME%\Favorites…}
Pictures        : \\SL-DC-01\RedirectedFolders\%USERNAME%\Pictures
Music           : \\SL-DC-01\RedirectedFolders\%USERNAME%\Music
Videos          : \\SL-DC-01\RedirectedFolders\%USERNAME%\Videos
Favorites       : \\SL-DC-01\RedirectedFolders\%USERNAME%\Favorites
Downloads       : \\SL-DC-01\RedirectedFolders\%USERNAME%\Downloads
Library         : {P, M, V, F…}

DisplayName     : RedirectedFolders_New01
GUID            : ea3025bc-034a-404a-b0ca-2f9cbf9d604c
Enabled         : true
Enforced        : true
Link            : savylabs.local
RedirectionPath : {\\SL-DC-01\RedirectedFolders\%USERNAME%\Desktop, \\SL-DC-01\RedirectedFolders\%USERNAME%\Documents}
Desktop         : \\SL-DC-01\RedirectedFolders\%USERNAME%\Desktop
Documents       : \\SL-DC-01\RedirectedFolders\%USERNAME%\Documents
Library         : {D, O}

        Searches all GPOs in the current domain and returns information about those found that have settings pertaining to Folder Redirections.
    .EXAMPLE
        PS C:\> Get-RedirectedFolderGPO | Where-Object {$_.Enabled -eq $true} | Select-Object DisplayName,Link

        Output:
DisplayName             Link
-----------             ----
RedirectedFolders_New02 savylabs.local
RedirectedFolders_New01 savylabs.local

        Retrieves all GPOs with Folder Redirection settings, returns only those found where the link to an Active Directory container is enabled, and returns the DisplayName of the GPO as well as where the GPO is linked in Active Directory.
    .EXAMPLE
        PS C:\> $Library = (Get-RedirectedFolderGPO -Name 'RedirectedFolders_New01').Library
        PS C:\> Get-RFH -ComputerName SL-COMPUTER-001 -Library $Library

        Output:
ComputerName : SL-COMPUTER-001
User         : Administrator
Desktop      : C:\Users\Administrator\Desktop
Documents    : C:\Users\Administrator\Documents
Downloads    :
Music        :
Pictures     :
Video        :
Favorites    :
AppData      :
StartMenu    :
Contacts     :
Links        :
Searches     :
SavedGames   :

ComputerName : SL-COMPUTER-001
User         : user1
Desktop      : \\SL-DC-01\RedirectedFolders\user1\Desktop
Documents    : \\SL-DC-01\RedirectedFolders\user1\Documents
Downloads    :
Music        :
Pictures     :
Video        :
Favorites    :
AppData      :
StartMenu    :
Contacts     :
Links        :
Searches     :
SavedGames   :

        Gets the letter abbreviations for the libraries redirected by the GPO called "RedirectedFolders_New01" and stores them in the variable called 'Library'.  The function Get-RFH is then called and uses $Library as the value for the -Library parameter to check those libraries for redirection on the computer called SL-COMPUTER-001.
    .EXAMPLE
        PS C:\> (Get-GPO -Name 'RedirectedFolders').Id.Guid | Get-RedirectedFolderGPO

        Output:
DisplayName     : RedirectedFolders
GUID            : 368bd0ae-e4ad-463a-b640-f4898f380cea
Enabled         : {false, false}
Enforced        : {false, false}
Link            : {savylabs.local/SavyLabs/SL_Users, savylabs.local/SavyLabs/SL_Groups}
RedirectionPath : {\\SL-DC-01\RedirectedFolders\%USERNAME%\Downloads, \\SL-DC-01\RedirectedFolders\%USERNAME%\Favorites, \\SL-DC-01\RedirectedFolders\%USERNAME%\Videos, \\SL-DC-01\RedirectedFolders\%USERNAME%\Music…}
Downloads       : \\SL-DC-01\RedirectedFolders\%USERNAME%\Downloads
Favorites       : \\SL-DC-01\RedirectedFolders\%USERNAME%\Favorites
Videos          : \\SL-DC-01\RedirectedFolders\%USERNAME%\Videos
Music           : \\SL-DC-01\RedirectedFolders\%USERNAME%\Music
Pictures        : \\SL-DC-01\RedirectedFolders\%USERNAME%\Pictures
Documents       : \\SL-DC-01\RedirectedFolders\%USERNAME%\Documents
Start Menu      : \\SL-DC-01\RedirectedFolders\%USERNAME%\Start Menu
Contacts        : \\SL-DC-01\RedirectedFolders\%USERNAME%\Contacts
Links           : \\SL-DC-01\RedirectedFolders\%USERNAME%\Links
Searches        : \\SL-DC-01\RedirectedFolders\%USERNAME%\Searches
Saved Games     : \\SL-DC-01\RedirectedFolders\%USERNAME%\Saved Games
Library         : {W, F, V, M…}

        Gets the GUID for the GPO called "RedirectedFolders" and passed it through the pipeline to Get-RedirectedFolderGPO to return information about the GPO's settings.
    .INPUTS
        System.String[]
        This cmdlet accepts a GUID as a string in order to specify which GPO needs to be targeted for getting Folder Redirection settings.
    .OUTPUTS
        System.Management.Automation.PSCustomObject
        By default, a PSCustomObject is returned by the cmdlet with the below members:

Name            MemberType   Definition
----            ----------   ----------
Equals          Method       bool Equals(System.Object obj)
GetHashCode     Method       int GetHashCode()
GetType         Method       type GetType()
ToString        Method       string ToString()
Contacts        NoteProperty System.String
DisplayName     NoteProperty System.String
Documents       NoteProperty System.String
Downloads       NoteProperty System.String
Enabled         NoteProperty System.Object[]
Enforced        NoteProperty System.Object[]
Favorites       NoteProperty System.String
GUID            NoteProperty System.Object[]
Library         NoteProperty System.Object[]
Link            NoteProperty System.Object[]
Links           NoteProperty System.String
Music           NoteProperty System.String
Pictures        NoteProperty System.String
RedirectionPath NoteProperty System.Object[]
Saved Games     NoteProperty System.String
Searches        NoteProperty System.String
Start Menu      NoteProperty System.String
Videos          NoteProperty System.String
    .NOTES
        Author: Robert Stapleton
        Version: 2.0.0
        Date: 2022-01-11
        Tested with PowerShell versions 3, 4, 5, and 7
    .LINK
        https://github.com/opschasingdev/Redirected.Folder.Health
    #>
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [Alias('Get-RFGPO')]
    Param (

        [Parameter(ParameterSetName = 'Name')]
        [string]$Name,

        [Parameter(ParameterSetName = 'Id',
                    ValueFromPipeline = $true)]
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
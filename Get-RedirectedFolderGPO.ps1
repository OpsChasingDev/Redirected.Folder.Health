# find best method to parse XML reports of the returned findings for which libraries they deal with and what their paths are
# the above parsed values need to be added as members to the object stored in $RedirectionGPO
# output
    # psobject with the below properties
        # GPO GUID
        # GPO Friendly Name
        # enabled or not
        # enforced or not
        # links to AD
        # gpostatus
        # description
        # each library found (name of each property here will be the name of the library); this property will be an object with more properties to drill into:
            # Name
            # Letter
            # Path
<#
[xml]$gpo = Get-GPO -Name 'RedirectedFolders_New01' | Get-GPOReport -ReportType xml
# basic info
$gpo.GPO
# enabled and enforced (Enabled and NoOverride, respectively)
$gpo.GPO.LinksTo
# returns 'Folder Redirection' if handling folder redirection settings
$gpo.GPO.User.ExtensionData.Name
#>

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

        # determine the libraries in the GPO
        $obj | Add-Member -Name 'Enabled' -MemberType NoteProperty -Value $Report.GPO.LinksTo.Enabled
        $obj | Add-Member -Name 'Enforced' -MemberType NoteProperty -Value $Report.GPO.LinksTo.NoOverride
        $obj | Add-Member -Name 'Desktop' -MemberType NoteProperty -Value $Report.GPO.User.ExtensionData.Extension.Folder.Location.DestinationPath

        Write-Output $obj
    }
}
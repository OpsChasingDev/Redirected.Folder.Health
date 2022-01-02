break
# objective here is to find which GPOs in the current domain have folder redirection settings
    # get all GPOs in an XML report
    # parse the XML for information only a GPO handling folder redirections would have
    # once discovered GPOs match the findings, parse those XML reports further to get the setting details
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

(Get-GPOReport -All -ReportType xml).count

[xml]$gpo = Get-GPO -Name 'RedirectedFolders_New01' | Get-GPOReport -ReportType xml

# basic info
$gpo.GPO
# enabled and enforced (Enabled and NoOverride, respectively)
$gpo.GPO.LinksTo
# returns 'Folder Redirection' if handling folder redirection settings
$gpo.GPO.User.ExtensionData.Name

# grabs all GPOs, gets an XML report of their config, and returns GPOs that have folder redirection settings
$AllGPO = Get-GPO -All
$AllGuid = $AllGPO.Id.Guid
ForEach ($Guid in $AllGuid) {
    [xml]$Report = Get-GPOReport -Guid $Guid -ReportType xml
    If ($Report.GPO.User.ExtensionData.Name -contains 'Folder Redirection'){
        Get-GPO -Guid $Guid
    }
}
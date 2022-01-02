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
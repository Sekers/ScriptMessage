DynamicParam
{
    # Initialize Parameter Dictionary
    $ParameterDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()

    # Make Mg* parameters appear only if messaging 'Service' is 'MgGraph'.
    if ($Service -eq 'MgGraph')
    { 
        $ParameterAttributes = [System.Management.Automation.ParameterAttribute]@{
            ParameterSetName = "MgGraph"
            Mandatory = $true
            ValueFromPipeline = $true
            ValueFromPipelineByPropertyName = $true
        }

        $AttributeCollection = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
        $AttributeCollection.Add($ParameterAttributes)

        $DynamicParameter1 = [System.Management.Automation.RuntimeDefinedParameter]::new(
            'MgPermissionType', [string], $AttributeCollection) # Delegated or Application. See: https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent#permission-types and https://docs.microsoft.com/en-us/graph/auth/auth-concepts#delegated-and-application-permissions.

        $DynamicParameter2 = [System.Management.Automation.RuntimeDefinedParameter]::new(
            'MgDisconnectWhenDone', [bool], $AttributeCollection)

        $DynamicParameter3 = [System.Management.Automation.RuntimeDefinedParameter]::new(
            'MgTenantID', [string], $AttributeCollection) 

        $DynamicParameter4 = [System.Management.Automation.RuntimeDefinedParameter]::new(
            'MgClientID', [string], $AttributeCollection) 

        $ParameterDictionary.Add('MgPermissionType', $DynamicParameter1)
        $ParameterDictionary.Add('MgDisconnectWhenDone', $DynamicParameter2)
        $ParameterDictionary.Add('MgTenantID', $DynamicParameter3)
        $ParameterDictionary.Add('MgClientID', $DynamicParameter4)
    }

    return $ParameterDictionary
}

begin
{
    # Set Variables From Dynamic Parameters
    switch ($Service)
    {
        'MgGraph' {
            $MgPermissionType = $PSBoundParameters['MgPermissionType']
            $MgDisconnectWhenDone = $PSBoundParameters['MgDisconnectWhenDone']
            $MgTenantID = $PSBoundParameters['MgTenantID']
            $MgClientID = $PSBoundParameters['MgClientID']
        }
    }
}

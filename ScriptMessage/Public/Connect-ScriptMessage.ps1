Function Connect-ScriptMessage
{
    <#
        .LINK
        https://github.com/Sekers/ScriptMessage/wiki

        .SYNOPSIS
        Connects to the specified messaging service ahead of sending the message, if required.

        .DESCRIPTION
        Connects to the specified messaging service ahead of sending the message, if required.

        .PARAMETER Service
        Specify the messaging service to connect to.

        .PARAMETER ReturnConnectionInfo
        Returns connection information after performing function.

        .EXAMPLE
        Connect-ScriptMessage -Service MgGraph
        .EXAMPLE
        Connect-ScriptMessage -Service MgGraph -ReturnConnectionInfo
    #>

    [CmdletBinding()]
    param(
        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        #[ValidateSet("MgGraph")]
        #[string]$Service,
        [MessagingService]$Service,

        [parameter(
        Position=1,
        Mandatory=$false,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [switch]$ReturnConnectionInfo
    )

    # Set the necesasary configuration variables.
    $ScriptMessageConfig = Get-ScriptMessageConfig
    
    # Set the connection parameters.
    $ConnectionParameters = @{
        ServiceConfig = $ScriptMessageConfig.$Service
    }

    # Connect to the proper service.
    switch ($Service)
    {
        MgGraph {Connect-ScriptMessage_MGGraph @ConnectionParameters}
    }

    # Return the connection information, if requested.
    if ($ReturnConnectionInfo)
    {  
        return Get-ScriptMessageContext -Service $Service
    }
}

function Disconnect-ScriptMessage
{
    <#
        .LINK
        https://github.com/Sekers/ScriptMessage/wiki

        .SYNOPSIS
        Disconnects from the specified messaging service ahead of sending the message, if possible.

        .DESCRIPTION
        Disconnects from the specified messaging service ahead of sending the message, if possible.

        .PARAMETER Service
        Specify the messaging service to disconnect from.
        .PARAMETER ReturnConnectionInfo
        Returns connection information after performing function.

        .EXAMPLE
        Disconnect-ScriptMessage -Service MgGraph
        .EXAMPLE
        Disconnect-ScriptMessage -Service MgGraph -ReturnConnectionInfo
    #>

    [CmdletBinding()]
    param(
        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [MessagingService]$Service,

        [parameter(
        Position=1,
        Mandatory=$false,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [switch]$ReturnConnectionInfo
    )

    # Disconnect from the proper service.
    $ServiceDisconnectReturnInfo = switch ($Service)
    {
        MgGraph {Disconnect-ScriptMessage_MGGraph}
    }

    # Return the disconnection information, if requested.
    if ($ReturnConnectionInfo)
    {  
        # Create the disconnect info object to return.
        $ScriptMessageDisconnectReturnInfo = New-Object System.Object

        # Retrieve any common disconnection info across services.
        $CommonConnectionInfo = [pscustomobject]@{
            Service = $Service.ToString()
        }
        foreach ($infoItem in $($CommonConnectionInfo.PSObject.Properties))
        {
            $ScriptMessageDisconnectReturnInfo | Add-Member -MemberType NoteProperty -Name "$($infoItem.Name)" -Value $($infoItem.Value)
        }

        # Add in disconnection information.
        switch ($Service)
        {
            MgGraph {
                if ([string]::IsNullOrEmpty($ServiceDisconnectReturnInfo))
                {
                    break # Terminate the switch statement.
                }
                foreach ($infoItem in $($ServiceDisconnectReturnInfo.PSObject.Properties))
                {
                    $ScriptMessageDisconnectReturnInfo | Add-Member -MemberType NoteProperty -Name "$($infoItem.Name)" -Value $($infoItem.Value)
                }
            }
        }
    }

    if ($ReturnConnectionInfo)
    {
        return $ScriptMessageDisconnectReturnInfo
    }
}

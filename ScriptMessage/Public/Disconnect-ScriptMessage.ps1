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
        Disconnect-ScriptMessage -Service MicrosoftGraph
        .EXAMPLE
        Disconnect-ScriptMessage -Service MicrosoftGraph -ReturnConnectionInfo
    #>

    # Named parameters only. With no Position declared anywhere, PowerShell would otherwise make -Service positional.
    [CmdletBinding(PositionalBinding = $false)]
    param(
        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [MessagingService]$Service,

        [parameter(
        Mandatory=$false)]
        [switch]$ReturnConnectionInfo
    )

    # Disconnect from the proper service.
    $Registration = Get-ScriptMessageServiceRegistration -Service $Service
    $ServiceDisconnectReturnInfo = & $Registration.DisconnectFunction

    # Drop this service's cached context. 'Get-ScriptMessageContext -ReturnCachedContext' reads that cache, so
    # leaving an entry behind would keep reporting the connection this call just ended. Removing a name the
    # cache does not hold does nothing, which is the case when nothing has asked for the context yet.
    $script:ScriptMessageCachedServiceContext.PSObject.Properties.Remove([string]$Service)

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

        # Add in whatever the service returned. A service that reports nothing back leaves the object holding
        # only the common information above.
        if (-not [string]::IsNullOrEmpty($ServiceDisconnectReturnInfo))
        {
            foreach ($infoItem in $($ServiceDisconnectReturnInfo.PSObject.Properties))
            {
                $ScriptMessageDisconnectReturnInfo | Add-Member -MemberType NoteProperty -Name "$($infoItem.Name)" -Value $($infoItem.Value)
            }
        }
    }

    if ($ReturnConnectionInfo)
    {
        return $ScriptMessageDisconnectReturnInfo
    }
}

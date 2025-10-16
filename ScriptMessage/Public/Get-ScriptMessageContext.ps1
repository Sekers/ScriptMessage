function Get-ScriptMessageContext
{
    <#
        .LINK
        https://github.com/Sekers/ScriptMessage/wiki
        
        .SYNOPSIS
        Retrieve the session context information for the specified messaging service.

        .DESCRIPTION
        Retrieve the session context information for the specified messaging service.

        .PARAMETER Service
        Specify the messaging service to retrieve your current session details from.

        .PARAMETER ReturnCachedContext
        Returns the cached context information if it exists to reduce API calls.

        .EXAMPLE
        Get-ScriptMessageContext -Service MicrosoftGraph

        .EXAMPLE
        Get-ScriptMessageContext -Service MicrosoftGraph -ReturnCachedContext
    #>

    [CmdletBinding()]
    param(
        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [MessagingService]$Service,

        [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [Switch]$ReturnCachedContext
    )
    
    # Create the context object to return.
    $ScriptMessageContext = New-Object System.Object

    # Enable a force refresh if no data exists for the specified service or if $ReturnCachedContext is not present.
    if (($null -eq $ScriptMessage_Global_CachedServiceContext.$Service) -or (-not $ReturnCachedContext.IsPresent))
    {
        $RefreshContext = $true
    }

    if ($RefreshContext)
    {
        # Retrieve any common connection info across services.
        $CommonConnectionInfo = [pscustomobject]@{
            Service = $Service.ToString()
        }
        foreach ($infoItem in $($CommonConnectionInfo.PSObject.Properties))
        {
            $ScriptMessageContext | Add-Member -MemberType NoteProperty -Name "$($infoItem.Name)" -Value $($infoItem.Value)
        }

        # Retrieve connection information.
        switch ($Service)
        {
            MicrosoftGraph {
                $MgContext = Get-MgContext
                if ([string]::IsNullOrEmpty($MgContext))
                {
                    break # Terminate the switch statement.
                }
                foreach ($infoItem in $($MgContext.PSObject.Properties))
                {
                    $ScriptMessageContext | Add-Member -MemberType NoteProperty -Name "$($infoItem.Name)" -Value $($infoItem.Value)
                }
            }
        }

        # Update cached context data for the specified service.
        $ScriptMessage_Global_CachedServiceContext | Add-Member -MemberType NoteProperty -Name $Service -Value $ScriptMessageContext -Force # Force allows overwriting existing members.
    }
    else
    {
        $ScriptMessageContext = $ScriptMessage_Global_CachedServiceContext.$Service
    }

    return $ScriptMessageContext
}
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

        .EXAMPLE
        Get-ScriptMessageContext -Service MgGraph
    #>

    [CmdletBinding()]
    param(
        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [MessagingService]$Service
    )
    
    begin
    {
        # Create the context object to return.
        $ScriptMessageContext = New-Object System.Object
    }

    process
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
            MgGraph {
                $MgContext = Get-MgContext
                if ([string]::IsNullOrEmpty($MgContext))
                {
                    return $null
                }
                foreach ($infoItem in $($MgContext.PSObject.Properties))
                {
                    $ScriptMessageContext | Add-Member -MemberType NoteProperty -Name "$($infoItem.Name)" -Value $($infoItem.Value)
                }
            }
        }
    }
    
    end
    {
        return $ScriptMessageContext
    }
}
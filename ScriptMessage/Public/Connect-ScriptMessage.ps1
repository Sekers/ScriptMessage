function Connect-ScriptMessage
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
        .PARAMETER ServiceConfig
        Optional. The messaging service's section of the configuration file, for a caller that has already read it.
        If not provided, the configuration file is read.
        .PARAMETER ReturnConnectionInfo
        Returns connection information after performing function.

        .EXAMPLE
        Connect-ScriptMessage -Service MicrosoftGraph
        .EXAMPLE
        Connect-ScriptMessage -Service MicrosoftGraph -ReturnConnectionInfo
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
        ValueFromPipelineByPropertyName = $true)]
        [pscustomobject]$ServiceConfig,

        [parameter(
        Position=1,
        Mandatory=$false,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [switch]$ReturnConnectionInfo
    )

    # Set the connection parameters. A caller that has already read the configuration file passes the service's
    # section, which is what keeps one send to a single read of the file.
    if (-not $PSBoundParameters.ContainsKey('ServiceConfig'))
    {
        $ServiceConfig = Get-ScriptMessageConfig -Service $Service
    }
    $ConnectionParameters = @{
        ServiceConfig = $ServiceConfig
    }

    # Connect to the proper service. Each service checks for the modules its allowed message types need.
    $Registration = Get-ScriptMessageServiceRegistration -Service $Service
    & $Registration.ConnectFunction @ConnectionParameters
    
    # Return the connection information, if requested.
    if ($ReturnConnectionInfo)
    {  
        return Get-ScriptMessageContext -Service $Service
    }
}

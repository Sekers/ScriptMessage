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

        [parameter(
        Position=1,
        Mandatory=$false,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [switch]$ReturnConnectionInfo
    )

    # Set the connection parameters. The file's path comes with the settings, so the service can resolve a relative
    # path in them from the file's folder.
    $ServiceConfig = Get-ScriptMessageConfig -Service $Service -ReturnConfigFilePath
    $ConnectionParameters = @{
        ServiceConfig  = $ServiceConfig
        ConfigFilePath = $ServiceConfig.ConfigFilePath
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

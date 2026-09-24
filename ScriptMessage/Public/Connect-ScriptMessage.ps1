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
        Settings passed here do not carry the configuration file's location, so a relative path in them is resolved from PowerShell's current location rather than from the configuration file's folder.
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

    # Set the connection parameters. A caller that has already read the configuration file can pass the service's
    # section, so the file is not read again. Only settings read here come with the file's path, which the service
    # uses to resolve a relative path in them.
    $ConfigFilePath = $null
    if (-not $PSBoundParameters.ContainsKey('ServiceConfig'))
    {
        $ServiceConfig = Get-ScriptMessageConfig -Service $Service -ReturnConfigFilePath
        $ConfigFilePath = $ServiceConfig.ConfigFilePath
    }
    $ConnectionParameters = @{
        ServiceConfig  = $ServiceConfig
        ConfigFilePath = $ConfigFilePath
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

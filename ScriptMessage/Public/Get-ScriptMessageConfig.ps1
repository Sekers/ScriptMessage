function Get-ScriptMessageConfig
{
    <#
        .LINK
        https://github.com/Sekers/ScriptMessage/wiki

        .SYNOPSIS
        Get the configuration and secrets to connect to the messaging service(s).

        .DESCRIPTION
       Get the configuration and secrets to connect to the messaging service(s).

        .PARAMETER ConfigPath
        Optional. If not provided, the function will use the path used in the current session (if set).

        .EXAMPLE
        Get-ScriptMessageConfig
        .EXAMPLE
        Get-ScriptMessageConfig -Path '.\Config\config_scriptmessage.json'
    #>

    [CmdletBinding()]
    param(
        [Parameter(
        Position=0,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [string]$Path = $ScriptMessage_Global_ConfigFilePath # If not entered will see if it can pull path from this variable.
    )
    
    # Make Sure Requested Path Isn't Null or Empty (better to catch it here than validating on the parameter of this function)
    if ([string]::IsNullOrEmpty($Path))
    {
        throw "`'`$ScriptMessage_Global_ConfigFilePath`' is not specified. Don't forget to first use the `'Set-ScriptMessageConfigFilePath`' cmdlet!"
    }

    # Get Config and Secrets
    try
    {
        $ScriptMessageConfig = Get-Content -Path "$Path" -ErrorAction 'Stop' | ConvertFrom-Json
        return $ScriptMessageConfig 
    }
    catch
    {
        throw "Can't find the JSON configuration file. Use 'Set-ScriptMessageConfigFilePath' to create one."
    }
}
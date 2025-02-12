function Set-ScriptMessageConfigFilePath
{
    <#
        .LINK
        https://github.com/Sekers/ScriptMessage/wiki
        
        .SYNOPSIS
        Set the path to your ScriptMessage configuration file.

        .DESCRIPTION
        Set the path to your ScriptMessage configuration file. The configuration holds general settings for the ScriptMessage module
        to use, as well as the connection information for the messaging service(s) you are using.

        .PARAMETER Service
        Specify the path to where your configuration file is located.

        .EXAMPLE
        Set-ScriptMessageConfigFilePath -Path "$PSScriptRoot\Config\config_scriptmessage.json"
    #>
    
    [CmdletBinding()]
    param (
        [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [string]$Path
    )
   
    begin {}

    process
    {
        New-Variable -Name 'ScriptMessage_Global_ConfigFilePath' -Value $Path -Scope Global -Force
    }

    end {}
}
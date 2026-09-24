function Get-ScriptMessageConfig
{
    <#
        .LINK
        https://github.com/Sekers/ScriptMessage/wiki

        .SYNOPSIS
        Get the configuration and secrets to connect to the messaging service(s).

        .DESCRIPTION
        Get the configuration and secrets to connect to the messaging service(s).

        .PARAMETER Path
        Optional. If not provided, the function will use the path used in the current session (if set).
        .PARAMETER Service
        Optional. Return only the info related to a specific service.
        .PARAMETER ReturnConfigFilePath
        Adds a ConfigFilePath property to the result, holding the full path of the configuration file that was read.

        .EXAMPLE
        Get-ScriptMessageConfig
        .EXAMPLE
        Get-ScriptMessageConfig -Path '.\Config\config_scriptmessage.json'
        .EXAMPLE
        Get-ScriptMessageConfig -Service MicrosoftGraph
        .EXAMPLE
        Get-ScriptMessageConfig -Service MicrosoftGraph -ReturnConfigFilePath
    #>

    [CmdletBinding()]
    param(
        [Parameter(
        Position=0,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [string]$Path = $script:ScriptMessageConfigFilePath, # The path set by Set-ScriptMessageConfigFilePath.

        [Parameter(
        Position=1,
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [MessagingService]$Service,

        [Parameter(
        Mandatory = $false)]
        [switch]$ReturnConfigFilePath
    )

    # Make Sure Requested Path Isn't Null or Empty (better to catch it here than validating on the parameter of this function)
    if ([string]::IsNullOrEmpty($Path))
    {
        throw "No ScriptMessage configuration file path is set. Run `'Set-ScriptMessageConfigFilePath`' first."
    }

    # Read the configuration file.
    try
    {
        # .NET resolves a relative path against the process directory, not PowerShell's current location.
        $FullPath = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($Path)

        # A byte order mark sets the encoding when the file has one. Otherwise the file is read as UTF-8, or in the
        # system's legacy code page when it is not valid UTF-8. See section 5 of
        # Research_Notes/File-Encoding-And-Line-Endings.md.
        $Reader = [System.IO.StreamReader]::new($FullPath, [System.Text.UTF8Encoding]::new($false, $true), $true)
        try
        {
            $ConfigText = $Reader.ReadToEnd()
        }
        catch [System.Text.DecoderFallbackException]
        {
            $ConfigText = [System.IO.File]::ReadAllText($FullPath, [System.Text.Encoding]::GetEncoding(0))
        }
        finally
        {
            $Reader.Dispose()
        }
    }
    catch [System.IO.FileNotFoundException], [System.IO.DirectoryNotFoundException], [System.Management.Automation.DriveNotFoundException]
    {
        throw "Can't find the JSON configuration file. $($_.Exception.GetBaseException().Message)"
    }
    catch
    {
        throw "Can't read the JSON configuration file. $($_.Exception.GetBaseException().Message)"
    }

    # Parse it.
    try
    {
        $ScriptMessageConfig = $ConfigText | ConvertFrom-Json -ErrorAction Stop
    }
    catch
    {
        # Windows PowerShell 5.1 ends many of these messages with the entire text it was parsing, which here is the
        # configuration file and its secrets, so keep only the reason in front of it. See section 11 of
        # Research_Notes/PowerShell-Language-Behavior.md.
        $Reason = $_.Exception.Message -replace '(?s) \(\d+\): .*$', ''
        throw "The configuration file '$FullPath' is not valid JSON. $Reason"
    }

    if (-not ($null -eq $Service))
    {
        $Result = $ScriptMessageConfig.$Service
    }
    else
    {
        $Result = $ScriptMessageConfig
    }

    # No -Force, so a configuration file that has its own 'ConfigFilePath' setting fails here rather than having it
    # silently replaced. A missing service section leaves nothing to add the path to.
    if ($ReturnConfigFilePath -and ($null -ne $Result))
    {
        $Result | Add-Member -MemberType NoteProperty -Name 'ConfigFilePath' -Value $FullPath
    }

    return $Result
}
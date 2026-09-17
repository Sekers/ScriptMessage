function Register-ScriptMessageService
{
    <#
        Called once by each file in Services/ as it is dot-sourced, so the public cmdlets can dispatch to a
        service without naming it. Handlers are recorded as function NAMES, not as command objects or script
        blocks: the name is resolved at call time, which is what lets Pester's mocks shadow a handler.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        # The generic, unprefixed configuration settings and parameters this service honors. Mandatory so that
        # every service has to state its answer: a caller who supplies one a service does not honor is told,
        # rather than having the value dropped without a word. A service's own prefixed settings are its
        # business and are not listed here.
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$SupportedSetting,

        [Parameter(Mandatory = $true)]
        [string]$ConnectFunction,

        [Parameter(Mandatory = $true)]
        [string]$DisconnectFunction,

        [Parameter(Mandatory = $true)]
        [string]$SendFunction,

        [Parameter(Mandatory = $true)]
        [string]$GetContextFunction
    )

    # Every public cmdlet types its '-Service' parameter as [MessagingService], so a service registered under a
    # name outside that enum could never be reached. Fail at import rather than leaving it silently unusable.
    if ($Name -notin [Enum]::GetNames([MessagingService]))
    {
        throw "Cannot register the messaging service `'$Name`': it is not a member of the MessagingService enum. Add it to the enum in ScriptMessage.psm1 first."
    }

    $script:ScriptMessageServiceTable[$Name] = [PSCustomObject]@{
        Name               = $Name
        SupportedSetting   = $SupportedSetting
        ConnectFunction    = $ConnectFunction
        DisconnectFunction = $DisconnectFunction
        SendFunction       = $SendFunction
        GetContextFunction = $GetContextFunction
    }
}

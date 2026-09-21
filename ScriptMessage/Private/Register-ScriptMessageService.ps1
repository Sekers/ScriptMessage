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
        ConnectFunction    = $ConnectFunction
        DisconnectFunction = $DisconnectFunction
        SendFunction       = $SendFunction
        GetContextFunction = $GetContextFunction
    }
}

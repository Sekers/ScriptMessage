function Get-ScriptMessageServiceRegistration
{
    <#
        Returns the service table entry a public cmdlet needs in order to reach a service's handlers. Takes the
        service as a string so that no caller has to hold the [MessagingService] type.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Service
    )

    $Registration = $script:ScriptMessageServiceTable[$Service]

    if ($null -eq $Registration)
    {
        $Registered = $script:ScriptMessageServiceTable.Keys -join ', '
        throw "`'$Service`' is not a registered messaging service. Registered services: $Registered."
    }

    return $Registration
}

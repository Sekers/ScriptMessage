function Get-ScriptMessageBooleanSetting
{
    [CmdletBinding()]
    param(
        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [AllowNull()]
        [Object]$Value,

        [Parameter(
        Mandatory = $true,
        ValueFromPipelineByPropertyName = $true)]
        [string]$Name
    )

    # An unset setting stays unset. The caller decides what that means for its own setting.
    if ($null -eq $Value)
    {
        return $null
    }

    # A JSON boolean arrives as a [bool]. A quoted "true" or "false" arrives as a string, and casting a string
    # to [bool] makes every non-empty one $true, so a quoted "false" would mean its opposite. Reject anything
    # that is not already a boolean rather than guessing at it.
    if ($Value -isnot [bool])
    {
        throw "The `'$Name`' setting in the ScriptMessage configuration file must be a boolean written as true or false without quotes. Found the $($Value.GetType().Name) value `'$Value`'."
    }

    return $Value
}

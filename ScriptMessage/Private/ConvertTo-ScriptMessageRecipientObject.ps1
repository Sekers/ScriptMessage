function ConvertTo-ScriptMessageRecipientObject
{
    [CmdletBinding()]
    param(
        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyCollection()]
        [AllowNull()]
        [pscustomobject]$Recipient
    )

    if (([string]::IsNullOrEmpty($Recipient)) -and ($Recipient.Count -lt 1))
    {
        return $null
    }

    [array]$ScriptMessageRecipientObject = foreach ($recipientItem in $Recipient)
    {
        # Check if 'recipientItem' is string (email address, etc.). If it is, turn into a PSobject.
        if ($recipientItem.GetType().Name -eq 'String')
        {
            [PSCustomObject]@{
                AddressObj = $recipientItem # Don't use 'Address' because it can conflict with the 'Address()' method.
            }
        }
        else # Return item as properly formatted PSObject that includes the 'Name' property.
        {
            if ([string]::IsNullOrEmpty($recipientItem.Name))
            {
                [PSCustomObject]@{
                    AddressObj = $recipientItem.Address # Don't use 'Address' because it can conflict with the 'Address()' method.
                }
            }
            else
            {
                [PSCustomObject]@{
                    Name = $recipientItem.Name
                    AddressObj = $recipientItem.Address # Don't use 'Address' because it can conflict with the 'Address()' method.
                }
            }
        }
    }
    
    return $ScriptMessageRecipientObject
}
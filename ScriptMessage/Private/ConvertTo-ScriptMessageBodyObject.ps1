function ConvertTo-ScriptMessageBodyObject
{
    [CmdletBinding()]
    param(
        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyCollection()]
        [AllowNull()]
        [pscustomobject]$Body
    )

    begin
    {
        if ([string]::IsNullOrEmpty($Body))
        {
            return $null
        }
    }

    process
    {
        # Check if 'Body' is string. If it is, turn into a PSobject.
        if ($Body.GetType().Name -eq 'String')
        {
            $ScriptMessageBodyObject = [PSCustomObject]@{
                ContentType = 'Text'
                Content = $Body
            }
        }
        else # Return item as properly formatted PSObject that includes the 'ContentType' property.
        {
            $ScriptMessageBodyObject = [PSCustomObject]@{
                ContentType = $Body.ContentType
                Content = $Body.Content
            }
        }
    }

    end
    {
        return $ScriptMessageBodyObject
    }
}
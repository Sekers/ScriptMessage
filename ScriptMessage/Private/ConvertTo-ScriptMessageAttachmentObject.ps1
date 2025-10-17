Function ConvertTo-ScriptMessageAttachmentObject
{
    [CmdletBinding()]
    param(
        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [AllowNull()]
        [array]$Attachment
    )

    if ([string]::IsNullOrEmpty($Attachment))
    {
        return $null
    }
    
    [array]$ScriptMessageAttachment = foreach ($currentAttachment in $Attachment)
    {       
        switch ($currentAttachment.GetType().Name)
        {
            'Hashtable' { # If direct file content is supplied.
                $AttachmentType = 'Content'
                if (($currentAttachment.ContainsKey('Name')) -and $currentAttachment.ContainsKey('Content'))
                {
                    [PSCustomObject]$ScriptMessageAttachmentItem = @{
                        Name        = $currentAttachment.Name
                        Content     = $currentAttachment.Content
                    }
                }
                else
                {
                    throw "The attachment hashtable object is improperly formatted. The hashtable requires the keys of `'Name`' and `'Content`'"
                }
            }
            'String' { # If a directory or file path is supplied.
                if (-not (Test-Path -Path $currentAttachment))
                {
                    throw 'Invalid path to attachment directory or file.'
                }

                switch ((Get-Item -Path $currentAttachment).GetType().Name)
                {
                    'FileInfo'{
                        $AttachmentType = 'FilePath'
                        $FileInfo = Get-Item -Path $currentAttachment
                        [PSCustomObject]$ScriptMessageAttachmentItem = @{
                            Name        = $FileInfo.Name
                            Content     = [System.IO.File]::ReadAllBytes($FileInfo.FullName)
                        }   
                    }
                    'DirectoryInfo' {
                        $AttachmentType = 'DirectoryPath'
                        $DirectoryContent = Get-ChildItem $currentAttachment -File -Recurse
                        [PSCustomObject]$ScriptMessageAttachmentItem = foreach ($file in $DirectoryContent)
                        {
                            @{
                                Name        = $file.Name
                                Content     = [System.IO.File]::ReadAllBytes($file.FullName)
                            }   
                        }
                    }
                    Default {throw 'Unexpected attachment object type.'}
                }
            }
            Default {throw 'Unexpected attachment object type.'}
        }
    
        $ScriptMessageAttachmentItem
    }            
    
    return $ScriptMessageAttachment
}

function Send-ScriptMessage
{
    <#
    .LINK
    https://github.com/Sekers/ScriptMessage/wiki

    .SYNOPSIS
    Sends a message using the specified messaging service.

    .DESCRIPTION
    Sends a message using the specified messaging service.

    Note: If necessary, you will be asked to authenticate to the messaging service.

    .PARAMETER Service
    Specify the messaging service(s) to send the message from. You can specify more than one service to send the same message for redundancy or other purposes.
    .PARAMETER Type
    Optionally specify the type(s) of message(s) (mail, chat, etc.) to use when sending the message. Defaults to the 'Mail' service type. You can specify more than one message type to send the same message for redundancy or other purposes.
    .PARAMETER From
    The messaging address you are sending from. Alternatively, provide an object with the 'Address' property value set to the messaging address and, optionally, include a 'Name' property and corresponding value.
    .PARAMETER ReplyTo
    The messaging address(es) you want recipients to reply to. Alternatively, provide an object (or array of objects) with the 'Address' property value set to the messaging address you want recipients to reply to and, optionally, include a 'Name' property and corresponding value.
    .PARAMETER To
    An array of addresses to send the message to. Alternatively, provide an object (or array of objects) with the 'Address' property value set to the messaging address you want to send to and, optionally, include a 'Name' property and corresponding value.
    Must have at least one of 'To', 'CC', or 'BCC' set, depending on the messaging service used.
    .PARAMETER CC
    An array of addresses to carbon copy (CC) the message to. Alternatively, provide an object (or array of objects) with the 'Address' property value set to the messaging address you want to send to and, optionally, include a 'Name' property and corresponding value.
    Must have at least one of 'To', 'CC', or 'BCC' set, depending on the messaging service used.
    .PARAMETER BCC
    An array of addresses to blind carbon copy (BCC) the message to. Alternatively, provide an object (or array of objects) with the 'Address' property value set to the messaging address you want to send to and, optionally, include a 'Name' property and corresponding value.
    Must have at least one of 'To', 'CC', or 'BCC' set, depending on the messaging service used.
    .PARAMETER SaveToSentItems
    Use this parameter to ask the messaging service to save the sent message to a 'Sent Items' location, if supported by the service.
    Defaults to '$true'.
    .PARAMETER Subject
    Specify the message subject, if supported by the messaging service.
    .PARAMETER Body
    An object with the 'Content' property value set to message you want to send. Optionally, include a 'ContentType' property and corresponding value ('Text' or 'HTML').
    'ContentType' defaults to 'Text'.
    .PARAMETER Attachment
    Specify any file attachments, if supported by the messaging service.
    You can submit an array of any of the following (you can mix types in the array):
        - Content (Hashtable). Submit with a 'Name' key value set to the filename you want and a 'Content' key value as the data (encoded as a byte stream).
        - File path (String)
        - Directory path (String)
    .PARAMETER SenderId
    Specify the account used to send the message request. This might be different than the 'From' parameter in the case of "Send As', "Send on Behalf", delegated mailboxes, etc.
    If not specified, defaults to the address inside of the 'From' parameter.

    .EXAMPLE 
    $MessageArguments = @{
        From = 'jdoe@domain.com'
        To = 'bmayes@domain.com'
        CC = @()
        Subject = "Test Message"
        Body = @{
            Content = "This is a test message.`n`nThank you!"
        }
    }

    Send-ScriptMessage -Service MgGraph -Type 'Mail' @MessageArguments
    .EXAMPLE 
    $MessageArguments = @{
        From = 'jdoe@domain.com'
        To = @('bmayes@domain.com')
        CC = @( "hcoonly@domain.com", "plittle@domain.com")
        Subject = "Test Message"
        Body = @{
            Content = "This is a test message.`n`nThank you!"
        }
    }

    Send-ScriptMessage -Service -Type 'Mail', 'Chat' MgGraph @MessageArguments
    .EXAMPLE
    $MessageArguments = @{
        From = @{
            Name = 'John Doe'
            Address = 'jdoe@domain.com'
        }
        ReplyTo= @{
            Name = "Lisa Maloney"
            Address = "lmaloney@domain.com"
        }
        To = @('bmayes@domain.com')
        CC = @( "hcoonly@domain.com", "plittle@domain.com")
        SaveToSentItems = $true
        Subject = "Test Message"
        Body = @{
            ContentType = 'Text'
            Content = "This is a test message.`n`nThank you!"
        }
        Attachment = @('C:\StuffToSend\', 'C:\Documents\AnotherFile.pdf')
        SenderId = 'senderaccount@domain.com'
    }

    Send-ScriptMessage -Service MgGraph @MessageArguments
    .EXAMPLE
    # Attachments From Variable - Option 1: PS Desktop or Core
    $Content1 = [System.IO.File]::ReadAllBytes('C:\Users\John\Downloads\MyPDF.pdf')
    $Content2 = [System.IO.File]::ReadAllBytes('C:\Users\John\Downloads\AMovie.mp4')

    # Attachments From Variable - Option 2: PS Desktop Only
    # $Content1 = Get-Content -Encoding Byte -Raw -Path 'C:\Users\John\Downloads\MyPDF.pdf'
    # $Content2 = Get-Content -Encoding Byte -Raw -Path 'C:\Users\John\Downloads\AMovie.mp4'

    # Attachments From Variable - Option 3: PS Core Only
    # $Content1 = Get-Content -AsByteStream -Raw -Path 'C:\Users\John\Downloads\MyPDF.pdf'
    # $Content2 = Get-Content -AsByteStream -Raw -Path 'C:\Users\John\Downloads\AMovie.mp4'

    $Attachment = @(
    @{
            Name    = 'MyPDF.pdf'
            Content = $Content1
        },
        @{
            Name    = 'funnyballgame.mp4'
            Content = $Content2
        }
    }

    $MessageArguments = @{
        From = @{
            Address = 'jdoe@domain.com'
        }
        To = @('bmayes@domain.com')
        Subject = "Test Message"
        Body = @{
            Content = "This is a test message.`n`nThank you!"
        }
        Attachment = $Attachment
    }
#>

    [CmdletBinding()]
    param(
        [Parameter(
        ParameterSetName = 'ServiceAndTypeSeparate',
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [MessagingService[]]$Service,

        [Parameter(
        ParameterSetName = 'ServiceAndTypeSeparate',
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [MessageType[]]$Type = 'Mail',

        [Parameter(
        ParameterSetName = 'ServiceAndTypeCombined',
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [MessageServiceType[]]$ServiceType,

        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [pscustomobject]$From,

        [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [pscustomobject]$ReplyTo,

        [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [pscustomobject]$To,

        [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [pscustomobject]$CC,

        [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [pscustomobject]$BCC,

        [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [bool]$SaveToSentItems = $true,

        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [string]$Subject,

        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [pscustomobject]$Body,

        [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [array]$Attachment, # Array of Content (bytes), File paths, and/or Directory paths,

        [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [string]$SenderId,

        [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [ChatType]$ChatType,

        [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [ValidateSet($null, $true, $false)]
        [Object]$IncludeBCCInGroupChat # Is an object so it can be set to $null
    )

    begin
    {
        # Set the necessary configuration variables.
        $ScriptMessageConfig = Get-ScriptMessageConfig
    }

    process
    {
        # Make sure that at least one of, To, CC, or BCC is provided.
        if ([string]::IsNullOrEmpty($To) -and [string]::IsNullOrEmpty($CC) -and [string]::IsNullOrEmpty($BCC))
        {
            throw 'Please provide at least one parameter value for any of the following: To, CC, or BCC'
        }

        # Convert recipient types into properly formatted PSObject.
        $From = ConvertTo-ScriptMessageRecipientObject -Recipient $From # Note that From is NOT an array. There should only be one.
        [array]$ReplyTo = ConvertTo-ScriptMessageRecipientObject -Recipient $ReplyTo
        [array]$To = ConvertTo-ScriptMessageRecipientObject -Recipient $To
        [array]$CC = ConvertTo-ScriptMessageRecipientObject -Recipient $CC
        [array]$BCC = ConvertTo-ScriptMessageRecipientObject -Recipient $BCC

        # Convert body into properly formatted PSObject.
        $Body = ConvertTo-ScriptMessageBodyObject -Body $Body

        # Convert attachments into properly formatted PSObject.
        $Attachment = ConvertTo-ScriptMessageAttachmentObject -Attachment $Attachment

        if ($null -ne $Service) # If ServiceAndTypeSeparate
        {
            # Remove message service & message type duplicates.
            $Service = $Service | Select-Object -Unique
            $Type = $Type | Select-Object -Unique

            # Create the ServiceType class object\hash.
            [MessageServiceType]$ServiceType = @{
                Service = $Service
                Type    = $Type
            }
        }

        foreach ($serviceTypeObj in $ServiceType)
        {
            # Set the connection parameters.
            $ConnectionParameters = @{
                ServiceConfig = $ScriptMessageConfig.$($serviceTypeObj.Service)
            }

            # Set default values if not specified by a parameter.
            if (-not $ChatType)
            {
                [ChatType]$ChatType = $ConnectionParameters.ServiceConfig.ChatType
            }
            if (-not $IncludeBCCInGroupChat)
            {
                [bool]$IncludeBCCInGroupChat = $ConnectionParameters.ServiceConfig.IncludeBCCInGroupChat
            }
    
            # Connect to the messaging service, if necessary (e.g., API service).
            Connect-ScriptMessage -Service $($serviceTypeObj.Service) -ErrorAction Stop
    
            switch ($($serviceTypeObj.Service))
            {
                'MgGraph'   {
                    $SendMessageParameters = [ordered]@{
                        From = $From
                        ReplyTo = $ReplyTo
                        To = $To
                        CC = $CC
                        BCC = $BCC
                        SaveToSentItems = $SaveToSentItems
                        Subject = $Subject
                        Body = $Body
                        Attachment = $Attachment
                        SenderId = $SenderId
                        Type = $serviceTypeObj.Type
                        ChatType = $ChatType
                        IncludeBCCInGroupChat = $IncludeBCCInGroupChat
                    }
    
                    Send-ScriptMessage_MgGraph @SendMessageParameters
    
                    # Disconnect from Microsoft Graph API, if enabled in config.
                    if ($ConnectionParameters.ServiceConfig.MgDisconnectWhenDone)
                    {
                        $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
                    }
                }
                Default {throw "Invalid `'Service`' value."}
            }
        }
    }
    
    end {}
}
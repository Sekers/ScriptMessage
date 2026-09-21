
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
    .PARAMETER ServiceType
    Alternatively, you can specify both the messaging service(s) and message type(s) in a single parameter. You can specify more than one service/type combination to send the same message for redundancy or other purposes.
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
    .PARAMETER MailType
    Override the default 'MailType' specified in the configuration file for the messaging service being used. Options are 'OneOnOne' or 'Group'.
    'Group' sends one message to all of the To, CC, and BCC recipients. 'OneOnOne' sends each recipient a separate message with only that recipient in To, so no recipient can see who else received it. A recipient listed more than once, even across To, CC, and BCC, receives one message.
    Each message counts separately against the messaging service's sending limits.
    If neither this parameter nor the configuration file sets a mail type, 'Group' is used.
    .PARAMETER ChatType
    Override the default 'ChatType' specified in the configuration file for the messaging service being used. Options are 'OneOnOne' or 'Group'.
    .PARAMETER IncludeBCCInGroupChat
    Specify whether to include BCC recipients in a group chat message.

    .EXAMPLE 
    $MessageArguments = @{
        From = 'jdoe@example.com'
        To = 'bmayes@example.com'
        CC = @()
        Subject = "Test Message"
        Body = @{
            Content = "This is a test message.`n`nThank you!"
        }
    }

    Send-ScriptMessage -Service MicrosoftGraph -Type 'Mail' @MessageArguments
    .EXAMPLE 
    $MessageArguments = @{
        From = 'jdoe@example.com'
        To = @('bmayes@example.com')
        CC = @( "hcoonly@example.com", "plittle@example.com")
        Subject = "Test Message"
        Body = @{
            Content = "This is a test message.`n`nThank you!"
        }
    }

    Send-ScriptMessage -Service MicrosoftGraph -Type 'Mail', 'Chat' @MessageArguments
    .EXAMPLE
    $MessageArguments = @{
        From = @{
            Name = 'John Doe'
            Address = 'jdoe@example.com'
        }
        ReplyTo= @{
            Name = "Lisa Maloney"
            Address = "lmaloney@example.com"
        }
        To = @('bmayes@example.com')
        CC = @( "hcoonly@example.com", "plittle@example.com")
        SaveToSentItems = $true
        Subject = "Test Message"
        Body = @{
            ContentType = 'Text'
            Content = "This is a test message.`n`nThank you!"
        }
        Attachment = @('C:\StuffToSend\', 'C:\Documents\AnotherFile.pdf')
        SenderId = 'senderaccount@example.com'
    }

    Send-ScriptMessage -Service MicrosoftGraph @MessageArguments
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
    )

    $MessageArguments = @{
        From = @{
            Address = 'jdoe@example.com'
        }
        To = @('bmayes@example.com')
        Subject = "Test Message"
        Body = @{
            Content = "This is a test message.`n`nThank you!"
        }
        Attachment = $Attachment
    }

    Send-ScriptMessage -Service MicrosoftGraph @MessageArguments
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
        [MailType]$MailType,

        [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [ChatType]$ChatType,

        [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [Nullable[bool]]$IncludeBCCInGroupChat # Nullable so it can be set to $null, meaning "use the configured value"
    )

    # Set the necessary configuration variables.
    $ScriptMessageConfig = Get-ScriptMessageConfig

    # Convert recipient types into properly formatted PSObject.
    $From = ConvertTo-ScriptMessageRecipientObject -Recipient $From # Note that From is NOT an array. There should only be one.
    [array]$ReplyTo = ConvertTo-ScriptMessageRecipientObject -Recipient $ReplyTo
    [array]$To = ConvertTo-ScriptMessageRecipientObject -Recipient $To
    [array]$CC = ConvertTo-ScriptMessageRecipientObject -Recipient $CC
    [array]$BCC = ConvertTo-ScriptMessageRecipientObject -Recipient $BCC

    # Make sure that at least one To, CC, or BCC recipient has an address.
    $RecipientsWithAddress = @(@($To) + @($CC) + @($BCC) | Where-Object {($null -ne $_) -and (-not [string]::IsNullOrWhiteSpace($_.AddressObj))})
    if ($RecipientsWithAddress.Count -eq 0)
    {
        throw 'Please provide at least one parameter value for any of the following: To, CC, or BCC'
    }

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

    foreach ($serviceTypeObj in $ServiceType) # TODO: Catch errors once we have multiple services so if one fails the other(s) can still be processed.
    {
        # Look the service up first, so an unregistered one stops the call before any other work.
        $Registration = Get-ScriptMessageServiceRegistration -Service ([string]$serviceTypeObj.Service)

        # Take this service's section out of the configuration read above, using the registered name rather
        # than the parameter's value: looking a property up by a name held in an array comes back empty.
        $ServiceConfig = $ScriptMessageConfig.$($Registration.Name)
        if ($null -eq $ServiceConfig)
        {
            throw "The ScriptMessage configuration file has no `'$($Registration.Name)`' section. Add one for that messaging service, or send from a service the file configures."
        }

        # Tell the caller about any generic setting this service does not honor. Services are meant to be
        # interchangeable, so one that cannot act on a setting is an ordinary case rather than an error: the
        # rest of the send goes ahead, which matters most when a single call goes to several services and only
        # some of them honor it.
        foreach ($settingName in @('MailType', 'ChatType', 'IncludeBCCInGroupChat'))
        {
            if (-not $PSBoundParameters.ContainsKey($settingName))
            {
                continue
            }
            # A $null means "use the configured value", so the caller has expressed no preference to drop.
            if ($null -eq $PSBoundParameters[$settingName])
            {
                continue
            }
            if ($Registration.SupportedSetting -notcontains $settingName)
            {
                Write-Warning -Message "The `'$($Registration.Name)`' messaging service does not support the `'$settingName`' parameter, so it was ignored."
            }
        }

        # Set the connection parameters.
        $ConnectionParameters = @{
            ServiceConfig = $ServiceConfig
        }

        # Set default values for anything the caller did not specify. These test the bound parameters rather than
        # the values because truthiness cannot tell either parameter apart from an unsupplied one: 'OneOnOne' is
        # ChatType enum value 0, and $false is a valid 'IncludeBCCInGroupChat' choice. The results go into
        # separate variables because a parameter keeps its type constraint through later assignments, and
        # [ChatType] cannot hold the $null of a configuration file that has no 'ChatType' setting.
        if ($PSBoundParameters.ContainsKey('ChatType'))
        {
            $ResolvedChatType = $ChatType
        }
        else
        {
            $ResolvedChatType = $ConnectionParameters.ServiceConfig.ChatType
        }

        # 'IncludeBCCInGroupChat' accepts $null to mean "use the configured value", which is also what an
        # unsupplied [Nullable[bool]] holds. The configured value is an opt-in flag, on only when it equals
        # $true, so anything else, including quoted text, leaves BCC recipients out. Keep the setting on the
        # left: a [bool] cast, or $true on the left, reads the text "false" as $true.
        if ($null -ne $IncludeBCCInGroupChat)
        {
            [bool]$ResolvedIncludeBCCInGroupChat = $IncludeBCCInGroupChat
        }
        else
        {
            [bool]$ResolvedIncludeBCCInGroupChat = ($ConnectionParameters.ServiceConfig.IncludeBCCInGroupChat -eq $true)
        }

        # 'MailType' is resolved only for a send that includes Mail, so a configuration value that only Mail would
        # use cannot stop a Chat send. 'Group' applies when neither the parameter nor the configuration file sets
        # one, including the empty string the configuration template uses for a setting it leaves unset, so a
        # configuration file without the setting sends one message to every recipient.
        $ResolvedMailType = $null
        if ($serviceTypeObj.Type -contains [MessageType]::Mail)
        {
            $ConfiguredMailType = $ConnectionParameters.ServiceConfig.MailType
            if ($PSBoundParameters.ContainsKey('MailType'))
            {
                $ResolvedMailType = $MailType
            }
            elseif ([string]::IsNullOrWhiteSpace($ConfiguredMailType))
            {
                $ResolvedMailType = [MailType]::Group
            }
            elseif ($ConfiguredMailType -in [Enum]::GetNames([MailType]))
            {
                $ResolvedMailType = [MailType]$ConfiguredMailType
            }
            else
            {
                throw "The `'MailType`' setting in the ScriptMessage configuration file must be one of: $([Enum]::GetNames([MailType]) -join ', '). Found `'$ConfiguredMailType`'."
            }
        }

        # Connect to the messaging service, if necessary (e.g., API service).
        Connect-ScriptMessage -Service $($serviceTypeObj.Service) -ServiceConfig $ServiceConfig -ErrorAction Stop

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
            IncludeBCCInGroupChat = $ResolvedIncludeBCCInGroupChat
            ServiceConfig = $ServiceConfig
        }

        # [ChatType] cannot be bound to $null, nor to the empty string a configuration file uses for a setting
        # it leaves unset, so pass it only when there is a real value. A Mail-only send does not need one, and
        # the service reports a Chat message that has none.
        if (-not [string]::IsNullOrWhiteSpace($ResolvedChatType))
        {
            $SendMessageParameters['ChatType'] = $ResolvedChatType
        }

        if ($null -ne $ResolvedMailType)
        {
            $SendMessageParameters['MailType'] = $ResolvedMailType
        }

        & $Registration.SendFunction @SendMessageParameters
    }
}
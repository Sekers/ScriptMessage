Function ConvertTo-IMicrosoftGraphRecipient
{
    [CmdletBinding()]
    param(
        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyCollection()]
        [AllowNull()]
        [pscustomobject]$EmailAddress,

        [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [string]$Name
    )

    # Return Null If Provided Recipient is Empty
    if (([string]::IsNullOrEmpty($EmailAddress)) -and ([string]::IsNullOrEmpty($EmailAddress.AddressObj)))
    {
        return $null
    }

    # Loop through each of the recipient parameter array objects
    $IMicrosoftGraphRecipient = foreach ($address in $EmailAddress)
    {
        # Check if string (email address) or object/hashtable/etc. If not, separate out.
        if (-not ($address.GetType().Name -eq 'String'))
        {
            # Verify object contains 'Address' key or property.
            if ([string]::IsNullOrEmpty($address.AddressObj))
            {
                throw "Improperly formatted from, recipient, or reply to address."
            }

            # Set 'Name' & update 'Address' (do 'Name' 1st!)
            $Name = $address.Name
            $address = $address.AddressObj
        }

        if ([string]::IsNullOrEmpty($Name))
        {
            @{
                EmailAddress = @{Address = $address}
            }
        }
        else
        {
            @{
                EmailAddress = [ordered]@{
                    Name = $Name
                    Address = $address}
            }
        }
    }

    return $IMicrosoftGraphRecipient
}

function ConvertTo-IMicrosoftGraphItemBody
{
    [CmdletBinding()]
    param(
        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [string]$Content,

        [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [ValidateSet('Text','HTML')]
        [string]$ContentType = 'Text' # The MIME type. See https://learn.microsoft.com/en-us/exchange/mail-flow-best-practices/message-format-and-transmission
    )

    $IMicrosoftGraphItemBody =
    @{
        ContentType = $ContentType
        Content = $Content
    }
    return $IMicrosoftGraphItemBody
}

Function ConvertTo-IMicrosoftGraphAttachment
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

    [array]$IMicrosoftGraphAttachment = foreach ($currentAttachment in $Attachment)
    {       
        if (($currentAttachment.ContainsKey('Name')) -and $currentAttachment.ContainsKey('Content'))
        {
            $Attachment_ByteEncoded = [System.Convert]::ToBase64String($currentAttachment.Content)
            $IMicrosoftGraphAttachmentItem = @{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                name          = $currentAttachment.Name
                contentBytes  = $Attachment_ByteEncoded
            }
            $IMicrosoftGraphAttachmentItem
        }
        else
        {
            throw "The attachment hashtable object is improperly formatted. The hashtable requires the keys of `'Name`' and `'Contents`'"
        }
    }

    return $IMicrosoftGraphAttachment
}

Function ConvertTo-IMicrosoftGraphChatMessageAttachment
{
    [CmdletBinding()]
    param(
        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [AllowNull()]
        [array]$MgDriveItem
    )

    if ([string]::IsNullOrEmpty($MgDriveItem))
    {
        return $null
    }
    
    [array]$IMicrosoftGraphChatMessageAttachment = foreach ($currentAttachment in $MgDriveItem)
    {       
        if ($currentAttachment.ContainsKey('name') -and $currentAttachment.ContainsKey('webUrl'))
        {
            $IMicrosoftGraphChatMessageAttachmentItem = @{
                contentType = 'reference'
                contentUrl  = $currentAttachment.webUrl
                name        = $currentAttachment.name
            }
            $IMicrosoftGraphChatMessageAttachmentItem
        }
        else
        {
            throw "The attachment hashtable object is improperly formatted. The hashtable requires the keys of `'webUrl`' and `'name`'"
        }
    }

    return $IMicrosoftGraphChatMessageAttachment
}

function ConvertTo-IMicrosoftGraphConversationMember
{
    [CmdletBinding()]
    param(
        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyCollection()]
        [AllowNull()]
        [pscustomobject]$EmailAddress
    )

    # Return Null If Provided Recipient is Empty
    if ([string]::IsNullOrEmpty($EmailAddress))
    {
        return $null
    }

    # Loop through each of the recipient parameter array objects
    $IMicrosoftGraphRecipient = foreach ($address in $EmailAddress)
    {
        # Check if string (email address) or object/hashtable/etc. If not, separate out.
        if (-not ($address.GetType().Name -eq 'String'))
        {
            throw "Improperly formatted from or recipient address."
        }

        # Return IMicrosoftGraphConversationMember
        @{
            '@odata.type'     = "#microsoft.graph.aadUserConversationMember"
            roles             = @(
                "owner"
            )
            "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$address')"
        }
    }
    
    return $IMicrosoftGraphRecipient
}

function ConvertTo-IMicrosoftGraphDriveInvite
{
    [CmdletBinding()]
    param(
        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyCollection()]
        [AllowNull()]
        [pscustomobject]$EmailAddress
    )

    # Return Null If Provided Recipient is Empty
    if ([string]::IsNullOrEmpty($EmailAddress))
    {
        return $null
    }

    # Loop through each of the recipient parameter array objects
    [array]$IMicrosoftGraphDriveRecipient = foreach ($address in $EmailAddress)
    {
        # Check if string (email address) or object/hashtable/etc. If not, separate out.
        if (-not ($address.GetType().Name -eq 'String'))
        {
            throw "Improperly formatted recipient address."
        }

        # Return IMicrosoftGraphDriveRecipient
        @{
            email = $address
        }
    }

    $IMicrosoftGraphDriveInvite = @{
        recipients     = $IMicrosoftGraphDriveRecipient
        requireSignIn  = $true
        sendInvitation = $false
        roles          = @(
            "read"
        )
    }

    return $IMicrosoftGraphDriveInvite
}

function Connect-ScriptMessage_MicrosoftGraph
{
    [CmdletBinding()]
    param(
        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [pscustomobject]$ServiceConfig,

        # The full path of the configuration file the settings came from, or empty when they did not come from one.
        [Parameter(
        Mandatory = $false,
        ValueFromPipelineByPropertyName = $true)]
        [string]$ConfigFilePath
    )

    # Collect the connection settings from the service configuration.
    $MgPermissionType = $ServiceConfig.MgPermissionType
    $MgTenantID = $ServiceConfig.MgTenantID
    $MgClientID = $ServiceConfig.MgClientID

    # Check For MicrosoftGraph Modules
    # Don't import the entire 'Microsoft.Graph' module. Only import the needed sub-modules.
    $RequiredModules = [System.Collections.Generic.List[Object]]::new()

    # Required For All Graph Modules
    [string]$ModuleName = 'Microsoft.Graph.Authentication' # Used for Connect-MgGraph, Disconnect-MgGraph, & Get-MgContext. A required module for all Graph modules.
    Import-Module -Name $ModuleName -ErrorAction SilentlyContinue 
    $RequiredModules.Add($ModuleName)

    # If Mail is enabled (based on the config item 'AllowableMessageTypes' containing 'Mail'), check for the needed module.
    if ($ServiceConfig.AllowableMessageTypes -contains 'Mail')
    {
        [string]$ModuleName = 'Microsoft.Graph.Users.Actions' # Used for Send-MgUserMail.
        Import-Module -Name $ModuleName -ErrorAction SilentlyContinue 
        $RequiredModules.Add($ModuleName)
    }
    
    # If Chat is enabled (based on the config item 'AllowableMessageTypes' containing 'Chat'), check for the needed module.
    if ($ServiceConfig.AllowableMessageTypes -contains 'Chat')
    {
        [string]$ModuleName = 'Microsoft.Graph.Teams' # Used for New-MgChat & New-MgChatMessage.
        Import-Module -Name $ModuleName -ErrorAction SilentlyContinue 
        $RequiredModules.Add($ModuleName)
    }

    # If uploads are enabled (based on the config item 'MgDelegatedPermission_RequestFilesReadWritePermission' being set to true), check for the needed module.
    if (($ServiceConfig.MgDelegatedPermission_RequestFilesReadWritePermission -eq $true) -and ($MgPermissionType -eq 'Delegated'))
    {
        [string]$ModuleName = 'Microsoft.Graph.Files' # Used for Get-MgUserDrive
        Import-Module -Name $ModuleName -ErrorAction SilentlyContinue 
        $RequiredModules.Add($ModuleName)
    }

    # Check for missing Graph modules.
    $ImportedModules = Get-Module
    $MissingModules = Compare-Object -ReferenceObject $RequiredModules -DifferenceObject $ImportedModules | Where-Object { $_.SideIndicator -eq '<=' } | Select-Object -ExpandProperty InputObject
    if ($MissingModules.Count -gt 0)
    {
        Write-Error "Please first install the following sub-modules from https://www.powershellgallery.com/packages/Microsoft.Graph/: $($MissingModules -join ', ')"
        Return
    }

    # Connect to the Microsoft Graph API.      
    switch ($MgPermissionType)
    {
        Delegated {
            # E.g. Connect-MgGraph -Scopes "User.Read.All","Group.ReadWrite.All"
            # You can add additional permissions by repeating the Connect-MgGraph command with the new permission scopes.
            # View the current scopes under which the PowerShell SDK is (trying to) execute cmdlets: Get-MgContext | select -ExpandProperty Scopes
            # List all the scopes granted on the service principal object (you cn also do it via the Azure AD UI): Get-MgServicePrincipal -Filter "appId eq '14d82eec-204b-4c2f-b7e8-296a70dab67e'" | % { Get-MgServicePrincipalOauth2PermissionGrant -ServicePrincipalId $_.Id } | fl
            # Find Graph permission needed. More info on permissions: https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent)
            #    E.g., Find-MgGraphPermission -SearchString "Teams" -PermissionType Delegated
            #    E.g., Find-MgGraphPermission -SearchString "Teams" -PermissionType Application

            # The Microsoft Authentication Library (MSAL) currently specifies offline_access, openid, profile, and email by default in authorization and token requests.
            $MicrosoftGraphScopes = @(
                'email' # Allows the app to read your users' primary email address
                'offline_access' # With the Microsoft identity platform v2.0 endpoint, you specify the offline_access scope in the scope parameter to explicitly request a refresh token when using the OAuth 2.0 or OpenID Connect protocols.
                'openid' # Allows users to sign in to the app with their work or school accounts and allows the app to see basic user profile information.
                'profile' # Allows the app to see your users' basic profile (e.g., name, picture, user name, email address)
            )
            if ($ServiceConfig.AllowableMessageTypes -contains 'Mail')
            {
                $MicrosoftGraphScopes += @(
                    'Mail.Send' # With the Mail.Send permission, an app can send mail and save a copy to the user's Sent Items folder, even if the app isn't granted the Mail.ReadWrite or Mail.ReadWrite.Shared permission. 
                    # 'Mail.Send.Shared' # This scope doesn't seem to be needed for sending as or on behalf of another user. I wonder if being able to do so using just 'Mail.Send' is a bug... > https://learn.microsoft.com/en-us/graph/outlook-send-mail-from-other-user
                )
            }
            if ($ServiceConfig.AllowableMessageTypes -contains 'Chat')
            {
                $MicrosoftGraphScopes += @(
                    'Chat.Create' # Allows the app to create chats on behalf of the signed-in user.
                    'ChatMessage.Send' # Allows an app to send one-to-one and group chat messages in Microsoft Teams, on behalf of the signed-in user.
                )

                if ($ServiceConfig.MgDelegatedPermission_RequestChatReadPermission -eq $true)
                {
                    $MicrosoftGraphScopes += @(
                        'Chat.Read' # Allows an app to read 1 on 1 or group chats threads, on behalf of the signed-in user.
                    )
                }
                else {
                    $MicrosoftGraphScopes += @(
                        'Chat.ReadBasic' # Allows an app to read the members and descriptions of one-to-one and group chat threads, on behalf of the signed-in user.
                    )
                }
                if ($ServiceConfig.MgDelegatedPermission_RequestFilesReadWritePermission -eq $true) # TODO: Do I want to allow 'Files.ReadWrite' for delegated email as well?
                {
                    $MicrosoftGraphScopes += @(
                        'Files.ReadWrite' # Allows the app to read, create, update and delete the signed-in user's files.
                    )
                }
            }
            $null = Connect-MgGraph -Scopes $MicrosoftGraphScopes -TenantId $MgTenantID -ClientId $MgClientID
        }
        Application {
            [string]$MgApp_AuthenticationType = $ServiceConfig.MgApp_AuthenticationType
            Write-Verbose -Message "Microsoft Graph App Authentication Type: $MgApp_AuthenticationType"

            switch ($MgApp_AuthenticationType)
            {
                CertificateFile {
                    # Expand only environment variables, so the configuration file cannot run code. This uses
                    # [regex]::Replace because Windows PowerShell 5.1 turns a -replace scriptblock into text.
                    $MgApp_CertificatePath = [regex]::Replace([string]$ServiceConfig.MgApp_CertificatePath, '\$\{env:([^}]+)\}|\$env:(\w+)',
                        [System.Text.RegularExpressions.MatchEvaluator]{
                            param($Match)
                            [Environment]::GetEnvironmentVariable($Match.Groups[1].Value + $Match.Groups[2].Value)
                        })

                    # A relative path is resolved from the configuration file's folder. .NET does not count a
                    # PowerShell drive as absolute and PowerShell does not count a UNC path, so a path is relative
                    # only when neither does; '~' is the home folder. Settings that did not come from a file keep
                    # resolving from the current location.
                    $DriveName = $null
                    $IsRelativePath = -not ([string]::IsNullOrWhiteSpace($MgApp_CertificatePath) -or
                        [System.IO.Path]::IsPathRooted($MgApp_CertificatePath) -or
                        $PSCmdlet.SessionState.Path.IsPSAbsolute($MgApp_CertificatePath, [ref]$DriveName) -or
                        $MgApp_CertificatePath.StartsWith('~'))
                    if ($IsRelativePath -and -not [string]::IsNullOrEmpty($ConfigFilePath))
                    {
                        $ConfigFolder = [System.IO.Path]::GetDirectoryName($ConfigFilePath)
                        $PathInConfigFolder = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($ConfigFolder, $MgApp_CertificatePath))

                        # Deprecated: a file found only relative to the current location is still used, with a
                        # warning. Remove this fallback in the next major version.
                        if (-not (Test-Path -LiteralPath $PathInConfigFolder -PathType Leaf) -and (Test-Path -LiteralPath $MgApp_CertificatePath -PathType Leaf))
                        {
                            Write-Warning -Message "The MgApp_CertificatePath file '$MgApp_CertificatePath' was found in PowerShell's current location, not in the configuration file's folder '$ConfigFolder'. Finding a relative path from the current location is deprecated, and the next major version of ScriptMessage will look only in the configuration file's folder. Move the file there, or set MgApp_CertificatePath to the file's full path."
                        }
                        else
                        {
                            $MgApp_CertificatePath = $PathInConfigFolder
                        }
                    }

                    # Load the file with .NET, because Windows PowerShell 5.1's Get-PfxCertificate cannot take a
                    # password. .NET resolves a relative path against the process directory and knows nothing of
                    # PowerShell drives, so it gets the full path. EphemeralKeySet keeps the private key in memory;
                    # otherwise it is written to the user profile, and PowerShell 7 leaves that file behind. New-Object
                    # rather than ::new() lets the tests mock the load. See Research_Notes/Certificate-File-Behavior.md.
                    $MgApp_CertificatePath = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($MgApp_CertificatePath)
                    $CertificateType = 'System.Security.Cryptography.X509Certificates.X509Certificate2'
                    $KeyStorage = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet
                    try
                    {
                        # A certificate file without a password.
                        $MgApp_Certificate = New-Object -TypeName $CertificateType -ArgumentList $MgApp_CertificatePath, '', $KeyStorage
                    }
                    catch # Otherwise use the password from the configuration file.
                    {
                        if (-not (Test-Path -LiteralPath $MgApp_CertificatePath -PathType Leaf))
                        {
                            throw "The Microsoft Graph certificate file named by MgApp_CertificatePath does not exist: '$MgApp_CertificatePath'."
                        }

                        $MgApp_EncryptedCertificatePassword = $ServiceConfig.MgApp_EncryptedCertificatePassword
                        if ([string]::IsNullOrEmpty($MgApp_EncryptedCertificatePassword))
                        {
                            $NewMessage = "Cannot access Microsoft Graph .pfx private key certificate file and no password has been provided."
                            throw $NewMessage
                        }

                        [SecureString]$MgApp_EncryptedCertificateSecureString = $MgApp_EncryptedCertificatePassword | ConvertTo-SecureString # Can only be decrypted by the same AD account on the same computer.
                        $MgApp_Certificate = New-Object -TypeName $CertificateType -ArgumentList $MgApp_CertificatePath, $MgApp_EncryptedCertificateSecureString, $KeyStorage
                    }

                    $null = Connect-MgGraph -TenantId $MgTenantID -ClientId $MgClientID -Certificate $MgApp_Certificate
                }
                CertificateName {
                    $MgApp_CertificateName = $ServiceConfig.MgApp_CertificateName
                    $null = Connect-MgGraph -TenantId $MgTenantID -ClientId $MgClientID -CertificateName $MgApp_CertificateName
                }
                CertificateThumbprint {
                    $MgApp_CertificateThumbprint = $ServiceConfig.MgApp_CertificateThumbprint
                    $null = Connect-MgGraph -TenantId $MgTenantID -ClientId $MgClientID -CertificateThumbprint $MgApp_CertificateThumbprint
                }
                ClientSecret {
                    $MgApp_EncryptedSecret = $ServiceConfig.MgApp_EncryptedSecret
                    $ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $MgClientID, $($MgApp_EncryptedSecret | ConvertTo-SecureString)
                    $null = Connect-MgGraph -TenantId $MgTenantID -ClientSecretCredential $ClientSecretCredential
                }
                Default {throw "Invalid `'MgApp_AuthenticationType`' value."}
            }
        }
        Default {throw "Invalid `'MgPermissionType`' value."}
    }
}

function Disconnect-ScriptMessage_MicrosoftGraph
{
    return Disconnect-MgGraph
}

function Get-ScriptMessageContext_MicrosoftGraph
{
    # Returns nothing when there is no active connection. The caller adds the properties of whatever comes back
    # to the context object it returns, so the shape of this is Microsoft Graph's to decide.
    return Get-MgContext
}

function Send-ScriptMessage_MicrosoftGraph
{
    [CmdletBinding()]
    param(
        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [MessageType[]]$Type,

        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [pscustomobject]$From,

        [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
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
        [array]$Attachment, # Array of Content(bytes), File paths, and/or Directory paths

        [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [string]$SenderId,

        [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [MailType]$MailType = 'Group',

        [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [ChatType]$ChatType,

        [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [bool]$IncludeBCCInGroupChat,

        [Parameter(
        Mandatory = $false,
        ValueFromPipelineByPropertyName = $true)]
        [pscustomobject]$ServiceConfig
    )

    # Set the Service ID.
    # Keep this as a STRING and not the ENUM type since it's returned to the caller (functions will convert to [MessagingService] type as needed).
    [string]$ServiceId = 'MicrosoftGraph'

    # Get the Service Config, unless the caller already read it and passed this service's section.
    if (-not $PSBoundParameters.ContainsKey('ServiceConfig'))
    {
        $ServiceConfig = Get-ScriptMessageConfig -Service $ServiceId
    }

    # Send the message on each supported service specified.
    foreach ($typeItem in $Type)
    {
        # Reset Warnings
        $MgWarningMessages = @()
        $MgErrorMessages = @()

        switch ($typeItem)
        {
            Mail {
                try
                {
                    # Check if AllowableMessageTypes contains '$typeItem'.
                    if ($ServiceConfig.AllowableMessageTypes -notcontains $typeItem)
                    {
                        $NewMessage = "The ScriptMessage configuration for '$ServiceId' does not allow sending messages of type: $typeItem"
                        Write-Warning -Message $NewMessage
                        $MgWarningMessages += "$NewMessage"
                    }
                    else
                    {
                        # Convert Parameters to IMicrosoft*
                        $Message = @{}
                        $Message['From'] = ConvertTo-IMicrosoftGraphRecipient -EmailAddress $From
                        [array]$Message['ReplyTo'] = ConvertTo-IMicrosoftGraphRecipient -EmailAddress $ReplyTo
                        [array]$Message['To'] = ConvertTo-IMicrosoftGraphRecipient -EmailAddress $To
                        [array]$Message['CC'] = ConvertTo-IMicrosoftGraphRecipient -EmailAddress $CC
                        [array]$Message['BCC'] = ConvertTo-IMicrosoftGraphRecipient -EmailAddress $BCC
                        if (-not [string]::IsNullOrEmpty($Body.Content))
                        {
                            if ([string]::IsNullOrEmpty($Body.ContentType)) # Don't send 'ContentType' if not provided. It will default to 'Text'
                            {
                                [hashtable]$Message['Body'] = ConvertTo-IMicrosoftGraphItemBody -Content $Body.Content
                            }
                            else
                            {
                                [hashtable]$Message['Body'] = ConvertTo-IMicrosoftGraphItemBody -Content $Body.Content -ContentType $Body.ContentType
                            }
                        }
                        [array]$Message['Attachment'] = ConvertTo-IMicrosoftGraphAttachment -Attachment $Attachment
                        
                        # TODO: Allow OPTION For Files To Be Shared Via OneDrive/SharePoint Instead Of As Direct Attachments. Perhaps add a switch parameter to override the default config. Maybe even have a config option of use OneDrive when over x Bytes. Set to 0 to always.
                        # Build Email
                        $EmailParams = [ordered]@{
                            SaveToSentItems = $SaveToSentItems
                            Message = [ordered]@{
                                From = $Message.From
                                ReplyTo = $Message.ReplyTo
                                ToRecipients = $Message.To
                                CcRecipients = $Message.CC
                                BccRecipients = $Message.BCC
                                Subject = $Subject
                                Body = $Message.Body
                                Attachments = $Message.Attachment
                            }
                        }
                        
                        # Check For Separate 'SenderID' Value. Make equal to 'From' if not provided.
                        if ([string]::IsNullOrEmpty($SenderId))
                        {
                            $SenderId = $Message.From.emailAddress.Address # Note: This is correct as 'xxxx.Address' (not 'AddressObj'). It is converted Microsoft's to IMicrosoftGraphRecipient.
                        }

                        # Send Email.
                        switch ($MailType)
                        {
                            Group
                            {
                                $SendEmailMessageResult = Send-MgUserMail -UserId $SenderId -BodyParameter $EmailParams -PassThru
                            }
                            OneOnOne
                            {
                                # Each recipient gets a message with only them in To, so nobody sees who else was sent
                                # one. An address listed more than once, even across To, CC, and BCC, gets one message;
                                # hashtable keys compare without case, as email addresses do.
                                $IndividualRecipients = [System.Collections.Generic.List[Object]]::new()
                                $SeenAddresses = @{}
                                foreach ($recipient in @($Message.To) + @($Message.CC) + @($Message.BCC))
                                {
                                    $RecipientAddress = $recipient.EmailAddress.Address
                                    if ([string]::IsNullOrWhiteSpace($RecipientAddress) -or $SeenAddresses.ContainsKey($RecipientAddress))
                                    {
                                        continue
                                    }
                                    $SeenAddresses[$RecipientAddress] = $true
                                    $IndividualRecipients.Add($recipient)
                                }

                                # One recipient's failure is reported and the rest are still sent, so Status is $true only
                                # when every message went out. -ErrorAction Stop makes a failure reach the catch whether or
                                # not the cmdlet reports it as a terminating error.
                                $SentCount = 0
                                foreach ($recipient in $IndividualRecipients)
                                {
                                    # A new request for each recipient rather than one edited in place, so no two
                                    # requests share an object.
                                    $IndividualMessage = [ordered]@{}
                                    foreach ($field in $EmailParams.Message.Keys)
                                    {
                                        $IndividualMessage[$field] = $EmailParams.Message[$field]
                                    }
                                    $IndividualMessage['ToRecipients'] = @($recipient)
                                    $IndividualMessage['CcRecipients'] = $null
                                    $IndividualMessage['BccRecipients'] = $null
                                    $IndividualEmailParams = [ordered]@{
                                        SaveToSentItems = $EmailParams.SaveToSentItems
                                        Message = $IndividualMessage
                                    }

                                    try
                                    {
                                        $null = Send-MgUserMail -UserId $SenderId -BodyParameter $IndividualEmailParams -PassThru -ErrorAction Stop
                                        $SentCount++
                                    }
                                    catch
                                    {
                                        $MgErrorMessages += "Mail not sent to `'$($recipient.EmailAddress.Address)`': $_"
                                    }
                                }

                                if ($SentCount -eq $IndividualRecipients.Count)
                                {
                                    $SendEmailMessageResult = $true
                                }
                            }
                        }
                    }
                }
                catch
                {
                    # Catch any errors and return as part of the $SendScriptMessageResult object.
                    $NewMessage = $_
                    $MgErrorMessages += "$NewMessage"
                }

                # Collect Return Info
                $SendScriptMessageResult_SentFrom = [PSCustomObject]@{
                    Name    = $From.Name
                    Address = $From.AddressObj
                }
                [array]$SendScriptMessageResult_Recipients_To = foreach ($i in $To)
                {
                    [PSCustomObject]@{
                        Name    = $i.Name
                        Address = $i.AddressObj
                    }
                }
                [array]$SendScriptMessageResult_Recipients_CC = foreach ($i in $CC)
                {
                    [PSCustomObject]@{
                        Name    = $i.Name
                        Address = $i.AddressObj
                    }
                }
                [array]$SendScriptMessageResult_Recipients_BCC = foreach ($i in $BCC)
                {
                    [PSCustomObject]@{
                        Name    = $i.Name
                        Address = $i.AddressObj
                    }
                }
                [array]$SendScriptMessageResult_Recipients_All = @( # Since Address is also a PSMethod we need to do some fun stuff (List<psobject> doesn't have a method called Address) so we don't get the dreaded 'OverloadDefinitions'.
                    if ($null -ne $SendScriptMessageResult_Recipients_To)
                    {
                        [System.Linq.Enumerable]::ToList([PSObject[]]$SendScriptMessageResult_Recipients_To).Address
                    }
                    if ($null -ne $SendScriptMessageResult_Recipients_CC)
                    {
                        [System.Linq.Enumerable]::ToList([PSObject[]]$SendScriptMessageResult_Recipients_CC).Address
                    }
                    if ($null -ne $SendScriptMessageResult_Recipients_BCC)
                    {
                        [System.Linq.Enumerable]::ToList([PSObject[]]$SendScriptMessageResult_Recipients_BCC).Address
                    }
                )
                [array]$SendScriptMessageResult_Recipients_All = $SendScriptMessageResult_Recipients_All | Sort-Object -Unique # Remove duplicate items.
                $SendScriptMessageResult_Recipients = [PSCustomObject]@{
                        To = $SendScriptMessageResult_Recipients_To
                        CC = $SendScriptMessageResult_Recipients_CC
                        BCC = $SendScriptMessageResult_Recipients_BCC
                        All = $SendScriptMessageResult_Recipients_All
                }

                # Compile Caught Errors and Warnings
                if ($MgWarningMessages.Count -gt 0 -or $MgErrorMessages.Count -gt 0)
                {
                    [array]$SendScriptMessageResult_Error = foreach ($mgWarningMessage in $MgWarningMessages)
                    {
                        [PSCustomObject]@{
                            Type    = 'Warning'
                            Message = $mgWarningMessage
                        }
                    }

                    [array]$SendScriptMessageResult_Error += foreach ($mgErrorMessage in $MgErrorMessages)
                    {
                        [PSCustomObject]@{
                            Type    = 'Error'
                            Message = $mgErrorMessage
                        }
                    }
                }
                else
                {
                    $SendScriptMessageResult_Error = $null
                }

                $SendScriptMessageResult = [PSCustomObject]@{
                    MessageService = $ServiceId
                    MessageType    = $typeItem
                    MailType       = $MailType
                    Status         = $SendEmailMessageResult # The SDK only returns $true and nothing else (and only that because of the 'PassThru')
                    Error          = $SendScriptMessageResult_Error
                    SentFrom       = $SendScriptMessageResult_SentFrom
                    Recipients = $SendScriptMessageResult_Recipients
                }

                # If successful, output result info.
                $SendScriptMessageResult
            }
            Chat { # TODO MgChat: If application permissions, then do a bot message. Maybe for delegated give option of direct or bot message.
                try
                {
                    # Check if AllowableMessageTypes contains '$typeItem'.
                    if ($ServiceConfig.AllowableMessageTypes -notcontains $typeItem)
                    {
                        $NewMessage = "The ScriptMessage configuration for '$ServiceId' does not allow sending messages of type: $typeItem"
                        Write-Warning -Message $NewMessage
                        $MgWarningMessages += "$NewMessage"
                    }
                    else
                    {
                        # Application CHAT permissions are only supported for migration into a Teams Channel.
                        if ($ServiceConfig.MgPermissionType -eq 'Application')
                        {
                            $NewMessage = "Chat not sent. Microsoft Graph does not support sending Chat messages using Application permissions. Application permissions are only supported for migration into a Teams Channel."
                            Write-Warning -Message $NewMessage
                            $MgWarningMessages += "$NewMessage"
                        }
                        elseif ($null -eq $ChatType)
                        {
                            $NewMessage = "Chat not sent. No chat type is set. Add the 'ChatType' setting to the '$ServiceId' section of the configuration file, or use the 'ChatType' parameter."
                            Write-Warning -Message $NewMessage
                            $MgWarningMessages += "$NewMessage"
                        }
                        else
                        {
                            # Grab the latest MicrosoftGraph service context.
                            $MicrosoftGraphContext = Get-ScriptMessageContext -Service $ServiceId

                            # Check For Separate 'SenderID' Value. Make equal to 'From' if not provided.
                            if ([string]::IsNullOrEmpty($SenderId))
                            {
                                $SenderId = $From.AddressObj
                            }

                            # Make sure SenderID is equal to From address because Microsoft Graph Chat doesn't support sending on behalf of others.
                            if ($SenderId -ne $From.AddressObj)
                            {
                                $NewMessage = "Chat not sent. Microsoft Graph does not support sending Chat messages on behalf of others."
                                Write-Warning -Message $NewMessage
                                $MgWarningMessages += "$NewMessage"
                            }
                            else
                            {
                                # Collect recipient email addresses
                                [array]$ChatRecipients_To = @(foreach ($i in $To.AddressObj){$i})
                                [array]$ChatRecipients_CC = @(foreach ($i in $CC.AddressObj){$i})
                                [array]$ChatRecipients_BCC = @(foreach ($i in $BCC.AddressObj){$i})

                                if (($ChatType -eq [ChatType]'Group') -and ($IncludeBCCInGroupChat -eq $false))
                                {
                                    [array]$ChatRecipients = 
                                        $ChatRecipients_To +
                                        $ChatRecipients_CC
                                }
                                else
                                {
                                    [array]$ChatRecipients = 
                                        $ChatRecipients_To +
                                        $ChatRecipients_CC +
                                        $ChatRecipients_BCC
                                }
                                
                                # Remove 'SenderID' address if it exists in the recipients list as well as duplicates.
                                # (Graph does not support sending direct chat messages to yourself since that's not a standard chat thread. I think it's some sort of "note" when used by Teams.)
                                [array]$ChatRecipients = $ChatRecipients | Sort-Object -Unique | Where-Object {$_ -ne $SenderId}

                                # Collect all chat participants.
                                [array]$AllChatParticipants = [array]$SenderId + [array]$ChatRecipients

                                # Process chat only there are recipients. Otherwise warn if no chat recipients
                                if ($ChatRecipients.Count -eq 0)
                                {
                                    $NewMessage = "Chat not sent. No chat recipients exist. If you are trying to send a chat message to yourself, please note that Microsoft doesn't support direct messaging to yourself via the Graph API."
                                    Write-Warning -Message $NewMessage
                                    $MgWarningMessages += "$NewMessage"
                                }
                                else
                                {
                                    # Add a warning that BCC recipients (not in Sender, To, or CC) are not included in the group chat.
                                    if (($ChatType -eq [ChatType]'Group') -and ($IncludeBCCInGroupChat -eq $false))
                                    {
                                        foreach ($chatRecipient_BCC in $ChatRecipients_BCC)
                                        {
                                            if ($chatRecipient_BCC -notin $AllChatParticipants)
                                            {
                                                $NewMessage = "The following BCC recipient is not included in the group chat: $chatRecipient_BCC"
                                                Write-Warning -Message $NewMessage
                                                $MgWarningMessages += "$NewMessage"
                                            }
                                        }
                                    }

                                    # Convert Parameters to IMicrosoft*
                                    $Message = @{}
                                    if (-not [string]::IsNullOrEmpty($Body.Content))
                                    {
                                        if ([string]::IsNullOrEmpty($Body.ContentType)) # Don't send 'ContentType' if not provided. It will default to 'Text'
                                        {
                                            [hashtable]$Message['Body'] = ConvertTo-IMicrosoftGraphItemBody -Content $Body.Content
                                        }
                                        else
                                        {
                                            [hashtable]$Message['Body'] = ConvertTo-IMicrosoftGraphItemBody -Content $Body.Content -ContentType $Body.ContentType
                                        }
                                    }

                                    # Upload and add any attachments, if needed. # TODO: Check for scope permissions.
                                    # Cannot use Set-MgDriveItemContent because it forces a filepath to be provided and we want to provide content directly sometimes.
                                    if (-not [string]::IsNullOrEmpty($Attachment))
                                    {
                                        # Upload the attached file(s) to OneDrive.
                                        $MgUserDrive = Get-MgUserDrive -UserId $($MicrosoftGraphContext.Account)
                                        $TeamsChatFolder = 'root:/Microsoft Teams Chat Files'
                                        # Upload files. This method only supports files up to 250 MB in size. For larger files, we would need to implement the "createUploadSession" method.
                                        [array]$MgDriveItem = foreach ($attachmentItem in $Attachment)
                                        {
                                            $MicrosoftGraphDriveEndpointUri = 'https://graph.microsoft.com/v1.0/drives/'
                                            $AttachmentFileName = $attachmentItem.Name
                                            
                                            # Get a list of existing files in the Teams Chat Files folder and rename if a file already exists with the same name.
                                            try # Need to check if the $TeamsChatFolder exists first. If not, Get-MgDriveItem will throw a terminating exception.
                                            {
                                                $ExistingFiles = (Get-MgDriveItem -DriveId $MgUserDrive.Id -DriveItemId $TeamsChatFolder -ExpandProperty 'Children').Children
                                            }
                                            catch
                                            {
                                                if ($_.Exception.Message -like "*itemNotFound*")
                                                {
                                                    $ExistingFiles = $null # Make sure this is null.
                                                }
                                                else
                                                {
                                                    # Handle other types of errors
                                                    Write-Error "An unexpected error occurred while uploading file attachment for chat: $($_.Exception.Message)"
                                                }
                                            }
                                            
                                            $FileNameCounter = 0
                                            while ($ExistingFiles.Name -contains $AttachmentFileName)
                                            { 
                                                $FileNameCounter++
                                                $FileBaseName = [System.IO.Path]::GetFileNameWithoutExtension($AttachmentFileName)
                                                $FileExtension = [System.IO.Path]::GetExtension($AttachmentFileName)
                                                $AttachmentFileName = "{0} {1}{2}" -f ($FileBaseName -replace ' \d+$',''), $FileNameCounter, $FileExtension
                                            }

                                            # Encode the filename to handle special characters in the filename when used in the URI.
                                            $AttachmentFileName_Encoded = [uri]::EscapeDataString($AttachmentFileName)

                                            # Upload File
                                            $DriveItemId = "$TeamsChatFolder/$($AttachmentFileName_Encoded):"
                                            $InvokeUri = $($MicrosoftGraphDriveEndpointUri + $MgUserDrive.Id + '/' + $DriveItemId + '/content')
                                            
                                            # Output the drive upload result.
                                            #Set-MgDriveItemContent -DriveId $MgUserDrive.Id -DriveItemId $DriveItemId -InFile $Attachment[0] # Overwrites file if it exists
                                            Invoke-MgGraphRequest -Method PUT -Uri $InvokeUri -Body $attachmentItem.Content -ContentType 'application/octet-stream'  # Overwrites file if it exists
                                        }
                                        
                                        # Update the file(s) sharing permissions.
                                        $DriveInviteParams = ConvertTo-IMicrosoftGraphDriveInvite -EmailAddress $ChatRecipients
                                        foreach ($UploadDriveItemResult in $MgDriveItem)
                                        {
                                            $DriveInviteResult = Invoke-MgInviteDriveItem -DriveId $MgUserDrive.Id -DriveItemId $UploadDriveItemResult.id -BodyParameter $DriveInviteParams
                                        }

                                        # Convert the uploaded files to chat message attachments.
                                        $Message['Attachment'] = [array](ConvertTo-IMicrosoftGraphChatMessageAttachment -MgDriveItem $MgDriveItem)

                                        $ChatParams = [ordered]@{
                                            Body = $Message.Body
                                            Attachments = $Message.Attachment
                                        }
                                    }
                                    else
                                    {
                                        $ChatParams = [ordered]@{
                                            Body = $Message.Body
                                        }
                                    }

                                    # Create a new chat object, if needed, & send the message.
                                    $Member_SenderID = [array](ConvertTo-IMicrosoftGraphConversationMember -EmailAddress $SenderId)

                                    switch ($ChatType)
                                    {
                                        OneOnOne
                                        {
                                            foreach ($chatRecipient in $ChatRecipients)
                                            {
                                                $Member_ChatRecipients = [array](ConvertTo-IMicrosoftGraphConversationMember -EmailAddress $chatRecipient)
                                                [array]$Message['Members'] = [array]$Member_SenderID + [array]$Member_ChatRecipients
                                                try
                                                {
                                                    $NewChatResult = New-MgChat -ChatType $ChatType.ToString() -Members $Message.Members
                                                    $SendChatMessageResult = New-MgChatMessage -ChatId $NewChatResult.Id -BodyParameter $ChatParams
                                                }
                                                catch
                                                {
                                                    $NewMessage = "Cannot create a chat with the recipient '$($chatRecipient)'."
                                                    Write-Warning -Message $NewMessage
                                                    $MgWarningMessages += "$NewMessage"
                                                }
                                            }
                                        }
                                        Group
                                        {
                                            # Collect Group Members
                                            [array]$Member_ChatRecipients = [array](ConvertTo-IMicrosoftGraphConversationMember -EmailAddress $ChatRecipients)
                                            [array]$Message['Members'] = [array]$Member_SenderID + [array]$Member_ChatRecipients

                                            # See if a group chat already exists with the same recipients.
                                            $MGChatProperties = @(
                                                'ChatType',
                                                'Id',
                                                'LastUpdatedDateTime'
                                            )

                                            # If the script has 'Chat.Read' or 'Chat.ReadWrite', then sort by the message preview (last time a message was sent). Otherwise, sort by the last time the chat OBJECT was updated.
                                            [array]$MicrosoftGraphScopes = $MicrosoftGraphContext | Select-Object -ExpandProperty Scopes
                                            if (@($MicrosoftGraphScopes) -contains 'Chat.Read' -or @($MicrosoftGraphScopes) -contains 'Chat.ReadWrite')
                                            {
                                                # It is slower, but we are using the -All parameter so that there is an accurate history of chats. Otherwise, it's possible that we can have multiple groups with the same members from your scripts.
                                                $ExistingGroupChats = Get-MgChat -All -Filter "ChatType eq 'group'" -Property $MGChatProperties -ExpandProperty 'Members', "LastMessagePreview"
                                                $ExistingGroupChats = $ExistingGroupChats | Sort-Object -Property {$_.LastMessagePreview.CreatedDateTime} -Descending
                                            }
                                            else # Only has Chat.ReadBasic so we can't see the last message preview.
                                            {
                                                # It is slower, but we are using the -All parameter so that there is an accurate history of chats. Otherwise, it's possible that we can have multiple groups with the same members from your scripts.
                                                $ExistingGroupChats = Get-MgChat -All -Filter "ChatType eq 'group'" -Property $MGChatProperties -ExpandProperty 'Members'
                                                $ExistingGroupChats = $ExistingGroupChats | Sort-Object -Property LastUpdatedDateTime -Descending
                                            }
                                            
                                            # Reset the variable and then do a compare\search
                                            $LatestExistingGroupChatMatch = $null
                                            foreach ($existingGroupChat in $ExistingGroupChats)
                                            {
                                                if (-not (Compare-Object -ReferenceObject @($existingGroupChat.Members.AdditionalProperties.email) -DifferenceObject $AllChatParticipants))
                                                {
                                                    $LatestExistingGroupChatMatch = $existingGroupChat
                                                }
                                            }

                                            # Send the chat message; create a new chat group if needed.
                                            if (-not $LatestExistingGroupChatMatch)
                                            {
                                                try
                                                {
                                                    $NewChatResult = New-MgChat -ChatType $ChatType.ToString() -Members $Message.Members
                                                    $ChatToUse = $NewChatResult
                                                    $SendChatMessageResult = New-MgChatMessage -ChatId $ChatToUse.Id -BodyParameter $ChatParams
                                                }
                                                catch
                                                {
                                                    $NewMessage = "Cannot create a new Teams group chat due to at least one recipient of the group: '$($ChatRecipients -join ', ')'."
                                                    Write-Warning -Message $NewMessage
                                                    $MgWarningMessages += "$NewMessage"
                                                }
                                            }
                                            else
                                            {
                                                $ChatToUse = $LatestExistingGroupChatMatch
                                                $SendChatMessageResult = New-MgChatMessage -ChatId $ChatToUse.Id -BodyParameter $ChatParams
                                            }
                                        }
                                    }
                                }
                            }   
                        }
                    }
                }
                catch
                {
                    # Catch any errors and return as part of the $SendScriptMessageResult object.
                    $NewMessage = $_
                    $MgErrorMessages += "$NewMessage"
                }
                
                # Collect Return Info
                $SendScriptMessageResult_SentFrom = [PSCustomObject]@{
                    Name    = $From.Name
                    Address = $From.AddressObj
                }
                [array]$SendScriptMessageResult_Recipients_To = foreach ($i in $To)
                {
                    [PSCustomObject]@{
                        Name    = $i.Name
                        Address = $i.AddressObj
                    }
                }
                [array]$SendScriptMessageResult_Recipients_CC = foreach ($i in $CC)
                {
                    [PSCustomObject]@{
                        Name    = $i.Name
                        Address = $i.AddressObj
                    }
                }
                [array]$SendScriptMessageResult_Recipients_BCC = foreach ($i in $BCC)
                {
                    [PSCustomObject]@{
                        Name    = $i.Name
                        Address = $i.AddressObj
                    }
                }
                [array]$SendScriptMessageResult_Recipients_All = @( # Since Address is also a PSMethod we need to do some fun stuff (List<psobject> doesn't have a method called Address) so we don't get the dreaded 'OverloadDefinitions'.
                    if ($null -ne $SendScriptMessageResult_Recipients_To)
                    {
                        [System.Linq.Enumerable]::ToList([PSObject[]]$SendScriptMessageResult_Recipients_To).Address
                    }
                    if ($null -ne $SendScriptMessageResult_Recipients_CC)
                    {
                        [System.Linq.Enumerable]::ToList([PSObject[]]$SendScriptMessageResult_Recipients_CC).Address
                    }
                    if ($null -ne $SendScriptMessageResult_Recipients_BCC)
                    {
                        [System.Linq.Enumerable]::ToList([PSObject[]]$SendScriptMessageResult_Recipients_BCC).Address
                    }
                )
                [array]$SendScriptMessageResult_Recipients_All = $SendScriptMessageResult_Recipients_All | Sort-Object -Unique # Remove duplicate items.
                [bool]$SendScriptMessageResult_Recipients_IncludeBCCInGroupChat = $IncludeBCCInGroupChat
                $SendScriptMessageResult_Recipients = [PSCustomObject]@{
                        To = $SendScriptMessageResult_Recipients_To
                        CC = $SendScriptMessageResult_Recipients_CC
                        BCC = $SendScriptMessageResult_Recipients_BCC
                        All = $SendScriptMessageResult_Recipients_All
                        IncludeBCCInGroupChat = $SendScriptMessageResult_Recipients_IncludeBCCInGroupChat
                }

                # Compile Caught Errors and Warnings
                if ($MgWarningMessages.Count -gt 0 -or $MgErrorMessages.Count -gt 0)
                {
                    [array]$SendScriptMessageResult_Error = foreach ($mgWarningMessage in $MgWarningMessages)
                    {
                        [PSCustomObject]@{
                            Type    = 'Warning'
                            Message = $mgWarningMessage
                        }
                    }

                    [array]$SendScriptMessageResult_Error += foreach ($mgErrorMessage in $MgErrorMessages)
                    {
                        [PSCustomObject]@{
                            Type    = 'Error'
                            Message = $mgErrorMessage
                        }
                    }
                }
                else
                {
                    $SendScriptMessageResult_Error = $null
                }
                
                $SendScriptMessageResult = [PSCustomObject]@{
                    MessageService = $ServiceId
                    MessageType    = $typeItem
                    ChatType        = $ChatType
                    Status         = $SendChatMessageResult
                    Error          = $SendScriptMessageResult_Error
                    SentFrom       = $SendScriptMessageResult_SentFrom
                    Recipients = $SendScriptMessageResult_Recipients
                }

                # If successful, output result info.
                $SendScriptMessageResult
            }
            Default {
                Write-Warning -Message "'$($typeItem)' is an invalid message type for service '$($ServiceId)'."
            }
        }
    }

    # Disconnect from the Microsoft Graph API, if enabled in the configuration file. The setting is an opt-in flag,
    # on only when it equals $true; keep it on the left, because a truthiness test reads the text "false" as $true.
    if ($ServiceConfig.MgDisconnectWhenDone -eq $true)
    {
        $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
    }
}

# Announce this service to the module. Keep this at the bottom of the file, after the functions it names.
Register-ScriptMessageService -Name 'MicrosoftGraph' `
    -ConnectFunction    'Connect-ScriptMessage_MicrosoftGraph' `
    -DisconnectFunction 'Disconnect-ScriptMessage_MicrosoftGraph' `
    -SendFunction       'Send-ScriptMessage_MicrosoftGraph' `
    -GetContextFunction 'Get-ScriptMessageContext_MicrosoftGraph'
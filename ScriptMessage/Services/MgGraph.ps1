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
    if (([string]::IsNullOrEmpty($EmailAddress)) -and ([string]::IsNullOrEmpty($EmailAddress.Address)))
    {
        return $null
    }

    # Loop through each of the recipient paramater array objects
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
                EmailAddress = @{
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
        switch ($currentAttachment.GetType().Name)
        {
            'Hashtable' { # If direct file content is supplied.
                $AttachmentType = 'Content'
                if (($currentAttachment.ContainsKey('Name')) -and $currentAttachment.ContainsKey('Content'))
                {
                    $Attachment_ByteEncoded = [System.Convert]::ToBase64String($currentAttachment.Content)
                }
                else
                {
                    throw "The attachment hashtable object is improperly formatted. The hashtable requires the keys of `'Name`' and `'Contents`'"
                }

                [array]$IMicrosoftGraphAttachmentItem = @{
                    "@odata.type" = "#microsoft.graph.fileAttachment"
                    Name          = $currentAttachment.Name
                    ContentBytes  = $Attachment_ByteEncoded
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
                        $Attachment_ByteEncoded = [convert]::ToBase64String([System.IO.File]::ReadAllBytes($FileInfo.FullName))
                        [array]$IMicrosoftGraphAttachmentItem = @{
                            "@odata.type" = "#microsoft.graph.fileAttachment"
                            Name          = $FileInfo.Name
                            ContentBytes  = $Attachment_ByteEncoded
                        }   
                    }
                    'DirectoryInfo' {
                        $AttachmentType = 'DirectoryPath'
                        $DirectoryContent = Get-ChildItem $currentAttachment -File -Recurse
                        [array]$IMicrosoftGraphAttachmentItem = foreach ($file in $DirectoryContent)
                        {
                            $Attachment_ByteEncoded = [convert]::ToBase64String([System.IO.File]::ReadAllBytes($file.FullName))
                            @{
                                "@odata.type" = "#microsoft.graph.fileAttachment"
                                Name          = $file.Name
                                ContentBytes  = $Attachment_ByteEncoded
                            }   
                        }
                    }
                    Default {throw 'Unexpected attachment object type.'}
                }
            }
            Default {throw 'Unexpected attachment object type.'}
        }
    
        $IMicrosoftGraphAttachmentItem
    }

    return $IMicrosoftGraphAttachment
}

function Send-ScriptMessage_MgGraph
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
        [string]$SenderId
    )

    # Set the Service ID.
    $ServiceId = 'MgGraph'
    
    # Set the necesasary configuration variables.
    $ScriptMessageConfig = Get-ScriptMessageConfig

    # Send the message on each supported service specified.
    foreach ($typeItem in $Type)
    {
        switch ($typeItem)
        {
            Mail {  
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
                
                # Check For Separate UserID Value
                if ([string]::IsNullOrEmpty($SenderId))
                {
                    $SenderId = $Message.From.emailAddress.Address
                }
    
                # Check if using beta Graph API & Send Email.
                if (-not ($ScriptMessageConfig.MgGraph.MgProfile -eq 'beta'))
                {
                    $SendEmailMessageResult = Send-MgUserMail -UserId $SenderId -BodyParameter $EmailParams -PassThru
                }
                else
                {
                    $SendEmailMessageResult = Send-MgBetaUserMail -UserId $SenderId -BodyParameter $EmailParams -PassThru
                }
    
                # Collect Return Info
                $SendScriptMessageResult = [ordered]@{}
                $SendScriptMessageResult.MessageService = $ServiceId
                $SendScriptMessageResult.MessageType = $typeItem
                $SendScriptMessageResult.Status = $SendEmailMessageResult # The SDK only returns $true and nothing else (and only that because of the 'PassThru')
                $SendScriptMessageResult.Error = $null
                $SendScriptMessageResult.SentFrom = @{}
                $SendScriptMessageResult.SentFrom.Name = $From.Name
                $SendScriptMessageResult.SentFrom.Address = $From.AddressObj
                $SendScriptMessageResult.Recipients = [ordered]@{}
                $SendScriptMessageResult.Recipients.All = $null # Create this before populating for ordered list purposes.
                [array]$SendScriptMessageResult.Recipients.To = @(($Message.To).EmailAddress | Sort-Object $_.Value)
                [array]$SendScriptMessageResult.Recipients.CC = @(($Message.CC).EmailAddress | Sort-Object $_.Value)
                [array]$SendScriptMessageResult.Recipients.BCC = @(($Message.BCC).EmailAddress | Sort-Object $_.Value)
                [array]$SendScriptMessageResult.Recipients.All = @( # Since Address is also a PSMethod we need to do some fun stuff (List<psobject> doesn't have a method called Address) so we don't get the dreaded 'OverloadDefinitions'.
                    [System.Linq.Enumerable]::ToList([psobject[]]$SendScriptMessageResult.Recipients.To).Address
                    [System.Linq.Enumerable]::ToList([psobject[]]$SendScriptMessageResult.Recipients.CC).Address
                    [System.Linq.Enumerable]::ToList([psobject[]]$SendScriptMessageResult.Recipients.BCC).Address
                )
                [array]$SendScriptMessageResult.Recipients.All = $SendScriptMessageResult.Recipients.All | Sort-Object -Unique # Remove duplicate items.
    
                # If successful, output result info.
                $SendScriptMessageResult
            }
            Chat {
                Write-Warning -Message "The '$($typeItem)' message type has not yet been implemented for service '$($ServiceId)'."
            }
            Default {
                Write-Warning -Message "'$($typeItem)' is an invalid message type for service '$($ServiceId)'."
            }
        }
    }
}

function Connect-ScriptMessage_MgGraph
{
    [CmdletBinding()]
    param(
        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [pscustomobject]$ServiceConfig
    )

    # Get Graph Profile From Config
    [string]$MgProfile = $ServiceConfig.MgProfile # 'beta' or 'v1.0'.

    # Check For Microsoft.Graph Module
    # Don't import the entire 'Microsoft.Graph' module. Only import the needed sub-modules.
    if (-not ($MgProfile -eq 'beta'))
    {
        Import-Module 'Microsoft.Graph.Authentication' -ErrorAction SilentlyContinue
        Import-Module 'Microsoft.Graph.Users.Actions' -ErrorAction SilentlyContinue
        if (!(Get-Module -Name "Microsoft.Graph.Users.Actions"))
        {
            # Module is not available.
            Write-Error @"
Please First Install the Microsoft.Graph.Users.Actions Module from https://www.powershellgallery.com/packages/Microsoft.Graph/ ".
Installing the main modules of the SDK, Microsoft.Graph, will install all sub modules for each module.
Consider only installing the necessary modules, including Microsoft.Graph.Authentication which is installed by default when you opt
to install the sub modules individually. For a list of available Microsoft Graph modules, use Find-Module Microsoft.Graph*.
Only cmdlets for the installed modules will be available for use.
"@
            Return
        }
    }
    else
    {
        Import-Module 'Microsoft.Graph.Authentication' -ErrorAction SilentlyContinue # No beta version available for this required module.
        Import-Module 'Microsoft.Graph.Beta.Users.Actions' -ErrorAction SilentlyContinue
        if (!(Get-Module -Name "Microsoft.Graph.Beta.Users.Actions"))
        {
            # Module is not available.
            Write-Error @"
Please First Install the Microsoft.Graph.Beta.Users.Actions Module from https://www.powershellgallery.com/packages/Microsoft.Graph.Beta/ ".
Installing the main modules of the SDK, Microsoft.Graph.Beta, will install all sub modules for each module.
Consider only installing the necessary modules, including Microsoft.Graph.Authentication which is installed by default when you opt
to install the sub modules individually. For a list of available Microsoft Graph modules, use Find-Module Microsoft.Graph*.
Only cmdlets for the installed modules will be available for use.
"@
            Return
        }
    }

    # Connect to the Microsoft Graph API.
    # E.g. Connect-MgGraph -Scopes "User.Read.All","Group.ReadWrite.All"
    # You can add additional permissions by repeating the Connect-MgGraph command with the new permission scopes.
    # View the current scopes under which the PowerShell SDK is (trying to) execute cmdlets: Get-MgContext | select -ExpandProperty Scopes
    # List all the scopes granted on the service principal object (you cn also do it via the Azure AD UI): Get-MgServicePrincipal -Filter "appId eq '14d82eec-204b-4c2f-b7e8-296a70dab67e'" | % { Get-MgServicePrincipalOauth2PermissionGrant -ServicePrincipalId $_.Id } | fl
    # Find Graph permission needed. More info on permissions: https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent)
    #    E.g., Find-MgGraphPermission -SearchString "Teams" -PermissionType Delegated
    #    E.g., Find-MgGraphPermission -SearchString "Teams" -PermissionType Application
    $MicrosoftGraphScopes = @(
        'Mail.Send'
        #'Mail.Send.Shared' # Scope is not needed at the moment.
    )
    
    $MgPermissionType = $ServiceConfig.MgPermissionType
    $MgTenantID = $ServiceConfig.MgTenantID
    $MgClientID = $ServiceConfig.MgClientID

    switch ($MgPermissionType)
    {
        Delegated {
            $null = Connect-MgGraph -Scopes $MicrosoftGraphScopes -TenantId $MgTenantID -ClientId $MgClientID
        }
        Application {
            [string]$MgApp_AuthenticationType = $ServiceConfig.MgApp_AuthenticationType
            if ($LoggingEnabled) {Write-PSFMessage -Message "Microsoft Graph App Authentication Type: $MgApp_AuthenticationType"}

            switch ($MgApp_AuthenticationType)
            {
                CertificateFile {
                    $MgApp_CertificatePath = $ExecutionContext.InvokeCommand.ExpandString($ServiceConfig.MgApp_CertificatePath)

                    # Try accessing private key certificate without password using current process credentials.
                    [X509Certificate]$MgApp_Certificate = $null
                    try
                    {
                        [X509Certificate]$MgApp_Certificate = Get-PfxCertificate -FilePath $MgApp_CertificatePath -NoPromptForPassword
                    }
                    catch # If that doesn't work try the included credentials.
                    {
                        $MgApp_EncryptedCertificatePassword = $ServiceConfig.MgApp_EncryptedCertificatePassword
                        if ([string]::IsNullOrEmpty($MgApp_EncryptedCertificatePassword))
                        {
                            if ($LoggingEnabled) {Write-PSFMessage -Level Error "Cannot access .pfx private key certificate file and no password has been provided."}
                            throw $_
                        }
                        else
                        {
                            [SecureString]$MgApp_EncryptedCertificateSecureString = $MgApp_EncryptedCertificatePassword | ConvertTo-SecureString # Can only be decrypted by the same AD account on the same computer.
                            [X509Certificate]$MgApp_Certificate = Get-PfxCertificate -FilePath $MgApp_CertificatePath -NoPromptForPassword -Password $MgApp_EncryptedCertificateSecureString
                        }
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
                    $MgApp_Secret = [System.Net.NetworkCredential]::new("", $($MgApp_EncryptedSecret | ConvertTo-SecureString)).Password # Can only be decrypted by the same AD account on the same computer.
                    $Body =  @{
                        Grant_Type    = "client_credentials"
                        Scope         = "https://graph.microsoft.com/.default"
                        Client_Id     = $MgClientID
                        Client_Secret = $MgApp_Secret
                    }
                    $Connection = Invoke-RestMethod `
                        -Uri https://login.microsoftonline.com/$MgTenantID/oauth2/v2.0/token `
                        -Method POST `
                        -Body $Body
                    $AccessToken = $Connection.access_token
                    $null = Connect-MgGraph -AccessToken $AccessToken
                }
                Default {throw "Invalid `'MgApp_AuthenticationType`' value."}
            }
        }
        Default {throw "Invalid `'MgPermissionType`' value."}
    }
}
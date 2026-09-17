# Pester 5 or later. Run from the repository root:
#   Invoke-Pester -Path .\Tests
# Nothing here connects to a tenant, and the Microsoft Graph modules do not need to be installed.

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\ScriptMessage\ScriptMessage.psd1') -Force

    # Pester can only mock a command that exists, so define stand-ins for the Microsoft Graph cmdlets inside the
    # module before mocking them. Each throws if a test reaches it without a mock.
    InModuleScope ScriptMessage {
        function script:Send-MgUserMail { param($UserId, $BodyParameter, [switch]$PassThru) throw 'Send-MgUserMail stand-in called without a mock.' }
        function script:Disconnect-MgGraph { throw 'Disconnect-MgGraph stand-in called without a mock.' }
        function script:New-MgChat { param($ChatType, $Members) throw 'New-MgChat stand-in called without a mock.' }
        function script:New-MgChatMessage { param($ChatId, $BodyParameter) throw 'New-MgChatMessage stand-in called without a mock.' }
        function script:Get-MgChat { param([switch]$All, $Filter, $Property, $ExpandProperty) throw 'Get-MgChat stand-in called without a mock.' }
        function script:Get-MgUserDrive { param($UserId) throw 'Get-MgUserDrive stand-in called without a mock.' }
        function script:Get-MgDriveItem { param($DriveId, $DriveItemId, $ExpandProperty) throw 'Get-MgDriveItem stand-in called without a mock.' }
        function script:Invoke-MgGraphRequest { param($Method, $Uri, $Body, $ContentType) throw 'Invoke-MgGraphRequest stand-in called without a mock.' }
        function script:Invoke-MgInviteDriveItem { param($DriveId, $DriveItemId, $BodyParameter) throw 'Invoke-MgInviteDriveItem stand-in called without a mock.' }
    }
}

AfterAll {
    Remove-Module ScriptMessage -Force -ErrorAction SilentlyContinue
}

Describe 'Send-ScriptMessage recipient check' {
    BeforeAll {
        Mock -ModuleName ScriptMessage Get-ScriptMessageConfig {
            $Config = [pscustomobject]@{
                MicrosoftGraph = [pscustomobject]@{
                    AllowableMessageTypes = @('Mail')
                    MailType              = 'Group'
                    ChatType              = 'Group'
                    IncludeBCCInGroupChat = $false
                    MgPermissionType      = 'Application'
                    MgDisconnectWhenDone  = $false
                }
            }
            # MicrosoftGraph is enum value 0, which is false, so compare with $null rather than testing $Service.
            if ($null -ne $Service) { $Config.$Service } else { $Config }
        }
        Mock -ModuleName ScriptMessage Connect-ScriptMessage { }
        Mock -ModuleName ScriptMessage Send-MgUserMail { $true }

        $MessageArguments = @{
            Service = 'MicrosoftGraph'
            Type    = 'Mail'
            From    = 'sender@example.org'
            Subject = 'Test subject'
            Body    = 'Test body'
        }
    }

    Context 'Recipients with an address are accepted' {
        It 'sends to <Name>' -ForEach @(
            @{ Name = 'a one-item array holding a ConvertFrom-Json object'; ExpectedType = 'Object[]'; Expected = 'a@example.org'
               To = @((ConvertFrom-Json '{"Name":"A","Address":"a@example.org"}')) }
            @{ Name = 'a one-item JSON array property'; ExpectedType = 'Object[]'; Expected = 'a@example.org'
               To = (ConvertFrom-Json '{"To":[{"Name":"A","Address":"a@example.org"}]}').To }
            @{ Name = 'a single object'; ExpectedType = 'PSCustomObject'; Expected = 'a@example.org'
               To = [pscustomobject]@{ Name = 'A'; Address = 'a@example.org' } }
            @{ Name = 'two objects'; ExpectedType = 'Object[]'; Expected = 'a@example.org,b@example.org'
               To = @([pscustomobject]@{ Name = 'A'; Address = 'a@example.org' }, [pscustomobject]@{ Name = 'B'; Address = 'b@example.org' }) }
            @{ Name = 'a hashtable'; ExpectedType = 'Hashtable'; Expected = 'a@example.org'
               To = @{ Name = 'A'; Address = 'a@example.org' } }
            @{ Name = 'a string'; ExpectedType = 'String'; Expected = 'a@example.org'
               To = 'a@example.org' }
            @{ Name = 'two strings'; ExpectedType = 'Object[]'; Expected = 'a@example.org,b@example.org'
               To = @('a@example.org', 'b@example.org') }
        ) {
            # Confirms the test data kept its shape; an unrolled one-item array would no longer test that case.
            $To.GetType().Name | Should -Be $ExpectedType

            $null = Send-ScriptMessage @MessageArguments -To $To

            Should -Invoke -ModuleName ScriptMessage Send-MgUserMail -Times 1 -Exactly -ParameterFilter {
                (@($BodyParameter.Message.ToRecipients | ForEach-Object { $_.EmailAddress.Address }) -join ',') -eq $Expected
            }
        }

        It 'sends to a one-item array in CC when To is not given' {
            $CC = @((ConvertFrom-Json '{"Name":"C","Address":"c@example.org"}'))

            $null = Send-ScriptMessage @MessageArguments -CC $CC

            Should -Invoke -ModuleName ScriptMessage Send-MgUserMail -Times 1 -Exactly -ParameterFilter {
                (@($BodyParameter.Message.CcRecipients | ForEach-Object { $_.EmailAddress.Address }) -join ',') -eq 'c@example.org'
            }
        }
    }

    Context 'Recipients without an address are rejected before connecting' {
        It 'throws for <Name>' -ForEach @(
            @{ Name = 'an empty string'; To = '' }
            @{ Name = 'a whitespace string'; To = '   ' }
            @{ Name = 'an object with a blank Address'; To = [pscustomobject]@{ Name = 'A'; Address = '' } }
            @{ Name = 'two objects with blank Addresses'; To = @([pscustomobject]@{ Name = 'A'; Address = '' }, [pscustomobject]@{ Name = 'B'; Address = ' ' }) }
            @{ Name = 'a hashtable with a blank Address'; To = @{ Name = 'A'; Address = '' } }
        ) {
            { Send-ScriptMessage @MessageArguments -To $To } | Should -Throw '*at least one parameter value*'
            Should -Invoke -ModuleName ScriptMessage Connect-ScriptMessage -Times 0 -Exactly
        }

        It 'throws when no To, CC, or BCC is given' {
            { Send-ScriptMessage @MessageArguments } | Should -Throw '*at least one parameter value*'
            Should -Invoke -ModuleName ScriptMessage Connect-ScriptMessage -Times 0 -Exactly
        }
    }
}

Describe 'Send-ScriptMessage chat message' {
    BeforeAll {
        Mock -ModuleName ScriptMessage Connect-ScriptMessage { }
        Mock -ModuleName ScriptMessage Get-ScriptMessageContext {
            [pscustomobject]@{
                Account = 'sender@example.org'
                Scopes  = @('Chat.ReadWrite')
            }
        }
        Mock -ModuleName ScriptMessage New-MgChat { [pscustomobject]@{ Id = 'new-chat-id' } }
        Mock -ModuleName ScriptMessage New-MgChatMessage { [pscustomobject]@{ Id = 'new-message-id' } }
        Mock -ModuleName ScriptMessage Get-MgChat { @() }
        Mock -ModuleName ScriptMessage Send-MgUserMail { $true }

        $MessageArguments = @{
            Service = 'MicrosoftGraph'
            Type    = 'Chat'
            From    = 'sender@example.org'
            To      = 'recipient@example.org'
            Subject = 'Test subject'
            Body    = 'Test body'
        }
    }

    Context 'A one-on-one chat with no attachments' {
        BeforeAll {
            Mock -ModuleName ScriptMessage Get-ScriptMessageConfig {
                $Config = [pscustomobject]@{
                    MicrosoftGraph = [pscustomobject]@{
                        AllowableMessageTypes = @('Mail', 'Chat')
                        MailType              = 'Group'
                        IncludeBCCInGroupChat = $false
                        MgPermissionType      = 'Delegated'
                        MgDisconnectWhenDone  = $false
                    }
                }
                # MicrosoftGraph is enum value 0, which is false, so compare with $null rather than testing $Service.
                if ($null -ne $Service) { $Config.$Service } else { $Config }
            }
        }

        It 'sends the message body and reports no error' {
            $Result = Send-ScriptMessage @MessageArguments -ChatType OneOnOne

            $Result.Error | Should -BeNullOrEmpty
            $Result.Status.Id | Should -Be 'new-message-id'
            Should -Invoke -ModuleName ScriptMessage New-MgChatMessage -Times 1 -Exactly -ParameterFilter {
                $BodyParameter.Body.Content -eq 'Test body'
            }
        }
    }

    Context 'A group chat with no attachments' {
        BeforeAll {
            Mock -ModuleName ScriptMessage Get-ScriptMessageConfig {
                $Config = [pscustomobject]@{
                    MicrosoftGraph = [pscustomobject]@{
                        AllowableMessageTypes = @('Mail', 'Chat')
                        MailType              = 'Group'
                        IncludeBCCInGroupChat = $false
                        MgPermissionType      = 'Delegated'
                        MgDisconnectWhenDone  = $false
                    }
                }
                if ($null -ne $Service) { $Config.$Service } else { $Config }
            }
        }

        It 'sends the message body and reports no error' {
            $Result = Send-ScriptMessage @MessageArguments -ChatType Group

            $Result.Error | Should -BeNullOrEmpty
            $Result.Status.Id | Should -Be 'new-message-id'
            Should -Invoke -ModuleName ScriptMessage New-MgChatMessage -Times 1 -Exactly -ParameterFilter {
                $BodyParameter.Body.Content -eq 'Test body'
            }
        }
    }

    Context 'Mail and Chat in the same call' {
        BeforeAll {
            Mock -ModuleName ScriptMessage Get-ScriptMessageConfig {
                $Config = [pscustomobject]@{
                    MicrosoftGraph = [pscustomobject]@{
                        AllowableMessageTypes = @('Mail', 'Chat')
                        MailType              = 'Group'
                        IncludeBCCInGroupChat = $false
                        MgPermissionType      = 'Delegated'
                        MgDisconnectWhenDone  = $false
                    }
                }
                if ($null -ne $Service) { $Config.$Service } else { $Config }
            }
        }

        It 'sends both, and the chat carries its own message content' {
            $Arguments = $MessageArguments.Clone()
            $Arguments['Type'] = @('Mail', 'Chat')
            $Arguments['ChatType'] = 'OneOnOne'

            $Result = @(Send-ScriptMessage @Arguments)

            @($Result | Where-Object { $_.MessageType -eq 'Mail' }).Count | Should -Be 1
            @($Result | Where-Object { $_.MessageType -eq 'Chat' }).Count | Should -Be 1
            # One result comes back per message type, so drop the empty Error properties instead of testing the
            # array of them, which is neither null nor empty.
            @($Result.Error | Where-Object { $null -ne $_ }) | Should -BeNullOrEmpty
            Should -Invoke -ModuleName ScriptMessage New-MgChatMessage -Times 1 -Exactly -ParameterFilter {
                $BodyParameter.Body.Content -eq 'Test body'
            }
        }
    }

    Context 'A one-on-one chat with an attachment' {
        BeforeAll {
            Mock -ModuleName ScriptMessage Get-ScriptMessageConfig {
                $Config = [pscustomobject]@{
                    MicrosoftGraph = [pscustomobject]@{
                        AllowableMessageTypes = @('Mail', 'Chat')
                        MailType              = 'Group'
                        IncludeBCCInGroupChat = $false
                        MgPermissionType      = 'Delegated'
                        MgDisconnectWhenDone  = $false
                    }
                }
                if ($null -ne $Service) { $Config.$Service } else { $Config }
            }
            Mock -ModuleName ScriptMessage Get-MgUserDrive { [pscustomobject]@{ Id = 'drive-id' } }
            Mock -ModuleName ScriptMessage Get-MgDriveItem { $null } # The chat files folder does not exist yet.
            Mock -ModuleName ScriptMessage Invoke-MgGraphRequest {
                # Graph returns the uploaded drive item as a hashtable, which the attachment converter reads by key.
                @{ id = 'drive-item-id'; name = 'My File+1.pdf'; webUrl = 'https://www.example.com/file' }
            }
            Mock -ModuleName ScriptMessage Invoke-MgInviteDriveItem { @{ id = 'invite-id' } }
        }

        It 'escapes the attachment filename for the upload path' {
            $Arguments = $MessageArguments.Clone()
            $Arguments['Attachment'] = @{ Name = 'My File+1.pdf'; Content = [byte[]]@(1, 2, 3) }
            $Arguments['ChatType'] = 'OneOnOne'

            $Result = Send-ScriptMessage @Arguments

            $Result.Error | Should -BeNullOrEmpty
            # A space has to arrive as %20. In a URL path a '+' is a literal plus sign, not a space.
            Should -Invoke -ModuleName ScriptMessage Invoke-MgGraphRequest -Times 1 -Exactly -ParameterFilter {
                $Uri -like '*/root:/Microsoft Teams Chat Files/My%20File%2B1.pdf:/content'
            }
        }
    }
}

Describe 'Send-ScriptMessage mail attachment' {
    BeforeAll {
        Mock -ModuleName ScriptMessage Get-ScriptMessageConfig {
            $Config = [pscustomobject]@{
                MicrosoftGraph = [pscustomobject]@{
                    AllowableMessageTypes = @('Mail')
                    MailType              = 'Group'
                    ChatType              = 'Group'
                    IncludeBCCInGroupChat = $false
                    MgPermissionType      = 'Application'
                    MgDisconnectWhenDone  = $false
                }
            }
            if ($null -ne $Service) { $Config.$Service } else { $Config }
        }
        Mock -ModuleName ScriptMessage Connect-ScriptMessage { }
        Mock -ModuleName ScriptMessage Send-MgUserMail { $true }

        $MailArguments = @{
            Service = 'MicrosoftGraph'
            Type    = 'Mail'
            From    = 'sender@example.org'
            To      = 'recipient@example.org'
            Subject = 'Test subject'
            Body    = 'Test body'
        }
    }

    It 'sends <Count> attachment(s) to Microsoft Graph' -ForEach @(
        @{ Count = 1; Files = @(@{ Name = 'One.pdf'; Content = [byte[]]@(1) }) }
        @{ Count = 2; Files = @(@{ Name = 'One.pdf'; Content = [byte[]]@(1) }, @{ Name = 'Two.pdf'; Content = [byte[]]@(2) }) }
    ) {
        $null = Send-ScriptMessage @MailArguments -Attachment $Files

        # An empty Attachments property is $null, and @($null).Count is 1, so filter before counting.
        Should -Invoke -ModuleName ScriptMessage Send-MgUserMail -Times 1 -Exactly -ParameterFilter {
            @($BodyParameter.Message.Attachments | Where-Object { $null -ne $_ }).Count -eq $Count
        }
    }
}

Describe 'Send-ScriptMessage configuration defaults' {
    BeforeAll {
        Mock -ModuleName ScriptMessage Connect-ScriptMessage { }
        Mock -ModuleName ScriptMessage Get-ScriptMessageContext {
            [pscustomobject]@{
                Account = 'sender@example.org'
                Scopes  = @('Chat.ReadWrite')
            }
        }
        Mock -ModuleName ScriptMessage New-MgChat { [pscustomobject]@{ Id = 'new-chat-id' } }
        Mock -ModuleName ScriptMessage New-MgChatMessage { [pscustomobject]@{ Id = 'new-message-id' } }
        Mock -ModuleName ScriptMessage Get-MgChat { @() }
        Mock -ModuleName ScriptMessage Send-MgUserMail { $true }

        $ChatArguments = @{
            Service = 'MicrosoftGraph'
            Type    = 'Chat'
            From    = 'sender@example.org'
            To      = 'recipient@example.org'
            Subject = 'Test subject'
            Body    = 'Test body'
        }
    }

    Context 'A configuration that sets ChatType to Group' {
        BeforeAll {
            Mock -ModuleName ScriptMessage Get-ScriptMessageConfig {
                $Config = [pscustomobject]@{
                    MicrosoftGraph = [pscustomobject]@{
                        AllowableMessageTypes = @('Mail', 'Chat')
                        MailType              = 'Group'
                        ChatType              = 'Group'
                        IncludeBCCInGroupChat = $false
                        MgPermissionType      = 'Delegated'
                        MgDisconnectWhenDone  = $false
                    }
                }
                if ($null -ne $Service) { $Config.$Service } else { $Config }
            }
        }

        # 'OneOnOne' is the first ChatType enum member, so its numeric value is 0. A truthiness test on the
        # parameter cannot tell it apart from an unsupplied value.
        It 'an explicit -ChatType OneOnOne wins over the configuration' {
            $null = Send-ScriptMessage @ChatArguments -ChatType OneOnOne

            Should -Invoke -ModuleName ScriptMessage New-MgChat -Times 1 -Exactly -ParameterFilter {
                $ChatType -eq 'OneOnOne'
            }
            # Only the Group branch looks for a chat to reuse.
            Should -Invoke -ModuleName ScriptMessage Get-MgChat -Times 0 -Exactly
        }

        It 'the configuration value is used when -ChatType is not supplied' {
            $null = Send-ScriptMessage @ChatArguments

            Should -Invoke -ModuleName ScriptMessage New-MgChat -Times 1 -Exactly -ParameterFilter {
                $ChatType -eq 'Group'
            }
            Should -Invoke -ModuleName ScriptMessage Get-MgChat -Times 1 -Exactly
        }
    }

    Context 'A configuration that sets IncludeBCCInGroupChat to true' {
        BeforeAll {
            Mock -ModuleName ScriptMessage Get-ScriptMessageConfig {
                $Config = [pscustomobject]@{
                    MicrosoftGraph = [pscustomobject]@{
                        AllowableMessageTypes = @('Mail', 'Chat')
                        MailType              = 'Group'
                        ChatType              = 'Group'
                        IncludeBCCInGroupChat = $true
                        MgPermissionType      = 'Delegated'
                        MgDisconnectWhenDone  = $false
                    }
                }
                if ($null -ne $Service) { $Config.$Service } else { $Config }
            }
        }

        It 'an explicit -IncludeBCCInGroupChat $false wins over the configuration' {
            $Result = Send-ScriptMessage @ChatArguments -BCC 'blind@example.org' -IncludeBCCInGroupChat $false -WarningAction SilentlyContinue

            @($Result.Error | Where-Object { $_.Type -eq 'Warning' -and $_.Message -like '*not included in the group chat*' }).Count |
                Should -Be 1
        }

        It 'the configuration value is used when -IncludeBCCInGroupChat is not supplied' {
            $Result = Send-ScriptMessage @ChatArguments -BCC 'blind@example.org'

            $Result.Error | Should -BeNullOrEmpty
        }
    }

    Context 'A configuration written before the ChatType setting existed' {
        BeforeAll {
            # A configuration file from 1.0.7 or earlier has no ChatType setting at all.
            Mock -ModuleName ScriptMessage Get-ScriptMessageConfig {
                $Config = [pscustomobject]@{
                    MicrosoftGraph = [pscustomobject]@{
                        AllowableMessageTypes = @('Mail')
                        MgPermissionType      = 'Application'
                        MgDisconnectWhenDone  = $false
                    }
                }
                if ($null -ne $Service) { $Config.$Service } else { $Config }
            }
        }

        It 'sends Mail without needing a ChatType setting' {
            $MailArguments = @{
                Service = 'MicrosoftGraph'
                Type    = 'Mail'
                From    = 'sender@example.org'
                To      = 'recipient@example.org'
                Subject = 'Test subject'
                Body    = 'Test body'
            }

            # A missing 'ChatType' must not stop a Mail send, so reaching the assertions at all is half of
            # what this checks.
            $Result = Send-ScriptMessage @MailArguments -ErrorAction Stop

            # A Mail send has no chat type and must not be told about one.
            $Result.Error | Should -BeNullOrEmpty
            Should -Invoke -ModuleName ScriptMessage Send-MgUserMail -Times 1 -Exactly
        }

        # The configuration allows Mail only, so that is what the caller has to change. A missing 'ChatType' is
        # not the reason this send cannot happen, and saying so would send them to the wrong setting.
        It 'reports the disallowed message type rather than the missing ChatType' {
            $ChatOnly = @{
                Service = 'MicrosoftGraph'
                Type    = 'Chat'
                From    = 'sender@example.org'
                To      = 'recipient@example.org'
                Subject = 'Test subject'
                Body    = 'Test body'
            }

            $Result = Send-ScriptMessage @ChatOnly -WarningAction SilentlyContinue

            @($Result.Error | Where-Object { $_.Message -like '*does not allow sending messages of type*' }).Count |
                Should -Be 1
            $Result.Error | Where-Object { $_.Message -like '*No chat type is set*' } | Should -BeNullOrEmpty
        }
    }

    Context 'A configuration that allows Chat but sets no ChatType' {
        BeforeAll {
            Mock -ModuleName ScriptMessage Get-ScriptMessageConfig {
                $Config = [pscustomobject]@{
                    MicrosoftGraph = [pscustomobject]@{
                        AllowableMessageTypes = @('Mail', 'Chat')
                        IncludeBCCInGroupChat = $false
                        MgPermissionType      = 'Delegated'
                        MgDisconnectWhenDone  = $false
                    }
                }
                if ($null -ne $Service) { $Config.$Service } else { $Config }
            }
        }

        It 'reports the missing chat type in the result instead of throwing' {
            $Result = Send-ScriptMessage @ChatArguments -WarningAction SilentlyContinue

            @($Result.Error | Where-Object { $_.Message -like '*No chat type is set*' }).Count | Should -Be 1
            Should -Invoke -ModuleName ScriptMessage New-MgChat -Times 0 -Exactly
        }

        It 'still sends Mail in the same call' {
            $Arguments = $ChatArguments.Clone()
            $Arguments['Type'] = @('Mail', 'Chat')

            $Result = @(Send-ScriptMessage @Arguments -WarningAction SilentlyContinue)

            @($Result | Where-Object { $_.MessageType -eq 'Mail' -and $null -eq $_.Error }).Count | Should -Be 1
            Should -Invoke -ModuleName ScriptMessage Send-MgUserMail -Times 1 -Exactly
        }
    }

    Context 'A configuration that leaves ChatType blank' {
        BeforeAll {
            # Templates/config_scriptmessage.json writes a setting it leaves unset as an empty string, so a
            # blank 'ChatType' has to read the same as a missing one.
            Mock -ModuleName ScriptMessage Get-ScriptMessageConfig {
                $Config = [pscustomobject]@{
                    MicrosoftGraph = [pscustomobject]@{
                        AllowableMessageTypes = @('Mail', 'Chat')
                        ChatType              = ''
                        IncludeBCCInGroupChat = $false
                        MgPermissionType      = 'Delegated'
                        MgDisconnectWhenDone  = $false
                    }
                }
                if ($null -ne $Service) { $Config.$Service } else { $Config }
            }
        }

        It 'reports the missing chat type rather than failing to convert an empty string' {
            $Result = Send-ScriptMessage @ChatArguments -WarningAction SilentlyContinue

            @($Result.Error | Where-Object { $_.Message -like '*No chat type is set*' }).Count | Should -Be 1
            $Result.Error | Where-Object { $_.Message -like '*identifier name*' } | Should -BeNullOrEmpty
            Should -Invoke -ModuleName ScriptMessage New-MgChat -Times 0 -Exactly
        }
    }
}

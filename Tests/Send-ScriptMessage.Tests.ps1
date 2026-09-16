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

    # Each context sets ChatType in the configuration rather than passing -ChatType, because the configuration
    # value currently wins over an explicitly passed parameter.
    Context 'A one-on-one chat with no attachments' {
        BeforeAll {
            Mock -ModuleName ScriptMessage Get-ScriptMessageConfig {
                $Config = [pscustomobject]@{
                    MicrosoftGraph = [pscustomobject]@{
                        AllowableMessageTypes = @('Mail', 'Chat')
                        MailType              = 'Group'
                        ChatType              = 'OneOnOne'
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
            $Result = Send-ScriptMessage @MessageArguments

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
                        ChatType              = 'Group'
                        IncludeBCCInGroupChat = $false
                        MgPermissionType      = 'Delegated'
                        MgDisconnectWhenDone  = $false
                    }
                }
                if ($null -ne $Service) { $Config.$Service } else { $Config }
            }
        }

        It 'sends the message body and reports no error' {
            $Result = Send-ScriptMessage @MessageArguments

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
                        ChatType              = 'OneOnOne'
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
}

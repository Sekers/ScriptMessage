# Pester 5 or later. Run from the repository root:
#   Invoke-Pester -Path .\Tests
# Nothing here connects to a tenant, and the Microsoft Graph modules do not need to be installed.

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\ScriptMessage\ScriptMessage.psd1') -Force
}

AfterAll {
    Remove-Module ScriptMessage -Force -ErrorAction SilentlyContinue
}

Describe 'The service table' {
    It 'holds the Microsoft Graph service with all four handlers' {
        InModuleScope ScriptMessage {
            $Registration = Get-ScriptMessageServiceRegistration -Service 'MicrosoftGraph'

            $Registration.ConnectFunction    | Should -Be 'Connect-ScriptMessage_MicrosoftGraph'
            $Registration.DisconnectFunction | Should -Be 'Disconnect-ScriptMessage_MicrosoftGraph'
            $Registration.SendFunction       | Should -Be 'Send-ScriptMessage_MicrosoftGraph'
            $Registration.GetContextFunction | Should -Be 'Get-ScriptMessageContext_MicrosoftGraph'
        }
    }

    It 'names the registered services when asked for one it does not have' {
        InModuleScope ScriptMessage {
            { Get-ScriptMessageServiceRegistration -Service 'NotAService' } |
                Should -Throw "*not a registered messaging service*MicrosoftGraph*"
        }
    }

    It 'refuses to register a service that is not a MessagingService member' {
        InModuleScope ScriptMessage {
            { Register-ScriptMessageService -Name 'Carrier Pigeon' -SupportedSetting @() `
                -ConnectFunction 'c' -DisconnectFunction 'd' -SendFunction 's' -GetContextFunction 'g' } |
                Should -Throw '*not a member of the MessagingService enum*'
        }
    }

    It 'records which generic settings Microsoft Graph honors, and which it does not' {
        InModuleScope ScriptMessage {
            $Supported = (Get-ScriptMessageServiceRegistration -Service 'MicrosoftGraph').SupportedSetting

            $Supported | Should -Contain 'ChatType'
            $Supported | Should -Contain 'IncludeBCCInGroupChat'
            $Supported | Should -Not -Contain 'MailType'
        }
    }
}

Describe 'A generic setting the service does not support' {
    BeforeAll {
        # Pester can only mock a command that exists, so the stand-in has to come before the mock.
        InModuleScope ScriptMessage {
            function script:Send-MgUserMail { param($UserId, $BodyParameter, [switch]$PassThru) throw 'Send-MgUserMail stand-in called without a mock.' }
        }

        Mock -ModuleName ScriptMessage Get-ScriptMessageConfig {
            $Config = [pscustomobject]@{
                MicrosoftGraph = [pscustomobject]@{
                    AllowableMessageTypes = @('Mail')
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

    It 'reports -MailType as unsupported and still sends' {
        $Warnings = @()
        $null = Send-ScriptMessage @MailArguments -MailType Group -WarningVariable Warnings -WarningAction SilentlyContinue

        ($Warnings | Out-String) | Should -BeLike "*does not support the 'MailType' parameter*"
        Should -Invoke -ModuleName ScriptMessage Send-MgUserMail -Times 1 -Exactly
    }

    It 'says nothing about a setting the service does support' {
        $Warnings = @()
        $null = Send-ScriptMessage @MailArguments -ChatType Group -WarningVariable Warnings -WarningAction SilentlyContinue

        ($Warnings | Out-String) | Should -Not -BeLike '*does not support*'
    }
}

Describe 'The public cmdlets dispatch through the service table' {
    # Re-point the Microsoft Graph entry at stand-in handlers and drive the public cmdlets with no Microsoft
    # Graph mocks loaded at all. If they still work, the core reaches a service only through the table. A fake
    # service cannot be registered under a new name while -Service is the MessagingService enum, which is why
    # this replaces an existing entry rather than adding one.
    BeforeAll {
        Mock -ModuleName ScriptMessage Get-ScriptMessageConfig {
            $Config = [pscustomobject]@{
                MicrosoftGraph = [pscustomobject]@{
                    AllowableMessageTypes = @('Mail')
                    IncludeBCCInGroupChat = $false
                    MgPermissionType      = 'Application'
                    MgDisconnectWhenDone  = $false
                }
            }
            if ($null -ne $Service) { $Config.$Service } else { $Config }
        }

        InModuleScope ScriptMessage {
            function script:Connect-StandInService { param($ServiceConfig) $script:StandInCalls += 'Connect' }
            function script:Disconnect-StandInService { $script:StandInCalls += 'Disconnect'; [pscustomobject]@{ StandIn = $true } }
            function script:Get-StandInServiceContext { $script:StandInCalls += 'GetContext'; [pscustomobject]@{ StandIn = $true } }
            function script:Send-StandInService { $script:StandInCalls += 'Send'; [pscustomobject]@{ StandIn = $true } }

            $script:OriginalRegistration = $script:ScriptMessageServiceTable['MicrosoftGraph']
            $script:ScriptMessageServiceTable['MicrosoftGraph'] = [PSCustomObject]@{
                Name               = 'MicrosoftGraph'
                ConnectFunction    = 'Connect-StandInService'
                DisconnectFunction = 'Disconnect-StandInService'
                SendFunction       = 'Send-StandInService'
                GetContextFunction = 'Get-StandInServiceContext'
            }
        }
    }

    AfterAll {
        # Leaving a stand-in in the table would break every later test file that expects the real service.
        InModuleScope ScriptMessage {
            $script:ScriptMessageServiceTable['MicrosoftGraph'] = $script:OriginalRegistration
        }
    }

    BeforeEach {
        InModuleScope ScriptMessage { $script:StandInCalls = @() }
    }

    It 'Connect-ScriptMessage reaches the registered connect handler' {
        $null = Connect-ScriptMessage -Service MicrosoftGraph

        InModuleScope ScriptMessage { $script:StandInCalls | Should -Contain 'Connect' }
    }

    It 'Disconnect-ScriptMessage reaches the registered disconnect handler' {
        $Info = Disconnect-ScriptMessage -Service MicrosoftGraph -ReturnConnectionInfo

        $Info.Service | Should -Be 'MicrosoftGraph'
        $Info.StandIn | Should -BeTrue
        InModuleScope ScriptMessage { $script:StandInCalls | Should -Contain 'Disconnect' }
    }

    It 'Get-ScriptMessageContext reaches the registered context handler' {
        $Context = Get-ScriptMessageContext -Service MicrosoftGraph

        $Context.Service | Should -Be 'MicrosoftGraph'
        $Context.StandIn | Should -BeTrue
        InModuleScope ScriptMessage { $script:StandInCalls | Should -Contain 'GetContext' }
    }

    It 'Send-ScriptMessage reaches the registered send handler, connecting on the way' {
        $Result = Send-ScriptMessage -Service MicrosoftGraph -Type Mail -From 'sender@example.org' `
            -To 'recipient@example.org' -Subject 'Test subject' -Body 'Test body'

        $Result.StandIn | Should -BeTrue
        InModuleScope ScriptMessage {
            $script:StandInCalls | Should -Contain 'Connect'
            $script:StandInCalls | Should -Contain 'Send'
        }
    }
}

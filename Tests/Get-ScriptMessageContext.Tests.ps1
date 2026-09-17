# Pester 5 or later. Run from the repository root:
#   Invoke-Pester -Path .\Tests
# Nothing here connects to a tenant, and the Microsoft Graph modules do not need to be installed.

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\ScriptMessage\ScriptMessage.psd1') -Force

    # Pester can only mock a command that exists, so define a stand-in for the Microsoft Graph cmdlet inside the
    # module before mocking it. It throws if a test reaches it without a mock.
    InModuleScope ScriptMessage {
        function script:Get-MgContext { throw 'Get-MgContext stand-in called without a mock.' }
    }
}

AfterAll {
    Remove-Module ScriptMessage -Force -ErrorAction SilentlyContinue
}

Describe 'Get-ScriptMessageContext' {
    Context 'A service with an active connection' {
        BeforeAll {
            Mock -ModuleName ScriptMessage Get-ScriptMessageContext_MicrosoftGraph {
                [pscustomobject]@{
                    Account  = 'sender@example.org'
                    Scopes   = @('Mail.Send', 'Chat.ReadWrite')
                    TenantId = 'tenant-id'
                }
            }
        }

        It 'returns the service name alongside the properties the service supplied' {
            $Context = Get-ScriptMessageContext -Service MicrosoftGraph

            $Context.Service  | Should -Be 'MicrosoftGraph'
            $Context.Account  | Should -Be 'sender@example.org'
            $Context.TenantId | Should -Be 'tenant-id'
            $Context.Scopes   | Should -Contain 'Chat.ReadWrite'
        }

        It 'asks the service rather than calling Microsoft Graph itself' {
            $null = Get-ScriptMessageContext -Service MicrosoftGraph

            Should -Invoke -ModuleName ScriptMessage Get-ScriptMessageContext_MicrosoftGraph -Times 1 -Exactly
        }

        It 'returns the cached context without asking the service again' {
            # The first call has no cache to read, so it refreshes and populates one.
            $null = Get-ScriptMessageContext -Service MicrosoftGraph
            $Cached = Get-ScriptMessageContext -Service MicrosoftGraph -ReturnCachedContext

            $Cached.Account | Should -Be 'sender@example.org'
            Should -Invoke -ModuleName ScriptMessage Get-ScriptMessageContext_MicrosoftGraph -Times 1 -Exactly
        }
    }

    Context 'A service with no active connection' {
        BeforeAll {
            Mock -ModuleName ScriptMessage Get-ScriptMessageContext_MicrosoftGraph { }
        }

        It 'still returns the service name and nothing else' {
            $Context = Get-ScriptMessageContext -Service MicrosoftGraph

            $Context.Service | Should -Be 'MicrosoftGraph'
            @($Context.PSObject.Properties).Count | Should -Be 1
        }
    }

    Context 'The Microsoft Graph service function' {
        BeforeAll {
            Mock -ModuleName ScriptMessage Get-MgContext { [pscustomobject]@{ Account = 'sender@example.org' } }
        }

        It 'is the only place that calls Get-MgContext' {
            $Context = Get-ScriptMessageContext -Service MicrosoftGraph

            $Context.Account | Should -Be 'sender@example.org'
            Should -Invoke -ModuleName ScriptMessage Get-MgContext -Times 1 -Exactly
        }
    }
}

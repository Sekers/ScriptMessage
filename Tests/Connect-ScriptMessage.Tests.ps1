# Pester 5 or later. Run from the repository root:
#   Invoke-Pester -Path .\Tests
# Nothing here connects to a tenant, and the Microsoft Graph modules do not need to be installed.

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\ScriptMessage\ScriptMessage.psd1') -Force

    # Pester can only mock a command that exists, so define a stand-in for the Microsoft Graph cmdlet inside the
    # module before mocking it. It throws if a test reaches it without a mock.
    InModuleScope ScriptMessage {
        function script:Connect-MgGraph { param($Scopes, $TenantId, $ClientId, $Certificate, $CertificateName, $CertificateThumbprint, $ClientSecretCredential) throw 'Connect-MgGraph stand-in called without a mock.' }
    }
}

AfterAll {
    Remove-Module ScriptMessage -Force -ErrorAction SilentlyContinue
}

Describe 'Connect-ScriptMessage Microsoft Graph module check' {
    BeforeAll {
        Mock -ModuleName ScriptMessage Connect-MgGraph { }
        Mock -ModuleName ScriptMessage Import-Module { }

        # Only the Microsoft Graph sub-modules a chat-only setup needs are loaded. New-Module builds real module
        # objects without importing them, so the check compares against the same type Get-Module returns.
        Mock -ModuleName ScriptMessage Get-Module {
            $Modules = @(
                New-Module -Name 'Microsoft.Graph.Authentication' -ScriptBlock { }
                New-Module -Name 'Microsoft.Graph.Teams' -ScriptBlock { }
            )
            if ($Name) { $Modules | Where-Object { $_.Name -in $Name } } else { $Modules }
        }
    }

    Context 'A configuration that allows only Chat' {
        BeforeAll {
            Mock -ModuleName ScriptMessage Get-ScriptMessageConfig {
                [pscustomobject]@{
                    AllowableMessageTypes = @('Chat')
                    MgPermissionType      = 'Delegated'
                    MgTenantID            = 'tenant-id'
                    MgClientID            = 'client-id'
                }
            }
        }

        It 'connects without Microsoft.Graph.Users.Actions' {
            { Connect-ScriptMessage -Service MicrosoftGraph -ErrorAction Stop } | Should -Not -Throw
            Should -Invoke -ModuleName ScriptMessage Connect-MgGraph -Times 1 -Exactly
        }
    }

    Context 'A configuration that allows Mail' {
        BeforeAll {
            Mock -ModuleName ScriptMessage Get-ScriptMessageConfig {
                [pscustomobject]@{
                    AllowableMessageTypes = @('Mail', 'Chat')
                    MgPermissionType      = 'Delegated'
                    MgTenantID            = 'tenant-id'
                    MgClientID            = 'client-id'
                }
            }
        }

        It 'reports Microsoft.Graph.Users.Actions as missing and does not connect' {
            { Connect-ScriptMessage -Service MicrosoftGraph -ErrorAction Stop } | Should -Throw '*: Microsoft.Graph.Users.Actions'
            Should -Invoke -ModuleName ScriptMessage Connect-MgGraph -Times 0 -Exactly
        }
    }
}

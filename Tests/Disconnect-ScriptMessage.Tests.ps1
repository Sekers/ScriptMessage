# Pester 5 or later. Run from the repository root:
#   Invoke-Pester -Path .\Tests
# Nothing here connects to a tenant, and the Microsoft Graph modules do not need to be installed. Mocking the
# service functions rather than the Microsoft Graph cmdlets they call keeps that true without stand-ins.

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\ScriptMessage\ScriptMessage.psd1') -Force
}

AfterAll {
    Remove-Module ScriptMessage -Force -ErrorAction SilentlyContinue
}

Describe 'Disconnect-ScriptMessage' {
    BeforeAll {
        Mock -ModuleName ScriptMessage Disconnect-ScriptMessage_MicrosoftGraph {
            [pscustomobject]@{ Account = 'sender@example.org' }
        }
        Mock -ModuleName ScriptMessage Get-ScriptMessageContext_MicrosoftGraph {
            [pscustomobject]@{ Account = 'sender@example.org' }
        }
    }

    It 'asks the service to disconnect' {
        $null = Disconnect-ScriptMessage -Service MicrosoftGraph

        Should -Invoke -ModuleName ScriptMessage Disconnect-ScriptMessage_MicrosoftGraph -Times 1 -Exactly
    }

    It 'returns the service name with the disconnection information when asked' {
        $Info = Disconnect-ScriptMessage -Service MicrosoftGraph -ReturnConnectionInfo

        $Info.Service | Should -Be 'MicrosoftGraph'
        $Info.Account | Should -Be 'sender@example.org'
    }

    It 'returns nothing unless -ReturnConnectionInfo is given' {
        Disconnect-ScriptMessage -Service MicrosoftGraph | Should -BeNullOrEmpty
    }

    It 'clears the cached context, so -ReturnCachedContext stops reporting the old connection' {
        # Populate the cache, disconnect, then read it back. A cleared cache forces the read to ask the service
        # again, which is the only way a caller can tell that the cached connection is gone.
        $null = Get-ScriptMessageContext -Service MicrosoftGraph
        $null = Disconnect-ScriptMessage -Service MicrosoftGraph
        $null = Get-ScriptMessageContext -Service MicrosoftGraph -ReturnCachedContext

        Should -Invoke -ModuleName ScriptMessage Get-ScriptMessageContext_MicrosoftGraph -Times 2 -Exactly
    }
}

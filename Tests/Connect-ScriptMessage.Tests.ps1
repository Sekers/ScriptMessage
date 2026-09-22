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

# Certificate file authentication throws before PowerShell 7.4, so there is no path to check there.
Describe 'Connect-ScriptMessage certificate file path' -Skip:($PSVersionTable.PSVersion -lt [version]'7.4') {
    BeforeAll {
        $env:SCRIPTMESSAGE_TEST_CERTS = 'C:\Certs'

        Mock -ModuleName ScriptMessage Connect-MgGraph { }
        Mock -ModuleName ScriptMessage Import-Module { }
        Mock -ModuleName ScriptMessage Get-PfxCertificate { }
        Mock -ModuleName ScriptMessage Get-Module {
            $Modules = @(
                New-Module -Name 'Microsoft.Graph.Authentication' -ScriptBlock { }
                New-Module -Name 'Microsoft.Graph.Teams' -ScriptBlock { }
            )
            if ($Name) { $Modules | Where-Object { $_.Name -in $Name } } else { $Modules }
        }
    }

    AfterAll {
        Remove-Item -Path 'Env:\SCRIPTMESSAGE_TEST_CERTS' -ErrorAction SilentlyContinue
    }

    It 'opens <Expected> when the setting is <Setting>' -ForEach @(
        @{ Setting = '$env:SCRIPTMESSAGE_TEST_CERTS\app.pfx';        Expected = 'C:\Certs\app.pfx' }
        @{ Setting = '${env:SCRIPTMESSAGE_TEST_CERTS}\app.pfx';      Expected = 'C:\Certs\app.pfx' }
        @{ Setting = 'C:\Certs\$([string]::Concat(''a'', ''b'')).pfx'; Expected = 'C:\Certs\$([string]::Concat(''a'', ''b'')).pfx' }
        @{ Setting = 'C:\Certs\$app\app.pfx';                        Expected = 'C:\Certs\$app\app.pfx' }
    ) {
        $ServiceConfig = [pscustomobject]@{
            AllowableMessageTypes    = @('Chat')
            MgPermissionType         = 'Application'
            MgTenantID               = 'tenant-id'
            MgClientID               = 'client-id'
            MgApp_AuthenticationType = 'CertificateFile'
            MgApp_CertificatePath    = $Setting
        }

        Connect-ScriptMessage -Service MicrosoftGraph -ServiceConfig $ServiceConfig -ErrorAction Stop
        Should -Invoke -ModuleName ScriptMessage Get-PfxCertificate -Times 1 -Exactly -ParameterFilter { $FilePath -eq $Expected }
    }
}

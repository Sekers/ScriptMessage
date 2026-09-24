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

    # Hands settings straight to the Microsoft Graph connect function, for tests of what it does with them.
    # Connect-ScriptMessage always reads them from the configuration file.
    function Invoke-MicrosoftGraphConnect
    {
        param([pscustomobject]$ServiceConfig)
        InModuleScope ScriptMessage -Parameters @{ ServiceConfig = $ServiceConfig } {
            param($ServiceConfig)
            Connect-ScriptMessage_MicrosoftGraph -ServiceConfig $ServiceConfig -ErrorAction Stop
        }
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

# The certificate is loaded with New-Object, so mocking it checks the path without a certificate file. The last Describe
# loads real files.
Describe 'Connect-ScriptMessage certificate file path' {
    BeforeAll {
        $env:SCRIPTMESSAGE_TEST_CERTS = 'C:\Certs'

        Mock -ModuleName ScriptMessage Connect-MgGraph { }
        Mock -ModuleName ScriptMessage Import-Module { }
        Mock -ModuleName ScriptMessage New-Object { [pscustomobject]@{ Thumbprint = 'stand-in' } } -ParameterFilter {
            $TypeName -eq 'System.Security.Cryptography.X509Certificates.X509Certificate2'
        }
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

        Invoke-MicrosoftGraphConnect -ServiceConfig $ServiceConfig
        Should -Invoke -ModuleName ScriptMessage New-Object -Times 1 -Exactly -ParameterFilter { $ArgumentList[0] -eq $Expected }
    }

    Context 'Opening the certificate file' {
        BeforeAll {
            $CertificateFile = Join-Path $TestDrive 'app.pfx'
            $null = New-Item -ItemType File -Path $CertificateFile
            $ServiceConfig = [pscustomobject]@{
                AllowableMessageTypes              = @('Chat')
                MgPermissionType                   = 'Application'
                MgTenantID                         = 'tenant-id'
                MgClientID                         = 'client-id'
                MgApp_AuthenticationType           = 'CertificateFile'
                MgApp_CertificatePath              = $CertificateFile
                MgApp_EncryptedCertificatePassword = ''
            }
        }

        # Any other storage option writes the private key to the user profile, and PowerShell 7 leaves it there.
        It 'keeps the private key in memory' {
            Invoke-MicrosoftGraphConnect -ServiceConfig $ServiceConfig

            Should -Invoke -ModuleName ScriptMessage New-Object -Times 1 -Exactly -ParameterFilter {
                $ArgumentList[2] -eq [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet
            }
        }

        Context 'A file that cannot be opened without a password' {
            BeforeAll {
                Mock -ModuleName ScriptMessage New-Object { throw 'The specified network password is not correct.' } -ParameterFilter {
                    $TypeName -eq 'System.Security.Cryptography.X509Certificates.X509Certificate2' -and $ArgumentList[1] -is [string]
                }
            }

            It 'opens it with the password in MgApp_EncryptedCertificatePassword' {
                $WithPassword = $ServiceConfig.PSObject.Copy()
                $WithPassword.MgApp_EncryptedCertificatePassword = ConvertTo-SecureString -String 'test-password' -AsPlainText -Force | ConvertFrom-SecureString

                Invoke-MicrosoftGraphConnect -ServiceConfig $WithPassword

                Should -Invoke -ModuleName ScriptMessage New-Object -Times 1 -Exactly -ParameterFilter { $ArgumentList[1] -is [securestring] }
                Should -Invoke -ModuleName ScriptMessage Connect-MgGraph -Times 1 -Exactly
            }

            It 'says so when MgApp_EncryptedCertificatePassword is empty' {
                { Invoke-MicrosoftGraphConnect -ServiceConfig $ServiceConfig } |
                    Should -Throw '*no password has been provided*'
                Should -Invoke -ModuleName ScriptMessage Connect-MgGraph -Times 0 -Exactly
            }

            It 'names a file that does not exist rather than asking for a password' {
                $Missing = $ServiceConfig.PSObject.Copy()
                $Missing.MgApp_CertificatePath = Join-Path $TestDrive 'missing.pfx'

                { Invoke-MicrosoftGraphConnect -ServiceConfig $Missing } |
                    Should -Throw "*does not exist: '$(Join-Path $TestDrive 'missing.pfx')'*"
                Should -Invoke -ModuleName ScriptMessage Connect-MgGraph -Times 0 -Exactly
            }
        }
    }

    Context 'A configuration file with a relative path' {
        BeforeAll {
            # The configuration file and PowerShell's current location are in different folders, so each test can
            # put the certificate file in either one, both, or neither.
            $ConfigFolder = Join-Path $TestDrive 'Config'
            $CurrentFolder = Join-Path $TestDrive 'Current'
            $null = New-Item -ItemType Directory -Path $ConfigFolder, $CurrentFolder
            $ConfigFile = Join-Path $ConfigFolder 'config_scriptmessage.json'
            Set-ScriptMessageConfigFilePath -Path $ConfigFile

            function Set-TestCertificatePath
            {
                param([string]$CertificatePath)
                @{
                    MicrosoftGraph = @{
                        AllowableMessageTypes    = @('Chat')
                        MgPermissionType         = 'Application'
                        MgTenantID               = 'tenant-id'
                        MgClientID               = 'client-id'
                        MgApp_AuthenticationType = 'CertificateFile'
                        MgApp_CertificatePath    = $CertificatePath
                    }
                } | ConvertTo-Json -Depth 3 | Set-Content -Path $ConfigFile
            }

            Push-Location -Path $CurrentFolder
        }

        AfterAll {
            Pop-Location
        }

        BeforeEach {
            Set-TestCertificatePath 'app.pfx'
            Remove-Item -Path (Join-Path $ConfigFolder 'app.pfx'), (Join-Path $CurrentFolder 'app.pfx') -ErrorAction SilentlyContinue
        }

        It 'opens the file in the configuration file''s folder' {
            $null = New-Item -ItemType File -Path (Join-Path $ConfigFolder 'app.pfx')

            Connect-ScriptMessage -Service MicrosoftGraph -ErrorAction Stop -WarningVariable Warnings

            $Warnings | Should -BeNullOrEmpty
            Should -Invoke -ModuleName ScriptMessage New-Object -Times 1 -Exactly -ParameterFilter {
                $ArgumentList[0] -eq (Join-Path $ConfigFolder 'app.pfx')
            }
        }

        It 'prefers the configuration file''s folder when the current location has the file too' {
            $null = New-Item -ItemType File -Path (Join-Path $ConfigFolder 'app.pfx'), (Join-Path $CurrentFolder 'app.pfx')

            Connect-ScriptMessage -Service MicrosoftGraph -ErrorAction Stop -WarningVariable Warnings

            $Warnings | Should -BeNullOrEmpty
            Should -Invoke -ModuleName ScriptMessage New-Object -Times 1 -Exactly -ParameterFilter {
                $ArgumentList[0] -eq (Join-Path $ConfigFolder 'app.pfx')
            }
        }

        It 'still opens a file found only in the current location, with a deprecation warning' {
            $null = New-Item -ItemType File -Path (Join-Path $CurrentFolder 'app.pfx')

            Connect-ScriptMessage -Service MicrosoftGraph -ErrorAction Stop -WarningVariable Warnings -WarningAction SilentlyContinue

            @($Warnings).Count | Should -Be 1
            "$Warnings" | Should -BeLike '*current location is deprecated*'
            Should -Invoke -ModuleName ScriptMessage New-Object -Times 1 -Exactly -ParameterFilter {
                $ArgumentList[0] -eq (Join-Path $CurrentFolder 'app.pfx')
            }
        }

        It 'names the configuration file''s folder when neither has the file' {
            Connect-ScriptMessage -Service MicrosoftGraph -ErrorAction Stop -WarningVariable Warnings

            $Warnings | Should -BeNullOrEmpty
            Should -Invoke -ModuleName ScriptMessage New-Object -Times 1 -Exactly -ParameterFilter {
                $ArgumentList[0] -eq (Join-Path $ConfigFolder 'app.pfx')
            }
        }

        It 'resolves a path that leaves the configuration file''s folder' {
            Set-TestCertificatePath '..\Certs\app.pfx'

            Connect-ScriptMessage -Service MicrosoftGraph -ErrorAction Stop

            Should -Invoke -ModuleName ScriptMessage New-Object -Times 1 -Exactly -ParameterFilter {
                $ArgumentList[0] -eq (Join-Path $TestDrive 'Certs\app.pfx')
            }
        }

        # .NET does not count a PowerShell drive as absolute, and PowerShell does not count a UNC path or '~'.
        It 'does not look for <CertificatePath> in the configuration file''s folder' -ForEach @(
            @{ CertificatePath = '~\app.pfx' }
            @{ CertificatePath = 'TestDrive:\app.pfx' }
            @{ CertificatePath = '\\server.example.com\share\app.pfx' }
        ) {
            Set-TestCertificatePath $CertificatePath
            $Expected = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CertificatePath)

            Connect-ScriptMessage -Service MicrosoftGraph -ErrorAction Stop

            Should -Invoke -ModuleName ScriptMessage New-Object -Times 1 -Exactly -ParameterFilter {
                $ArgumentList[0] -eq $Expected
            }
        }

        It 'resolves from the configuration file''s folder when Send-ScriptMessage connects' {
            Mock -ModuleName ScriptMessage Send-ScriptMessage_MicrosoftGraph { }
            $null = New-Item -ItemType File -Path (Join-Path $ConfigFolder 'app.pfx')

            Send-ScriptMessage -Service MicrosoftGraph -Type Chat -From 'sender@example.org' -To 'recipient@example.org' `
                -Subject 'Test subject' -Body 'Test body' -ErrorAction Stop

            Should -Invoke -ModuleName ScriptMessage New-Object -Times 1 -Exactly -ParameterFilter {
                $ArgumentList[0] -eq (Join-Path $ConfigFolder 'app.pfx')
            }
            Should -Invoke -ModuleName ScriptMessage Send-ScriptMessage_MicrosoftGraph -Times 1 -Exactly
        }
    }
}

# Real certificate files, so the load itself is checked in each edition.
Describe 'Connect-ScriptMessage certificate file loading' {
    BeforeAll {
        Mock -ModuleName ScriptMessage Connect-MgGraph { }
        Mock -ModuleName ScriptMessage Import-Module { }
        Mock -ModuleName ScriptMessage Get-Module {
            $Modules = @(
                New-Module -Name 'Microsoft.Graph.Authentication' -ScriptBlock { }
                New-Module -Name 'Microsoft.Graph.Teams' -ScriptBlock { }
            )
            if ($Name) { $Modules | Where-Object { $_.Name -in $Name } } else { $Modules }
        }

        # A throwaway self-signed certificate, made in memory and saved with and without a password.
        $Rsa = [System.Security.Cryptography.RSA]::Create(2048)
        $Request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new('CN=ScriptMessage Test', $Rsa,
            [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $TestCertificate = $Request.CreateSelfSigned([DateTimeOffset]::Now.AddMinutes(-5), [DateTimeOffset]::Now.AddHours(1))
        $TestThumbprint = $TestCertificate.Thumbprint
        $Pfx = [System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx
        [System.IO.File]::WriteAllBytes((Join-Path $TestDrive 'with-password.pfx'), $TestCertificate.Export($Pfx, 'test-password'))
        [System.IO.File]::WriteAllBytes((Join-Path $TestDrive 'no-password.pfx'), $TestCertificate.Export($Pfx))
        $TestCertificate.Dispose()
        $Rsa.Dispose()
        $EncryptedPassword = ConvertTo-SecureString -String 'test-password' -AsPlainText -Force | ConvertFrom-SecureString

        # The folders Windows keeps a user's private keys in, so a test can see whether loading wrote one.
        $Sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        $KeyFolders = @("$env:APPDATA\Microsoft\Crypto\Keys", "$env:APPDATA\Microsoft\Crypto\RSA\$Sid")
        function Get-KeyFileNames
        {
            @(foreach ($Folder in $KeyFolders) { if (Test-Path -LiteralPath $Folder) { Get-ChildItem -LiteralPath $Folder -Force -File | ForEach-Object FullName } })
        }
    }

    It 'signs in with <File>, keeping its private key out of the user profile' -ForEach @(
        @{ File = 'with-password.pfx'; NeedsPassword = $true }
        @{ File = 'no-password.pfx'; NeedsPassword = $false }
    ) {
        $ServiceConfig = [pscustomobject]@{
            AllowableMessageTypes              = @('Chat')
            MgPermissionType                   = 'Application'
            MgTenantID                         = 'tenant-id'
            MgClientID                         = 'client-id'
            MgApp_AuthenticationType           = 'CertificateFile'
            MgApp_CertificatePath              = Join-Path $TestDrive $File
            MgApp_EncryptedCertificatePassword = $(if ($NeedsPassword) { $EncryptedPassword } else { '' })
        }
        $Before = Get-KeyFileNames

        Invoke-MicrosoftGraphConnect -ServiceConfig $ServiceConfig

        # The Connect-MgGraph mock still holds the certificate, so a key file written while loading it would be here.
        @(Get-KeyFileNames | Where-Object { $_ -notin $Before }) | Should -BeNullOrEmpty
        Should -Invoke -ModuleName ScriptMessage Connect-MgGraph -Times 1 -Exactly -ParameterFilter {
            $Certificate.Thumbprint -eq $TestThumbprint -and $Certificate.HasPrivateKey
        }
    }
}

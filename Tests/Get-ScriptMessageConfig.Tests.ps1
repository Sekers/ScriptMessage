# Pester 5 or later. Run from the repository root:
#   Invoke-Pester -Path .\Tests
# Nothing here connects to a tenant. Every configuration file is written to Pester's TestDrive.

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\ScriptMessage\ScriptMessage.psd1') -Force

    # U+00E9 (e with an acute accent) exists in most legacy code pages. U+4E2D (a CJK character) exists in no
    # single-byte code page, so it only survives a Unicode encoding.
    $AccentedName = 'Soci' + [char]0x00E9 + 't' + [char]0x00E9
    $UnicodeName = $AccentedName + ' ' + [char]0x4E2D

    # Writes a configuration file holding one MgApp_CertificateName, as exactly the given preamble and encoding.
    function New-TestConfigFile
    {
        param(
            [string]$Path,
            [string]$CertificateName,
            [System.Text.Encoding]$Encoding,
            [byte[]]$Preamble = @()
        )

        $Json = '{ "MicrosoftGraph": { "MgApp_CertificateName": "' + $CertificateName + '" } }'
        [System.IO.File]::WriteAllBytes($Path, [byte[]]($Preamble + $Encoding.GetBytes($Json)))
    }
}

AfterAll {
    Remove-Module ScriptMessage -Force -ErrorAction SilentlyContinue
}

Describe 'Get-ScriptMessageConfig file encodings' {
    It 'reads UTF-8 without a byte order mark' {
        $Path = Join-Path $TestDrive 'utf8-no-bom.json'
        New-TestConfigFile -Path $Path -CertificateName $UnicodeName -Encoding ([System.Text.UTF8Encoding]::new($false))

        (Get-ScriptMessageConfig -Path $Path -Service MicrosoftGraph).MgApp_CertificateName | Should -BeExactly $UnicodeName
    }

    It 'reads UTF-8 with a byte order mark' {
        $Path = Join-Path $TestDrive 'utf8-bom.json'
        New-TestConfigFile -Path $Path -CertificateName $UnicodeName -Encoding ([System.Text.UTF8Encoding]::new($false)) -Preamble 0xEF, 0xBB, 0xBF

        (Get-ScriptMessageConfig -Path $Path -Service MicrosoftGraph).MgApp_CertificateName | Should -BeExactly $UnicodeName
    }

    It 'reads UTF-16 LE with a byte order mark' {
        $Path = Join-Path $TestDrive 'utf16le-bom.json'
        New-TestConfigFile -Path $Path -CertificateName $UnicodeName -Encoding ([System.Text.UnicodeEncoding]::new($false, $false)) -Preamble 0xFF, 0xFE

        (Get-ScriptMessageConfig -Path $Path -Service MicrosoftGraph).MgApp_CertificateName | Should -BeExactly $UnicodeName
    }

    It 'reads a file saved in the system legacy code page' {
        $Legacy = [System.Text.Encoding]::GetEncoding(0)
        $LegacyBytes = $Legacy.GetBytes($AccentedName)
        if (($Legacy.CodePage -eq 65001) -or -not ($LegacyBytes | Where-Object { $_ -gt 0x7F }))
        {
            Set-ItResult -Skipped -Because "code page $($Legacy.CodePage) does not store U+00E9 as a byte that is invalid UTF-8"
        }

        $Path = Join-Path $TestDrive 'legacy-code-page.json'
        New-TestConfigFile -Path $Path -CertificateName $AccentedName -Encoding $Legacy

        (Get-ScriptMessageConfig -Path $Path -Service MicrosoftGraph).MgApp_CertificateName | Should -BeExactly $AccentedName
    }
}

Describe 'Get-ScriptMessageConfig paths' {
    BeforeAll {
        # ASCII only, so these tests depend on how the path is found and not on how the file is decoded.
        $AsciiName = 'Path test'
    }

    It 'reads a path relative to the current PowerShell location' {
        New-TestConfigFile -Path (Join-Path $TestDrive 'relative.json') -CertificateName $AsciiName -Encoding ([System.Text.UTF8Encoding]::new($false))

        Push-Location -LiteralPath $TestDrive
        try
        {
            (Get-ScriptMessageConfig -Path '.\relative.json' -Service MicrosoftGraph).MgApp_CertificateName | Should -BeExactly $AsciiName
        }
        finally
        {
            Pop-Location
        }
    }

    It 'reads a path on a PowerShell drive' {
        New-TestConfigFile -Path (Join-Path $TestDrive 'drive.json') -CertificateName $AsciiName -Encoding ([System.Text.UTF8Encoding]::new($false))

        (Get-ScriptMessageConfig -Path 'TestDrive:\drive.json' -Service MicrosoftGraph).MgApp_CertificateName | Should -BeExactly $AsciiName
    }

    It 'reads a path containing square brackets' {
        $Folder = [System.IO.Directory]::CreateDirectory((Join-Path $TestDrive 'Config [old]')).FullName
        $Path = Join-Path $Folder 'brackets.json'
        New-TestConfigFile -Path $Path -CertificateName $AsciiName -Encoding ([System.Text.UTF8Encoding]::new($false))

        (Get-ScriptMessageConfig -Path $Path -Service MicrosoftGraph).MgApp_CertificateName | Should -BeExactly $AsciiName
    }

    It 'reports a file that does not exist' {
        { Get-ScriptMessageConfig -Path (Join-Path $TestDrive 'missing.json') } | Should -Throw -ExpectedMessage "Can't find the JSON configuration file*"
    }
}

Describe 'Get-ScriptMessageConfig errors' {
    It 'reports a folder that does not exist' {
        { Get-ScriptMessageConfig -Path (Join-Path $TestDrive 'No Such Folder\config.json') } |
            Should -Throw -ExpectedMessage "Can't find the JSON configuration file*No Such Folder*"
    }

    It 'reports a path it cannot read as a file' {
        { Get-ScriptMessageConfig -Path $TestDrive } | Should -Throw -ExpectedMessage "Can't read the JSON configuration file*"
    }

    # Windows PowerShell 5.1 ends its message for this mistake with the whole file, secrets included.
    It 'reports a file that is not valid JSON, without repeating its contents' {
        $Path = Join-Path $TestDrive 'invalid.json'
        [System.IO.File]::WriteAllText($Path, '{ "MicrosoftGraph": { "MgClientID" "not-a-real-secret" } }')

        $ErrorRecord = { Get-ScriptMessageConfig -Path $Path } | Should -Throw -ExpectedMessage "The configuration file '*invalid.json' is not valid JSON.*" -PassThru
        $ErrorRecord.Exception.Message | Should -Not -BeLike '*not-a-real-secret*'
    }
}

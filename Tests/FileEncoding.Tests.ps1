# Pester 5 or later. Run .\Tests\Invoke-Tests.ps1 to test under both PowerShell editions, or
# Invoke-Pester -Path .\Tests for the current one. Nothing here imports the module or connects to a tenant: it
# reads the repository's files from disk and checks them against the file encoding rules in AGENTS.md.
#
# Git will not report these problems itself. It normalizes line endings before diffing, so a file whose endings
# are wrong produces no diff, and it says nothing about a byte order mark or a non-ASCII character.

BeforeAll {
    $RepoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).ProviderPath

    # Tracked files plus untracked files that are not ignored, so a new file is checked before it is committed
    # while .gitignore (which keeps real configuration files out) is still honored. -z stops Git quoting paths.
    $Listed = @((& git -C $RepoRoot ls-files -z --cached --others --exclude-standard) -join '' -split "`0" |
        Where-Object { $_ } | Sort-Object -Unique)
    if (($LASTEXITCODE -ne 0) -or ($Listed.Count -eq 0))
    {
        throw "git ls-files listed nothing under '$RepoRoot' (exit code $LASTEXITCODE). These tests need Git and a clone of the repository."
    }

    # ISO-8859-1 maps each byte to the character with the same number, so a regex over the result sees the raw
    # bytes, and a byte above 0x7F is a non-ASCII character in any encoding this repository could hold.
    $Latin1 = [System.Text.Encoding]::GetEncoding(28591)

    $Checked = [System.Collections.Generic.List[string]]::new()
    $WithBom = [System.Collections.Generic.List[string]]::new()
    $WithCR = [System.Collections.Generic.List[string]]::new()
    $NonAsciiPowerShell = [System.Collections.Generic.List[string]]::new()

    foreach ($Relative in $Listed)
    {
        $Path = Join-Path $RepoRoot $Relative
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf))
        {
            continue # Tracked, but deleted in the working tree.
        }

        $Bytes = [System.IO.File]::ReadAllBytes($Path)

        # A NUL byte is Git's own test for a binary file. Git applies no line-ending rule to one, and none of
        # these rules apply either.
        if ([Array]::IndexOf($Bytes, [byte]0) -ge 0)
        {
            continue
        }
        $Checked.Add($Relative)

        $HasBom = ($Bytes.Length -ge 3) -and ($Bytes[0] -eq 0xEF) -and ($Bytes[1] -eq 0xBB) -and ($Bytes[2] -eq 0xBF)
        $ContentStart = if ($HasBom) { 3 } else { 0 }
        $HasNonAscii = $Latin1.GetString($Bytes, $ContentStart, $Bytes.Length - $ContentStart) -match '[^\x00-\x7F]'

        # The one exception to both rules: a module manifest cannot write a non-ASCII character as an escape, so
        # one that needs such a character is saved with a BOM, which both editions read correctly.
        $IsManifestNeedingBom = ($Relative -like '*.psd1') -and $HasBom -and $HasNonAscii

        if ($HasBom -and -not $IsManifestNeedingBom)
        {
            $WithBom.Add($Relative)
        }
        if ([Array]::IndexOf($Bytes, [byte]13) -ge 0)
        {
            $WithCR.Add($Relative)
        }
        if (($Relative -match '\.(ps1|psm1|psd1|ps1xml)$') -and $HasNonAscii -and -not $IsManifestNeedingBom)
        {
            $NonAsciiPowerShell.Add($Relative)
        }
    }
}

Describe 'Repository file encoding' {
    # A check that silently reads nothing would pass every test below.
    It 'reads the repository''s text files' {
        $Checked | Should -Contain 'ScriptMessage/ScriptMessage.psm1'
        $Checked | Should -Contain 'CHANGELOG.md'
    }

    It 'no text file starts with a byte order mark' {
        $WithBom | Should -BeNullOrEmpty -Because 'text files are UTF-8 without a BOM, except a module manifest that holds a non-ASCII character'
    }

    It 'every text file uses LF line endings' {
        $WithCR | Should -BeNullOrEmpty -Because '.gitattributes pins every text file to eol=lf; AGENTS.md shows how to convert one'
    }

    It 'PowerShell files contain only ASCII' {
        $NonAsciiPowerShell | Should -BeNullOrEmpty -Because 'Windows PowerShell 5.1 misreads a non-ASCII character in a file without a BOM; write it as an escape such as [char]0x00E9'
    }
}

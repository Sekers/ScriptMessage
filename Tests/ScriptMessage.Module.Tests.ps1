# Pester 5 or later. Run .\Tests\Invoke-Tests.ps1 to test under both PowerShell editions, or
# Invoke-Pester -Path .\Tests for the current one. Nothing here connects to a tenant.

BeforeAll {
    $ModuleRoot = Join-Path $PSScriptRoot '..\ScriptMessage'
    $ManifestPath = Join-Path $ModuleRoot 'ScriptMessage.psd1'
}

Describe 'ScriptMessage module' {
    It 'has a valid module manifest' {
        { Test-ModuleManifest -Path $ManifestPath -ErrorAction Stop } | Should -Not -Throw
    }

    It 'exports exactly the functions defined in Public' {
        $PublicFunctions = @(Get-ChildItem -Path (Join-Path $ModuleRoot 'Public') -Filter '*.ps1' | ForEach-Object { $_.BaseName } | Sort-Object)
        $ExportedFunctions = @((Import-PowerShellDataFile -Path $ManifestPath).FunctionsToExport | Sort-Object)

        ($ExportedFunctions -join ', ') | Should -Be ($PublicFunctions -join ', ')
    }

    # PSScriptAnalyzer is optional locally. The check runs under PowerShell 7 only, because Windows PowerShell 5.1
    # cannot see the copy the CI workflow installs; the analyzer has the same rules in both editions, and syntax
    # that only PowerShell 7 accepts already fails the 5.1 run when the module imports.
    #
    # PSSCRIPTANALYZER_VERSION, when set, names the exact version to use. The workflow sets it to the version it
    # installs, because the runner image ships its own copy and PowerShell would otherwise load whichever copy it
    # finds first. Without it, the newest installed version is used. In CI, where GitHub sets CI=true, a missing
    # analyzer fails rather than skips, so the check cannot quietly disappear if the install step breaks.
    #
    # ParseError is its own severity, so without it a syntax error would report nothing.
    #
    # Research_Notes/Test-Tooling-Behavior.md records the measurements behind these choices.
    It 'has no PSScriptAnalyzer errors' {
        if ($PSVersionTable.PSEdition -ne 'Core')
        {
            Set-ItResult -Skipped -Because 'the analyzer check runs under PowerShell 7 only'
        }

        $Installed = @(Get-Module -ListAvailable -Name PSScriptAnalyzer)
        if ($env:PSSCRIPTANALYZER_VERSION)
        {
            $Analyzer = $Installed | Where-Object { $_.Version -eq [version]$env:PSSCRIPTANALYZER_VERSION } | Select-Object -First 1
            $Missing = "PSScriptAnalyzer $env:PSSCRIPTANALYZER_VERSION is not installed"
        }
        else
        {
            $Analyzer = $Installed | Sort-Object -Property Version -Descending | Select-Object -First 1
            $Missing = 'PSScriptAnalyzer is not installed'
        }
        if (-not $Analyzer)
        {
            if ($env:CI -eq 'true')
            {
                throw "$Missing, and CI must run this check."
            }
            Set-ItResult -Skipped -Because $Missing
        }

        Import-Module -Name PSScriptAnalyzer -RequiredVersion $Analyzer.Version
        $Findings = @(Invoke-ScriptAnalyzer -Path $ModuleRoot -Recurse -Severity Error, ParseError)

        @($Findings | ForEach-Object { '{0}:{1} {2}' -f $_.ScriptName, $_.Line, $_.RuleName }) | Should -BeNullOrEmpty
    }
}

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

    # A new runspace has its own global scope, so variables left by the other test files' imports cannot hide one
    # created here. The script runs in a local scope so that its own variables are not global either.
    It 'creates no global variables' {
        $Script = {
            param($ManifestPath, $ConfigPath)

            $Before = @(Get-Variable -Scope Global | ForEach-Object { $_.Name })
            Import-Module $ManifestPath
            Set-ScriptMessageConfigFilePath -Path $ConfigPath
            @(Get-Variable -Scope Global | Where-Object { $_.Name -notin $Before } | ForEach-Object { $_.Name })
        }

        $PowerShell = [PowerShell]::Create()
        try
        {
            $null = $PowerShell.AddScript($Script.ToString(), $true).AddArgument($ManifestPath).AddArgument((Join-Path $TestDrive 'config.json'))
            $NewGlobals = $PowerShell.Invoke()
            $PowerShell.Streams.Error | Should -BeNullOrEmpty
            $NewGlobals | Should -BeNullOrEmpty
        }
        finally
        {
            $PowerShell.Dispose()
        }
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
            Set-ItResult -Skipped -Because 'this is Windows PowerShell 5.1, and the analyzer check runs only in the PowerShell 7 run'
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

# Rules for the public parameters, so breaking one takes a deliberate change here. Section 6 of
# Research_Notes/PowerShell-Language-Behavior.md has the measurements behind them.
Describe 'ScriptMessage public parameters' {
    BeforeAll {
        Import-Module $ManifestPath -Force
        $Common = [System.Management.Automation.PSCmdlet]::CommonParameters + [System.Management.Automation.PSCmdlet]::OptionalCommonParameters
        $PublicParameters = foreach ($Command in (Get-Command -Module ScriptMessage))
        {
            foreach ($Parameter in ($Command.Parameters.Values | Where-Object { $_.Name -notin $Common }))
            {
                foreach ($Attribute in $Parameter.Attributes | Where-Object { $_ -is [System.Management.Automation.ParameterAttribute] })
                {
                    [pscustomobject]@{
                        Name       = "$($Command.Name) -$($Parameter.Name)"
                        IsSwitch   = $Parameter.SwitchParameter
                        Pipeline   = $Attribute.ValueFromPipeline -or $Attribute.ValueFromPipelineByPropertyName
                        Positional = $Command.Parameters[$Parameter.Name].ParameterSets.Values.Position -ne [int]::MinValue
                    }
                }
            }
        }
    }

    AfterAll {
        Remove-Module ScriptMessage -Force -ErrorAction SilentlyContinue
    }

    # A piped or positional $true or $false binds to a switch that allows it, so a switch never does.
    It 'gives no switch pipeline input or a position' {
        @($PublicParameters | Where-Object { $_.IsSwitch -and ($_.Pipeline -or $_.Positional) } | ForEach-Object Name) |
            Should -BeNullOrEmpty
    }

    # One piped object used to bind to several parameters at once. Taking pipeline input by property name, which
    # can be added later, needs the process block to stop reusing values from one piped object for the next.
    It 'gives Send-ScriptMessage no pipeline input' {
        @($PublicParameters | Where-Object { $_.Name -like 'Send-ScriptMessage *' -and $_.Pipeline } | ForEach-Object Name) |
            Should -BeNullOrEmpty
    }

    # Named only, as in 1.1.1. A position would be new public input.
    It 'gives Connect-ScriptMessage and Disconnect-ScriptMessage no positional parameters' {
        @($PublicParameters | Where-Object { $_.Name -match '^(Connect|Disconnect)-ScriptMessage ' -and $_.Positional } | ForEach-Object Name) |
            Should -BeNullOrEmpty
    }
}

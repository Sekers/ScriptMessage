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

    # PSScriptAnalyzer is optional locally. The CI workflow installs it for PowerShell 7, so the check always runs
    # there.
    It 'has no PSScriptAnalyzer errors' -Skip:(-not (Get-Module -ListAvailable -Name PSScriptAnalyzer)) {
        $Findings = @(Invoke-ScriptAnalyzer -Path $ModuleRoot -Recurse -Severity Error)

        @($Findings | ForEach-Object { '{0}:{1} {2}' -f $_.ScriptName, $_.Line, $_.RuleName }) | Should -BeNullOrEmpty
    }
}

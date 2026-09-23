# Test tooling behavior

How PSScriptAnalyzer, PowerShell's module loading, and the GitHub Actions runner image behave where this
repository's tests depend on them. None of this is module behavior; it constrains how the test suite and the Tests
workflow are written, chiefly the `has no PSScriptAnalyzer errors` test in `Tests/ScriptMessage.Module.Tests.ps1`
and the `pester` job in `.github/workflows/Tests.yml`.

Measured on **2026-09-23** on Windows 11 (10.0.26200), with Windows PowerShell **5.1.26100.9444**, PowerShell
**7.6.6**, PSScriptAnalyzer **1.25.0**, and Pester **6.1.0**. Nothing connected to a tenant.

**Claims are labelled with their evidence.**

- **Measured** means observed on the date given.
- **From source** means read from a file in this repository.
- **From documentation** means read from the linked page on the date given.
- **Unverified** means nobody has tested it.

> **The one-line summary:** PSScriptAnalyzer reports a syntax error only when asked for `ParseError`, it parses
> with the edition running it, and PowerShell loads the first copy of a module it finds rather than the newest.
> Name the analyzer version explicitly, because the runner image ships its own copy.

**When changing `PSSCRIPTANALYZER_VERSION`** in `.github/workflows/Tests.yml`, repeat sections 1 and 2 with the new
version.

## 1. Measured: syntax errors are reported only under the `ParseError` severity

`Invoke-ScriptAnalyzer -ScriptDefinition 'function f { if ($true { } }'` gave the same result in both editions:

| `-Severity` | Findings |
| --- | --- |
| `Error` | none |
| `Error, ParseError` | `UnexpectedToken` and `MissingEndParenthesisAfterStatement`, both with severity `ParseError` |
| not given | the same two |

Under PowerShell 7, with a missing `)` planted in an `if` statement in
`ScriptMessage/Private/ConvertTo-ScriptMessageBodyObject.ps1`, `Invoke-ScriptAnalyzer -Path ScriptMessage -Recurse`
reported nothing with `-Severity Error` and `MissingEndParenthesisAfterStatement` on that line with
`-Severity Error, ParseError`. Without the plant, the module had no finding of either severity.

This is why the analyzer test passes `-Severity Error, ParseError`.

## 2. Measured: the analyzer's rules are the same in both editions, but its parsing is not

- `Get-ScriptAnalyzerRule` listed 75 rules in each edition, with identical names.
- `Invoke-ScriptAnalyzer -ScriptDefinition '$a = $true ? 1 : 2' -Severity Error, ParseError` returned nothing under
  PowerShell 7 and `UnexpectedToken` under Windows PowerShell 5.1. The analyzer accepts whatever syntax the edition
  running it accepts (**Unverified:** that it uses that edition's parser; its source was not read).

So an analyzer run under PowerShell 7 alone cannot flag syntax that only PowerShell 7 accepts. The rest of the
suite covers that gap: with the ternary `return $true ? $null : $null` planted in
`ConvertTo-ScriptMessageBodyObject`, `.\Tests\Invoke-Tests.ps1 -Edition Desktop` failed 79 tests and passed 7, and
six of the eight test files failed as a whole.

This is one of the two reasons the analyzer test runs under PowerShell 7 only; section 3 gives the other.

## 3. Measured: which copy of a module PowerShell loads

Two folders were put at the front of `PSModulePath` in a fresh process of each edition. The first held a module
`FakeMod` at version 1.0.0 and the second held `FakeMod` 2.0.0, each exporting one function that returned its
version.

| Load | Windows PowerShell 5.1 | PowerShell 7 |
| --- | --- | --- |
| Calling the function (command autoloading) | 1.0.0 | 1.0.0 |
| `Import-Module -Name FakeMod` | 1.0.0 | 1.0.0 |

PowerShell loaded the copy in the first `PSModulePath` folder that held the module, not the newest version.

The folders each edition searched on the test machine, in order:

- PowerShell 7: the user's `Documents\PowerShell\Modules`, `C:\Program Files\PowerShell\Modules`,
  `$PSHOME\Modules`, `C:\Program Files\WindowsPowerShell\Modules`, and
  `C:\Windows\system32\WindowsPowerShell\v1.0\Modules`.
- Windows PowerShell 5.1, started from PowerShell 7 the way `Tests/Invoke-Tests.ps1` starts it: the user's
  `Documents\WindowsPowerShell\Modules`, `C:\Program Files\WindowsPowerShell\Modules`, and
  `C:\Windows\system32\WindowsPowerShell\v1.0\Modules`. None of PowerShell 7's own folders.

**From documentation** ([Install-PSResource](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.psresourceget/install-psresource),
read 2026-09-23): the default scope is `CurrentUser`, which installs to a folder such as
`$home\Documents\PowerShell\Modules`, and `AllUsers` installs to a folder such as
`$env:ProgramFiles\PowerShell\Modules`. A bare version such as `1.25.0` is treated as the required version, not
as a minimum.

So the copy the Tests workflow installs with `Install-PSResource` lands in one of PowerShell 7's own folders, where
a Windows PowerShell 5.1 test run cannot load it. That is the other reason the analyzer test runs under PowerShell
7 only. It imports the version named in `PSSCRIPTANALYZER_VERSION` with `Import-Module -RequiredVersion` rather
than relying on which copy PowerShell finds first. With a fake PSScriptAnalyzer 9.9.9 placed first on
`PSModulePath`, the test loaded the fake when the variable was not set (it then takes the newest installed copy)
and 1.25.0 when the variable was `1.25.0`.

## 4. The runner image's own copy

**From documentation**, read 2026-09-23:

- The [actions/runner-images README](https://github.com/actions/runner-images) maps `windows-latest` to the Windows
  Server 2025 with Visual Studio 2026 image, described in
  [Windows2025-VS2026-Readme.md](https://github.com/actions/runner-images/blob/main/images/windows/Windows2025-VS2026-Readme.md).
  For image version 20260907.229.1, that readme lists `PSScriptAnalyzer: 1.25.0` and `Pester: 3.4.0, 5.9.0`.
- The image's
  [toolset-2025-vs2026.json](https://github.com/actions/runner-images/blob/main/images/windows/toolsets/toolset-2025-vs2026.json)
  lists PSScriptAnalyzer with no version, and
  [Install-PowerShellModules.ps1](https://github.com/actions/runner-images/blob/main/images/windows/scripts/build/Install-PowerShellModules.ps1)
  installs a module listed that way with `Install-Module -Scope AllUsers` and no version. So each image build takes
  whatever PSScriptAnalyzer version is newest on the PowerShell Gallery at the time.

**Measured** from the logs of two Tests workflow runs, on 2026-09-22 at commit 453ce45 and on 2026-09-23 at
dba8cd7: the `Set up job` step printed `Image: windows-2025-vs2026` and `Version: 20260907.229.1`. In the
2026-09-22 run, before the analyzer test chose its own version, the test passed under both editions. Since Windows
PowerShell 5.1 cannot see the copy the workflow installs (section 3), that 5.1 run must have loaded the image's
copy.

So the analyzer version the image provides can change with any image update, with no change in this repository.
`PSSCRIPTANALYZER_VERSION` is what holds the version the test uses.

## What this does not establish

- Only PSScriptAnalyzer 1.25.0 was tested. Another version may have different rules, or may report syntax errors
  differently; that is why sections 1 and 2 need repeating when the version changes.
- Section 2 tested one construct that only PowerShell 7 accepts (a ternary), planted in one private helper file.
- Section 3 tested loading by name with no version parameter, in processes that had no copy loaded yet.
- The runner's own `PSModulePath` was not measured, and the image may add folders to it. So, before the analyzer
  test chose its version, whether the PowerShell 7 run in CI loaded the workflow's copy or the image's is
  **Unverified**. So is whether the install step installed a second copy of 1.25.0 or skipped: it printed nothing.
- Which folder the image's copy is in was not established; the 2026-09-22 run shows only that Windows PowerShell
  5.1 could load it.

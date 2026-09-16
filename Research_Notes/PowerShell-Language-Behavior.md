# PowerShell language behavior

PowerShell behaviors that this module's code depends on or trips over. None of this is Microsoft Graph
behavior; it is the language and runtime, recorded here because each item below either explains a bug found
in the module or constrains how a fix has to be written.

Measured on **2026-09-15** on Windows 11 (10.0.26200), against Windows PowerShell **5.1.26100.9444** and
PowerShell **7.6.6**, using small standalone scripts. The module was not imported and no tenant was involved.
Both editions agreed on every result unless a section says otherwise.

**Claims are labelled with their evidence.**

- **Measured** means run on both editions on the date above and the output recorded.
- **From source** means read out of a file in this repository; it describes what the module does, not what
  PowerShell does.
- **Unverified** means nobody has tested it.

> **The one-line summary:** PowerShell's truthiness, variable type constraints, and member enumeration each do
> something reasonable that is not what the code next to them assumed. Test the value you actually have, not
> the value you expect to have.

## 1. Measured: `-not` cannot tell "not passed" from "passed a zero-valued enum or `$false`"

With `enum ChatType { OneOnOne; Group }`:

| Expression | Result |
| --- | --- |
| `-not [ChatType]::OneOnOne` (enum value 0) | `True` |
| `-not [ChatType]::Group` (enum value 1) | `False` |
| A `[ChatType]` parameter the caller did not pass | value is empty (`$null`); `-not` gives `True` |
| A `[bool]` parameter the caller did not pass | value is `False` |
| `-not $false` | `True` |
| `$PSBoundParameters.ContainsKey('ChatType')`, passed `OneOnOne` / not passed | `True` / `False` |
| A `[Nullable[bool]]` variable that was never assigned | `$null` |

So `if (-not $ChatType) { <apply default> }` applies the default both when the caller passed nothing and when
the caller explicitly passed `OneOnOne`, and the same test on a `[bool]` treats an explicit `$false` as "not
passed". `$PSBoundParameters.ContainsKey()` is the test that tells them apart. A parameter that genuinely needs
three states (unset, true, false) should be `[Nullable[bool]]`.

**From source:** `Send-ScriptMessage` fills in `ChatType` and `IncludeBCCInGroupChat` from the configuration
file using exactly this `-not` test, which is why an explicit `-ChatType OneOnOne` or
`-IncludeBCCInGroupChat $false` is replaced by the configuration file's value.

## 2. Measured: a `[PSCustomObject]` constraint on a variable does not convert a hashtable

| Code | Resulting type |
| --- | --- |
| `[PSCustomObject]$h = @{ Name = 'x' }` | `System.Collections.Hashtable` |
| `$h = [PSCustomObject]@{ Name = 'x' }` | `System.Management.Automation.PSCustomObject` |

Only the cast applied directly to a hashtable literal builds a `PSCustomObject`. Putting the type on the
variable looks equivalent and is not.

**From source:** `ConvertTo-ScriptMessageAttachmentObject` uses the variable form, so its attachment items are
hashtables. The Microsoft Graph service code then calls `.ContainsKey()` on them, which works only because they
are hashtables, so changing either side without the other breaks attachments.

## 3. Measured: `.Address` on an array reaches the array's own method

For an array of objects that each have an `Address` property:

| Expression | Result |
| --- | --- |
| `$Array.Address` | a `PSMethod`, not the addresses |
| `$Array.Name` (a member arrays do not have) | the names, as expected |
| `$Array.ForEach('Address')` | the addresses |
| `$Array \| Select-Object -ExpandProperty Address` | the addresses |

.NET arrays have an `Address` method, and member access on a collection finds that before it enumerates the
elements. The printed type name differs by edition (`PSMethod` on 5.1, ``PSMethod`1`` on 7); the behavior is
the same.

**From source:** this is why the module's internal recipient objects use an `AddressObj` property instead of
`Address`, and why result building in `ScriptMessage/Services/MicrosoftGraph.ps1` goes through
`System.Linq.Enumerable`. `.ForEach('Address')` or `Select-Object -ExpandProperty Address` avoids the collision
without renaming the property.

## 4. Measured: `ExpandString` runs code; `ExpandEnvironmentVariables` does not

- `$ExecutionContext.InvokeCommand.ExpandString('C:\certs\$([math]::Sqrt(16)).pfx')` returned
  `C:\certs\4.pfx`. The `$(...)` subexpression was executed, so any command written inside one runs the same
  way.
- `[Environment]::ExpandEnvironmentVariables('%SystemRoot%\x')` returned `C:\Windows\x`. It only substitutes
  environment variables.

**From source:** `Connect-ScriptMessage_MicrosoftGraph` passes the `MgApp_CertificatePath` configuration value
through `ExpandString`, so that setting can execute code written into the configuration file.

## 5. Measured: looking up a property by a name held in an array

With `$Config = [pscustomobject]@{ MicrosoftGraph = 'cfg' }`:

| Code | Result |
| --- | --- |
| `$Config.$(@('MicrosoftGraph'))` | `cfg` |
| `$Config.$(@('MicrosoftGraph','Slack'))` | `$null` |

A one-element array works as the property name, so code that looks a property up by a variable holding service
names works while there is one service and silently returns nothing once there are two.

**From source:** `Send-ScriptMessage` looks up each service's configuration with
`$ScriptMessageConfig.$($serviceTypeObj.Service)`, where `Service` is an array holding every requested service.

## 6. Measured: without a `process` block, piped input binds only the last item

Piping `'a','b','c'` into an advanced function whose parameter has `ValueFromPipeline = $true`:

| Function body | Output |
| --- | --- |
| no `process` block | runs once, with `c` |
| a `process { }` block | runs three times, with `a`, `b`, and `c` |

**From source:** every function in `ScriptMessage/Public/` declares `ValueFromPipeline = $true` on its
parameters and none has a `process` block, so piping several objects into any of them acts on the last object
only.

## 7. Measured: `System.Web.HttpUtility` and URL encoding

| Expression | Windows PowerShell 5.1 | PowerShell 7 |
| --- | --- | --- |
| `[System.Web.HttpUtility]::UrlEncode('My File+1.pdf')` | fails on first use in a session: "Unable to find type [System.Web.HttpUtility]"; `My+File%2b1.pdf` later in that same session | `My+File%2b1.pdf` |
| `[uri]::EscapeDataString('My File+1.pdf')` | `My%20File%2B1.pdf` | `My%20File%2B1.pdf` |

`UrlEncode` produces form encoding, which turns a space into `+`. That is correct in a query string, but in a
URL path a `+` is a literal plus sign. `EscapeDataString` needs no extra assembly in either edition and encodes
a space as `%20`.

**Measured (2026-09-16; Windows PowerShell 5.1.26100.9444 and PowerShell 7.6.6, Pester 6.1.0, Windows 11 build
26200):** sending a chat attachment through `Send-ScriptMessage` with every Microsoft Graph cmdlet mocked, so no
Graph module was imported, and with the upload still calling `UrlEncode` as it did before the switch to
`EscapeDataString`. Under 5.1, the first upload in the session raised "Unable to find type
[System.Web.HttpUtility]", which arrived in the result's `Error` property with `Status` left null, and a second
upload later in that same session succeeded, building the path `My+File%2b1.pdf` from the filename
`My File+1.pdf`. Under PowerShell 7 both uploads succeeded and built the same path. A failure here depends on
what has already run in the session, so it can look intermittent.

**From source:** the Microsoft Graph chat attachment upload in `Send-ScriptMessage_MicrosoftGraph` builds its
drive item path with `[uri]::EscapeDataString`.

### What this does not establish

- **Unverified:** whether importing the Microsoft Graph modules loads `System.Web` on Windows PowerShell 5.1.
  An attempt on 2026-09-15 could not import `Microsoft.Graph.Authentication` 2.25.0 in a non-interactive 5.1
  session on the test machine (its format file failed the execution policy check), so this is still open. The
  2026-09-16 run above mocked the Graph cmdlets instead of importing the modules, so it says nothing about this
  either.
- **Unverified:** why a later reference to the type succeeded in a session where the first had failed. Only the
  outcome was observed, not the cause.
- **Unverified:** what name OneDrive gives a file uploaded through a path containing `+`. Nothing was uploaded.

## 8. Measured: .NET file methods and PowerShell paths

Measured on **2026-09-16** with the same editions, in a standalone script whose PowerShell location
(`Push-Location`) was a test folder while the process's working directory was the repository root. A PowerShell
drive was created with `New-PSDrive -PSProvider FileSystem`.

| Path given | `Get-Content -Path` | `[System.IO.File]::ReadAllText()` on the path as given | `ReadAllText()` on `$PSCmdlet.GetUnresolvedProviderPathFromPSPath()` |
| --- | --- | --- | --- |
| relative, `.\file.json` | read | failed | read |
| on a PowerShell drive | read | not tested | read |
| folder name containing `[x]` | failed | not tested | read |
| wildcard matching one file, `file*.json` | read | not tested | failed |
| `~\file.json` | not tested | not tested | resolved under `$HOME` (not read) |

- .NET resolves a relative path against the process's working directory, which `Set-Location` and
  `Push-Location` do not change, so a relative path has to be resolved by PowerShell first.
- `Get-Content -Path` treats `[` and `]` as wildcard characters, so a folder or file name containing them does
  not match itself. `GetUnresolvedProviderPathFromPSPath` takes the path literally: it handles relative paths,
  drives, and `~`, never expands wildcards, and does not check that the file exists.

**From source:** `Get-ScriptMessageConfig` resolves its `-Path` with `GetUnresolvedProviderPathFromPSPath` and then
opens the file with .NET, so a configuration file path containing brackets works and a wildcard in it does not.
`1.1.0` opened the file with `Get-Content -Path`.

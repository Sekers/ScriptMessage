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
file when the caller did not supply them, and tests `$PSBoundParameters.ContainsKey()` for exactly this reason.
The `-not` test would drop an explicit `-ChatType OneOnOne` or `-IncludeBCCInGroupChat $false` in favor of the
configuration file's value. Section 10 covers why the resolved values go into separate variables rather than
back into the parameters.

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

## 9. Measured: a `PSCustomObject` in a one-item array converts to an empty string

Measured on **2026-09-16** with the same editions, in a standalone script with `$OFS` unset. With
`$Obj = [pscustomobject]@{ Name = 'A'; AddressObj = 'a@domain.com' }`, `$One = @($Obj)`, and
`$Two = @($Obj, $Obj)`:

| Expression | Result |
| --- | --- |
| `[string]$Obj` or `"$Obj"` | `@{Name=A; AddressObj=a@domain.com}` |
| `$Obj.ToString()` | an empty string |
| `[string]$One` or `"$One"` | an empty string |
| `[string]$Two` | a single space |
| `[string]::IsNullOrEmpty($Obj)` | `False` |
| `[string]::IsNullOrEmpty($One)` | `True` |
| `[string]::IsNullOrEmpty($Two)` | `False` (`IsNullOrWhiteSpace` gives `True`) |
| `[string]::IsNullOrEmpty($One.AddressObj)` | `False` |
| `[string]@(@{ Name = 'x' })`, a hashtable in a one-item array | `System.Collections.Hashtable`, so `IsNullOrEmpty` gives `False` |
| `[string]@('a@domain.com')` | `a@domain.com` |
| `[string]::IsNullOrEmpty(@())` | `True` |

So `[string]::IsNullOrEmpty` cannot tell an empty array from one holding a single `PSCustomObject`, and for an
array of those objects its answer changes between one item and two.

The same holds for values arriving through a parameter:

| Parameter | Argument | `[string]::IsNullOrEmpty()` | `.Count` |
| --- | --- | --- | --- |
| `[pscustomobject]$Recipient` | `$Obj` | `False` | `$null` on 5.1, `1` on 7 |
| `[pscustomobject]$Recipient` | `$One` | `True` | `1` |
| `[array]$Attachment` | a lone `PSCustomObject` | `True` | `1` |
| `[array]$Attachment` | a lone hashtable | `False` | `1` |

An `[array]` parameter wraps a lone argument in a one-item array (both arguments above arrived as `Object[]`), so
a single `PSCustomObject` passed to one is empty by this test. `.Count` on a lone `PSCustomObject` is the only
result in this section where the editions differ.

**From source:**

- `ConvertTo-ScriptMessageRecipientObject` and `ConvertTo-IMicrosoftGraphRecipient` both take a `[pscustomobject]`
  parameter and pair `[string]::IsNullOrEmpty` with a second test: `.Count -lt 1` in the first,
  `[string]::IsNullOrEmpty($EmailAddress.AddressObj)` in the second. `Send-ScriptMessage` keeps the converted
  `To`, `CC`, `BCC`, and `ReplyTo` lists in `[array]` variables, so a single recipient reaches
  `ConvertTo-IMicrosoftGraphRecipient` as a one-item array, and only the second test keeps it from being returned
  as `$null`. For a lone object the `IsNullOrEmpty` half is already `False`, so the `.Count` difference between
  editions does not change either function's result.
- `ConvertTo-IMicrosoftGraphConversationMember` and `ConvertTo-IMicrosoftGraphDriveInvite` use
  `[string]::IsNullOrEmpty` alone, on email address strings, which are unaffected.
- `ConvertTo-ScriptMessageAttachmentObject` uses it alone on its `[array]` parameter, which receives
  `Send-ScriptMessage -Attachment` as given. Hashtables and path strings, the supported types, are unaffected. A
  single `PSCustomObject`, which is not supported, comes back as `$null`, the same as no attachment, while two of
  them reach the function's "Unexpected attachment object type." error. `ConvertTo-IMicrosoftGraphAttachment` and
  `ConvertTo-IMicrosoftGraphChatMessageAttachment` also use it alone, on items they call `.ContainsKey()` on, so
  they expect hashtables.

### What this does not establish

- **Unverified:** why an array converts differently. The results fit converting each element with `ToString()`
  and joining the results with a space, but that is inferred from the outputs, not read from PowerShell's source.
- Only `[string]` casts, string expansion, `[string]::IsNullOrEmpty`, and `[string]::IsNullOrWhiteSpace` were
  tested. Truthiness tests such as `if ($One)` or `-not $One`, and other .NET methods taking a string argument,
  were not.
- A script that sets `$OFS` was not tested.

## 10. Measured: an enum cannot hold `$null`, and a parameter keeps its type constraint

Measured on **2026-09-17**, same machine and editions as the header. Both editions agreed on every row.

With `enum CT { OneOnOne; Group }`:

| Code | Result |
| --- | --- |
| `[CT]$x = $null`, fresh variable, inside a function | throws "Cannot convert null to type" |
| `param([CT]$x)` then `$x = $null` later in the body | throws the same way |
| `param([CT]$x)` then `$x = 'Group'` later in the body | converts; `$x` is `[CT]` |
| `Test-Thing -ChatType $null` against a `[CT]` parameter | throws "Cannot process argument transformation" |
| Splatting `@{ ChatType = $null }` at a `[CT]` parameter | throws the same way |
| Splatting `@{ B = $null }` at a `[bool]` parameter | throws `Cannot convert value "" to type "System.Boolean"` |
| `[bool]$x = $null`, fresh variable, inside a function | converts; `$x` is `False` |
| `[bool]$x = $true` then `$x = $null` later, inside a function | converts; `$x` is `False` |
| Splatting `@{ ChatType = '' }` at a `[CT]` parameter | throws `The identifier name  cannot be processed because it is either too similar or identical to the following enumerator names` |
| `[string]::IsNullOrWhiteSpace()` given a `[CT]` value | `False`; the enum converts to its member name first |
| `[bool]` of the string `'false'` | `True`; every non-empty string is truthy |

Two consequences. A parameter's type constraint is not just an entry gate: it applies to every later assignment
to that variable in the body, so a parameter cannot be reused as a scratch variable for a value its type
cannot hold. And an enum has no representation for "no value", so a missing configuration setting cannot be
carried in an enum-typed variable or passed to an enum-typed parameter; the call has to omit the argument
instead.

A `[bool]` behaves differently from an enum in the one case that matters: a `[bool]`-constrained variable
inside a function accepts `$null` and becomes `False`, while a `[bool]` parameter rejects it at binding.

**From source:** `Send-ScriptMessage` resolves `ChatType` and `IncludeBCCInGroupChat` into separate
`$Resolved...` variables rather than assigning back to the parameters, because a parameter cannot be reassigned
a value its type cannot hold. It adds `ChatType` to the `Send-ScriptMessage_MicrosoftGraph` splat only when
`[string]::IsNullOrWhiteSpace` says there is a real value, because neither `$null` nor `''` can bind to the
service's `[ChatType]` parameter, and `Templates/config_scriptmessage.json` writes a setting it leaves unset as
an empty string. `Send-ScriptMessage_MicrosoftGraph` then reports a Chat message that arrived without one.
`IncludeBCCInGroupChat` needs no such guard: it is resolved through a `[bool]`-constrained variable inside a
function, so a configuration file with no `IncludeBCCInGroupChat` setting yields `False`.

The last row is a trap this module has not addressed: a configuration file that quotes the value, as
`"IncludeBCCInGroupChat": "false"`, yields `True`, because `[bool]` of any non-empty string is `True`. JSON
`false` is unquoted and converts correctly, so only a hand-edited file with the quotes is affected.

### What this does not establish

- Why the same `[bool]` assignment that converts `$null` inside a function throws at script scope. Both
  editions do it, and the module is unaffected because all of its code lives in functions, but the cause was
  not traced.
- `[Nullable[bool]]` and `[CT?]` parameters were not tested here; section 1 covers an unassigned
  `[Nullable[bool]]` variable only.
- Only `$null` was tested as the rejected value. Empty strings and other unconvertible values were not.

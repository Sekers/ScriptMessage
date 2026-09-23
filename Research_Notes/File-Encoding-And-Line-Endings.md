# File encoding and line endings

How PowerShell reads script files and configuration files, how Git stores and reports line endings, and what
happened when this repository was normalized to LF. The rules themselves live in `AGENTS.md` and
`.gitattributes`; this file is the evidence behind them.

Sections 1 to 4 were measured on **2026-09-15** on Windows 11 (10.0.26200), with Windows PowerShell
**5.1.26100.9444**, PowerShell **7.6.6**, and Git **2.54.0.windows.1** with `core.autocrlf=true` set in the
system-wide Git configuration. Section 5, and the module manifest measurement in section 1, give their own
dates.

**Claims are labelled with their evidence.**

- **Measured** means observed on the date given.
- **From source** means read from a file in this repository.
- **Unverified** means nobody has tested it.

> **The one-line summary:** Windows PowerShell 5.1 silently misreads a non-ASCII character in a script file
> that has no byte order mark, and Git hides line-ending changes from `git diff` while still listing the files
> in `git status`. Keep PowerShell files ASCII, and check line endings with `git ls-files --eol`.

## 1. Measured: a non-ASCII character in a script file without a BOM

A two-line script that assigns one em dash (U+2014) to a string and prints the string's length, saved as UTF-8:

| Saved as | Windows PowerShell 5.1 | PowerShell 7 |
| --- | --- | --- |
| UTF-8 without BOM | `length=3` | `length=1` |
| UTF-8 with BOM | `length=1` | `length=1` |

Windows PowerShell 5.1 read the em dash's three UTF-8 bytes as three characters, with no error or warning.
PowerShell 7 read both files correctly. This is why PowerShell files in this repository are ASCII only rather
than UTF-8 with a BOM: with no non-ASCII characters, whether a file has a BOM makes no difference to either
edition.

**Measured on 2026-09-23**, on the same machine and PowerShell versions, with ANSI code page 1252: the same
question for a module manifest, the one PowerShell file that may need a BOM, since a `.psd1` cannot contain an
escape such as `[char]0x00E9`. A minimal manifest whose `Author` was `Caf` followed by U+00E9 was saved as UTF-8
and read with `Import-PowerShellDataFile`, `Test-ModuleManifest`, and `Import-Module`:

| Saved as | Windows PowerShell 5.1 | PowerShell 7 |
| --- | --- | --- |
| UTF-8 without BOM | 5 characters, `Caf` then U+00C3 U+00A9 | 4 characters, `Caf` then U+00E9 |
| UTF-8 with BOM | 4 characters, `Caf` then U+00E9 | 4 characters, `Caf` then U+00E9 |

All three commands gave the same result in every case, and none raised an error or warning. So a manifest that
genuinely needs a non-ASCII character is saved as UTF-8 with a BOM, which both editions read correctly. Only the
`Author` value and code page 1252 were tested.

Section 5 measures the same question for the JSON configuration file.

## 2. Measured: the repository before normalization

As of commit ff31824, before `.gitattributes` existed:

- Git's index held all 24 tracked files with LF line endings.
- With `core.autocrlf=true`, 22 of them were checked out with CRLF. `CHANGELOG.md` and `.gitignore` were LF in
  the working tree (**Unverified:** why those two differed).
- `ScriptMessage/ScriptMessage.psm1` was the only file with a UTF-8 BOM.
- No tracked file contained a non-ASCII character.

## 3. Measured: `git status` lists files whose content did not change

After adding `.gitattributes` (every text type `eol=lf`) and rewriting the 22 CRLF working files as LF:

| Check | Result |
| --- | --- |
| `git diff --stat` | only `ScriptMessage/ScriptMessage.psm1`, one line (its BOM removal) |
| `git status --short` | all 22 rewritten files listed as modified |
| `git update-index --refresh` | reported each of them as `needs update`; `git status` did not change |
| `git hash-object <file>` compared with the blob in `git ls-files -s` | identical for the 21 files other than `ScriptMessage.psm1` |
| `git add <file>` on those 21 | removed them from `git status`; `git diff --cached` stayed empty |

So after a line-ending or `.gitattributes` change, a file that `git status` lists as modified may have no
content change at all, and refreshing the index does not clear it. Compare `git hash-object` with
`git ls-files -s` before assuming there is something to commit, and after staging, `git diff --cached --stat`
shows what a commit would actually contain.

## 4. Measured: starting Windows PowerShell 5.1 as a child process on the test machine

`powershell.exe -NoProfile -NonInteractive -Command "Import-Module .\ScriptMessage\ScriptMessage.psd1"` failed
with "running scripts is disabled on this system". Adding `-ExecutionPolicy Bypass` to that command line, which
applies to that process only, let the module import. PowerShell 7 imported the module without it.

This comes from the machine's execution policy, not from the module, but it stops any test run that starts 5.1
in a child process, so a cross-edition test runner needs to pass `-ExecutionPolicy Bypass` or run under a
policy that allows local scripts.

## 5. Measured: how the configuration file is decoded

Measured on **2026-09-16** on the same machine, with Windows PowerShell **5.1.26100.9444** and PowerShell
**7.6.6**. The system's legacy (ANSI) code page was **1252** and the culture `en-US`. The tests only read local
files; nothing connected to Microsoft Graph.

**From source:** in `1.1.0`, `Get-ScriptMessageConfig` (`ScriptMessage/Public/Get-ScriptMessageConfig.ps1`)
reads the file with `Get-Content` and no `-Encoding`, and pipes the lines to `ConvertFrom-Json`. The tables call
that the 1.1.0 reader; it was measured with `develop` at c88393d, whose reader is identical. `Get-ScriptMessageConfig`
now uses the reader the tables call "Current".

**Method.** Every test file held the same two string settings. `MgApp_CertificateName` was `Soci`, U+00E9, `t`,
U+00E9, a space, and U+20AC, all of which exist in code page 1252. A second setting held U+0141, U+4E2D, and the
surrogate pair for U+1F600, none of which exist in code page 1252. Six files were written byte for byte, and six
more by piping `ConvertTo-Json` to `Set-Content`, `Out-File`, and `>` in each edition. Both editions read every
file with the module's own `Get-ScriptMessageConfig` and with three alternatives, and each value was compared with
what the file's bytes actually hold.

What each edition's cmdlets wrote:

| Cmdlet | Windows PowerShell 5.1 | PowerShell 7 |
| --- | --- | --- |
| `Set-Content` | code page 1252, no BOM; characters outside it were lost while writing (U+0141 became `L`, the rest `?`) | UTF-8 without BOM |
| `Out-File` and `>` | UTF-16 LE with BOM | UTF-8 without BOM |

`ConvertTo-Json` escaped none of the test characters in either edition.

What each reader returned. "wrong" means different text with no error; no read in the whole test raised an error.

| File encoding | 1.1.0 reader, 5.1 | 1.1.0 reader, 7 | `-Encoding UTF8`, 5.1 | `-Encoding UTF8`, 7 | Current, 5.1 | Current, 7 |
| --- | --- | --- | --- | --- | --- | --- |
| UTF-8 without BOM | **wrong** | ok | ok | ok | ok | ok |
| UTF-8 with BOM | ok | ok | ok | ok | ok | ok |
| UTF-16 LE with BOM | ok | ok | ok | ok | ok | ok |
| UTF-16 BE with BOM | ok | ok | ok | ok | ok | ok |
| Code page 1252, no BOM | ok | **wrong** | **wrong** | **wrong** | ok | ok |
| ASCII, non-ASCII written as JSON escape sequences | ok | ok | ok | ok | ok | ok |

- "1.1.0 reader" is `Get-Content` with no `-Encoding` piped to `ConvertFrom-Json`, called through
  `Get-ScriptMessageConfig`.
- "`-Encoding UTF8`" is `Get-Content -Encoding UTF8` piped to `ConvertFrom-Json`. It still honored a UTF-16 BOM.
  `[System.IO.File]::ReadAllText()` with no encoding argument gave the same result for every file in both editions.
- "Current" is `Get-ScriptMessageConfig` as it reads the file now. It opens a `System.IO.StreamReader` with byte
  order mark detection and a strict `UTF8Encoding($false, $true)`. When that throws `DecoderFallbackException`, it
  rereads the file with `[System.Text.Encoding]::GetEncoding(0)`.
- The files written by cmdlets gave the same results as the row for the encoding they were written in.

The misreads were silent. Windows PowerShell 5.1 decoded each byte of a UTF-8 sequence as a separate code page
1252 character (U+00E9 came back as U+00C3 U+00A9). A UTF-8 read of the code page 1252 file turned each
non-ASCII byte into U+FFFD.

Which code page each API used, first in the process's own culture (`en-US`) and then after setting the thread's
culture to `ja-JP` in the same process:

| Measurement | 5.1, `en-US` | 5.1, `ja-JP` | 7, `en-US` | 7, `ja-JP` |
| --- | --- | --- | --- | --- |
| `[System.Text.Encoding]::GetEncoding(0).CodePage` | 1252 | 1252 | 1252 | 1252 |
| `[System.Text.Encoding]::Default.CodePage` | 1252 | 1252 | 65001 | 65001 |
| `[System.Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage` | 1252 | 932 | 1252 | 932 |
| `Get-Content` with no `-Encoding` read the code page 1252 file correctly | yes | yes | no | no |

So `GetEncoding(0)` follows the system's code page in both editions, as Windows PowerShell 5.1's `Get-Content`
does, and ignores the culture. `Encoding.Default` is UTF-8 on PowerShell 7, and the culture's code page changes
with the culture.

A file starting with a UTF-8 byte order mark and containing an invalid UTF-8 byte raised no exception from the
strict `StreamReader` in either edition; the byte came back as U+FFFD. So the reader only falls back to the legacy
code page for a file with no byte order mark.

**What this means for the module:**

- Adding `-Encoding UTF8` would make Windows PowerShell 5.1 read exactly as PowerShell 7 does with the 1.1.0
  reader: the two gave identical results for all twelve files. It fixes UTF-8 files without a BOM on 5.1, which
  is what PowerShell 7's cmdlets write. It breaks 5.1's reading of files in the legacy code page, which is what
  5.1's `Set-Content` writes and which PowerShell 7 already misreads. The module does not use it for that reason.
- The current reader read every file correctly in both editions. `Tests/Get-ScriptMessageConfig.Tests.ps1`
  covers UTF-8 with and without a BOM, UTF-16 LE with a BOM, and the legacy code page.
- **From source:** the only configuration settings that hold free text are `MgApp_CertificateName` and
  `MgApp_CertificatePath`; the rest are GUIDs, a thumbprint, booleans, names from a fixed set, and encrypted
  strings. `Connect-ScriptMessage_MicrosoftGraph` uses `MgApp_CertificatePath` only for `CertificateFile`
  authentication, which throws before PowerShell 7.4, so on 5.1 only `MgApp_CertificateName` can be misread in a
  way that matters. On PowerShell 7 both settings can be misread from a legacy code page file.

**What this does not establish:**

- Only code page 1252 was tested. Another legacy code page, or Windows' "Use Unicode UTF-8 for worldwide
  language support" option (which sets the legacy code page to 65001), is **Unverified**.
- The culture test changed only the thread's culture. A machine whose system locale uses another code page was
  not tested.
- What `GetEncoding(0)` returns on Linux and macOS is **Unverified**. The reader only reaches it for a file that
  is not valid UTF-8.
- A legacy code page file whose non-ASCII bytes happen to form valid UTF-8 would be decoded as UTF-8. Not tested.
- What `Connect-MgGraph` does with a misread `MgApp_CertificateName` was not tested; it needs a certificate in
  the store and a connection.
- Which encoding Notepad, VS Code, or other editors use when a user saves a copy of the template is
  **Unverified**.
- UTF-16 without a BOM and UTF-32 were not tested.

# File encoding and line endings

How PowerShell reads script files, how Git stores and reports line endings, and what happened when this
repository was normalized to LF. The rules themselves live in `AGENTS.md` and `.gitattributes`; this file is
the evidence behind them.

Measured on **2026-09-15** on Windows 11 (10.0.26200), with Windows PowerShell **5.1.26100.9444**, PowerShell
**7.6.6**, and Git **2.54.0.windows.1** with `core.autocrlf=true` set in the system-wide Git configuration.

**Claims are labelled with their evidence.**

- **Measured** means observed on the date above.
- **Unverified** means nobody has tested it.

> **The one-line summary:** Windows PowerShell 5.1 silently misreads a non-ASCII character in a script file
> that has no byte order mark, and Git hides line-ending changes from `git diff` while still listing the files
> in `git status`. Keep the repository ASCII, and check line endings with `git ls-files --eol`.

## 1. Measured: a non-ASCII character in a script file without a BOM

A two-line script that assigns one em dash (U+2014) to a string and prints the string's length, saved as UTF-8:

| Saved as | Windows PowerShell 5.1 | PowerShell 7 |
| --- | --- | --- |
| UTF-8 without BOM | `length=3` | `length=1` |
| UTF-8 with BOM | `length=1` | `length=1` |

Windows PowerShell 5.1 read the em dash's three UTF-8 bytes as three characters, with no error or warning.
PowerShell 7 read both files correctly. This is why the repository rule is "ASCII only" rather than "UTF-8
with a BOM": with no non-ASCII characters, whether a file has a BOM makes no difference to either edition.

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

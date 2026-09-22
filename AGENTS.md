# Guidance for AI Assistants Working in This Repository

This is the single source of truth for AI assistant guidance. `CLAUDE.md` at the repository root is a pointer
to this file, so Claude Code and Codex both pick up the same rules. Edit this file, not the pointer.

**What this repository is:** `ScriptMessage`, a PowerShell module that lets scripts send email and chat
messages through messaging services (currently Microsoft Graph) with one set of cmdlets. Public functions live
one per file in `ScriptMessage/Public/`, private helpers in `ScriptMessage/Private/`, and each messaging
service's implementation in `ScriptMessage/Services/`; `ScriptMessage/ScriptMessage.psm1` dot-sources all
three folders. Work happens on `develop`; `main` is the released branch. The GitHub wiki is a separate
repository (`Sekers/ScriptMessage.wiki`), not part of this one.

## Safety rules that override convenience

- **Never send a real message without explicit permission.** Sending is this module's whole job, so running
  `Send-ScriptMessage` against a real configuration delivers real email or Teams chat to real people, and a
  chat with attachments also uploads the files to the sender's OneDrive and shares them with the recipients.
  `Connect-ScriptMessage` can open an interactive sign-in, so ask before running that too. The scripts in
  `Tests/` use whichever configuration their `Set-ScriptMessageConfigFilePath` line names; read that line, and
  the recipients, before running anything.
- **Never name the production tenant in any file.** No organization name, abbreviation, or email domain in
  code, comments, help examples, tests, or commit messages. Every placeholder address and URL uses a domain
  reserved by [RFC 2606](https://www.rfc-editor.org/rfc/rfc2606): `example.com`, `example.net`, or
  `example.org`. None of them can ever be registered, so a placeholder copied out of a help example and run
  unedited, or a test whose mock is missing, cannot reach a real recipient. Prefer `example.com` for anything
  new, but all three are equally correct: never rename existing placeholders to match that preference.
- **Never commit a real configuration file.** Configuration files hold the tenant ID, client ID, certificate
  details, and encrypted secrets. `Tests/Config/` and `@Local Only/` are gitignored for that reason, and
  `Templates/config_scriptmessage.json` must only ever contain placeholder values.

## CHANGELOG.md

**The changelog is the release notes.** Its entries are what users actually read when a version ships, so it
is the destination for user-facing change, not a staging area for it. Never move that content into
`README.md`. If something genuinely needs more room than an entry allows, it goes in the wiki, and only when
it is a big enough issue that people need to be warned about it.

### The baseline is the latest RELEASE, never the development branch

An entry describes what changed for someone **upgrading from the last released version**. Before writing any
"Fixed ..." entry, verify the bug actually existed in that release.

**Derive the release. Do not trust a version number written down anywhere, including here or in
`ScriptMessage/ScriptMessage.psd1`.** A release is an annotated tag; pushing it runs
`.github/workflows/Release.yml`, which publishes that build to the PowerShell Gallery and then creates the GitHub
release ([RELEASING.md](./RELEASING.md) describes the whole process and the versioning rules). `origin/main` can
sit past the last tag (1.1.0 got a changelog-only commit after its tag, for example), so ask for the nearest tag
rather than an exact match:

```powershell
git fetch origin --tags --quiet                        # only if the local copy might be behind
$Release = git describe --tags --abbrev=0 origin/main
```

Work happens on `develop`. The **local `main` branch is stale**, so compare against `$Release` or
`origin/main`, never local `main`, and never against `HEAD` or `develop`.

Verify before claiming a fix:

```powershell
git cat-file -e "${Release}:ScriptMessage/Public/<Name>.ps1"   # nonzero exit = did not exist in the release
git show "${Release}:ScriptMessage/Public/<Name>.ps1" | Select-String '<pattern>'   # was the bug there?
git show "${Release}:ScriptMessage/Services/MicrosoftGraph.ps1" | Select-String '<pattern>'
git tag --contains <commit>   # no output = the commit that introduced the bug never shipped
```

**A bug introduced on `develop` and fixed before release is not a changelog entry.** It never reached a user.
This is easy to get wrong while a release is in progress, because a lot of churn happens on `develop`, and a
fix to something that itself landed after the last tag is invisible to users. The checks above settle it.

**A function that does not exist in the last release** belongs under **Added** as a new function only. It can
never also appear as a "Fixed" entry, and it should not be listed among the functions a fix "affects." The same
goes for a new configuration setting or message type.

**A change the module adapted to is not a fix.** "Fixed" claims a defect in this module. If the code was never
wrong and something underneath it changed (the Microsoft Graph API, the Microsoft Graph PowerShell SDK, or
PowerShell itself), the entry belongs in **Changed**, or in **Added** when it gives callers something new. The
giveaway is that the change in behavior already reached users on the last release without any change in this
repository.

### Write for the end user, never for the module developer

An entry answers one question: **what changes for someone using this module?** Anything that does not help a
caller decide "does this affect me, and do I need to do something?" belongs in the commit message or a code
comment, not here.

Include:

- What went wrong from the caller's side, described in terms of what they saw.
- What the behavior is now.
- The public functions, parameters, and configuration file settings affected, and any action the user has to
  take, such as editing their configuration file or installing another Microsoft Graph sub-module.

**Verify the symptom, not just the bug.** A real defect still produces a false entry if the impact is
overstated. Trace what a caller actually experienced before describing it. For example, `Send-ScriptMessage`
catches errors raised while a service sends and reports them in the returned result's `Error` property instead
of throwing, so a bug that broke sending usually showed up to the caller as an `Error` entry, not a crash.

Leave out:

- **Root cause and mechanism.** Why the bug happened is developer information. "A `-not` check treated the
  enum value `OneOnOne` as false" tells a user nothing they can act on; "`-ChatType OneOnOne` was ignored when
  the configuration file set `ChatType` to `Group`" does.
- **Private helpers, internal variables and code paths.** Only name things a caller can actually use; check
  `FunctionsToExport` in `ScriptMessage/ScriptMessage.psd1` before naming a function. Service functions such
  as `Send-ScriptMessage_MicrosoftGraph` are private.
- **How it was found, measured or diagnosed**, including pointers to `Research_Notes/`. Those notes are for
  contributors, so link them from the commit message or the code, not from a user-facing entry.
- **Anything with no observable effect on a caller.** If no script a user could reasonably write would have
  hit the bug, there is nothing to log.

Two or three sentences is a normal entry. Length is not thoroughness: someone scanning a release to see
whether it affects them should get the answer in the first line.

**One entry, one audience.** If a single change produces two facts aimed at different readers, write two
entries. A fix for delegated permissions alongside a new requirement for application permissions reads as a
contradiction in one bullet; as two bullets, each reader finds their half immediately.

### Structure and style

The format is [Keep a Changelog 1.1.0](https://keepachangelog.com/en/1.1.0/). [RELEASING.md](./RELEASING.md)
covers when `[Unreleased]` becomes a version and how that version is chosen.

- **New entries go under `## [Unreleased]`** at the top of the file. Never invent a version number for them;
  the release-prep pull request turns `[Unreleased]` into a dated version heading.
- **Add the `[Unreleased]` section only with an entry.** Right after a release there is none, so the first new
  entry also adds `## [Unreleased](https://github.com/Sekers/ScriptMessage/compare/<release>...develop)` at the
  top, where `<release>` is `$Release` from the baseline commands above. Never add an empty one, and never leave
  one in a release: the Release workflow refuses to publish while it is there.
- Newest version first. A released version uses this heading shape, and a `---` line separates it from the next
  (older) version:

  ```markdown
  ## [1.2.0](https://github.com/Sekers/ScriptMessage/tree/1.2.0) - 2026-09-20

  ### Fixed

  - ...
  ```

- Sections in order: `### Added`, `### Changed`, `### Deprecated`, `### Removed`, `### Fixed`, `### Security`.
  Omit a section that has no entries.
- Versions 1.1.0 and earlier use `### Fixes`, `### Features`, and `### Other`, with an author line. Leave them
  as they are.
- Prefix an entry with `Minor:` inside **Fixed** when a user probably never noticed it: help text wording, a
  message, an edge case needing unusual conditions to hit, or a cost the module absorbed itself such as an
  extra Microsoft Graph request. List the `Minor:` entries after the rest of the Fixed entries.
- Prefix an entry with `BREAKING CHANGE:` when it breaks the public contract as [RELEASING.md](./RELEASING.md)
  defines it, so an existing script or configuration file can stop working, and say what the user has to
  change. Such an entry usually belongs under **Changed** or **Removed**.
- A new public function goes under **Added** and reads `- New Function: <Function-Name> > <what it does>`.
- **No em dashes.** Use a semicolon, colon, parentheses, comma, or a new sentence. This applies to every file
  in the repo, not just the changelog.
- In new entries, put code identifiers in backticks in Markdown, including parameter names, configuration
  settings, and .NET property names, so a bare `Type` does not read as ordinary prose. Older entries quote
  identifiers with single quotes; leave them as they are. In PowerShell comment-based help, do **not** use
  backticks; `Get-Help` renders them literally.

## File encoding and line endings

**This is already decided. Do not re-derive it, and do not "tidy" files to match a different opinion.**
`.gitattributes` is the authority and explains its own reasoning; this section exists so you know the answer
without going to look.

- **LF everywhere.** Every text type is pinned `eol=lf`, which is what Git stores anyway, so a checked-out
  file is byte for byte the repository's copy on any platform and under any `core.autocrlf`.
- **UTF-8 with no byte order mark.** Every tracked text file is UTF-8 without a BOM.
- **PowerShell files contain only ASCII.** This covers `.ps1`, `.psm1`, `.psd1`, and `.ps1xml` files. Windows
  PowerShell 5.1 reads a PowerShell file that has no BOM using the system's legacy code page, so a single
  non-ASCII character (a curly quote or an em dash pasted into a message string, for example) is silently
  misread there while working fine in PowerShell 7. In code, write a non-ASCII character as an escape such as
  `[char]0x00E9`. A `.psd1` cannot contain expressions, so if a module manifest genuinely needs a non-ASCII
  character, save that one file as UTF-8 with a BOM, which both editions read correctly
  (`Research_Notes/File-Encoding-And-Line-Endings.md` section 1).
- **Other text files may contain non-ASCII characters**, such as a contributor's name in `CHANGELOG.md`. The
  no-em-dash rule in the changelog's "Structure and style" section still applies to every file.
- **The module writes no files at runtime.** It only reads the configuration file, which users create
  themselves. If a feature ever writes a file, choose its encoding deliberately: `Out-File -Encoding utf8`
  writes a BOM on Windows PowerShell 5.1 and none on PowerShell 7, and 5.1 has no `utf8NoBOM`.
- **Skip anything containing a NUL byte**, which is Git's own binary test. Git reports such a file as
  `-text`, so no line-ending rule applies to it, and a bulk "read text, write text" pass over one re-encodes
  it and corrupts it.

Git will not show you a line-ending mistake: it normalizes before diffing, so a wrong file produces no diff
and leaves `git status` clean. **`Tests/FileEncoding.Tests.ps1` enforces all of these rules** as part of the
test suite: no byte order mark (a module manifest that holds a non-ASCII character is the one exception), LF line
endings, and only ASCII in PowerShell files. It checks untracked files that are not ignored as well as tracked
ones, so it catches a new file before it is committed, and it skips any file containing a NUL byte. A failure
names each offending file.

New files can pick up CRLF on Windows (VS Code's default line ending follows the operating system), so run the
tests after creating a file, not only after changing code. Never change the line endings of a file you are not
otherwise editing. To fix one that is wrong:

```powershell
$Path = (Resolve-Path -LiteralPath '<file>').ProviderPath   # .NET needs a full path
$Text = [System.IO.File]::ReadAllText($Path)
[System.IO.File]::WriteAllText($Path, ($Text -replace "`r`n","`n"))   # writes UTF-8 without a BOM
```

## Code

- **A new public function needs an explicit entry in `FunctionsToExport` in
  `ScriptMessage/ScriptMessage.psd1`.** Every `ScriptMessage/Public/*.ps1` is dot-sourced automatically, so a
  function with no manifest entry loads but stays invisible to callers. Aliases work the same way through
  `AliasesToExport`.
- **Comments describe the code as it is now**, never what it used to do or what a fix changed. Change history
  belongs in the commit message and the changelog.
- **A boolean configuration setting is an opt-in flag: on only when it equals `$true`.** Read it as
  `$ServiceConfig.<Name> -eq $true`, with the setting on the left, so missing, `false`, typos, and any quoted
  text except `"true"` (in any letter case, which counts as on) all mean off. Never read one with a `[bool]` cast, a truthiness test, or `$true -eq`: each of those reads the
  text `"false"` as true. Name a new boolean setting so that `true` opts in and off is the safe state.
- **A service that cannot honor a setting or parameter it receives says so in its result and sends what it
  can.** Add a `Warning` entry to that message's `Error` property, as the Microsoft Graph service does for a
  chat sent on behalf of someone else, rather than failing the whole call or dropping the value silently.
  `Send-ScriptMessage` passes every message parameter to every service, so each service's send function must
  accept all of them.

## Testing

- **Run `.\Tests\Invoke-Tests.ps1` from the repository root before proposing a commit.** It runs every
  `Tests/*.Tests.ps1` file under both Windows PowerShell 5.1 and PowerShell 7, each in a fresh process, and the
  CI workflow runs the same script. `-Edition Core` runs PowerShell 7 alone for a quicker check; finish with both,
  because the editions behave differently in places.
- **The Pester files never connect to a tenant** and need no Microsoft Graph modules. Mock the service functions
  such as `Send-ScriptMessage_MicrosoftGraph`, or the Microsoft Graph cmdlets they call. Pester can only mock a
  command that exists, so a test that mocks a Graph cmdlet first defines a stand-in for it inside the module that
  throws if reached, as the existing files do.
- **The other scripts in `Tests/` are not tests.** `ScriptMessage_Testing.ps1` and
  `ScriptMessage_Testing (Installed Module).ps1` are scratch scripts that load the real configuration from
  `Tests/Config/`, so anything run from them sends real messages; the safety rules above apply.

## Research notes

Researched or measured behavior belongs in `Research_Notes/`, one file per behavior category (PowerShell
language behavior, Microsoft Graph behavior, file encoding, and so on); list the directory to see what already
exists. Add to the matching file, or create a new file for a new category rather than stretching an existing
one.

- **Label every claim with its evidence:** **Measured**, **From source** (read from a file in this
  repository), **From documentation** (with the page linked), or **Unverified**. An unverified assumption the
  module depends on is still worth recording, labelled as such.
- **Date everything that can change.** Give the date and environment (PowerShell editions and versions, module
  versions) for anything measured, and the date read for anything taken from documentation.
- **Say what a result does not establish**, so a later reader does not stretch it past what was tested.
- **Tie each behavior to the code it affects**, by file and function name rather than line number.
- **The notes are for contributors.** Link them from commit messages and code comments, never from
  `CHANGELOG.md` or other user-facing text.
- **The safety rules above apply while gathering evidence:** no live sends without permission, and no tenant
  name or email domain in a note.

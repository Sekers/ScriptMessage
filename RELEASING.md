# Releasing ScriptMessage

This is the release policy for ScriptMessage: how versions are numbered, how branches are used, how the
changelog is kept, and how a release is published. It applies to every maintainer and contributor.
[CONTRIBUTING.md](./CONTRIBUTING.md) covers everyday contributions, and the `CHANGELOG.md` section of
[AGENTS.md](./AGENTS.md) holds the detailed rules for writing changelog entries.

## Versioning

ScriptMessage follows [Semantic Versioning 2.0.0](https://semver.org/spec/v2.0.0.html) for every release after
1.1.0. Earlier releases did not follow it strictly (some patch releases contained breaking changes), and their
numbers stay as they are.

| Change to the public contract | Increment | Example |
|---|---|---|
| Breaking: an existing script or configuration file can stop working | Major | `1.4.2` to `2.0.0` |
| New backward-compatible functionality, including a new optional parameter | Minor | `1.4.2` to `1.5.0` |
| Backward-compatible fix | Patch | `1.4.2` to `1.4.3` |

The PowerShell Gallery's own publishing guidelines give new parameters as a patch-level example. ScriptMessage
uses the stricter Semantic Versioning rule: new functionality, including a new parameter, is a minor release.

### The public contract

The public contract is the behavior ScriptMessage offers its callers:

- Exported functions and aliases, their parameters and parameter sets, and each parameter's requirements,
  defaults, validation, and meaning.
- Pipeline input and binding.
- Output objects, including the properties of the `Send-ScriptMessage` result.
- Errors, and whether they terminate.
- Supported PowerShell editions and versions (Windows PowerShell 5.1 and PowerShell 7) and supported operating
  systems.
- Configuration file settings, their meaning, and their defaults.
- The external modules the module requires, such as the Microsoft Graph sub-modules each message type needs.
- The permissions the module requests or needs, such as Microsoft Graph scopes.

A side effect of how the module happens to be built is not part of it, even when a caller can observe it. See
[Deciding hard cases](#deciding-hard-cases).

### Deciding hard cases

- **A fix that changes observable behavior** is a patch when the old behavior could not have worked for anyone,
  such as rejecting recipients that have no address, which never sent anything. It is a breaking change when a
  reasonable script could have depended on the old behavior. Either way, the changelog entry says what changed.
- **Correcting an unintended side effect is not a breaking change.** Behavior that exists only because of how the
  module was built, and that no documentation ever offered (the help, the README, the wiki, the sample
  configuration file, or the changelog), is outside the public contract even though a caller can observe or use
  it. Examples: a variable the module kept in the global scope, where any script could read or change it, and a
  true/false setting written as quoted text. Changing or removing such behavior needs no deprecation, and the
  "reasonable script" test above does not apply. Naming something in an error message does not document it.
  Behavior the module intends callers to use stays in the contract even where its documentation is thin, such as
  the properties of the `Send-ScriptMessage` result. The change gets a changelog entry only when a script that
  follows the documentation could notice it.
- **Requiring a new module or permission is a breaking change**, because existing installations and app
  registrations stop working until someone acts.
- **Deprecate before removing.** Mark functionality as deprecated in a minor release (documentation, a warning,
  and a `Deprecated` changelog entry), and remove it no earlier than the next major release.
- **If a breaking change ships in a minor or patch release by mistake,** do not replace or unlist that version.
  Release a new version that restores compatibility as soon as possible.
- **Do not choose the version early.** Changes collect under `[Unreleased]` in `CHANGELOG.md`, and the version
  is chosen at release time from what is there: anything breaking means major, anything added means minor,
  and only fixes means patch.

### Where the version lives

- `ModuleVersion` in `ScriptMessage/ScriptMessage.psd1` is the source of truth, and has exactly three numeric
  parts. It keeps the last released version until a release-prep pull request changes it.
- A prerelease also sets `Prerelease` under `PrivateData.PSData`. For a stable release, leave that key
  commented out rather than setting it to an empty string.
- The Git tag, the changelog heading, the GitHub release, and the PowerShell Gallery version all use the same
  string, with no `v` prefix: `1.2.0`, or `2.0.0-rc01` for a prerelease. The release workflow refuses to
  publish when they disagree.

### Prereleases

- A prerelease label may contain only letters, digits, and hyphens. The PowerShell Gallery rejects dots and `+`.
- Use `alpha`, then `beta`, then `rc`, each with a two-digit number: `2.0.0-alpha01`, `2.0.0-beta01`,
  `2.0.0-rc01`. The Gallery compares labels as text, so without the leading zero `rc10` would sort before `rc2`.
- Installing a prerelease needs PowerShellGet 1.6.0 or later with `-AllowPrerelease`, or PSResourceGet with
  `-Prerelease`. The PowerShellGet built into Windows PowerShell 5.1 cannot install one. A prerelease and the
  final version install into the same folder, so the final version replaces the prerelease.

### Published versions never change

A released version is never republished, re-tagged, or replaced, and its number is never reused. The PowerShell
Gallery enforces this on its side: packages cannot be deleted, only unlisted, and each publish must be a higher
version than the last. If a release is defective, fix it in a new version.

## Branches

ScriptMessage uses the
[Simple Modified Gitflow](https://www.grimadmin.com/article.php/simple-modified-gitflow-workflow) workflow:
Gitflow without release branches, because only one version is maintained at a time.

| Branch | Created from | Merges into | Purpose |
|---|---|---|---|
| `main` | Permanent | None | Always the latest release. Every merge into `main` is a release and is tagged. |
| `develop` | Permanent | `main` | The integration branch, holding the complete history of the project. |
| `feature/<short-name>` | `develop` | `develop` | New features, and bug fixes that can wait for the next release. |
| `hotfix/<version>` | `main` | `main`, then `develop` | Urgent fixes to the released version. |

Rules for the shared branches:

- **Changes reach `develop` and `main` only through pull requests**, and the tests must pass first.
- **Merge into `main` with a merge commit, never a squash or a rebase.** A squashed commit on `main` does not
  exist on `develop`, so every later release pull request would conflict with it.
- **Never rebase or force-push `main` or `develop`.** Release tags cannot move, and rewriting shared history
  breaks every contributor's clone. Rebase only your own feature branch, and only before it merges.
- **After every release or hotfix, merge `main` back into `develop` with a merge commit, never a squash or a
  rebase**, so `develop` contains the tagged release commit. A squash or a rebase copies `main`'s changes into
  new commits instead, so `develop` never contains the tagged commit, and a later release pull request can
  conflict with a hotfix that `develop` already holds under a different commit. `develop` also allows squash
  merges, for consolidating a feature branch, and GitHub's merge button offers the method used last, so check it
  before merging `main` back.
- **Stable release tags go only on `main`.** Prerelease tags may go on `develop`.

## Changelog

`CHANGELOG.md` follows [Keep a Changelog 1.1.0](https://keepachangelog.com/en/1.1.0/). Every pull request with a
user-visible change adds its entry under `## [Unreleased]`, in one of these sections, in this order: `Added`,
`Changed`, `Deprecated`, `Removed`, `Fixed`, `Security`. The rules for what an entry says are in the
`CHANGELOG.md` section of [AGENTS.md](./AGENTS.md), and they apply to everyone, not only to AI assistants.

`[Unreleased]` exists only on `develop`, and only while it has entries. A release renames it to the version
heading, so right after a release there is none. The next pull request that adds an entry also adds the section at
the top, linked to a comparison from the last release:
`## [Unreleased](https://github.com/Sekers/ScriptMessage/compare/<release>...develop)`, where `<release>` is the
latest release tag. The Release workflow refuses to publish a tag whose `CHANGELOG.md` still has an `[Unreleased]`
section.

## Releasing a version

1. **Choose the version** from the `[Unreleased]` entries, following [Deciding hard cases](#deciding-hard-cases).
2. **Open a release-prep pull request into `develop`** that changes only:
   - `ModuleVersion`, and `Prerelease` for a prerelease, in `ScriptMessage/ScriptMessage.psd1`.
   - `CHANGELOG.md`: rename `## [Unreleased](...)` to a dated heading such as
     `## [1.2.0](https://github.com/Sekers/ScriptMessage/tree/1.2.0) - 2026-09-20`. Do not add a new
     `[Unreleased]` section.
3. **Open a pull request from `develop` into `main`**, and merge it with a merge commit once the checks pass. Merge
   nothing else into `develop` until then, so no new `[Unreleased]` section reaches `main`.
4. **Tag the merge commit with an annotated tag** (add `-s` to sign it when you have signing set up), and push
   the tag:

   ```powershell
   git fetch origin
   git tag -a 1.2.0 origin/main -m 'ScriptMessage 1.2.0'
   git push origin 1.2.0
   ```

5. **The Release workflow** (`.github/workflows/Release.yml`) runs these jobs in order and stops at the first
   failure:
   1. Confirms the tag is annotated, matches the manifest version, has a dated `CHANGELOG.md` section and no
      `[Unreleased]` section, and points at a commit on `main` (or on `develop`, for a prerelease).
   2. Runs the test suite (`.github/workflows/Tests.yml`).
   3. Builds the package once and publishes it to the PowerShell Gallery, in the `psgallery` environment.
   4. Creates the GitHub release from the version's changelog section, attaches the same package, and
      publishes the release.
6. **Merge `main` back into `develop` with a merge commit**, not a squash ([Branches](#branches) explains why).

The GitHub release is created last so that a failed check or publish never leaves behind a public release, which
immutable releases would lock, for a version that did not ship.

### If the Release workflow fails

- **Job 1, 2, or 3 failed:** the Gallery has not accepted the package and no GitHub release exists. Fix the
  problem through pull requests, delete the tag (a repository admin must bypass the tag ruleset), and tag the new
  merge commit with the same version.
- **Only job 4 failed:** the version has shipped to the Gallery. Use **Re-run failed jobs** so only the GitHub
  release is retried. Do not re-run all jobs, because the Gallery rejects a second publish of the same version.

## Hotfixes

1. Create `hotfix/<version>` from `main`, for example `hotfix/1.2.1`.
2. Fix the problem, add tests, bump the patch version, and add a dated changelog section for the hotfix.
3. Open a pull request into `main`, merge it with a merge commit, and tag and push as in
   [Releasing a version](#releasing-a-version), steps 4 and 5.
4. Merge `main` into `develop` with a merge commit. In `CHANGELOG.md`, keep the `[Unreleased]` section from
   `develop` at the top, with the hotfix's section below it.

## Prereleases

A prerelease follows the same steps, with three differences: the release-prep pull request also sets
`Prerelease`, the changelog heading names the prerelease (`## [2.0.0-rc01](...) - 2026-09-20`), and the tag goes
on the release-prep merge commit on `develop` instead of on `main`. When the final version ships, its changelog
section covers everything since the last stable release, not only the changes since the last prerelease.

## Repository settings

A repository admin configures these once in the GitHub repository settings. They are what make the rules above
enforceable:

- **Branch rulesets for `main` and `develop`:** require a pull request, require the `Workflow lint` and `Pester`
  checks to pass, and block force pushes and deletion. For `main`, allow only merge commits. For `develop`, also
  allow squash merges, for consolidating a feature branch.
- **A tag ruleset for release tags:** target tags that look like versions, and block updates and deletion, with
  bypass only for admins.
- **Immutable releases:** enabled, so a published release's tag and assets cannot change.
- **A `psgallery` environment:** holds the `PS_GALLERY_KEY` secret, is limited to release tags, and can require a
  maintainer's approval. Remove any repository-level copy of the secret. Use a PowerShell Gallery API key scoped
  to the ScriptMessage package with only the "Push new or update packages" permission. Gallery API keys expire
  after at most 365 days, the Gallery emails the key's owner 10 days before expiry, and an expired key fails the
  publish job.

## Release checklist

- [ ] Every change listed under `[Unreleased]` has been reviewed and merged into `develop`.
- [ ] The version increment follows [Versioning](#versioning).
- [ ] The release-prep pull request changed only the manifest version and the changelog heading, and added no
      `[Unreleased]` section.
- [ ] `develop` was merged into `main` with a merge commit.
- [ ] The tag is annotated, has no `v` prefix, and is on the merge commit.
- [ ] Every job in the Release workflow passed.
- [ ] The version appears on the PowerShell Gallery, and the GitHub release has the right notes and package.
- [ ] The wiki (`Sekers/ScriptMessage.wiki`) describes this release, and its update was pushed after the Release
      workflow passed.
- [ ] `main` was merged back into `develop` with a merge commit.
- [ ] Anything that needs a live Microsoft Graph test was tested separately in a test tenant. The automated tests
      never connect to a tenant.

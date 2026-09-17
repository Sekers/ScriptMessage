# Contributing to ScriptMessage

Thank you for helping improve ScriptMessage. This guide covers the everyday workflow.
[RELEASING.md](./RELEASING.md) covers versioning, branches, and releases in full.

## Workflow

1. Fork the repository, or create a branch in it if you have write access.
2. Branch from `develop` as `feature/<short-name>`. Urgent fixes to the released version go on a
   `hotfix/<version>` branch from `main` instead; check with a maintainer first.
3. Make your change, following the conventions of the surrounding code.
4. Run the tests under both PowerShell editions (Windows PowerShell 5.1 and PowerShell 7):

   ```powershell
   .\Tests\Invoke-Tests.ps1
   ```

5. If users will notice the change, add an entry to `CHANGELOG.md` under `[Unreleased]`. If that section is not
   there yet, as right after a release, add it at the top ([RELEASING.md](./RELEASING.md#changelog) shows the
   heading).
6. Open a pull request into `develop` (or into `main` for a hotfix) and complete the pull request template,
   including the release classification.

## Rules

- **Never commit a real configuration file, secret, or certificate.** Never name a real organization, tenant, or
  email domain in code, tests, examples, or commit messages; use an RFC 2606 reserved domain instead, such as
  `example.com`. Those can never be registered, so a placeholder that gets run unedited reaches nobody.
- **The automated tests never connect to a tenant.** If a change needs a live Microsoft Graph test, run it in
  your own test tenant and describe what you tested in the pull request.
- **Text files are UTF-8 with LF line endings and no byte order mark, and PowerShell files contain only ASCII.**
  Windows PowerShell 5.1 misreads non-ASCII characters in PowerShell files saved without a byte order mark.
  `.gitattributes` handles line endings, and [AGENTS.md](./AGENTS.md) explains the rules and how to check them.
- **Rebase only your own branch, and only before it merges.** Never force-push `develop` or `main`.

AI assistants working in this repository follow [AGENTS.md](./AGENTS.md), which also holds the detailed rules
for writing changelog entries.

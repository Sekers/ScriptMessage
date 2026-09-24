## Summary

<!-- What changes for someone using the module, and why? Link any related issue. -->

## Release classification

<!-- Check one. RELEASING.md explains what counts as part of the public contract. -->

- [ ] Major: breaks the public contract (an existing script or configuration file can stop working)
- [ ] Minor: adds backward-compatible functionality
- [ ] Patch: a backward-compatible fix
- [ ] None: no user-visible change (tests, CI, or contributor documentation only)

## Checklist

- [ ] Targets `develop`, or `main` for a hotfix
- [ ] `CHANGELOG.md` has an entry in the right section under `[Unreleased]` (added if missing), or the change is not user-visible
- [ ] `.\Tests\Invoke-Tests.ps1` passes
- [ ] Contains no real tenant names, email domains, secrets, or configuration files
- [ ] Any live Microsoft Graph testing is described above, and sent messages and files only to accounts you control

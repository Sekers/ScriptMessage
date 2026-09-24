# Certificate file behavior

How PowerShell and .NET load a `.pfx` certificate file, where that matters for `CertificateFile` authentication in
`Connect-ScriptMessage_MicrosoftGraph` (`ScriptMessage/Services/MicrosoftGraph.ps1`). Through `1.1.1` that
authentication type threw before PowerShell 7.4. This file records what stood in the way of Windows PowerShell 5.1,
what loading a certificate leaves behind in either edition, and how the module loads the file now (section 6).

Measured on **2026-09-23** on Windows 11 (10.0.26200), with Windows PowerShell **5.1.26100.9444** (.NET Framework
4.8.9345.0), PowerShell **7.6.6**, and `Microsoft.Graph.Authentication` **2.25.0**, in standalone scripts. The
certificates were throwaway self-signed RSA 2048 certificates, built in memory with PowerShell 7's
`CertificateRequest` and exported as `.pfx` files with and without a password. No certificate store was written to.
Only sections 5 and 6 connected to a tenant, with application permissions, and nothing was sent. For part of
section 5 and all of section 6, one throwaway certificate's public key was registered on the app temporarily.

**Claims are labelled with their evidence.**

- **Measured** means observed on the date above.
- **From source** means read from a file in this repository, or from a release tag.
- **Unverified** means nobody has tested it.

> **The one-line summary:** Windows PowerShell 5.1's `Get-PfxCertificate` cannot take a password, but .NET's
> `X509Certificate2` loads the same file in both editions, and `Connect-MgGraph` signs in with it in both. Loading
> normally writes the private key to the user profile, and under PowerShell 7 that file stays behind when the
> process exits, as the module's sign-in did through `1.1.1`. `EphemeralKeySet` writes nothing and signs in the
> same.

## 1. Measured: `Get-PfxCertificate` differs by edition

| Edition | `Get-Command Get-PfxCertificate -Syntax` |
| --- | --- |
| Windows PowerShell 5.1 | `[-FilePath] <string[]>` or `-LiteralPath <string[]>`, nothing else |
| PowerShell 7.6.6 | the same, plus `[-Password <securestring>] [-NoPromptForPassword]` |

In Windows PowerShell 5.1, `Get-PfxCertificate -FilePath` loaded a `.pfx` with no password, with a private key that
could sign. For a `.pfx` with a password it printed `Enter password:` and waited. With `-NonInteractive`, as a
scheduled task runs, it failed instead: "Windows PowerShell is in NonInteractive mode. Read and Prompt functionality is
not available."

**From source:** every release from `1.0.0` to `1.1.1` called `Get-PfxCertificate -NoPromptForPassword`, then again
with `-Password` from `MgApp_EncryptedCertificatePassword` if that failed, and threw before PowerShell 7.4.

## 2. Measured: loading with `X509Certificate2` works in both editions

`[System.Security.Cryptography.X509Certificates.X509Certificate2]::new(<path>, <SecureString>, <X509KeyStorageFlags>)`,
with PowerShell's location set to the folder holding the files and the process's working directory elsewhere. Both
editions gave the same results unless the table says otherwise.

| Case | Result |
| --- | --- |
| password file, right password | loads; `HasPrivateKey` is `True`, and the key signs SHA-256 data with PKCS #1 padding |
| password file, no password argument | `CryptographicException`: "The specified network password is not correct." |
| password file, empty `SecureString` | the same |
| password file, wrong password | the same |
| no-password file, no password argument | loads, with a private key |
| no-password file, a password given | `CryptographicException`: "The specified network password is not correct." |
| a file that does not exist | `CryptographicException`: "The system cannot find the file specified." |
| a bare file name, relative to PowerShell's location | the same; the process's working directory is used |
| that name after `$PSCmdlet.GetUnresolvedProviderPathFromPSPath()` | loads |
| a PowerShell drive path, `MyCerts:\with-password.pfx` | 5.1: `NotSupportedException`, "The given path's format is not supported." 7: `CryptographicException`, "The filename, directory name, or volume label syntax is incorrect." |

So .NET needs a full file system path, the same as `System.IO.File` in section 8 of
[PowerShell-Language-Behavior.md](./PowerShell-Language-Behavior.md), and it reports a missing password and a wrong
one with the same message.

## 3. Measured: the private key is written to the user profile while the certificate is loaded

The folders watched were `%APPDATA%\Microsoft\Crypto\Keys` and `%APPDATA%\Microsoft\Crypto\RSA\<user SID>`, by file
name only. Every key loaded as `RSACng`, and every file written went to `Keys`.

| Loader | Edition | While loaded | After `Dispose()` | After the process exits without `Dispose()` |
| --- | --- | --- | --- | --- |
| `X509Certificate2`, `DefaultKeySet` | 7 | one key file | removed | **left behind** (1,735 bytes) |
| `Get-PfxCertificate -Password` | 7 | one key file | removed | **left behind** (1,735 bytes) |
| `X509Certificate2`, `EphemeralKeySet` | 7 | none | none | none written |
| `X509Certificate2`, `DefaultKeySet` | 5.1 | one key file | removed | removed |
| `Get-PfxCertificate`, no-password file | 5.1 | one key file | removed | not tested |
| `X509Certificate2`, `EphemeralKeySet` | 5.1 | none | none | none written |

With `EphemeralKeySet` the private key still signed in both editions.

Through the module itself, under PowerShell 7: a child process imported this repository's module, replaced
`Connect-MgGraph` inside the module with a stand-in that kept the certificate it was given, and ran
`Connect-ScriptMessage` with `CertificateFile` settings and an encrypted password. After the process exited, one key
file (1,735 bytes) was left in `Keys`. The key files these runs left behind were deleted afterwards.

**From source:** `Connect-ScriptMessage_MicrosoftGraph` never disposes the certificate; it passes it to
`Connect-MgGraph -Certificate`, which keeps it. So through `1.1.1`, a script that connected with `CertificateFile`
authentication under PowerShell 7 left a copy of the app's private key in the profile of the account that ran it.
Section 5 measured this with the real `Connect-MgGraph` after a successful sign-in, with and without
`Disconnect-ScriptMessage`.

## 4. Measured: `Connect-MgGraph` takes a loaded certificate in both editions

With `Microsoft.Graph.Authentication` 2.25.0 imported in each edition, and no connection made, `Connect-MgGraph` had
the same three certificate parameters in both: `-Certificate` (`X509Certificate2`), `-CertificateSubjectName`, and
`-CertificateThumbprint`. None takes a file path, so the module has to load the file itself.

## 5. Measured: a certificate loaded from a file signs in, in both editions

Each attempt ran in a new process with application permissions. The direct attempts ran
`Connect-MgGraph -TenantId -ClientId -NoWelcome`, then `Get-MgContext`, then `Disconnect-MgGraph`. The module
attempts ran this repository's `Connect-ScriptMessage -ReturnConnectionInfo` with `CertificateFile` settings, an
encrypted password, and the real `Connect-MgGraph`, before section 6's change, so they loaded the file as every
release through `1.1.1` did.

The throwaway certificate was tried twice: first while it was not registered on the app, then after its public key
was registered. The first run shows that a sign-in uses the certificate it is given, since Entra rejected it; the
second shows the sign-in itself. As a control, the app's own certificate from `Cert:\CurrentUser\My`, used by
thumbprint, signed in under both editions.

| Loaded with | Edition | Not registered | Registered | Key file left after the process exited |
| --- | --- | --- | --- | --- |
| the module, `Connect-ScriptMessage` (`Get-PfxCertificate`) | 7 | not run | signed in | **one** |
| the module, then `Disconnect-ScriptMessage` | 7 | not run | signed in | **one** |
| `Get-PfxCertificate -Password` | 7 | `AADSTS700027` | signed in | **one** |
| `X509Certificate2`, `DefaultKeySet` | 7 | `AADSTS700027` | signed in | **one** |
| `X509Certificate2`, `EphemeralKeySet` | 7 | `AADSTS700027` | signed in | none |
| `X509Certificate2`, `DefaultKeySet` | 5.1 | `AADSTS700027` | signed in | none |
| `X509Certificate2`, `EphemeralKeySet` | 5.1 | `AADSTS700027` | signed in | none |

"Signed in" means `Get-MgContext` reported `AuthType` `AppOnly` and the app's `Mail.Send` role. The key file column
was the same in both runs where both ran, and a key file was left behind whether or not the attempt disconnected.
Every `AADSTS700027` arrived as a `Microsoft.Identity.Client.MsalServiceException`: "The certificate with identifier
used to sign the client assertion is not registered on application. [Reason - The key was not found. ...]". No
attempt failed on the machine.

Separately, in both editions, a signature over test data made with the private key, loaded with either
`DefaultKeySet` or `EphemeralKeySet`, verified with a copy of the certificate holding only the public key.

## 6. From source: how the module loads the file now

`Connect-ScriptMessage_MicrosoftGraph` has no PowerShell version check for `CertificateFile` authentication. It:

- Expands environment variables in `MgApp_CertificatePath` with `[regex]::Replace`, because a `-replace` scriptblock
  turns into text under Windows PowerShell 5.1 ([PowerShell-Language-Behavior.md](./PowerShell-Language-Behavior.md)
  section 4).
- Decides whether the path is relative with `IsPathRooted()` and `IsPSAbsolute()`, which gave the same answers in
  both editions ([PowerShell-Language-Behavior.md](./PowerShell-Language-Behavior.md) section 13).
- Resolves the result with `$PSCmdlet.GetUnresolvedProviderPathFromPSPath()`, since .NET needs a full file system
  path (section 2).
- Loads it with `New-Object` for `X509Certificate2` and `EphemeralKeySet`: first with an empty password, which opens a
  file that has none, then with `MgApp_EncryptedCertificatePassword`. When the first attempt fails, it reports a file
  that does not exist by its path before looking for a password, since .NET gives the same message for a missing
  password and a wrong one but names no file when the file is missing. `New-Object` rather than `::new()` lets the
  tests mock the load.

`Tests/Connect-ScriptMessage.Tests.ps1` loads real `.pfx` files, with and without a password, in both editions, and
fails if loading writes a key file to either folder in section 3. A copy of the module that loaded with
`DefaultKeySet` failed those tests in both editions.

**Measured:** after the change, each edition signed in through the module with the registered throwaway certificate,
each attempt in a new process: once with `Connect-ScriptMessage -ServiceConfig` and the full path, and once from a
configuration file whose `MgApp_CertificatePath` was the file name alone, with PowerShell's current location in
another folder. All four signed in with no warning, and none left a key file after the process exited.

## What this does not establish

- Only a sign-in was tried, with one certificate and the app's `Mail.Send` role. No message was sent, so a token's use
  after the sign-in was not exercised.
- Only one kind of `.pfx` was tried: a key created by .NET, which landed in `Keys` in both editions. A `.pfx` whose
  key names a legacy CryptoAPI provider, such as one exported from the Windows certificate store, may be written to
  `RSA\<user SID>` instead.
- PowerShell 7.0 to 7.3 were not run, so why releases through `1.1.1` required 7.4 is not established, and certificate
  file authentication was not tried on those versions.
- Linux and macOS were not run.
- Whether Windows ever removes a key file left behind this way was not examined.

# Certificate file behavior

How PowerShell and .NET load a `.pfx` certificate file, where that matters for `CertificateFile` authentication in
`Connect-ScriptMessage_MicrosoftGraph` (`ScriptMessage/Services/MicrosoftGraph.ps1`). That authentication type
throws before PowerShell 7.4, and this file records what stands in the way of Windows PowerShell 5.1 and what loading
a certificate leaves behind in either edition.

Measured on **2026-09-23** on Windows 11 (10.0.26200), with Windows PowerShell **5.1.26100.9444** (.NET Framework
4.8.9345.0), PowerShell **7.6.6**, and `Microsoft.Graph.Authentication` **2.25.0**, in standalone scripts. The
certificates were throwaway self-signed RSA 2048 certificates, built in memory with PowerShell 7's
`CertificateRequest` and exported as `.pfx` files with and without a password. No certificate store was written to.
Only section 5 connected to a tenant, with application permissions, and nothing was sent. For part of it, one
throwaway certificate's public key was registered on the app temporarily.

**Claims are labelled with their evidence.**

- **Measured** means observed on the date above.
- **From source** means read from a file in this repository, or from a release tag.
- **Unverified** means nobody has tested it.

> **The one-line summary:** Windows PowerShell 5.1's `Get-PfxCertificate` cannot take a password, but .NET's
> `X509Certificate2` loads the same file in both editions, and `Connect-MgGraph` signs in with it in both. Loading
> normally writes the private key to the user profile, and under PowerShell 7 that file stays behind when the
> process exits, which the module's own sign-in does today. `EphemeralKeySet` writes nothing and signs in the same.

## 1. Measured: `Get-PfxCertificate` differs by edition

| Edition | `Get-Command Get-PfxCertificate -Syntax` |
| --- | --- |
| Windows PowerShell 5.1 | `[-FilePath] <string[]>` or `-LiteralPath <string[]>`, nothing else |
| PowerShell 7.6.6 | the same, plus `[-Password <securestring>] [-NoPromptForPassword]` |

In Windows PowerShell 5.1, `Get-PfxCertificate -FilePath` loaded a `.pfx` with no password, with a private key that
could sign. For a `.pfx` with a password it printed `Enter password:` and waited. With `-NonInteractive`, as a
scheduled task runs, it failed instead: "Windows PowerShell is in NonInteractive mode. Read and Prompt functionality is
not available."

**From source:** `Connect-ScriptMessage_MicrosoftGraph` calls `Get-PfxCertificate -NoPromptForPassword`, then again
with `-Password` from `MgApp_EncryptedCertificatePassword` if that fails, and throws before PowerShell 7.4. Every
release from `1.0.0` to `1.1.1` loads the file with `Get-PfxCertificate`.

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
`Connect-MgGraph -Certificate`. So under PowerShell 7, a script that connects with `CertificateFile` authentication
leaves a copy of the app's private key in the profile of the account that runs it. Section 5 measured this with the
real `Connect-MgGraph` after a successful sign-in, with and without `Disconnect-ScriptMessage`.

## 4. Measured: `Connect-MgGraph` takes a loaded certificate in both editions

With `Microsoft.Graph.Authentication` 2.25.0 imported in each edition, and no connection made, `Connect-MgGraph` had
the same three certificate parameters in both: `-Certificate` (`X509Certificate2`), `-CertificateSubjectName`, and
`-CertificateThumbprint`. None takes a file path, so the module has to load the file itself.

## 5. Measured: a certificate loaded from a file signs in, in both editions

Each attempt ran in a new process with application permissions. The direct attempts ran
`Connect-MgGraph -TenantId -ClientId -NoWelcome`, then `Get-MgContext`, then `Disconnect-MgGraph`. The module
attempts ran this repository's `Connect-ScriptMessage -ReturnConnectionInfo` with `CertificateFile` settings, an
encrypted password, and the real `Connect-MgGraph`, which is what every release does to load the file.

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

## 6. From source: what else in `Connect-ScriptMessage_MicrosoftGraph` needs PowerShell 7

- Expanding environment variables in `MgApp_CertificatePath` uses `-replace` with a scriptblock, which Windows
  PowerShell 5.1 turns into text instead of running
  ([PowerShell-Language-Behavior.md](./PowerShell-Language-Behavior.md) section 4).
- Deciding whether `MgApp_CertificatePath` is relative uses `IsPathRooted()` and `IsPSAbsolute()`, which gave the same
  results in both editions ([PowerShell-Language-Behavior.md](./PowerShell-Language-Behavior.md) section 13).

## What this does not establish

- Only a sign-in was tried, with one certificate and the app's `Mail.Send` role. No message was sent, so a token's use
  after the sign-in was not exercised.
- Only one kind of `.pfx` was tried: a key created by .NET, which landed in `Keys` in both editions. A `.pfx` whose
  key names a legacy CryptoAPI provider, such as one exported from the Windows certificate store, may be written to
  `RSA\<user SID>` instead.
- PowerShell 7.0 to 7.3 were not run, so why the module requires 7.4 rather than an earlier 7.x is not established.
- Linux and macOS were not run.
- Whether Windows ever removes a key file left behind this way was not examined.

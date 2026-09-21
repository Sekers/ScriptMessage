# Changelog for ScriptMessage PowerShell Module

## [Unreleased](https://github.com/Sekers/ScriptMessage/compare/1.1.1...develop)

### Added

- New Parameter: `Send-ScriptMessage -MailType` > Chooses how an email with more than one recipient is sent. It overrides the configuration file's `MailType` setting, and `Group` is used when neither one sets it.
  - `Group` sends one email to all of the To, CC, and BCC recipients.
  - `OneOnOne` sends each To, CC, and BCC recipient a separate email with only that recipient in To, so no recipient can see who else received it. A recipient listed more than once gets one email.

### Fixed

- `Send-ScriptMessage` now honors an explicit `-ChatType OneOnOne`. The configuration file's `ChatType` setting was used in its place, so a call asking for one-on-one chats sent a single group chat when that setting was `Group`. An explicit `-ChatType Group` was unaffected.
- `Send-ScriptMessage` now honors an explicit `-IncludeBCCInGroupChat $false`. The configuration file's `IncludeBCCInGroupChat` setting was used in its place, so a call asking to keep BCC recipients out of a group chat added them anyway when that setting was `true`, making those addresses visible to everyone in the chat. An explicit `-IncludeBCCInGroupChat $true` was unaffected.
- `Send-ScriptMessage` now sends mail without a `ChatType` setting in the configuration file. A Mail-only send failed without it, reporting "Cannot convert null to type ChatType", even though `ChatType` only affects chat; a configuration file listing only `Mail` in `AllowableMessageTypes`, or one written before the setting was added in version 1.0.8, was affected. A `Chat` send with no chat type, because the setting is missing or blank and `-ChatType` was not passed, now reports that in the returned result's `Error` property and names both places it can be set.
- `Get-ScriptMessageContext -ReturnCachedContext` no longer reports a connection that `Disconnect-ScriptMessage` has already ended. Disconnecting left the cached context in place, so a cached read kept returning the account and scopes of the finished session, and only a read without `-ReturnCachedContext` showed the truth. Disconnecting now clears that service's cached context, so the next cached read asks the service again.
- Minor: `Send-ScriptMessage` now reads every true/false setting in the configuration file the same way, so a value other than `true` no longer turns a setting on.
- Minor: The help for `Get-ScriptMessageConfig` and `Set-ScriptMessageConfigFilePath` now describes the `-Path` parameter, which it left blank.
- Minor: The fourth `Send-ScriptMessage` help example, which sends attachments held in variables, now runs as written. It failed with "Missing closing ')' in subexpression" when copied, and it ended without calling `Send-ScriptMessage`.

---
## [1.1.1](https://github.com/Sekers/ScriptMessage/tree/1.1.1) - 2026-09-16

### Changed

- The sample configuration file `Templates/config_scriptmessage.json` now includes the `MgDelegatedPermission_RequestFilesReadWritePermission` setting, which requests permission to upload Teams chat attachments to OneDrive when using delegated permissions. Earlier versions already read this setting, so it can be added to an existing configuration file.

### Fixed

- `Send-ScriptMessage` no longer rejects recipients given as a one-item array holding a recipient object, such as a single `Name` and `Address` entry read from a JSON configuration file. When those were the only recipients, it failed with "Please provide at least one parameter value for any of the following: To, CC, or BCC" and sent nothing.
- `Send-ScriptMessage` now sends Teams chat messages that have no attachments. A call using `-Type Chat` delivered nothing and returned "Cannot index into a null array" in the result's `Error` property, unless the same call also sent mail or the chat included an attachment. Both `OneOnOne` and `Group` chats were affected.
- On Windows PowerShell 5.1, `Send-ScriptMessage` no longer fails to send Teams chat messages that have an attachment. Such a send could deliver nothing and report "Unable to find type [System.Web.HttpUtility]" in the result's `Error` property.
- In both PowerShell editions, `Send-ScriptMessage` now escapes spaces in chat attachment filenames correctly when uploading them. A filename containing a space was uploaded to a path holding a plus sign in place of each space.
- Minor: `Send-ScriptMessage` now stops with that same error when the only recipients have blank addresses, such as an object or hashtable whose `Address` is empty, or a string of only spaces.
- Minor: `Connect-ScriptMessage` and `Send-ScriptMessage` no longer require the `Microsoft.Graph.Files` module when `MgPermissionType` is `Application`. If `MgDelegatedPermission_RequestFilesReadWritePermission` was also `true` and the module was not installed, they failed with "Please first install the following sub-modules" and nothing was sent. With delegated permissions, that setting still requires the module.
- Minor: With `MgApp_AuthenticationType` set to `CertificateFile`, the error reported when the certificate needs a password and `MgApp_EncryptedCertificatePassword` is empty now names Microsoft Graph.
- Minor: On Windows PowerShell 5.1, a configuration file saved as UTF-8 without a byte order mark is now read correctly when it contains non-ASCII characters, such as an accented letter in `MgApp_CertificateName`. `Get-ScriptMessageConfig`, `Connect-ScriptMessage`, and `Send-ScriptMessage` silently read those characters as different ones. This is how PowerShell 7's `Set-Content` and `Out-File` save a file.
- Minor: On PowerShell 7, a configuration file saved in the Windows ANSI code page is now read correctly when it contains non-ASCII characters. `Get-ScriptMessageConfig`, `Connect-ScriptMessage`, and `Send-ScriptMessage` silently read those characters as replacement characters. This is how Windows PowerShell 5.1's `Set-Content` saves a file.
- Minor: A configuration file whose path contains square brackets, such as a folder named `Scripts [old]`, can now be used. `Get-ScriptMessageConfig`, `Connect-ScriptMessage`, and `Send-ScriptMessage` stopped with "Can't find the JSON configuration file" and sent nothing.

---
## [1.1.0](https://github.com/Sekers/ScriptMessage/tree/1.1.0) - (2025-10-16)

### Fixes

- ScriptMessage now properly allows for for NULL in TO, CC, BCC parameters in Microsoft Graph.
- Reverted recently added code formatting that included Begin\End process blocks as the way they were implemented caused bugs.
- Catch if a Teams Chat attachments folder does not already exists in the user's OneDrive.
- The 'AllowableMessageTypes' Microsoft Graph configuration is now being applied properly.

### Features

- Much better error handling with the Microsoft Graph service. Warnings and Errors are now collected and if one service type has an error the command will still process the other service type (when applicable).
- Get-ScriptMessageConfig can now optionally pull data for a specific service only. Returning all services configurations is still the default.
- Sending multiple attachments in a single email or chat now works with Microsoft Graph.
- Sending attachments with non-standard filenames (uploaded to the user's OneDrive) using Microsoft Graph Chat is now supported.

### Other

- BREAKING CHANGE: Renamed the Microsoft Graph service from 'MgGraph' to 'MicrosoftGraph'. Please update your configuration files as needed.
- Changes made to only import necessary Microsoft Graph PowerShell modules.

Author: [**@Sekers**](https://github.com/Sekers)

---
## [1.0.8](https://github.com/Sekers/ScriptMessage/tree/1.0.8) - (2025-09-18)

### Fixes

- Get-ScriptMessageContext will return common info (like service name) rather than NULL if nothing comes back from the service (e.g., if it's not connected).

### Features

- BREAKING CHANGE: Updates to the return info when sending messages. This is to make the responses more consistent across various services going forward.
- New Function: Disconnect-ScriptMessage > Disconnects from the specified messaging service ahead of sending the message, if possible.
- Enabled support for Graph Module Chat (delegated permissions only since Graph does not support application permissions for chat).
- ScriptMessage now caches service context information. Use Get-ScriptMessageContext with the 'ReturnCachedContext' parameter to have the cmdlet return the cached context information (if it exists to) reduce API calls.
- Attachments are now supported in messages.

### Other

- README and built-in help updates.
- Added a few warning messages when using unsupported actions for the Microsoft Graph SDK.

Author: [**@Sekers**](https://github.com/Sekers)

---
## [1.0.7](https://github.com/Sekers/ScriptMessage/tree/1.0.7) - (2024-12-26)

### Features

- New messaging class "MessageServiceType" and corresponding parameters for the Send-ScriptMessage Cmdlet to allow multiple/combined service & type parameters in one call.

### Other

- Results returned from Send-ScriptMessage has recipients adjusted from Hashtables to PSCustomObjects. This better handles collections (arrays) of addresses than hashtables being returned. Originally it was hashtables to mimic what Graph uses but that's not great for our purpose.

Author: [**@Sekers**](https://github.com/Sekers)

---
## [1.0.6](https://github.com/Sekers/ScriptMessage/tree/1.0.6) - (2024-12-19)

### Fixes

- Support for Client Secret authentication in Microsoft Graph 2.x and newer.

### Features

- Added in an "AllowableMessageTypes" configuration setting for specific services to configure which types of messaging (Mail, Chat, etc.) are available for use with those services.

### Other

- BREAKING CHANGE: Removed support for Microsoft Graph SDK version 1.x when using Client Secret authentication. Use version 2.x of the SDK or newer.
- Removed Graph SDK beta support since the way beta works has been changed by Microsoft and supporting it adds unnecessary complexity.
- Adjusted code formatting to include Begin\End process blocks to better accommodate future code updates.
- Minor code adjustments and spelling fixes to assist with debugging.

Author: [**@Sekers**](https://github.com/Sekers)

---
## [1.0.5](https://github.com/Sekers/ScriptMessage/tree/1.0.5) - (2024-05-22)

### Features

- Added in an "All" recipients field when returning message send results, for convenience.

Author: [**@Sekers**](https://github.com/Sekers)

---
## [1.0.4](https://github.com/Sekers/ScriptMessage/tree/1.0.4) - (2024-05-21)

### Fixes

- Renamed 'Sender' parameters & variables to 'SenderId' since 'Sender' is an [automatic variable](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_automatic_variables) that is built into PowerShell and assigning to it might have undesired side effects in some circumstances.

Author: [**@Sekers**](https://github.com/Sekers)

---
## [1.0.3](https://github.com/Sekers/ScriptMessage/tree/1.0.3) - (2024-05-21)

### Features

- Added an optional message type parameter to 'Send-ScriptMessage' to differentiate types of messages (chat, mail, etc.) for message services that support multiple kinds of messages. Defaults to 'Mail' if not specified.

Author: [**@Sekers**](https://github.com/Sekers)

---
## [1.0.2](https://github.com/Sekers/ScriptMessage/tree/1.0.2) - (2024-05-21)

### Fixes

- Resolved issue with an internal function where a Microsoft Graph recipient returns empty when a recipient is inside an array of arrays.

Author: [**@Sekers**](https://github.com/Sekers)

---
## [1.0.1](https://github.com/Sekers/ScriptMessage/tree/1.0.1) - (2024-05-20)

### Fixes

- Removed default command prefix leftover from initial testing.

Author: [**@Sekers**](https://github.com/Sekers)

---
## [1.0.0](https://github.com/Sekers/ScriptMessage/tree/1.0.0) - (2024-05-20)

### Features

- Initial public release

Author: [**@Sekers**](https://github.com/Sekers)

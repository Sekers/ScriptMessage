# Changelog for ScriptMessage PowerShell Module

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
## [1.0.0](https://github.com/Sekers/SKYAPI/tree/1.0.0) - (2024-05-20)

### Features

- Initial public release

Author: [**@Sekers**](https://github.com/Sekers)

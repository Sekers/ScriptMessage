# Changelog for ScriptMessage PowerShell Module

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

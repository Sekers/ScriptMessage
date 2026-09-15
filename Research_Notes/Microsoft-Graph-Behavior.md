# Microsoft Graph behavior

What Microsoft Graph and the Microsoft Graph PowerShell SDK do in the areas this module relies on:
permissions, what a successful send returns, chat creation, attachment and upload limits, and how SDK sign-in
persists. Each claim is tied to the part of the module it affects.

**Nothing here was measured against a tenant.** Documentation was read on **2026-09-15** from the pages listed
under Sources. Microsoft revises these pages, so re-read a page before relying on a limit or a permission.

**Claims are labelled with their evidence.**

- **From documentation** means stated on a page listed under Sources, on the date above.
- **From source** means read out of `ScriptMessage/Services/MicrosoftGraph.ps1` or
  `ScriptMessage/Public/Send-ScriptMessage.ps1`; it describes what this module does, which is not the same as
  what Graph does.
- **Inference** means reasoned from documentation or source without a test.
- **Unverified** means nobody has tested it or found it documented. Several of these are assumptions the module
  makes today.

> **The one-line summary:** a successful mail send means Exchange accepted the request, not that the message
> was delivered; posting a Teams chat message needs delegated permissions; and Microsoft's terms of use forbid
> using Teams as a log file, which matters for a module built to send messages from scripts.

## Sources

| Page | Last updated (page metadata) |
| --- | --- |
| [user: sendMail](https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0) | 2026-06-19 |
| [Send message in a chat](https://learn.microsoft.com/en-us/graph/api/chat-post-messages?view=graph-rest-1.0) | 2026-05-19 |
| [Create chat](https://learn.microsoft.com/en-us/graph/api/chat-post?view=graph-rest-1.0) | 2025-12-03 |
| [Attach large files to Outlook messages or events](https://learn.microsoft.com/en-us/graph/outlook-large-attachments) | 2025-08-06 |
| [Upload small files](https://learn.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0) | 2026-08-11 |
| [Use Microsoft Graph PowerShell authentication commands](https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands?view=graph-powershell-1.0) | 2026-05-06 |

## 1. From documentation: permissions for each operation the module performs

| Operation | Call in the module | Delegated (least privileged) | Application (least privileged) |
| --- | --- | --- | --- |
| Send mail | `Send-MgUserMail` | `Mail.Send` | `Mail.Send` |
| Create a chat | `New-MgChat` | `Chat.Create` | `Chat.Create` |
| Send a chat message | `New-MgChatMessage` | `ChatMessage.Send` | `Teamwork.Migrate.All` only |
| Upload a file to OneDrive | `Invoke-MgGraphRequest -Method PUT` | `Files.ReadWrite` | `Files.ReadWrite.All` |

`Teamwork.Migrate.All` belongs to the message import (migration) flow, and the page's import example requires
the target chat to be in migration mode. So an app using application permissions can create a chat but cannot
post an ordinary message into it.

**From source:** with delegated permissions the module requests `email`, `offline_access`, `openid`, and
`profile`; plus `Mail.Send` when Mail is allowed; plus `Chat.Create`, `ChatMessage.Send`, and either
`Chat.Read` or `Chat.ReadBasic` when Chat is allowed; plus `Files.ReadWrite` when
`MgDelegatedPermission_RequestFilesReadWritePermission` is true. With application permissions the module
requests no scopes (the app registration decides) and refuses to send chat, adding a warning to the result.

**Unverified:** the permissions needed by the other calls the module makes: `Get-MgChat` (finding an existing
group chat), `Get-MgUserDrive` and `Get-MgDriveItem` (the attachment folder), and `Invoke-MgInviteDriveItem`
(sharing uploaded files). Those pages were not read.

## 2. From documentation: Teams must not be used as a log file

The Send message in a chat page states: "It's a violation of the terms of use to use Microsoft Teams as a log
file. Only send messages that people will read." The same page says the API "can't create a new chat", so the
chat has to exist before a message is posted.

**Inference:** a script that posts a chat message for every item it processes is the pattern this forbids.
Chat suits alerts a person should act on; high-volume status output belongs in mail or a log file. User-facing
documentation for the Chat message type should say so.

**From source:** the module creates or finds the chat before posting, consistent with the page.

## 3. From documentation: what a successful send returns

- **sendMail** returns `202 Accepted` with an empty body. The page: "A `202 Accepted` response code indicates
  that the request has been accepted; however, it doesn't indicate that the request processing has completed.
  Delivery of the message is subject to Exchange Online limitations and throttling." `saveToSentItems` defaults
  to true.
- **Send message in a chat** returns `201 Created` with the new chatMessage object.

**From source:** for Mail, the result's `Status` holds what `Send-MgUserMail -PassThru` returned. For Chat,
`Status` holds the chatMessage object from `New-MgChatMessage`.

**Inference:** a Mail result with no errors means Exchange accepted the message, not that it was delivered, and
the module has no way to report a later delivery failure.

**Unverified:** that `Send-MgUserMail -PassThru` returns only `$true`. A comment in the module says so; it was
not checked.

## 4. From documentation: creating one-on-one and group chats

The Create chat page: "Only one one-on-one chat can exist between two members. If a one-on-one chat already
exists, this operation returns the existing chat and doesn't create a new one."

**From source:** for `ChatType` `OneOnOne`, the module calls `New-MgChat` for each recipient on every send and
relies on this to land in the existing chat. For `Group`, it first lists all of the sender's group chats
(`Get-MgChat -All`, expanding members) and posts into one whose members exactly match the sender plus the
recipients, creating a new group chat only when none matches.

**Unverified:** whether `New-MgChat` for a group chat always creates a new chat even when one with the same
members already exists. The page does not say, and the module's lookup exists on the assumption that it does.

## 5. From documentation: attachment and upload size limits

- **Attachments on an Outlook message item:** "you can attach files up to 150 MB to an Outlook message". Under
  3 MB, "do a single POST on the attachments navigation property of the Outlook item"; between 3 MB and 150 MB,
  "create an upload session". Creating an upload session for a message needs `Mail.ReadWrite`. Both routes add
  the attachment to an existing message (a draft), not to a `sendMail` call.
- **sendMail:** "When using JSON format, you can include a file attachment in the same sendMail action call."
  The page gives no size limit for that.
- **OneDrive simple upload** (`PUT .../content`): "This method only supports files up to 250 MB in size."
  Larger files need an upload session.

**From source:** mail attachments go inline in the `sendMail` request as base64 `contentBytes`, and the module
requests `Mail.Send` only. Chat attachments are uploaded with a simple `PUT` to
`root:/Microsoft Teams Chat Files/<name>:/content` in the sender's OneDrive (with a number added to the name if
that name is taken), shared with the chat recipients as read-only with sign-in required and no invitation
email, and then referenced from the chat message.

### What this does not establish

- **Unverified:** the largest attachment, or total request size, that `sendMail` accepts inline. The 3 MB figure
  is documented for adding an attachment to an existing message, so do not quote it as the `sendMail` limit.
- **Inference:** supporting mail attachments past whatever the inline limit is would mean sending through a
  draft message and an upload session, which needs `Mail.ReadWrite`, a broader permission than the module
  requests today.
- **Unverified:** what the simple `PUT` does when a file with that name already exists. The page does not say; a
  comment in the module says it overwrites, and the module renames the file to avoid the question.
- **Unverified:** that `Microsoft Teams Chat Files` is the folder Teams itself uses for chat attachments. The
  module assumes it; none of the pages above mention it.

## 6. From documentation: how SDK sign-in persists, and other clouds

- "sign-in persists across PowerShell sessions because Microsoft Graph PowerShell securely caches the token when
  using the default `CurrentUser` context scope. If you use the `-ContextScope Process` parameter with
  `Connect-MgGraph`, sign-in only persists for the current PowerShell session."
- "`Disconnect-MgGraph` clears the cached token and ends the session for the current context scope, requiring
  you to authenticate again for future commands."
- "By default, `Connect-MgGraph` targets the global public cloud." The `-Environment` parameter selects another
  cloud, and `Get-MgEnvironment` lists each cloud's Graph endpoint (for example `https://graph.microsoft.us` for
  `USGov`).
- The page recommends PowerShell 7 or later when connecting with client secret credentials.

**From source:** the module calls `Connect-MgGraph` without `-ContextScope` or `-Environment`, and when
`MgDisconnectWhenDone` is true (the template's default) `Send-ScriptMessage` calls `Disconnect-MgGraph` after
sending. The chat member binding and the OneDrive upload URL are built on a hard-coded
`https://graph.microsoft.com`. Client secret authentication is offered on both PowerShell editions.

**Inference:** with the default `CurrentUser` scope, `MgDisconnectWhenDone` clears a token cache that other
PowerShell sessions of the same Windows user also rely on. And because the endpoint is hard-coded, chat and chat
attachments would not work in another cloud even if `-Environment` were passed.

### What this does not establish

- **Unverified:** whether the SDK can hold more than one connection in a process. The page describes "your
  current session" but does not say what a second `Connect-MgGraph` does to an existing connection. The concern
  that ScriptMessage replaces a calling script's own Graph connection needs a live test, which needs a tenant
  and permission.
- **Unverified:** whether `Disconnect-MgGraph` in one session signs out another running session that shares the
  `CurrentUser` token cache.

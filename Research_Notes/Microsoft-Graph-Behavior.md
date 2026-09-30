# Microsoft Graph behavior

What Microsoft Graph and the Microsoft Graph PowerShell SDK do in the areas this module relies on:
permissions, what a successful send returns, chat creation and reuse, attachment and upload limits, how SDK sign-in
persists, and what disconnecting after a send changes. Each claim is tied to the part of the module it affects.

**Only sections 7 and 9 were measured against a tenant**, on **2026-09-16** and **2026-09-30**; each section
gives its environment. Documentation was read on **2026-09-15** from the pages listed under Sources, and the SDK
source on **2026-09-16**, except for section 8, whose pages and SDK source were read on **2026-09-21**, and
section 9, whose pages were read on **2026-09-30**. Microsoft revises these pages, so re-read a page before
relying on a limit or a permission.

**Claims are labelled with their evidence.**

- **Measured** means observed in a test, in the environment the section describes.
- **From documentation** means stated on a page listed under Sources, on the date above.
- **From source** means read out of `ScriptMessage/Services/MicrosoftGraph.ps1` or
  `ScriptMessage/Public/Send-ScriptMessage.ps1`; it describes what this module does, which is not the same as
  what Graph does.
- **From SDK source** means read out of the Microsoft Graph PowerShell SDK source at the release tag or branch
  listed under Sources; a later SDK release can behave differently.
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
| [Microsoft Graph PowerShell SDK authentication source, tag `2.25.0`](https://github.com/microsoftgraph/msgraph-sdk-powershell/tree/2.25.0/src/Authentication) | none (release tag) |
| [Connect-MgGraph](https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.authentication/connect-mggraph?view=graph-powershell-1.0) | 2026-03-02 |
| [Install the Microsoft Graph PowerShell SDK](https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0) | 2025-07-23 |
| [Microsoft Graph permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference) | 2026-09-15 |
| [Scopes and permissions in the Microsoft identity platform](https://learn.microsoft.com/en-us/entra/identity-platform/scopes-oidc) | 2026-06-25 |
| [Configure how users consent to applications](https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/configure-user-consent) | 2026-08-04 |
| [How to register an app in Microsoft Entra ID](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app) | 2026-06-15 |
| [How to add a redirect URI to your application](https://learn.microsoft.com/en-us/entra/identity-platform/how-to-add-redirect-uri) | 2026-06-15 |
| [Add and manage app credentials in Microsoft Entra ID](https://learn.microsoft.com/en-us/entra/identity-platform/how-to-add-credentials) | 2026-06-15 |
| [Microsoft Entra authentication and authorization error codes](https://learn.microsoft.com/en-us/entra/identity-platform/reference-error-codes) | 2026-06-15 |
| [Role Based Access Control for Applications in Exchange Online](https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac) | 2026-08-21 |
| [ConvertFrom-SecureString](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/convertfrom-securestring) | 2026-08-10 |
| [about_Signing](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_signing) | 2026-09-21 |
| [Microsoft Graph PowerShell SDK authentication source, branch `main`](https://github.com/microsoftgraph/msgraph-sdk-powershell/tree/main/src/Authentication) | none (branch, read 2026-09-21) |
| [List chats](https://learn.microsoft.com/en-us/graph/api/chat-list?view=graph-rest-1.0) | 2026-04-13 |
| [chat resource type](https://learn.microsoft.com/en-us/graph/api/resources/chat?view=graph-rest-1.0) | 2026-08-05 |

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

**From source:** with `MailType` `Group`, `Send-ScriptMessage_MicrosoftGraph` calls `Send-MgUserMail` once, with
no `-ErrorAction`, inside a `try` whose `catch` puts the error in the result's `Error` property. With `OneOnOne`
it calls `Send-MgUserMail` once per unique recipient with `-ErrorAction Stop`, so each failure reaches that
recipient's own `catch` and the remaining recipients are still sent.

**Unverified:** whether a failed `Send-MgUserMail` raises a terminating error. The `Group` send depends on it: if
the error is non-terminating, it goes to the caller's error stream instead of `Error`, and `Status` is empty.

**Unverified:** how Exchange Online throttles many `sendMail` requests from one mailbox in quick succession,
which a `OneOnOne` send to a long recipient list makes. The sendMail page quoted above gives no figures.

## 4. From documentation: creating one-on-one and group chats

The Create chat page: "Only one one-on-one chat can exist between two members. If a one-on-one chat already
exists, this operation returns the existing chat and doesn't create a new one."

**From source:** for `ChatType` `OneOnOne`, the module calls `New-MgChat` for each recipient on every send and
relies on this to land in the existing chat. For `Group`, it first lists all of the sender's group chats
(`Get-MgChat -All`, expanding members) and posts into one whose members exactly match the sender plus the
recipients, creating a new group chat only when none matches.

**Measured** (section 9): creating a group chat with the same members as an existing one created a second chat
rather than returning the first. The page does not say, and the module's lookup exists because of it.

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
`MgDisconnectWhenDone` is true `Send-ScriptMessage` calls `Disconnect-MgGraph` after sending. The chat member
binding and the OneDrive upload URL are built on a hard-coded `https://graph.microsoft.com`. Client secret
authentication is offered on both PowerShell editions.

**From SDK source:** the `CurrentUser` default applies only to delegated sign-in. Certificate and client secret
connections, which the module uses for application permissions, default to `Process`, and section 7 measures
what that means. `Disconnect-MgGraph` also does less than the page's wording suggests: it clears the in-memory
token cache and deletes `~/.mg/mg.authrecord.json` (the file a delegated sign-in reads to find its cached
account), but leaves the on-disk token cache in place.

**Inference:** because the endpoint is hard-coded, chat and chat attachments would not work in another cloud even
if `-Environment` were passed.

### What this does not establish

- **Unverified, delegated permissions only:** what `Disconnect-MgGraph`, and so `MgDisconnectWhenDone`, does to
  other PowerShell sessions of the same Windows user. Reading the SDK source suggests a session already signed in
  keeps working, while the next `Connect-MgGraph` in any session prompts again because `mg.authrecord.json` is
  gone; neither was tested. Application-permission connections share nothing with other sessions (section 7).

## 7. Measured: application-permission connections and `MgDisconnectWhenDone`

**Environment:** measured on 2026-09-16 in PowerShell 7.6.6 and Windows PowerShell 5.1.26100.9444 on Windows 11,
with `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users.Actions`, and `Microsoft.Graph.Teams` 2.25.0 in
both editions. The configuration used `MgPermissionType` `Application` and `MgApp_AuthenticationType`
`CertificateThumbprint`, for an app registration whose token carried only the `Mail.Send` application role.
**No message was sent:** the tests ran `Connect-ScriptMessage` and then, where needed, the `Disconnect-MgGraph`
call that `Send-ScriptMessage` makes after a send, and never called `Send-ScriptMessage`. Both editions gave
the same results.

**Measured:**

- After `Connect-ScriptMessage`, `Get-MgContext` reported `AuthType` `AppOnly`, `ContextScope` `Process`, and
  `TokenCredentialType` `ClientCertificate`.
- A separate PowerShell process started while connected reported no Graph connection.
- Connecting wrote no token file: `~/.mg` and `%LOCALAPPDATA%\.IdentityService` were unchanged.
- Connecting again in the same process without disconnecting reused the access token from memory (the SDK's
  debug output reported `source: Cache`). After `Disconnect-MgGraph`, the next connect requested a new token
  (`source: IdentityProvider`).
- `Disconnect-MgGraph` deleted a placeholder `~/.mg/mg.authrecord.json` even though the connection was app-only.
- **A script's own Graph connection does not survive a send.** The test script first connected with the same app
  and certificate by subject name, so its connection could be told apart from ScriptMessage's, which uses the
  thumbprint. After `Connect-ScriptMessage`, `Get-MgContext` showed ScriptMessage's connection. With the
  `Disconnect-MgGraph` step (`MgDisconnectWhenDone` true), no connection remained, and the script's next
  `Invoke-MgGraphRequest` failed with "Authentication needed. Please call Connect-MgGraph." Without that step
  (`MgDisconnectWhenDone` false), ScriptMessage's connection remained.

**From SDK source:** `ConnectMgGraph.cs` defaults `ContextScope` to `Process` for the certificate, client secret,
managed identity, and environment variable sign-ins, and to `CurrentUser` only for delegated sign-in.
`GraphSession.cs` keeps one static session holding one connection, so each `Connect-MgGraph` replaces the
process's connection whatever kind it was. `AuthenticationHelpers.LogoutAsync`, which `Disconnect-MgGraph` calls,
clears the in-memory token cache, clears the connection, and deletes `mg.authrecord.json`.

**From source:** `Send-ScriptMessage` runs `Connect-ScriptMessage_MicrosoftGraph` before every send, and
`Disconnect-MgGraph` after it only when the configuration's `MgDisconnectWhenDone` is true, so a configuration
without the setting stays connected. `Connect-ScriptMessage_MicrosoftGraph` passes no `-ContextScope`, so every
application authentication type gets `Process` scope.

**Inference, for `MgDisconnectWhenDone` with application permissions:**

- A script that runs in its own process and then exits keeps nothing either way. The only lasting effect of
  true is deleting `mg.authrecord.json`, which a delegated sign-in by the same Windows user relies on.
- Neither setting preserves a calling script's own Graph connection. With true, the script is left with no
  connection after every send. With false, it is left with ScriptMessage's connection, which works when the
  script used the same app and otherwise runs its later Graph commands with ScriptMessage's app and permissions.
  `Templates/config_scriptmessage.json` sets `MgDisconnectWhenDone` to false for this reason.
- With true, every send requests a new token. With false, sends in the same process reuse one token until it
  expires.

### What this does not establish

- **Unverified:** replacement of a calling script's connection made with a different app, delegated permissions,
  or a managed identity. Only a connection with the same app was measured; the SDK source says any connection is
  replaced.
- **Unverified:** the `CertificateFile`, `CertificateName`, and `ClientSecret` authentication types. The SDK
  source gives them `Process` scope too; they were not measured.
- **Unverified:** runspaces sharing a process, such as `ForEach-Object -Parallel` or thread jobs. The static
  session suggests they share one connection; this was not tested.
- Nothing in this section was measured with delegated permissions or with an SDK release other than 2.25.0.

## 8. From documentation: app registration, consent, credentials, and limiting mailbox access

These pages back the setup steps on the wiki's Microsoft Graph page and several setting entries on its Home page.
Nothing in this section was tested against a tenant.

**From documentation:**

- **Redirect URIs for delegated sign-in with your own app.** The authentication commands page's steps for a
  custom application register `http://localhost` as a **Public client/native** redirect URI, and add: "An
  additional Redirect URI is required for Windows Authentication Manager (WAM) broker-based sign-in", entered
  under **Mobile and desktop applications** as `ms-appx-web://Microsoft.AAD.BrokerPlugin/<YOUR_APP_CLIENT_ID>`.
  The same page turns WAM off with `Set-MgGraphOption -DisableLoginByWAM $true`. The error code page describes
  `AADSTS50011` as "The reply address is missing, misconfigured, or doesn't match reply addresses configured for
  the app."
- **Admin consent.** The permissions reference lists admin consent as not required for the delegated `Mail.Send`,
  `Chat.Create`, `ChatMessage.Send`, `Chat.Read`, `Chat.ReadBasic`, and `Files.ReadWrite`, and as required for the
  application `Mail.Send`. The user consent page: "By default, all users are allowed to consent to applications
  for permissions that don't require administrator consent", and "Applications that require users to be assigned
  to the application must have their permissions consented by an administrator".
- **Certificates and client secrets.** The credentials page: "Microsoft recommends that you use a certificate
  instead of a client secret before moving the application to a production environment"; an uploaded certificate
  must be a `.cer`, `.pem`, or `.crt` file; "Client secret lifetime is limited to two years (24 months) or less",
  and "Microsoft recommends that you set an expiration value of less than 12 months"; the secret's **Value** "is
  *never displayed again* after you leave this page".
- **Certificate stores.** The authentication commands page says `-CertificateThumbprint` and `-CertificateName`
  load the certificate "from either `Cert:\CurrentUser\My\` or `Cert:\LocalMachine\My\`". The `Connect-MgGraph`
  reference page says instead that the certificate "will be retrieved from the current user's certificate store".
- **Limiting mailbox access.** The RBAC for Applications page says the feature "replaces Application Access
  Policies", and that "The permissions assigned using Application RBAC act in addition to grants you make in
  Microsoft Entra ID", so an unscoped `Mail.Send` grant in Microsoft Entra has to be removed before a scope limits
  anything. Assigning the roles needs the Organization Management role group, or the Exchange Administrator role
  in Microsoft Entra ID. Changes "are subject to cache maintenance that varies between 30 minutes and 2 hours".
  `New-ServicePrincipal` takes the IDs from the Enterprise applications page, not from App registrations.
- **Encrypted standard strings.** The `ConvertFrom-SecureString` page: without a key, "the Windows Data
  Protection API (DPAPI) is used", and "The contents of a SecureString aren't encrypted on non-Windows systems".
- **Graph sub-modules.** The installation page: `Microsoft.Graph.Authentication` "is installed by default when you
  opt to install the sub modules individually"; installing in one version of PowerShell "doesn't install it for
  the other"; Windows PowerShell needs .NET Framework 4.7.2 or later.
- **Blocked downloads.** `about_Signing`: under the RemoteSigned policy, a downloaded unsigned script fails with
  `The file <file-name> cannot be loaded. The file <file-name> is not digitally signed.`, and `Unblock-File` lets
  it run.

**From SDK source:** at tag `2.25.0` and on `main`, the certificate lookup by thumbprint and by subject name
searches `StoreLocation.CurrentUser` and then `StoreLocation.LocalMachine`. That agrees with the authentication
commands page, not the reference page. At `2.25.0`, WAM is used only when `EnableWAMForMSGraph` is set and the
platform is Windows; on `main`, when the authentication context's `WamEnabled` is true on Windows.

**From source:** `Connect-ScriptMessage_MicrosoftGraph` calls `Connect-MgGraph` with `-ClientId` and `-TenantId`
for delegated sign-in, so the app registration's redirect URIs apply to it; with `-CertificateThumbprint`,
`-CertificateName`, or `-Certificate` for the three certificate options; and with a credential built from
`MgApp_EncryptedSecret` for `ClientSecret`. It decrypts `MgApp_EncryptedSecret` and
`MgApp_EncryptedCertificatePassword` with `ConvertTo-SecureString`, so both have to be encrypted by the account
that runs the script, on the computer that runs it.

### What this does not establish

- **Unverified:** which SDK release made WAM the default on Windows, and whether a delegated sign-in without the
  broker redirect URI fails or falls back to the browser.
- **Unverified:** that `openid`, `profile`, `email`, and `offline_access` need no admin consent. The permissions
  reference did not return them. The scopes page describes them as standard OpenID Connect scopes shown on the
  user consent page, so treating them as user-consentable is **Inference**.
- **Unverified:** that an `Application Mail.Send` assignment in RBAC for Applications works with the
  `Send-MgUserMail` call the module makes. The page lists Microsoft Graph as a supported protocol for that role;
  no send was tested.
- **Unverified:** the message `ConvertTo-SecureString` gives for a string encrypted by another account or on
  another computer. The wiki's troubleshooting entry uses "Key not valid for use in specified state.", the
  Windows message usually reported for it; it was not reproduced here.

## 9. Measured: finding a group chat to reuse

**Environment:** measured on 2026-09-30 in PowerShell 7.6.6 on Windows 11, with `Microsoft.Graph.Authentication`
and `Microsoft.Graph.Teams` 2.25.0. The sign-in was delegated, through an app registration with `Chat.Read`
consented, and `Send-ScriptMessage` came from `develop` at `bcc28ea`. Every chat and message involved only three
accounts the tester controls: the signed-in sender and two recipients.

**Measured:**

- `Get-MgChat -All -Filter "ChatType eq 'group'"` with `-ExpandProperty 'Members', 'LastMessagePreview'`, the
  call the module makes, succeeded: Graph accepts both expansions in one request. Each of the signed-in account's
  6 group chats came back with its members' `email` values and a `LastMessagePreview`, and no two of them had the
  same members.
- **A group chat with the same members as an existing one is a new chat.** `New-MgChat -BodyParameter`, with
  `chatType` `group` and the same three members as an existing group chat, returned a chat with a new ID and a
  `CreatedDateTime` of that moment.
- **Posting a message does not change `LastUpdatedDateTime`.** After three messages were posted into the existing
  chat, its `LastUpdatedDateTime` still equalled its `CreatedDateTime`, more than a year earlier, while
  `LastMessagePreview.CreatedDateTime` moved to each new message.
- With one matching chat, `Send-ScriptMessage -ChatType Group` posted into it. With two, where the newer chat held
  the latest message, it posted into the older one, whose last message was older too.
- **A delegated sign-in gets every delegated permission already granted to the app, not only the ones asked
  for.** With `MgDelegatedPermission_RequestChatReadPermission` false, `AllowableMessageTypes` only `Chat`, and
  `MgDelegatedPermission_RequestFilesReadWritePermission` false, `Connect-ScriptMessage` asked for
  `Chat.ReadBasic` but no `Chat.Read`, `Mail.Send`, or `Files.ReadWrite`. The connection's scopes, as
  `Get-ScriptMessageContext` reported them, included all four, because the app registration had all of them
  granted. The group chat lookup therefore took the `Chat.Read` sort.
- **`Chat.ReadBasic` cannot read `lastMessagePreview`.** With `Chat.Read`'s grant revoked, so the connection's
  chat scopes were only `Chat.Create`, `Chat.ReadBasic`, and `ChatMessage.Send`, the same `Get-MgChat` call with
  `-ExpandProperty 'Members', 'LastMessagePreview'` failed with 403 "Expansion on lastMessagePreview requires one
  of the 'Chat.Read,Chat.ReadWrite' additional permissions". Expanding only `Members` succeeded.
- On that connection, with the same two matching chats, `Send-ScriptMessage -ChatType Group` from this change's
  code posted into the newer chat, by `LastUpdatedDateTime`, although the older one held the latest message.
- Before `Chat.ReadBasic` was granted to the app, that sign-in stopped at an "Approval required" prompt listing it,
  in a tenant whose user consent settings require an administrator. The earlier sign-ins, which asked for
  `Chat.Read`, already granted, had no prompt.
- `New-MgChat` failed with 400 `'user@odata.bind' field is missing in the request.` when a member's
  `user@odata.bind` value was present but named a user that does not exist (a quoted list of addresses passed as
  one address). The message says the field is missing when it is there but wrong. Both the `-Members` form and
  the `-BodyParameter` form gave it.

**From documentation:** the chat resource type page defines `lastUpdatedDateTime` as "Date and time at which the
chat was renamed or the list of members was last changed", and says `lastMessagePreview` is "Null if no messages
were sent in the chat" and that only the list chats operation supports it. The List chats page lists
`Chat.ReadBasic` as the least privileged delegated permission and supports `$orderby` only on
`lastMessagePreview/createdDateTime`, descending. It also says `$expand=members` "returns a maximum of 25 member
items". Neither page says which permission is needed to expand `lastMessagePreview`.

**From source:** the `Group` branch of `Send-ScriptMessage_MicrosoftGraph` lists every group chat and sorts it
most recent first: by `LastMessagePreview.CreatedDateTime` when the connection's scopes include `Chat.Read` or
`Chat.ReadWrite`, and otherwise by `LastUpdatedDateTime`. It posts into the first chat whose member `email`
values are exactly the sender plus the recipients, and creates a new group chat only when none matches.

**Inference:**

- Without `Chat.Read`, "most recent" means the chat most recently created, renamed, or given different members;
  messages never change which chat is chosen. That sort is used only when the app registration has not been
  granted `Chat.Read` (or `Chat.ReadWrite`), whatever `MgDelegatedPermission_RequestChatReadPermission` says;
  the same holds for every setting that only chooses which delegated permission is asked for.
- A group chat with more than 25 members never matches, if the 25-member cap applies to `-All` paging too, so the
  module would create a new chat on every send to it.
- A recipient given by a UPN that differs from their primary email address never matches either, with the same
  result: the module names each member as `users('<address>')`, which going by Graph's
  `/users/{id | userPrincipalName}` form takes a user ID or UPN, but matches on the member's `email`, the primary
  address. Matching on each member's `userId` would avoid it. A true alias, neither the UPN nor the primary
  address, probably fails when the chat is created instead.

### What this does not establish

- **Unverified:** the module's own `New-MgChat -Members` call creating a group chat with valid members. A matching
  chat already existed, so that path was not exercised.
- **Unverified:** whether the Teams client ever creates a second group chat with the same members, and whether
  anything besides a rename or a member change moves `LastUpdatedDateTime` (only the documentation says so).
- **Unverified:** the 25-member cap, the UPN mismatch, and the alias failure in the inferences above. None was
  tested.
- **Unverified:** whether permissions a user consented to for themselves, rather than ones an administrator granted
  for the organization, are carried the same way. Only administrator grants were measured.
- Nothing in this section was measured on Windows PowerShell 5.1 or with an SDK release other than 2.25.0.

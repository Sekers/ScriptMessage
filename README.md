# ScriptMessage PowerShell Module <!-- omit in toc -->

## Table of Contents  <!-- omit in toc -->

- [Versioning](#versioning)
- [Overview](#overview)
- [Supported Services](#supported-services)
- [What's New](#whats-new)
- [Documentation](#documentation)
- [Developing and Contributing](#developing-and-contributing)

---
## Versioning
This module can be used in production environments. Releases after 1.1.0 follow [Semantic Versioning 2.0.0](https://semver.org/spec/v2.0.0.html): breaking changes increase the major version, new backward-compatible functionality increases the minor version, and backward-compatible fixes increase the patch version. Check [CHANGELOG.md](./CHANGELOG.md) before updating to a new major version, and see [RELEASING.md](./RELEASING.md) for the full versioning and release policy.

## Overview

PowerShell Module to Connect Scripts to Messaging Services (Email, Chat, etc.).

ScriptMessage is designed to simplify the use of messaging services in PowerShell scripts. For example, you can:

- Take advantage advantage of the Microsoft Graph SDK PowerShell module (Graph API) and other messaging modules or APIs without having to learn all the object formatting that they require to send email and chat messages.
- Specify more than one service in a single command to easily send the same message multiple ways for redundancy or other purposes. For example, you might want to send an email using Microsoft Graph and a chat message using both Teams & Slack for the same alert.
- Easily switch the desired messaging service(s) in your scripts by updating simple config files. Whether your current messaging service is deprecated, you need to add a new service, or switch to a different service, you will no longer need to rewrite all of your scripts.

## Supported Services

- [**Microsoft Graph SDK PowerShell:**](https://learn.microsoft.com/en-us/powershell/microsoftgraph/overview?view=graph-powershell-1.0) Take advantage of the Microsoft Graph SDK PowerShell module to send email and chat messages using the Graph API without having to learn all the object formatting that the API requires (which unfortunately the SDK doesn't simplify).
  - Since the Microsoft Graph API only supports Teams Chat when using delegated [permissions](https://learn.microsoft.com/en-us/graph/permissions-overview), we are looking into [Teams Bots](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/overview) support for future releases to allow for application permissions.
  - Currently, the ScriptMessage module only supports one-on-one and group chats. Teams Channel chats will be enabled in a future release.
  - Specific Modules Used (if you want to minimize footprint):
    - Microsoft.Graph.Authentication (Mail & Chat)
    - Microsoft.Graph.Users.Actions (Mail Only)
    - Microsoft.Graph.Teams (Chat Only)
    - Microsoft.Graph.Files (Mail & Chat - Only if OneDrive Uploads are Needed)
- [**Mailozaurr:**](https://github.com/EvotecIT/MailoZaurr) Support is planned in future releases. Mailozaurr is a PowerShell module that aims to provide SMTP, POP3, IMAP and few other ways to interact with Email. Underneath it uses MimeKit and MailKit and EmailValidation libraries.
- [**Slack:**](https://api.slack.com/) Support is planned in future releases. Send messages into channels ([including ephemeral messages](https://api.slack.com/surfaces/messages#ephemeral)) or directly to users.
- [**PSGSuite:**](https://github.com/SCRT-HQ/PSGSuite) Support is planned in future releases. Send Google Workspace Gmail & Chat messages. PSGSuite is a PowerShell module wrapping Google's .NET SDKs.

---

## What's New

See [CHANGELOG.md](./CHANGELOG.md) for information on the latest updates, as well as past releases.

---

## Documentation

The ScriptMessage module documentation is hosted in the [ScriptMessage Wiki](https://github.com/Sekers/ScriptMessage/wiki). Examples are included in the comment-based help for each function/cmdlet (e.g., Get-Help Connect-ScriptMessage -Examples).

---

## Developing and Contributing

This project is developed using a [simplified Gitflow workflow](https://www.grimadmin.com/article.php/simple-modified-gitflow-workflow) that cuts out the release branches, which are unnecessary when maintaining only a single version for production. The Main branch is always the latest stable version released, and it is tagged with the new version number each time the Develop branch is merged into it. Shared branches are never rebased or force-pushed. See [CONTRIBUTING.md](./CONTRIBUTING.md) for how to contribute, and [RELEASING.md](./RELEASING.md) for branching, versioning, and releases.

You are welcome to [fork](https://guides.github.com/activities/forking/) the project and then offer your changes back using a [pull request](https://guides.github.com/activities/forking/#making-a-pull-request).
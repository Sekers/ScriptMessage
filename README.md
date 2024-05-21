# ScriptMessage PowerShell Module <!-- omit in toc -->

## Table of Contents  <!-- omit in toc -->

- [Overview](#overview)
- [What's New](#whats-new)
- [Current API Support](#current-api-support)
- [Documentation](#documentation)
- [Developing and Contributing](#developing-and-contributing)

---

## Overview

PowerShell Module to Connect Scripts to Messaging Services (Email, Chat, etc.).

ScriptMessage is designed to simplify the use of messaging services in PowerShell scripts. For example, you can:
- Take advantage of the Microsoft Graph SDK PowerShell module to send email messages using the Graph API without having to learn all the object formatting that the API requires.
- Easily switch messaging services in your scripts by changing the service used without having to rewrite your scripts to use the new service.
- Specify more than one service in a single command to send the same message multiple ways for redundancy or other purposes.

Note: The module currently only supports the Microsoft Graph SDK PowerShell module, but other messaging services will be added in.

---

## What's New

See [CHANGELOG.md](./CHANGELOG.md) for information on the latest updates, as well as past releases.

---

## Documentation

The ScriptMessage module documentation is hosted in the [ScriptMessage Wiki](https://github.com/Sekers/ScriptMessage/wiki). Examples are included in the [Sample Usage Scripts folder](./Sample_Usage_Scripts) as well as in the comment-based help for each function/cmdlet (e.g., Get-Help Connect-ScriptMessage).

---

## Developing and Contributing

This project is developed using a [simplified Gitflow workflow](https://www.grimadmin.com/article.php/simple-modified-gitflow-workflow) that cuts out the release branches, which are unnecessary when maintaining only a single version for production. The Master/Main branch will always be the latest stable version released and tagged with an updated version number anytime the Develop branch is merged into it. [Rebasing](https://www.atlassian.com/git/tutorials/merging-vs-rebasing) will occur if we need to streamline complex history.

You are welcome to [fork](https://guides.github.com/activities/forking/) the project and then offer your changes back using a [pull request](https://guides.github.com/activities/forking/#making-a-pull-request).
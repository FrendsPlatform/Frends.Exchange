# Frends.Exchange.ReadMail

![Frends.Exchange.Send Main](https://github.com/FrendsPlatform/Frends.Exchange/actions/workflows/Send_main.yml/badge.svg)
![MyGet](https://img.shields.io/myget/frends-tasks/v/Frends.Exchange?label=NuGet)
![GitHub](https://img.shields.io/github/license/FrendsPlatform/Frends.Exchange?label=License)
![Coverage](https://app-github-custom-badges.azurewebsites.net/Badge?key=FrendsPlatform/Frends.Exchange/Frends.Exchange.ReadMail|main)

Frends task for reading emails via Microsoft Exchange. 

Targets .NET 6.0.

## Installing

You can install the Task via Frends UI Task View, or you can find the NuGet package from the following NuGet feed: 
https://www.myget.org/F/frends-tasks/api/v2

## Specifications

### Settings

| Property              | Type                                              | Description                                                  |
| --------------------- | ------------------------------------------------- | ------------------------------------------------------------ |
| ExchangeServerVersion | [ExchangeServerVersion](###ExchangeServerVersion) | Specifies the expected Exchange Server version.              |
| UseAutoDiscover       | bool                                              | If set to true, the Task will attempt to auto-discover the server address from username. |
| ServerAddress         | string                                            | Hostname or IP address of the Exchange Server                |
| UseAgentAccount       | bool                                              | If set to true, the Task will attempt to log in with the Agent account. |
| Username              | string                                            | Email account to use for logging in.                         |
| Password              | string                                            | Password to use for logging in.                              |
| Mailbox               | string                                            | Box to read emails from. If left empty, the Task will read from default Inbox mailbox. |

### Options

| Property                     | Type   | Description                                                  |
| ---------------------------- | ------ | ------------------------------------------------------------ |
| MaxEmails                    | int    | Maximum number of emails to retrieve.                        |
| GetOnlyUnreadEmails          | bool   | If set to true, will retrieve only unread emails.            |
| MarkEmailsAsRead             | bool   | If set to true, will mark retrieved emails as read.          |
| DeleteReadEmails             | bool   | If set to true, the read emails will be deleted after retrieving. |
| EmailSenderFilter            | string | If a value is set, emails only from the certain sender are retrieved. |
| EmailSubjectFilter           | string | If a value is set, emails only with the certain subject are retrieved. |
| ThrowErrorIfNoMessagesFound  | bool   | If set to true, will throw an Exception if no messages are found |
| IgnoreAttachments            | bool   | If set to true, the task will retrieve emails without their attachments. |
| GetOnlyEmailsWithAttachments | bool   | If set to true, the task will retrieve only the mails with attachments. |
| AttachmentSaveDirectory      | string | Path of the directory where to save the attachments in.      |
| OverwriteAttachment          | bool   | If set to true, the task will overwrite an attachment file if there already exists one with the same name. |

### Other specifications

#### ExchangeServerVersion

An enumeration consisting of a set of named Exchange Server versions. If you're setting up Exchange settings for CI/CD tests, you'll need the numeric representations seen below.

| Name             | Numeric representation |
| ---------------- | ---------------------- |
| Exchange2007_SP1 | 0                      |
| Exchange2010     | 1                      |
| Exchange2010_SP1 | 2                      |
| Exchange2010_SP2 | 3                      |
| Exchange2013     | 4                      |
| Exchange2013_SP1 | 5                      |
| Office365        | 6                      |

## Testing

Tests (or actually one particular test) in the project expects you to provide [Exchange settings](###Settings) via environment variable `EXCHANGE_SETTINGS_FOR_TESTING` as base64-encoded JSON. This can be set thru operating system, or if the tests are run via Github Actions, thru repository secrets.

Below is an example of settings for testing against an Office365 server.

```json
{
    "ExchangeServerVersion": 6,
    "UseAutoDiscover": false,
    "ServerAddress": "https://example.server.com/exchange.asmx",
    "UseAgentAccount": false,
    "Username": "example.user@example.server.com",
    "Password": "example-password-123"
  }
```


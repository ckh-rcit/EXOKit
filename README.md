# EXOKit

EXOKit is a WinUI 3 (.NET 8) replacement for the PowerShell-based **Exchange Admin Toolkit**
WinForms GUI. It hosts a real Exchange Online PowerShell session and Microsoft Graph client
in-process to provide the same mailbox/group/Bookings administration workflows through a modern
Windows app shell.

## Scope

Ported with feature parity from the original toolkit:

- **Mailboxes** — Add/remove Full Access, Send As, and Send on Behalf permissions on shared
  mailboxes, with existence checks, batched `Set-Mailbox` calls for Send on Behalf, and a
  1-second/10-attempt validation poll before reporting final status.
- **Resources** — Same permission workflow as Mailboxes, scoped to Room/Equipment resource
  mailboxes.
- **Groups** — Add/remove Member and Owner roles on both Distribution Groups (via EXO cmdlets)
  and Microsoft 365 (Unified) Groups (via Microsoft Graph), including group-type resolution,
  connection-requirement checks, and the same validation poll pattern.
- **Bookings** — Adds users to the configured Bookings license group and sets their OWA mailbox
  policy to enable Microsoft Bookings, with already-exists/already-set short-circuits.
- **Recipient Lookup** — Resolves a recipient identity and reports a friendly recipient type
  (User/Shared/Room/Equipment Mailbox, Distribution Group, Microsoft 365 Group, or a raw
  `RecipientType (RecipientTypeDetails)` fallback).
- **Ticket notes / logging** — Central `Logger` service mirrors the script's
  `Write-OutputLog` / `Write-SummaryLog` behavior, including a "Copy Last Ticket Notes" action
  that extracts the most recent `--- For IT Ticket ---` section from the log.
- **ServiceNow integration** (optional) — Retrieves service-account credentials from Azure Key
  Vault and closes/updates a ServiceNow ticket by number, matching `Close-ServiceNowTask`.

### Explicitly excluded

The following sections from the original toolkit are **not** ported to EXOKit, per project scope:

- **User Provisioning** — new-user mailbox/license provisioning workflows.
- **Calendar Tools** — calendar permission management.

## Architecture

| Service | Responsibility |
|---|---|
| `Services/ExoPowerShellService.cs` | Hosts a dedicated STA runspace running `ExchangeOnlineManagement` cmdlets in-process. |
| `Services/AuthService.cs` | Interactive MSAL sign-in for Microsoft Graph, using scopes from `config.json`. |
| `Services/GraphService.cs` / `TokenAuthenticationProvider.cs` | Microsoft Graph SDK client and M365 group/user operations. |
| `Services/MailboxPermissionService.cs` | Mailbox/Resource permission add/remove orchestration. |
| `Services/GroupMembershipService.cs` | Distribution Group / M365 Group role add/remove orchestration. |
| `Services/BookingsService.cs` | Bookings license group + OWA policy enablement. |
| `Services/RecipientLookupService.cs` | Recipient type lookup for the Recipient Lookup tab. |
| `Services/ServiceNowService.cs` | Optional ServiceNow ticket closure via Key Vault-backed credentials. |
| `Services/ConfigService.cs` | Loads and validates `config.json`. |
| `Services/Logger.cs` | Central log/ticket-note tracking consumed by the UI. |

The app enforces the same connection order as the original script: Exchange Online must be
connected before Microsoft Graph can be connected.

## Configuration

`config.json` (copied to the output directory) must define:

```json
{
  "Settings": {
	"LicenseGroups": { "Bookings": { "GroupName": "<Bookings license group display name>" } },
	"OwaPolicies": { "BookingsCreators": "<OWA mailbox policy name>" },
	"GraphApi": { "Scopes": [ "User.Read.All", "Group.ReadWrite.All", "..." ] },
	"ServiceNow": { "Enabled": false }
  }
}
```

`ServiceNow` is optional; set `Enabled: false` (or omit the section) if ticket auto-close is not
needed.

# EXOKit

EXOKit is a WinUI 3 (.NET 10) replacement for the PowerShell-based **Exchange Admin Toolkit**
WinForms GUI. It hosts a real Exchange Online PowerShell session and Microsoft Graph client
in-process to provide the same mailbox/group/Bookings administration workflows through a modern
Windows app shell.

## Scope

Ported with feature parity from the original toolkit:

- **Mailboxes** — Add/remove Full Access, Send As, and Send on Behalf permissions on shared
  mailboxes, with existence checks, batched `Set-Mailbox` calls for Send on Behalf, and a
  2-second/15-attempt validation poll before reporting final status. Permission reads and writes
  retry the known EXO server-side null-reference failure; unreadable state remains unconfirmed.
- **Resources** — Same permission workflow as Mailboxes, scoped to Room/Equipment resource
  mailboxes.
- **Groups** — Add/remove Member and Owner roles on both Distribution Groups (via EXO cmdlets)
  and Microsoft 365 (Unified) Groups (via Microsoft Graph), including group-type resolution,
  connection-requirement checks, and the same validation poll pattern.
- **Bookings** — Adds users to the configured Bookings license group and sets their OWA mailbox
  policy to enable Microsoft Bookings, with already-exists/already-set short-circuits.
- **Recipient Lookup** — Resolves multiple recipient identities and reports a friendly recipient type
  (User/Shared/Room/Equipment Mailbox, Distribution Group, Microsoft 365 Group, or a raw
  `RecipientType (RecipientTypeDetails)` fallback). Paste one identity per line or import a TXT/CSV
  file. Names containing spaces are preserved and duplicate identities are removed. Results appear
  per recipient, can be copied, and remain available after cancellation. A failed lookup does not
  stop the rest of the batch; cancellation stops before the next lookup.
- **User Search** - Searches Entra users by name, employee ID, UPN, email, or SAM account,
  with optional deleted-user results and read-only EXO mailbox checks. Includes the user
  details and clipboard support from Entra Scout within EXOKit's existing connections.
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

### Recipient lookup files

TXT files contain one identity per line. CSV files can contain a single headerless identity column,
or a header named `Identity`, `Recipient`, `Email`, `EmailAddress`, `PrimarySmtpAddress`,
`UserPrincipalName`, or `UPN` (case-insensitive). For CSVs with multiple columns, the first recognized
identity column is used and other columns are ignored. Comma and semicolon delimiters are supported;
quote names containing either delimiter. Imports append to the existing input without duplicating it.

### User search

Connect EXO and Graph, then open User Search. Enter a name, `Last, First`, employee ID,
UPN/email, or SAM account and press Enter or Search. Numeric employee IDs first try an
`E` prefix, then the raw number if no match is found. Name searches use display-name prefixes;
UPN/email searches use exact matches. All Graph result pages are read.

Include deleted users adds soft-deleted directory users. Both live and deleted-user queries
use the existing delegated `User.Read.All` permission with administrator consent; no new
Graph scope is added. A denied or incomplete Graph read fails the search and clears its results.

Check EXO mailboxes queries active mailboxes by directory object ID and verifies the returned
identity. Turn it off for directory-only searches. Lookup failures remain Unknown with an
explanation in user details; Not found means no active mailbox was returned within the
connected administrator's Exchange scope. Deleted directory users are not checked for
mailboxes, and their status makes no claim about retained or inactive mailboxes.

Select a result and use View details or Copy details, or open details by double-clicking or
pressing Enter. Details include email, employee/SAM IDs, job title, department, account status,
mailbox status, object ID, creation time (UTC), and business phone numbers.
Cancel retains completed directory results and marks unchecked mailboxes. Editing the query
or search options, disconnecting, or saving connection settings clears old results.
Search and mailbox checks make no directory changes. Entra Scout remains a separate, unchanged app.

### Reporting

Choose Group membership and owners or Mailbox permissions, enter the target identity, then
select Generate report or press Enter. Target validation runs automatically and rejects an
incompatible recipient type before reading the report. The results area shows the resolved
target, entry count, and per-entry status. Export CSV saves the current result with its report
type in the filename; partial reports retain their warnings.

Changing the report type, target, or connected EXO tenant clears the previous result and disables
export. Available report types and their target requirements are defined in
`ReportingService.ReportTypes`; both types share the results and export workflow.

The Log / Ticket Notes area is hidden on Settings, User Search, and Reporting. Logs continue
to be collected. Operational pages restore its previous height and collapsed state; the divider
supports dragging and Up/Down keys. Clickable controls use a hand cursor, and the divider uses
the native vertical-resize cursor.

## Configuration

The package includes only the sanitized sample configuration. Runtime settings are stored at
`%LOCALAPPDATA%\EXOKit\config.json`. Saves atomically replace the file and retain the previous
version as `config.json.bak`. Invalid configuration opens Settings without overwriting the original.

Settings groups fields under Connections, Bookings, ServiceNow Integration, and Updates. The Updates
section contains the installed version, update feed URL, and update check. Save Settings applies the
edited configuration; the update check uses the saved feed URL.

Configure these values in Settings:

```json
{
  "Settings": {
	"LicenseGroups": { "Bookings": { "GroupName": "<Bookings license group display name>" } },
	"OwaPolicies": { "BookingsCreators": "<OWA mailbox policy name>" },
  "GraphApi": { "ClientId": "<EXOKit app registration ID>", "TenantId": "<tenant ID>", "Scopes": [ "User.Read.All", "Group.ReadWrite.All" ] },
  "UpdateFeedUrl": "",
	"ServiceNow": { "Enabled": false }
  }
}
```

`ServiceNow` is optional. Set `Enabled: false` or omit it when ticket closure is not needed.

## Requirements

### Authentication and Diagnostics

Graph uses system-browser MSAL authentication with the configured app registration and its `http://localhost` desktop redirect. EXO keeps its interactive sign-in and retries once in the browser only for WAM/window-handle failures. The existing browser-sign-in setting also selects that compatibility path directly. No device-code sign-in or fallback is used. Graph and EXO must use the configured tenant; independent consent or MFA prompts are possible.

Before EXO sign-in, the embedded PowerShell 7.6.6 host checks stable PSGallery versions and installs missing or newer PackageManagement, PowerShellGet, and ExchangeOnlineManagement modules under CurrentUser. Install PowerShell 7.6+ if PSResourceGet 1.2+ is unavailable. An older loaded module requires restarting the app. Repository URL validation and normal PowerShell authorization remain enabled; publisher decisions require explicit user input. An offline update check can use suitable installed modules.

Connection failures show a dialog with the service error. Unexpected managed startup/UI/background failures are reported, with diagnostic logs under `%LOCALAPPDATA%\EXOKit\Logs`. Fatal errors are not marked handled to keep an unknown application state running. Applied changes are not automatically rolled back.

Local validation covers service regressions and an x64 application build. Live MFA, Conditional Access, tenant consent, Key Vault access, and admin workflows still require operator testing with the newly built application.

References: [EXO prerequisites](https://learn.microsoft.com/powershell/exchange/exchange-online-powershell-v2), [WAM compatibility](https://learn.microsoft.com/troubleshoot/exchange/administration/wam-integration-issues), [MSAL browser configuration](https://learn.microsoft.com/entra/msal/dotnet/acquiring-tokens/using-web-browsers).

- Windows desktop; the release workflow builds x64. Local validation covers x64 only.
- ExchangeOnlineManagement 3.10.1 or later installed where the embedded PowerShell can discover it.
  The app embeds PowerShell 7.6.6 on .NET 10 and Windows App SDK 2.4.0. Module security checks remain enabled.
- An EXOKit-owned, single-tenant public-client Entra app registration. Configure the Mobile and
  Desktop platform with the `http://localhost` redirect URI. Do not configure a client secret.
- Delegated Graph permissions `User.Read.All` and `Group.ReadWrite.All`, with administrator consent.
  The signed-in administrator also needs the appropriate directory and Exchange roles. Additional
  tenant-specific policies, hidden memberships, and role-assignable groups may need extra rights;
  denied reads stop the operation rather than being treated as absent membership.

Connect EXO first. Graph authentication is bound to the configured tenant, which must match the
tenant returned by EXO. Saving configuration disconnects Graph and rebuilds its dependent services;
reconnect before using them. Tokens are sent only to the public-cloud Graph endpoint.

EXO uses standard interactive sign-in with Windows Web Account Manager (WAM) enabled by default,
without requesting device-code authentication. Graph retains its separate MSAL client and token
cache; EXO and Graph access tokens are not interchangeable. Run EXOKit in the signed-in Windows
user's context, not with RunAs under a different Windows account.

If WAM causes a sign-in error, select Browser sign-in for Exchange (WAM compatibility) in Settings
and reconnect. This explicitly uses `-DisableWAM`, which Microsoft recommends only as a temporary
workaround. The choice applies to the next connection and is persisted by Save Settings. It does
not bypass MFA or tenant policy and does not fix shared-process assembly conflicts. Verify EXO
access again after connecting Graph when testing authentication changes.

PowerShell publisher-trust prompts appear in a confirmation dialog with no choice preselected.
Review the publisher and file before selecting a response. `Run once` approves that execution;
`Always run` persistently trusts the publisher through PowerShell. Cancel stops the prompt without
choosing for you. EXOKit does not bypass execution policy or automatically install publisher
certificates. These prompts can occur during module loading, before Microsoft sign-in starts.

The embedded PowerShell host explicitly requests process-scoped `RemoteSigned`, following
[Microsoft's C# EXO example](https://learn.microsoft.com/powershell/exchange/connect-to-exo-powershell-c-sharp).
It does not inherit the separate `pwsh.exe` installation's local configuration reliably. This
setting does not persist to CurrentUser or LocalMachine, but takes precedence over those scopes
inside EXOKit. MachinePolicy and UserPolicy still override it. Downloaded unsigned scripts remain
blocked, and publisher prompts still require a user choice; the authorization manager stays enabled.

For ServiceNow, the same registration is used for interactive Key Vault authentication. Configure
Azure Key Vault delegated `user_impersonation` access and give the administrator secret-read access
to the required vault secrets, such as the Key Vault Secrets User role at the appropriate scope.
The integration accepts HTTPS origin URLs only and currently supports public-cloud
`*.vault.azure.net` vaults. The subscription ID is retained for configuration compatibility; it is
not used for data-plane secret access.

Set Bookings Group Object ID when possible. It takes precedence over the display name. A name
matching multiple groups is rejected. Group membership does not itself prove license assignment
has completed; check licensing separately.

Bookings checks the organization's `BookingsEnabled` setting and the selected OWA policy's
`BookingsMailboxCreationEnabled` setting before making changes. The confirmation identifies that
assigning the configured policy replaces the user's entire OWA policy, not just Bookings settings.
Use a compatible policy and verify effective licensing separately. These reads require access to
`Get-OrganizationConfig` and `Get-OwaMailboxPolicy`.

## Operation Safety

Only one service operation runs at a time. Disconnect and Settings are disabled during work, and
window closure is blocked. Cancel requests an asynchronous stop of the active Exchange command
and prevents remaining calls. The button shows "Cancelling..." until cleanup finishes. Each embedded
Exchange command has a five-minute timeout. Cancellation and timeouts do not roll back completed
mutations; reload and verify unconfirmed results before retrying or closing a ticket.

Permission and role removals capture a durable, tenant-bound before-state snapshot under
`%LOCALAPPDATA%\EXOKit\Snapshots` before mutation. A failed snapshot write blocks that removal.
Absent permissions are left unchanged rather than generating speculative recovery rights. Restore
supports mailbox permissions, group membership/ownership, and group delegates. It adds missing
rights without replacing unrelated delegates, verifies each result, and rejects unknown roles,
empty records, and cross-tenant records. Legacy snapshots without a tenant ID require manual review
and restoration. Snapshot files contain identities and access information; protect and retain them
according to your organization's policy.

Group Settings must be loaded for the exact identity before saving. Recipient lists are normalized
to SMTP addresses, stale loaded state is rejected, and saved values are read back. EXO does not offer
an atomic compare-and-set for these settings; avoid concurrent administration of the same group.
Mail-enabled security group joining/leaving is owner-managed (Closed).
The Group Settings banner confirms when a save has been read back and verified. Save failures show
the service error in a dialog. After failure or cancellation, saving is disabled until settings are
reloaded; some changes may already have reached Exchange.

Starting in v1.0.33, PowerShell collections are read without assuming a generic collection type.
Missing or malformed settings cannot silently become empty lists. Group Settings comparisons
ignore list ordering and preserve an independent copy of the loaded values. DG changes use
`BypassSecurityGroupManagerCheck`, which still requires Microsoft's documented administrator RBAC.

M365 recipient lookup explicitly requests `GroupMailbox` when the ordinary recipient lookup
finds no match. Graph user input accepts UPN, primary email, SMTP proxy address, or object ID and
rejects ambiguous matches. Send As checks read the full ACL and compare canonical identities and
available current/historical SIDs. An unresolved trustee blocks a claim of absence; resolve the
ACL/SID history manually before retrying. Reports retain unresolved SID grants instead of hiding
them. Unresolved group owners and Send-on-Behalf delegates appear with their original identities
and an "Unresolved" status. An unreadable mailbox permission section gets an "Unavailable" row;
readable sections remain in the report. These report-only fallbacks do not relax mutation checks.

Mailbox permission reports resolve the target's primary SMTP address before querying and reject
returned identities outside that recipient's known identifiers. Changing the report type or target,
or starting another report, clears the previous results and disables export. The report banner
distinguishes complete, empty, partial, failed, and cancelled results. Partial CSV exports have
"Partial" in the suggested filename and include the same Status details as the table.

CSV imports in every section use the recognized identity-column rules described above. Arbitrary
multi-column files without a recognized identity header are rejected. Exports quote carriage
returns and line feeds and retain formula-injection protection.

Creation results distinguish partial completion from success. A failed follow-up does not mean the
new mailbox or group was removed. After a Teams replication failure, wait at least 15 minutes after
group creation and use Resume Team Provisioning with the existing group's email or object ID.
Do not recreate the group. Shared mailbox archive use still requires appropriate licensing.

Dynamic distribution-group reports use `Get-DynamicDistributionGroupMember`, which returns the
calculated membership list. This differs from previewing the filter and can lag directory changes
by the service's refresh interval.

Rendering or copying ticket notes never closes a ticket. Fully verified mailbox/group role batches
offer ServiceNow closure only after the administrator reviews the notes and confirms. Other
workflows retain copyable notes for manual ticket handling. Partial, cancelled, and unconfirmed
results do not qualify for closure.

## Version and Updates

The title bar and Settings show the installed MSIX version, currently `v1.0.36`; unpackaged builds
use the assembly version. Release tags must use `vMajor.Minor.Build`. The workflow stamps that
version into the manifest and assembly while retaining the existing package name and publisher.
Use a version higher than the installed version for an update.

### Publisher migration in v1.0.31

Starting with v1.0.31, the publisher is `CN=CKH-RCIT`. This changes the Windows package family;
it is not an in-place update of v1.0.30 or earlier. Back up settings and snapshots from the old
installation before removing it, including any package-redirected application data. Install the
new release's public certificate using its trust bundle after verifying the certificate fingerprint,
then install the new package and confirm configuration and snapshots before retiring the old app.
Do not assume Windows will migrate data or sign-in state between package families.

Existing releases and their original trust bundles are retained for legacy installations. Their
signed binaries were not rebuilt when Git history was sanitized; historical source tags now contain
the sanitized publisher text and are not exact identity matches for those archived binaries. Do not
rerun an old release tag using the new signing secrets. Future automatic updates apply within the
new package family after installation through its `.appinstaller` feed.

The stable update feed is hosted at
`https://ckh-rcit.github.io/EXOKit/EXOKit.appinstaller`. Download and open that file with Windows
App Installer once to enroll an existing v1.0.31 installation or install the current release.
Verify and trust the current release certificate first. In EXOKit Settings, set Update Feed URL to
that same address and save it so Check for Updates opens the feed. Existing saved settings are not
overwritten when a new package supplies a default feed URL.

The repository Actions variable `APPINSTALLER_FEED_URL` must equal the address above, and GitHub
Pages must use GitHub Actions as its publishing source. The Publish Update Feed workflow runs after
a successful Build and Release workflow, or manually through Run workflow. It checks out the default
branch, reads the latest stable release's published MSIX manifest, and generates the feed using that
package's identity and version. The publisher rejects package-family mismatches and feed downgrades.
Deployments are serialized. The generated `.appinstaller` and download page are uploaded to Pages;
installers and public certificate trust bundles stay in GitHub Releases. The download page at
`https://ckh-rcit.github.io/EXOKit/` links to the validated release assets and installation guidance.
Its version and download links refresh with the feed. No signing material or local configuration
is included in the site.

The release workflow also generates an `.appinstaller` asset when the variable is configured.
An existing release can be enrolled by running Publish Update Feed without rebuilding its signed
MSIX. Release signing secrets are not available to the Pages job. Private/authenticated release
downloads require a separate distribution solution.

Without that variable, releases remain direct MSIX downloads with no update enrollment. No hosting
or app registration is created by this repository. Existing direct-MSIX users must install once
through the configured `.appinstaller` file to enroll. Check for Updates opens the configured feed
in Windows; it does not silently install a package itself. App Installer checks on launch and with
its background task, without blocking activation or forcing an active operation to stop.

Signing uses the existing publisher identity and a timestamp server. Keep the signing key secure.
The trust script imports the public code-signing certificate into LocalMachine TrustedPeople only;
verify its fingerprint through a trusted channel and run the script yourself as administrator.
It does not add a certificate authority to Trusted Root. Certificates previously added to Root are
not removed automatically; assess those separately before changing machine trust.

## Validation

```powershell
dotnet test tests/EXOKit.Tests/EXOKit.Tests.csproj --configuration Release
dotnet build EXOKit.csproj -p:Platform=x64 -p:AppxPackageSigningEnabled=false
dotnet list EXOKit.csproj package --vulnerable --include-transitive
```

The isolated tests cover retry/verification, cancellation, snapshot persistence and restore policy,
configuration preservation, and ServiceNow request validation. Mock runspaces exercise the actual
EXO service and Group Settings/mailbox orchestration, including unchanged list round trips and
snapshot-before-removal failures. HTTP handlers exercise actual Graph SDK identity requests and
prove cancellation after a ticket lookup prevents its closure PATCH. They do not mutate a live tenant.
Reporting tests cover explicit SMTP scoping, unexpected mailbox identities, unresolved delegates,
partial permission sections, and CSV warning details. Cancellation tests stop running pipelines,
exercise command timeouts, and verify runspace reuse. WinUI layout and real Exchange response shapes
still require an interactive test before release.
NuGet advisories fail restore. Before production rollout, test EXO/Graph sign-in, one reversible
permission change and restore, ServiceNow closure, and signed feed enrollment in a test tenant.

## Microsoft References

- [Full Access report parameters](https://learn.microsoft.com/powershell/module/exchangepowershell/get-exomailboxpermission)
- [Send As report parameters](https://learn.microsoft.com/powershell/module/exchangepowershell/get-exorecipientpermission)
- [Asynchronous PowerShell stop](https://learn.microsoft.com/dotnet/api/system.management.automation.powershell.stopasync)
- [Interactive EXO sign-in](https://learn.microsoft.com/powershell/exchange/connect-to-exchange-online-powershell#connect-to-exchange-online-powershell-with-an-interactive-sign-in-prompt)
- [WAM compatibility guidance](https://learn.microsoft.com/troubleshoot/exchange/administration/wam-integration-issues)
- [EXO module compatibility](https://learn.microsoft.com/powershell/exchange/exchange-online-powershell-v2)
- [PowerShell support lifecycle](https://learn.microsoft.com/powershell/scripting/install/powershell-support-lifecycle)
- [Windows App SDK release channels](https://learn.microsoft.com/windows/apps/windows-app-sdk/release-channels)
- [Dynamic group calculated membership](https://learn.microsoft.com/exchange/recipients-in-exchange-online/manage-dynamic-distribution-groups/modern-dynamic-distribution-groups)
- [Create a Team from a group](https://learn.microsoft.com/graph/api/team-put-teams)
- [App Installer update settings](https://learn.microsoft.com/windows/msix/app-installer/update-settings)
- [Create an App Installer file](https://learn.microsoft.com/windows/msix/app-installer/how-to-create-appinstaller-file)

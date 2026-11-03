using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EXOKit.Services
{
    public enum ReportKind
    {
        GroupMembership,
        MailboxPermissions,
        AccountActivity
    }

    public sealed record ReportDefinition(ReportKind Kind, string Name, ReportTargetKind TargetKind, string TargetLabel, string Placeholder);

    public sealed record GeneratedReport(ReportTargetInfo Target, IReadOnlyList<ReportRow> Rows);

    /// <summary>
    /// A single row of a generated report (membership/ownership role or a mailbox permission entry).
    /// </summary>
    public class ReportRow
    {
        public string ObjectName { get; set; } = string.Empty;
        public string ObjectPrimarySmtpAddress { get; set; } = string.Empty;
        public string ObjectType { get; set; } = string.Empty;
        public string MemberOrDelegate { get; set; } = string.Empty;
        public string RoleOrPermission { get; set; } = string.Empty;
        public string Status { get; set; } = "Read";
        public bool HasWarning => Status != "Read";
    }

    /// <summary>
    /// Broad classification of a resolved report target identity, used by the UI to validate input
    /// and steer the user toward the correct report action (membership/ownership report for groups,
    /// mailbox delegate report for mailboxes) before running it.
    /// </summary>
    public enum ReportTargetKind
    {
        Unsupported,
        Group,
        Mailbox,
        Any
    }

    /// <summary>
    /// Result of resolving and classifying a report target identity, surfaced by the Validate action
    /// in the Reporting UI.
    /// </summary>
    public class ReportTargetInfo
    {
        public string Identity { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string PrimarySmtpAddress { get; set; } = string.Empty;
        public string RecipientTypeDetails { get; set; } = string.Empty;
        public string FriendlyType { get; set; } = string.Empty;
        public ReportTargetKind Kind { get; set; }
    }

    /// <summary>
    /// Ports the reporting logic from M365_Reporting_Tool.ps1 (per-object membership/permission report,
    /// exportable to CSV) and DelegateReporter.ps1 (quick mailbox delegate lookup), combined into a
    /// single Reporting section backed by the existing EXO PowerShell connection.
    /// </summary>
    public class ReportingService
    {
        private readonly ExoPowerShellService _exo;

        public static IReadOnlyList<ReportDefinition> ReportTypes { get; } = Array.AsReadOnly(new[]
        {
            new ReportDefinition(ReportKind.GroupMembership, "Group membership and owners", ReportTargetKind.Group, "Group", "Group email, alias, or object ID"),
            new ReportDefinition(ReportKind.MailboxPermissions, "Mailbox permissions", ReportTargetKind.Mailbox, "Mailbox", "Mailbox email, alias, or object ID"),
            new ReportDefinition(ReportKind.AccountActivity, "Account activity (is it in use?)", ReportTargetKind.Any, "Group or mailbox", "Shared mailbox, group, or distribution list email")
        });

        public ReportingService(ExoPowerShellService exo)
        {
            _exo = exo;
        }

        public async Task<GeneratedReport> GenerateAsync(ReportKind kind, string identity)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(identity);
            var definition = ReportTypes.FirstOrDefault(report => report.Kind == kind)
                ?? throw new ArgumentOutOfRangeException(nameof(kind));
            _exo.OperationCancellationToken.ThrowIfCancellationRequested();
            var target = await ClassifyReportTargetAsync(identity.Trim());
            var targetMatches = definition.TargetKind == ReportTargetKind.Any
                ? target.Kind != ReportTargetKind.Unsupported
                : target.Kind == definition.TargetKind;
            if (!targetMatches)
                throw new InvalidOperationException($"{definition.Name} requires a {definition.TargetLabel.ToLowerInvariant()}. '{target.Name}' is a {target.FriendlyType}.");
            var rows = kind switch
            {
                ReportKind.GroupMembership => await GenerateObjectReportAsync(target.Identity),
                ReportKind.MailboxPermissions => await GetMailboxDelegatesAsync(target.Identity),
                ReportKind.AccountActivity => await GenerateAccountActivityAsync(target),
                _ => throw new ArgumentOutOfRangeException(nameof(kind))
            };
            _exo.OperationCancellationToken.ThrowIfCancellationRequested();
            return new GeneratedReport(target, rows);
        }

        public static IEnumerable<string> CreateCsvLines(IEnumerable<ReportRow> rows)
        {
            yield return "ObjectName,ObjectPrimarySmtpAddress,ObjectType,MemberOrDelegate,RoleOrPermission,Status";
            foreach (var row in rows)
                yield return string.Join(",", new[] { row.ObjectName, row.ObjectPrimarySmtpAddress, row.ObjectType,
                    row.MemberOrDelegate, row.RoleOrPermission, row.Status }.Select(InputParsingHelpers.EscapeCsv));
        }

        /// <summary>
        /// Generates the full membership/ownership or permission report for a single Group/DL/DDG/
        /// Security Group/Mailbox, mirroring M365_Reporting_Tool.ps1's Get-Recipient dispatch logic.
        /// </summary>
        public async Task<List<ReportRow>> GenerateObjectReportAsync(string identity)
        {
            var recipient = await _exo.GetRecipientAsync(identity);
            if (recipient == null)
            {
                throw new InvalidOperationException($"Could not find a recipient matching '{identity}'.");
            }

            var objectName = recipient.Name ?? identity;
            var objectSmtp = recipient.PrimarySmtpAddress ?? identity;
            var recipientType = recipient.RecipientTypeDetails ?? string.Empty;
            var rows = new List<ReportRow>();

            void AddUnresolvedOwner(string owner, string error) => rows.Add(new ReportRow
            {
                ObjectName = objectName,
                ObjectPrimarySmtpAddress = objectSmtp,
                ObjectType = recipientType,
                MemberOrDelegate = owner,
                RoleOrPermission = "Owner",
                Status = $"Unresolved: {error}"
            });

            switch (recipientType)
            {
                case "GroupMailbox":
                    {
                        var links = await _exo.GetUnifiedGroupLinksAsync(recipient.Identity ?? identity);
                        rows.AddRange(links.Select(l => new ReportRow
                        {
                            ObjectName = objectName,
                            ObjectPrimarySmtpAddress = objectSmtp,
                            ObjectType = "Microsoft 365 Group",
                            MemberOrDelegate = $"{l.DisplayName} ({l.PrimarySmtpAddress})",
                            RoleOrPermission = l.Role
                        }));
                        break;
                    }
                case "MailUniversalDistributionGroup":
                    {
                        var links = await _exo.GetDistributionGroupReportLinksAsync(recipient.Identity ?? identity, isDynamic: false, onUnresolvedOwner: AddUnresolvedOwner);
                        rows.AddRange(links.Select(l => new ReportRow
                        {
                            ObjectName = objectName,
                            ObjectPrimarySmtpAddress = objectSmtp,
                            ObjectType = "Distribution List",
                            MemberOrDelegate = $"{l.DisplayName} ({l.PrimarySmtpAddress})",
                            RoleOrPermission = l.Role
                        }));
                        break;
                    }
                case "DynamicDistributionGroup":
                    {
                        var links = await _exo.GetDistributionGroupReportLinksAsync(recipient.Identity ?? identity, isDynamic: true, onUnresolvedOwner: AddUnresolvedOwner);
                        rows.AddRange(links.Select(l => new ReportRow
                        {
                            ObjectName = objectName,
                            ObjectPrimarySmtpAddress = objectSmtp,
                            ObjectType = "Dynamic Distribution Group",
                            MemberOrDelegate = $"{l.DisplayName} ({l.PrimarySmtpAddress})",
                            RoleOrPermission = l.Role
                        }));
                        break;
                    }
                case "MailUniversalSecurityGroup":
                    {
                        var links = await _exo.GetDistributionGroupReportLinksAsync(recipient.Identity ?? identity, isDynamic: false, onUnresolvedOwner: AddUnresolvedOwner);
                        rows.AddRange(links.Select(l => new ReportRow
                        {
                            ObjectName = objectName,
                            ObjectPrimarySmtpAddress = objectSmtp,
                            ObjectType = "Mail-Enabled Security Group",
                            MemberOrDelegate = $"{l.DisplayName} ({l.PrimarySmtpAddress})",
                            RoleOrPermission = l.Role
                        }));
                        break;
                    }
                case "SharedMailbox":
                case "RoomMailbox":
                case "EquipmentMailbox":
                case "UserMailbox":
                case "TeamMailbox":
                case "LegacyMailbox":
                case "LinkedMailbox":
                case "RemoteUserMailbox":
                case "RemoteSharedMailbox":
                case "RemoteRoomMailbox":
                case "RemoteEquipmentMailbox":
                case "RemoteTeamMailbox":
                    {
                        rows.AddRange(await BuildMailboxPermissionRowsAsync(recipient.Identity ?? identity, objectName, objectSmtp, recipientType));
                        break;
                    }
                default:
                    throw new InvalidOperationException($"The object type '{recipientType}' is not supported for reporting.");
            }

            return rows;
        }

        /// <summary>
        /// Resolves an identity and classifies it as a group-type or mailbox-type report target,
        /// so the UI can validate the input and offer the correct report action before running it.
        /// </summary>
        public async Task<ReportTargetInfo> ClassifyReportTargetAsync(string identity)
        {
            var recipient = await _exo.GetRecipientAsync(identity);
            if (recipient == null)
            {
                throw new InvalidOperationException($"Could not find a recipient matching '{identity}'.");
            }

            var recipientType = recipient.RecipientTypeDetails ?? string.Empty;
            var (kind, friendlyType) = recipientType switch
            {
                "GroupMailbox" => (ReportTargetKind.Group, "Microsoft 365 Group"),
                "MailUniversalDistributionGroup" => (ReportTargetKind.Group, "Distribution List"),
                "DynamicDistributionGroup" => (ReportTargetKind.Group, "Dynamic Distribution Group"),
                "MailUniversalSecurityGroup" => (ReportTargetKind.Group, "Mail-Enabled Security Group"),
                "SharedMailbox" => (ReportTargetKind.Mailbox, "Shared Mailbox"),
                "RoomMailbox" => (ReportTargetKind.Mailbox, "Room Mailbox"),
                "EquipmentMailbox" => (ReportTargetKind.Mailbox, "Equipment Mailbox"),
                "UserMailbox" => (ReportTargetKind.Mailbox, "User Mailbox"),
                "TeamMailbox" => (ReportTargetKind.Mailbox, "Team Mailbox"),
                "LegacyMailbox" => (ReportTargetKind.Mailbox, "Legacy Mailbox"),
                "LinkedMailbox" => (ReportTargetKind.Mailbox, "Linked Mailbox"),
                "RemoteUserMailbox" => (ReportTargetKind.Mailbox, "Remote (Hybrid) User Mailbox"),
                "RemoteSharedMailbox" => (ReportTargetKind.Mailbox, "Remote (Hybrid) Shared Mailbox"),
                "RemoteRoomMailbox" => (ReportTargetKind.Mailbox, "Remote (Hybrid) Room Mailbox"),
                "RemoteEquipmentMailbox" => (ReportTargetKind.Mailbox, "Remote (Hybrid) Equipment Mailbox"),
                "RemoteTeamMailbox" => (ReportTargetKind.Mailbox, "Remote (Hybrid) Team Mailbox"),
                _ => (ReportTargetKind.Unsupported, string.IsNullOrEmpty(recipientType) ? "Unknown" : recipientType)
            };

            return new ReportTargetInfo
            {
                Identity = recipient.Identity ?? identity,
                Name = recipient.Name ?? identity,
                PrimarySmtpAddress = recipient.PrimarySmtpAddress ?? identity,
                RecipientTypeDetails = recipientType,
                FriendlyType = friendlyType,
                Kind = kind
            };
        }

        public const int ActivityLookbackDays = 180;

        /// <summary>
        /// Read-only "is this account still used" check: newest Inbox/Sent Items dates and logon times for
        /// mailboxes, plus a 90-day message trace for every target type. Failed checks are shown, never assumed empty.
        /// </summary>
        public async Task<List<ReportRow>> GenerateAccountActivityAsync(ReportTargetInfo target, DateTime? utcNow = null)
        {
            var now = utcNow ?? DateTime.UtcNow;
            var cutoff = now.AddDays(-ActivityLookbackDays);
            var address = target.PrimarySmtpAddress;
            var details = new List<ReportRow>();
            var evidence = new List<DateTime>();
            var anyFailed = false;

            ReportRow Row(string check, string result, string status = "Read") => new()
            {
                ObjectName = target.Name,
                ObjectPrimarySmtpAddress = address,
                ObjectType = target.FriendlyType,
                MemberOrDelegate = check,
                RoleOrPermission = result,
                Status = status
            };

            async Task<bool> CheckAsync(string check, string noneText, bool isMailActivity, Func<Task<DateTime?>> read)
            {
                _exo.OperationCancellationToken.ThrowIfCancellationRequested();
                try
                {
                    var date = await read();
                    details.Add(Row(check, date.HasValue ? date.Value.ToString("yyyy-MM-dd HH:mm 'UTC'") : noneText));
                    if (isMailActivity && date.HasValue) evidence.Add(date.Value);
                    return true;
                }
                catch (Exception ex) when (ex is not OperationCanceledException && !_exo.OperationCancellationToken.IsCancellationRequested)
                {
                    anyFailed = true;
                    details.Add(Row(check, string.Empty, $"Unavailable: {ex.Message}"));
                    Logger.Log($"{check} unavailable: {ex.Message}", LogType.Warning);
                    return false;
                }
            }

            var hasMailbox = target.Kind == ReportTargetKind.Mailbox || target.RecipientTypeDetails == "GroupMailbox";
            var foldersChecked = false;
            if (hasMailbox)
            {
                var inbox = await CheckAsync("Newest Inbox item", "No items", true, () => _exo.GetNewestFolderItemDateAsync(address, "Inbox"));
                var sent = await CheckAsync("Newest Sent Items item", "No items", true, () => _exo.GetNewestFolderItemDateAsync(address, "SentItems"));
                foldersChecked = inbox && sent;
            }
            if (target.Kind == ReportTargetKind.Mailbox)
            {
                (DateTime? LastLogon, DateTime? LastUserAction) usage = default;
                var usageRead = await CheckAsync("Last logon", "Never", false, async () => { usage = await _exo.GetMailboxUsageAsync(address); return usage.LastLogon; });
                if (usageRead)
                    await CheckAsync("Last user action", "None recorded", false, () => Task.FromResult(usage.LastUserAction));
            }
            await CheckAsync("Last message sent (trace)", "None in the last 89 days", true, () => _exo.GetNewestMessageTraceDateAsync(address, true, utcNow: now));
            await CheckAsync("Last message received (trace)", "None in the last 89 days", true, () => _exo.GetNewestMessageTraceDateAsync(address, false, utcNow: now));

            var recent = evidence.Where(date => date >= cutoff).ToList();
            var verdict = recent.Count > 0 ? $"Active - most recent mail activity {recent.Max():yyyy-MM-dd HH:mm 'UTC'}"
                : anyFailed ? "Unconfirmed - one or more checks were unavailable"
                : foldersChecked ? "No mail activity found in the last 6 months"
                : "No mail activity in the last 89 days; older activity cannot be checked for this account type";

            var rows = new List<ReportRow> { Row("Verdict", verdict) };
            rows.AddRange(details);
            rows.Add(Row("Coverage", hasMailbox
                ? "Inbox/Sent Items dates include subfolders; items deleted or moved elsewhere are not counted. Message trace covers 89 days."
                : "No mailbox to inspect; only the 89-day message trace is available."));
            return rows;
        }

        /// <summary>
        /// Quick single-mailbox delegate lookup (Full Access / Send As / Send on Behalf), ported from
        /// DelegateReporter.ps1's Get-MailboxDelegates.
        /// </summary>
        public async Task<List<ReportRow>> GetMailboxDelegatesAsync(string mailboxIdentity)
        {
            var exists = await _exo.MailboxExistsAsync(mailboxIdentity);
            if (!exists)
            {
                throw new InvalidOperationException($"Could not find or access mailbox '{mailboxIdentity}'.");
            }

            var recipient = await _exo.GetRecipientAsync(mailboxIdentity);
            var objectName = recipient?.Name ?? mailboxIdentity;
            var objectSmtp = recipient?.PrimarySmtpAddress ?? mailboxIdentity;
            var recipientType = recipient?.RecipientTypeDetails ?? string.Empty;

            return await BuildMailboxPermissionRowsAsync(recipient?.Identity ?? mailboxIdentity, objectName, objectSmtp, recipientType);
        }

        private async Task<List<ReportRow>> BuildMailboxPermissionRowsAsync(string mailboxIdentity, string objectName, string objectSmtp, string recipientType)
        {
            var rows = new List<ReportRow>();
            ReportRow CreateRow(string principal, string permission, string status = "Read") => new()
            {
                ObjectName = objectName,
                ObjectPrimarySmtpAddress = objectSmtp,
                ObjectType = recipientType,
                MemberOrDelegate = principal,
                RoleOrPermission = permission,
                Status = status
            };
            var sections = new (string Permission, Func<Task<List<string>>> Read)[]
            {
                ("Full Access", () => _exo.GetAllFullAccessDelegatesAsync(mailboxIdentity)),
                ("Send As", () => _exo.GetAllSendAsDelegatesAsync(mailboxIdentity)),
                ("Send on Behalf", () => _exo.GetAllSendOnBehalfDelegatesAsync(mailboxIdentity,
                    (principal, error) => rows.Add(CreateRow(principal, "Send on Behalf", $"Unresolved: {error}"))))
            };
            foreach (var section in sections)
            {
                _exo.OperationCancellationToken.ThrowIfCancellationRequested();
                Logger.Log($"Getting {section.Permission} permissions...");
                try
                {
                    var principals = await section.Read();
                    rows.AddRange(principals.Select(principal => CreateRow(principal, section.Permission)));
                }
                catch (Exception ex) when (ex is not OperationCanceledException && !_exo.OperationCancellationToken.IsCancellationRequested)
                {
                    rows.Add(CreateRow(string.Empty, section.Permission, $"Unavailable: {ex.Message}"));
                    Logger.Log($"{section.Permission} report incomplete: {ex.Message}", LogType.Warning);
                }
            }

            return rows;
        }
    }
}

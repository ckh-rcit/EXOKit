using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EXOKit.Services
{
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
        Mailbox
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

        public ReportingService(ExoPowerShellService exo)
        {
            _exo = exo;
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
                        var links = await _exo.GetDistributionGroupReportLinksAsync(recipient.Identity ?? identity, isDynamic: false);
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
                        var links = await _exo.GetDistributionGroupReportLinksAsync(recipient.Identity ?? identity, isDynamic: true);
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
                        var links = await _exo.GetDistributionGroupReportLinksAsync(recipient.Identity ?? identity, isDynamic: false);
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

            return await BuildMailboxPermissionRowsAsync(mailboxIdentity, objectName, objectSmtp, recipientType);
        }

        private async Task<List<ReportRow>> BuildMailboxPermissionRowsAsync(string mailboxIdentity, string objectName, string objectSmtp, string recipientType)
        {
            var rows = new List<ReportRow>();

            var fullAccess = await _exo.GetAllFullAccessDelegatesAsync(mailboxIdentity);
            rows.AddRange(fullAccess.Select(u => new ReportRow
            {
                ObjectName = objectName,
                ObjectPrimarySmtpAddress = objectSmtp,
                ObjectType = recipientType,
                MemberOrDelegate = u,
                RoleOrPermission = "Full Access"
            }));

            var sendAs = await _exo.GetAllSendAsDelegatesAsync(mailboxIdentity);
            rows.AddRange(sendAs.Select(u => new ReportRow
            {
                ObjectName = objectName,
                ObjectPrimarySmtpAddress = objectSmtp,
                ObjectType = recipientType,
                MemberOrDelegate = u,
                RoleOrPermission = "Send As"
            }));

            var sendOnBehalf = await _exo.GetAllSendOnBehalfDelegatesAsync(mailboxIdentity);
            rows.AddRange(sendOnBehalf.Select(u => new ReportRow
            {
                ObjectName = objectName,
                ObjectPrimarySmtpAddress = objectSmtp,
                ObjectType = recipientType,
                MemberOrDelegate = u,
                RoleOrPermission = "Send on Behalf"
            }));

            return rows;
        }
    }
}

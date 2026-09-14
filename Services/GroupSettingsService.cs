using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EXOKit.Services
{
    /// <summary>
    /// UI-facing snapshot of a Distribution Group's settings, mirroring the EAC "Delivery management",
    /// "Edit delegates", "Edit message approval", and "Edit membership approvals" panels.
    /// </summary>
    public class GroupSettingsSnapshot
    {
        // Delivery management
        public bool AllowExternalSenders { get; set; }
        public string[] SpecifiedSenders { get; set; } = Array.Empty<string>();

        // Manage delegates
        public string[] SendAsDelegates { get; set; } = Array.Empty<string>();
        public string[] SendOnBehalfDelegates { get; set; } = Array.Empty<string>();

        // Message approval
        public bool RequireModeratorApproval { get; set; }
        public string[] Moderators { get; set; } = Array.Empty<string>();
        public string[] BypassModerationSenders { get; set; } = Array.Empty<string>();
        public string NotifySenderMode { get; set; } = "Always";

        // Membership approvals
        public string JoinRestriction { get; set; } = "Open";
        public string DepartRestriction { get; set; } = "Open";
    }

    /// <summary>
    /// Loads and saves Distribution Group / Mail-Enabled Security Group settings for the "Group Settings"
    /// section, ported from the Exchange Admin Center's "Delivery management", "Edit delegates", "Edit
    /// message approval", and "Edit membership approvals" panels. Scoped to Distribution Groups only
    /// (Microsoft 365 Groups expose similar options through Microsoft Graph and are out of scope for now).
    /// </summary>
    public class GroupSettingsService
    {
        private static readonly string[] SupportedDistributionGroupTypes = { "MailUniversalDistributionGroup", "MailNonUniversalGroup", "MailUniversalSecurityGroup" };

        private readonly ExoPowerShellService _exo;
        private readonly SnapshotService _snapshots;
        private readonly Dictionary<string, GroupSettingsSnapshot> _loaded = new(StringComparer.OrdinalIgnoreCase);

        public GroupSettingsService(ExoPowerShellService exo, SnapshotService? snapshots = null)
        {
            _exo = exo;
            _snapshots = snapshots ?? new SnapshotService();
        }

        /// <summary>
        /// Confirms the identity resolves to a Distribution Group or Mail-Enabled Security Group before
        /// loading/saving settings, since Microsoft 365 Groups are not supported by these EXO cmdlets.
        /// </summary>
        public async Task<string?> ValidateDistributionGroupAsync(string identity)
        {
            var groupType = await _exo.GetDistributionGroupTypeAsync(identity);
            if (groupType == null)
            {
                Logger.Log($"  ERROR: Group '{identity}' not found.", LogType.Error);
                return null;
            }

            if (!SupportedDistributionGroupTypes.Contains(groupType, StringComparer.OrdinalIgnoreCase))
            {
                Logger.Log($"  ERROR: '{identity}' is a '{groupType}', not a Distribution Group or Mail-Enabled Security Group. Microsoft 365 Groups are not supported in this section.", LogType.Error);
                return null;
            }

            return groupType;
        }

        public async Task<GroupSettingsSnapshot?> LoadSettingsAsync(string identity)
        {
            var groupType = await ValidateDistributionGroupAsync(identity);
            if (groupType == null) return null;

            Logger.Log($"Loading group settings for '{identity}'...");
            var settings = await _exo.GetDistributionGroupSettingsAsync(identity);
            if (settings == null)
            {
                Logger.Log($"  ERROR: Could not read settings for '{identity}'.", LogType.Error);
                return null;
            }

            var sendAsDelegates = await _exo.GetSendAsDelegatesAsync(identity);

            var snapshot = new GroupSettingsSnapshot
            {
                AllowExternalSenders = !settings.RequireSenderAuthenticationEnabled,
                SpecifiedSenders = await _exo.ResolveRecipientAddressesAsync(settings.AcceptMessagesOnlyFromSendersOrMembers),
                SendAsDelegates = await _exo.ResolveRecipientAddressesAsync(sendAsDelegates),
                SendOnBehalfDelegates = await _exo.ResolveRecipientAddressesAsync(settings.GrantSendOnBehalfTo),
                RequireModeratorApproval = settings.ModerationEnabled,
                Moderators = await _exo.ResolveRecipientAddressesAsync(settings.ModeratedBy),
                BypassModerationSenders = await _exo.ResolveRecipientAddressesAsync(settings.BypassModerationFromSendersOrMembers),
                NotifySenderMode = settings.SendModerationNotifications,
                JoinRestriction = settings.MemberJoinRestriction,
                DepartRestriction = settings.MemberDepartRestriction
            };

            Logger.Log($"  SUCCESS: Loaded settings for '{identity}'.", LogType.Success);
            _loaded[identity] = System.Text.Json.JsonSerializer.Deserialize<GroupSettingsSnapshot>(System.Text.Json.JsonSerializer.Serialize(snapshot))!;
            return snapshot;
        }

        private async Task EnsureUnchangedAsync(string identity)
        {
            if (!_loaded.TryGetValue(identity, out var loaded)) throw new InvalidOperationException("Load the group's settings before saving.");
            var current = await LoadSettingsAsync(identity) ?? throw new InvalidOperationException("Group settings unavailable.");
            if (loaded.AllowExternalSenders != current.AllowExternalSenders
                || loaded.RequireModeratorApproval != current.RequireModeratorApproval
                || !string.Equals(loaded.NotifySenderMode, current.NotifySenderMode, StringComparison.OrdinalIgnoreCase)
                || !string.Equals(loaded.JoinRestriction, current.JoinRestriction, StringComparison.OrdinalIgnoreCase)
                || !string.Equals(loaded.DepartRestriction, current.DepartRestriction, StringComparison.OrdinalIgnoreCase)
                || !SameSet(loaded.SpecifiedSenders, current.SpecifiedSenders)
                || !SameSet(loaded.SendAsDelegates, current.SendAsDelegates)
                || !SameSet(loaded.SendOnBehalfDelegates, current.SendOnBehalfDelegates)
                || !SameSet(loaded.Moderators, current.Moderators)
                || !SameSet(loaded.BypassModerationSenders, current.BypassModerationSenders))
            {
                _loaded.Remove(identity);
                throw new InvalidOperationException("Group settings changed since loading. Reload and review before saving.");
            }
        }

        private async Task VerifySettingsAsync(string identity, Func<GroupSettingsSnapshot, bool> matches)
        {
            for (var attempt = 0; attempt < 15; attempt++)
            {
                var current = await LoadSettingsAsync(identity);
                if (current != null && matches(current)) return;
                await Task.Delay(TimeSpan.FromSeconds(2), _exo.OperationCancellationToken);
            }
            _loaded.Remove(identity);
            throw new InvalidOperationException("Settings change is unconfirmed. Reload and verify before retrying.");
        }

        private static bool SameSet(IEnumerable<string> first, IEnumerable<string> second) =>
            new HashSet<string>(first, StringComparer.OrdinalIgnoreCase).SetEquals(second);

        public async Task<bool> SaveDeliveryManagementAsync(string identity, bool allowExternalSenders, string[] specifiedSenders)
        {
            if (await ValidateDistributionGroupAsync(identity) == null) return false;

            try
            {
                Logger.Log($"Executing: Set-DistributionGroup -Identity '{identity}' -RequireSenderAuthenticationEnabled {!allowExternalSenders} -AcceptMessagesOnlyFromSendersOrMembers ...");
                await EnsureUnchangedAsync(identity);
                specifiedSenders = await _exo.ResolveRecipientAddressesAsync(specifiedSenders);
                await _exo.SetDeliveryManagementAsync(identity, allowExternalSenders, specifiedSenders);
                await VerifySettingsAsync(identity, current => current.AllowExternalSenders == allowExternalSenders && SameSet(current.SpecifiedSenders, specifiedSenders));
                Logger.Log($"  SUCCESS: Delivery management settings updated for '{identity}'.", LogType.Success);

                WriteGroupSettingsTicketNote(identity, "Delivery Management", new List<string>
                {
                    $"Allow Senders Inside and Outside the Organization: {(allowExternalSenders ? "Yes" : "No")}",
                    $"Specified Senders: {(specifiedSenders.Length > 0 ? string.Join(", ", specifiedSenders) : "(none)")}"
                });

                return true;
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                Logger.Log($"  ERROR: Failed to update delivery management settings for '{identity}'. DETAILS: {ex.Message}", LogType.Error);
                return false;
            }
        }

        public async Task<bool> SaveDelegatesAsync(string identity, string[] desiredSendAsDelegates, string[] desiredSendOnBehalfDelegates)
        {
            if (await ValidateDistributionGroupAsync(identity) == null) return false;

            try
            {
                desiredSendAsDelegates = await _exo.ResolveRecipientAddressesAsync(desiredSendAsDelegates);
                await EnsureUnchangedAsync(identity);
                desiredSendOnBehalfDelegates = await _exo.ResolveRecipientAddressesAsync(desiredSendOnBehalfDelegates);
                var currentSendAs = await _exo.ResolveRecipientAddressesAsync(await _exo.GetSendAsDelegatesAsync(identity));
                var toAddSendAs = desiredSendAsDelegates.Where(d => !currentSendAs.Contains(d, StringComparer.OrdinalIgnoreCase)).ToArray();
                var toRemoveSendAs = currentSendAs.Where(d => !desiredSendAsDelegates.Contains(d, StringComparer.OrdinalIgnoreCase)).ToArray();

                var currentSettings = await _exo.GetDistributionGroupSettingsAsync(identity);
                if (currentSettings == null) throw new InvalidOperationException("Group settings unavailable; no delegates changed.");
                var currentSendOnBehalf = await _exo.ResolveRecipientAddressesAsync(currentSettings.GrantSendOnBehalfTo);
                var toRemoveSendOnBehalf = currentSendOnBehalf.Where(d => !desiredSendOnBehalfDelegates.Contains(d, StringComparer.OrdinalIgnoreCase)).ToArray();

                var snapshotItems = new List<SnapshotItem>();
                foreach (var d in toRemoveSendAs)
                {
                    snapshotItems.Add(new SnapshotItem { User = d, Role = "Send As" });
                }
                foreach (var d in toRemoveSendOnBehalf)
                {
                    snapshotItems.Add(new SnapshotItem { User = d, Role = "Send on Behalf" });
                }

                if (snapshotItems.Count > 0)
                {
                    try
                    {
                        _snapshots.SaveSnapshot(new SnapshotRecord
                        {
                            OperationType = "GroupDelegateRemoval",
                            Target = identity,
                            Description = $"Before removing {snapshotItems.Count} delegate(s) from group '{identity}'",
                            Items = snapshotItems
                        });
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"Failed to save group delegate removal snapshot for '{identity}'. DETAILS: {ex.Message}", LogType.Error);
                        throw;
                    }
                }

                foreach (var d in toAddSendAs)
                {
                    Logger.Log($"Executing: Add-RecipientPermission -Identity '{identity}' -Trustee '{d}' -AccessRights SendAs");
                    try
                    {
                        await _exo.AddSendAsDelegateAsync(identity, d);
                    }
                    catch (Exception ex) when (ex.Message.Contains("Object reference not set to an instance of an object", StringComparison.OrdinalIgnoreCase))
                    {
                        Logger.Log($"  Add Send As for '{d}' returned server error (Object reference not set). Retrying once...", LogType.Warning);
                        await Task.Delay(2000, _exo.OperationCancellationToken);
                        await _exo.AddSendAsDelegateAsync(identity, d);
                    }
                }

                foreach (var d in toRemoveSendAs)
                {
                    Logger.Log($"Executing: Remove-RecipientPermission -Identity '{identity}' -Trustee '{d}' -AccessRights SendAs");
                    try
                    {
                        await _exo.RemoveSendAsDelegateAsync(identity, d);
                    }
                    catch (Exception ex) when (ex.Message.Contains("Object reference not set to an instance of an object", StringComparison.OrdinalIgnoreCase))
                    {
                        Logger.Log($"  Remove Send As for '{d}' returned server error (Object reference not set). Retrying once...", LogType.Warning);
                        await Task.Delay(2000, _exo.OperationCancellationToken);
                        await _exo.RemoveSendAsDelegateAsync(identity, d);
                    }
                }

                Logger.Log($"Executing: Set-DistributionGroup -Identity '{identity}' -GrantSendOnBehalfTo ...");
                await _exo.SetSendOnBehalfDelegatesAsync(identity, desiredSendOnBehalfDelegates);
                await VerifySettingsAsync(identity, current => SameSet(current.SendAsDelegates, desiredSendAsDelegates) && SameSet(current.SendOnBehalfDelegates, desiredSendOnBehalfDelegates));

                Logger.Log($"  SUCCESS: Delegates updated for '{identity}'.", LogType.Success);

                WriteGroupSettingsTicketNote(identity, "Delegates", new List<string>
                {
                    $"Send As Added: {(toAddSendAs.Length > 0 ? string.Join(", ", toAddSendAs) : "(none)")}",
                    $"Send As Removed: {(toRemoveSendAs.Length > 0 ? string.Join(", ", toRemoveSendAs) : "(none)")}",
                    $"Send on Behalf: {(desiredSendOnBehalfDelegates.Length > 0 ? string.Join(", ", desiredSendOnBehalfDelegates) : "(none)")}",
                    $"Send on Behalf Removed: {(toRemoveSendOnBehalf.Length > 0 ? string.Join(", ", toRemoveSendOnBehalf) : "(none)")}"
                });

                return true;
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                Logger.Log($"  ERROR: Failed to update delegates for '{identity}'. DETAILS: {ex.Message}", LogType.Error);
                return false;
            }
        }

        public async Task<bool> SaveMessageApprovalAsync(string identity, bool requireModeratorApproval, string[] moderators, string[] bypassSenders, string notifySenderMode)
        {
            if (await ValidateDistributionGroupAsync(identity) == null) return false;

            if (requireModeratorApproval && moderators.Length == 0)
            {
                Logger.Log("  ERROR: At least one moderator is required when moderator approval is enabled.", LogType.Error);
                return false;
            }

            try
            {
                Logger.Log($"Executing: Set-DistributionGroup -Identity '{identity}' -ModerationEnabled {requireModeratorApproval} -ModeratedBy ... -SendModerationNotifications {notifySenderMode}");
                await EnsureUnchangedAsync(identity);
                moderators = await _exo.ResolveRecipientAddressesAsync(moderators);
                bypassSenders = await _exo.ResolveRecipientAddressesAsync(bypassSenders);
                await _exo.SetMessageApprovalAsync(identity, requireModeratorApproval, moderators, bypassSenders, notifySenderMode);
                await VerifySettingsAsync(identity, current => current.RequireModeratorApproval == requireModeratorApproval
                    && SameSet(current.Moderators, moderators) && SameSet(current.BypassModerationSenders, bypassSenders) && current.NotifySenderMode == notifySenderMode);
                Logger.Log($"  SUCCESS: Message approval settings updated for '{identity}'.", LogType.Success);

                WriteGroupSettingsTicketNote(identity, "Message Approval", new List<string>
                {
                    $"Require Moderator Approval: {(requireModeratorApproval ? "Yes" : "No")}",
                    $"Moderators: {(moderators.Length > 0 ? string.Join(", ", moderators) : "(none)")}",
                    $"Bypass Moderation Senders: {(bypassSenders.Length > 0 ? string.Join(", ", bypassSenders) : "(none)")}",
                    $"Notify Sender Mode: {notifySenderMode}"
                });

                return true;
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                Logger.Log($"  ERROR: Failed to update message approval settings for '{identity}'. DETAILS: {ex.Message}", LogType.Error);
                return false;
            }
        }

        public async Task<bool> SaveMembershipApprovalAsync(string identity, string joinRestriction, string departRestriction)
        {
            var groupType = await ValidateDistributionGroupAsync(identity);
            if (groupType == null) return false;
            if (groupType == "MailUniversalSecurityGroup")
            {
                joinRestriction = "Closed";
                departRestriction = "Closed";
            }

            try
            {
                Logger.Log($"Executing: Set-DistributionGroup -Identity '{identity}' -MemberJoinRestriction {joinRestriction} -MemberDepartRestriction {departRestriction}");
                await EnsureUnchangedAsync(identity);
                await _exo.SetMembershipApprovalAsync(identity, joinRestriction, departRestriction);
                await VerifySettingsAsync(identity, current => current.JoinRestriction == joinRestriction && current.DepartRestriction == departRestriction);
                Logger.Log($"  SUCCESS: Membership approval settings updated for '{identity}'.", LogType.Success);

                WriteGroupSettingsTicketNote(identity, "Membership Approval", new List<string>
                {
                    $"Join Restriction: {joinRestriction}",
                    $"Depart Restriction: {departRestriction}"
                });

                return true;
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                Logger.Log($"  ERROR: Failed to update membership approval settings for '{identity}'. DETAILS: {ex.Message}", LogType.Error);
                return false;
            }
        }

        private static void WriteGroupSettingsTicketNote(string identity, string sectionLabel, List<string> details)
        {
            Logger.Log("--- For IT Ticket ---", LogType.Ticket);
            Logger.Log($"Group Settings Change: {sectionLabel}", LogType.Ticket);
            Logger.Log($"Group: {identity}", LogType.Ticket);
            foreach (var detail in details)
            {
                Logger.Log(detail, LogType.Ticket);
            }
            Logger.Log("---------------------", LogType.Ticket);
        }
    }
}

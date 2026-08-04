using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EXOKit.Services
{
    public enum GroupKind
    {
        M365,
        DistributionGroup
    }

    public enum GroupRole
    {
        Member,
        Owner
    }

    public class GroupOperationContext
    {
        public bool IsValid { get; set; }
        public string? GroupType { get; set; }
        public GroupKind GroupKind { get; set; }
        public string? M365GroupId { get; set; }
        public string? FailureStatus { get; set; }
        public string? FailureMessage { get; set; }
    }

    public class GroupRoleSelections
    {
        public bool Member { get; set; }
        public bool Owner { get; set; }
    }

    public class GroupRoleResult
    {
        public string User { get; set; } = string.Empty;
        public string Group { get; set; } = string.Empty;
        public List<string> Statuses { get; } = new();
    }

    /// <summary>
    /// Ports Resolve-GroupOperationContext, Test-GroupRolePresence, and Invoke-GroupMembershipOperation
    /// from the toolkit script: resolves whether a group identity is a Distribution Group or M365
    /// (Unified) Group, adds/removes Member and Owner roles via the correct backend (Graph for M365,
    /// EXO cmdlets for Distribution Groups), and polls once per second for up to 10 seconds to verify
    /// the change actually applied before reporting final status.
    /// </summary>
    public class GroupMembershipService
    {
        private static readonly string[] SupportedDistributionGroupTypes = { "MailUniversalDistributionGroup", "MailNonUniversalGroup" };

        private readonly ExoPowerShellService _exo;
        private readonly GraphService _graph;

        public GroupMembershipService(ExoPowerShellService exo, GraphService graph)
        {
            _exo = exo;
            _graph = graph;
        }

        public async Task<GroupOperationContext> ResolveGroupOperationContextAsync(string groupEmail, bool exoConnected, bool graphConnected)
        {
            var groupRecipient = await _exo.GetRecipientAsync(groupEmail);
            if (groupRecipient == null)
            {
                throw new InvalidOperationException($"Group '{groupEmail}' not found.");
            }

            var groupType = groupRecipient.RecipientTypeDetails ?? string.Empty;
            var isM365Group = string.Equals(groupType, "GroupMailbox", StringComparison.OrdinalIgnoreCase);
            var isDistributionGroup = SupportedDistributionGroupTypes.Contains(groupType, StringComparer.OrdinalIgnoreCase);

            if (!isM365Group && !isDistributionGroup)
            {
                return new GroupOperationContext
                {
                    IsValid = false,
                    GroupType = groupType,
                    FailureStatus = $"Skipped - Unsupported Group Type ({groupType})",
                    FailureMessage = $"ERROR: Group '{groupEmail}' is type '{groupType}', which is not supported by this tool. Supported types: Distribution Groups and Microsoft 365 Groups only."
                };
            }

            if (isM365Group)
            {
                if (!graphConnected)
                {
                    return new GroupOperationContext
                    {
                        IsValid = false,
                        GroupType = groupType,
                        FailureStatus = "Skipped - Graph Connection Required for M365 Group",
                        FailureMessage = $"ERROR: Graph connection required for M365 Group '{groupEmail}'. Skipping."
                    };
                }

                var m365GroupId = await _graph.ResolveM365GroupIdAsync(groupEmail);
                if (string.IsNullOrEmpty(m365GroupId))
                {
                    throw new InvalidOperationException($"M365 Group '{groupEmail}' not found in Graph.");
                }

                return new GroupOperationContext
                {
                    IsValid = true,
                    GroupType = groupType,
                    GroupKind = GroupKind.M365,
                    M365GroupId = m365GroupId
                };
            }

            if (!exoConnected)
            {
                return new GroupOperationContext
                {
                    IsValid = false,
                    GroupType = groupType,
                    FailureStatus = "Skipped - EXO Connection Required for Distribution Group",
                    FailureMessage = $"ERROR: EXO connection required for Distribution Group '{groupEmail}'. Skipping."
                };
            }

            return new GroupOperationContext
            {
                IsValid = true,
                GroupType = groupType,
                GroupKind = GroupKind.DistributionGroup,
                M365GroupId = null
            };
        }

        private async Task<bool> TestGroupRolePresenceAsync(GroupKind groupKind, GroupRole role, string groupIdentity, string? m365GroupId, string? userId, RecipientInfo? userRecipient, string userIdentityFallback)
        {
            if (groupKind == GroupKind.M365)
            {
                if (string.IsNullOrWhiteSpace(m365GroupId) || string.IsNullOrWhiteSpace(userId)) return false;
                return role == GroupRole.Member
                    ? await _graph.IsGroupMemberAsync(m365GroupId, userId)
                    : await _graph.IsGroupOwnerAsync(m365GroupId, userId);
            }

            if (userRecipient == null) return false;
            return role == GroupRole.Member
                ? await _exo.IsDistributionGroupMemberAsync(groupIdentity, userRecipient, userIdentityFallback)
                : await _exo.IsDistributionGroupOwnerAsync(groupIdentity, userRecipient, userIdentityFallback);
        }

        public async Task<List<GroupRoleResult>> InvokeGroupMembershipOperationAsync(
            PermissionOperationType operationType,
            IEnumerable<string> targetGroups,
            IEnumerable<string> users,
            GroupRoleSelections roles,
            bool exoConnected,
            bool graphConnected)
        {
            var groupList = targetGroups.ToList();
            var userList = users.ToList();
            var resultMap = new Dictionary<string, Dictionary<string, List<string>>>(StringComparer.OrdinalIgnoreCase);

            var actionTitle = operationType == PermissionOperationType.Add ? "Assignment" : "Removal";
            Logger.Log($"--- Starting Group Role {actionTitle} ---");

            foreach (var groupEmail in groupList)
            {
                Logger.Log($"Processing Group: {groupEmail}");
                var validationQueue = new List<(string User, GroupRole Role, GroupKind GroupKind, string? GroupId, string? UserId, RecipientInfo? UserRecipient)>();

                GroupOperationContext? groupContext;
                try
                {
                    groupContext = await ResolveGroupOperationContextAsync(groupEmail, exoConnected, graphConnected);
                }
                catch (Exception ex)
                {
                    Logger.Log($"ERROR: Group '{groupEmail}' could not be resolved. Skipping. DETAILS: {ex.Message}", LogType.Error);
                    foreach (var userEmail in userList)
                    {
                        GetOrCreateResultList(resultMap, userEmail, groupEmail, out var statuses);
                        statuses.Add("Group Not Found");
                    }
                    continue;
                }

                if (!groupContext.IsValid)
                {
                    Logger.Log(groupContext.FailureMessage ?? "ERROR: Group could not be resolved.", LogType.Error);
                    foreach (var userEmail in userList)
                    {
                        GetOrCreateResultList(resultMap, userEmail, groupEmail, out var statuses);
                        statuses.Add(groupContext.FailureStatus ?? "Skipped");
                    }
                    continue;
                }

                foreach (var userEmail in userList)
                {
                    Logger.Log($"  Processing User: {userEmail} for Group: {groupEmail} ({groupContext.GroupType})");
                    GetOrCreateResultList(resultMap, userEmail, groupEmail, out var statuses);

                    string? mgUserId = null;
                    RecipientInfo? userRecipient = null;

                    if (groupContext.GroupKind == GroupKind.M365)
                    {
                        mgUserId = await _graph.GetUserIdAsync(userEmail);
                        if (string.IsNullOrEmpty(mgUserId))
                        {
                            Logger.Log($"    ERROR: User '{userEmail}' not found (Graph). Skipping.", LogType.Error);
                            statuses.Add("User Not Found (Graph)");
                            continue;
                        }
                    }
                    else
                    {
                        userRecipient = await _exo.GetRecipientAsync(userEmail);
                        if (userRecipient == null)
                        {
                            Logger.Log($"    ERROR: User '{userEmail}' not found (EXO). Skipping.", LogType.Error);
                            statuses.Add("User Not Found (EXO)");
                            continue;
                        }
                    }

                    foreach (var role in new[] { GroupRole.Member, GroupRole.Owner })
                    {
                        if ((role == GroupRole.Member && !roles.Member) || (role == GroupRole.Owner && !roles.Owner))
                        {
                            continue;
                        }

                        Logger.Log($"    Attempting {operationType} {role}...");
                        var isPresent = await TestGroupRolePresenceAsync(groupContext.GroupKind, role, groupEmail, groupContext.M365GroupId, mgUserId, userRecipient, userEmail);

                        if (operationType == PermissionOperationType.Add)
                        {
                            await ProcessAddRoleAsync(groupEmail, groupContext, userEmail, role, isPresent, mgUserId, userRecipient, statuses, validationQueue);
                        }
                        else
                        {
                            await ProcessRemoveRoleAsync(groupEmail, groupContext, userEmail, role, isPresent, mgUserId, userRecipient, statuses, validationQueue);
                        }
                    }
                }

                if (validationQueue.Count > 0)
                {
                    await ValidateGroupRoleChangesAsync(operationType, groupEmail, validationQueue, resultMap);
                }
            }

            return resultMap.SelectMany(userEntry => userEntry.Value.Select(groupEntry =>
            {
                var result = new GroupRoleResult { User = userEntry.Key, Group = groupEntry.Key };
                result.Statuses.AddRange(groupEntry.Value);
                return result;
            })).ToList();
        }

        private async Task ProcessAddRoleAsync(
            string groupEmail, GroupOperationContext groupContext, string userEmail, GroupRole role, bool isPresent,
            string? mgUserId, RecipientInfo? userRecipient, List<string> statuses,
            List<(string User, GroupRole Role, GroupKind GroupKind, string? GroupId, string? UserId, RecipientInfo? UserRecipient)> validationQueue)
        {
            if (isPresent)
            {
                Logger.Log($"    STATUS: Already {role}.", LogType.Warning);
                statuses.Add($"{role} (Already Exists)");
                return;
            }

            try
            {
                if (groupContext.GroupKind == GroupKind.M365)
                {
                    if (role == GroupRole.Member)
                    {
                        await _graph.AddGroupMemberAsync(groupContext.M365GroupId!, mgUserId!);
                    }
                    else
                    {
                        await _graph.AddGroupOwnerAsync(groupContext.M365GroupId!, mgUserId!);
                    }
                }
                else
                {
                    if (role == GroupRole.Member)
                    {
                        await _exo.AddDistributionGroupMemberAsync(groupEmail, userEmail);
                    }
                    else
                    {
                        await _exo.AddDistributionGroupOwnerAsync(groupEmail, userEmail);
                    }
                }

                Logger.Log($"    INFO: {role} add command submitted. Pending validation.");
                statuses.Add($"{role} (Pending Validation)");
                validationQueue.Add((userEmail, role, groupContext.GroupKind, groupContext.M365GroupId, mgUserId, userRecipient));
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("already", StringComparison.OrdinalIgnoreCase))
                {
                    Logger.Log($"    STATUS: Already {role}.", LogType.Warning);
                    statuses.Add($"{role} (Already Exists)");
                }
                else
                {
                    Logger.Log($"    ERROR: Failed Add {role}. DETAILS: {ex.Message}", LogType.Error);
                    statuses.Add($"{role} (Error)");
                }
            }
        }

        private async Task ProcessRemoveRoleAsync(
            string groupEmail, GroupOperationContext groupContext, string userEmail, GroupRole role, bool isPresent,
            string? mgUserId, RecipientInfo? userRecipient, List<string> statuses,
            List<(string User, GroupRole Role, GroupKind GroupKind, string? GroupId, string? UserId, RecipientInfo? UserRecipient)> validationQueue)
        {
            if (!isPresent)
            {
                Logger.Log($"    STATUS: {role} not found.", LogType.Warning);
                statuses.Add($"{role} (Not Found)");
                return;
            }

            try
            {
                if (groupContext.GroupKind == GroupKind.M365)
                {
                    if (role == GroupRole.Member)
                    {
                        await _graph.RemoveGroupMemberAsync(groupContext.M365GroupId!, mgUserId!);
                    }
                    else
                    {
                        await _graph.RemoveGroupOwnerAsync(groupContext.M365GroupId!, mgUserId!);
                    }
                }
                else
                {
                    if (role == GroupRole.Member)
                    {
                        await _exo.RemoveDistributionGroupMemberAsync(groupEmail, userEmail);
                    }
                    else
                    {
                        await _exo.RemoveDistributionGroupOwnerAsync(groupEmail, userEmail);
                    }
                }

                Logger.Log($"    INFO: {role} remove command submitted. Pending validation.");
                statuses.Add($"{role} (Pending Validation)");
                validationQueue.Add((userEmail, role, groupContext.GroupKind, groupContext.M365GroupId, mgUserId, userRecipient));
            }
            catch (Exception ex)
            {
                if (role == GroupRole.Owner && ex.Message.Contains("last owner", StringComparison.OrdinalIgnoreCase))
                {
                    Logger.Log($"    WARNING: Cannot remove the last owner from M365 Group '{groupEmail}'.", LogType.Warning);
                    statuses.Add("Owner (Skipped - Last Owner Cannot Remove)");
                }
                else if (IsNotFoundError(ex.Message))
                {
                    Logger.Log($"    STATUS: {role} not found.", LogType.Warning);
                    statuses.Add($"{role} (Not Found)");
                }
                else
                {
                    Logger.Log($"    ERROR: Failed Remove {role}. DETAILS: {ex.Message}", LogType.Error);
                    statuses.Add($"{role} (Error)");
                }
            }
        }

        private async Task ValidateGroupRoleChangesAsync(
            PermissionOperationType operationType, string groupEmail,
            List<(string User, GroupRole Role, GroupKind GroupKind, string? GroupId, string? UserId, RecipientInfo? UserRecipient)> validationQueue,
            Dictionary<string, Dictionary<string, List<string>>> resultMap)
        {
            Logger.Log($"Validating group role changes for '{groupEmail}' once per second for up to 10 seconds...");
            var pending = new List<(string User, GroupRole Role, GroupKind GroupKind, string? GroupId, string? UserId, RecipientInfo? UserRecipient)>(validationQueue);

            for (var attempt = 1; attempt <= 10 && pending.Count > 0; attempt++)
            {
                await Task.Delay(1000);
                var remaining = new List<(string User, GroupRole Role, GroupKind GroupKind, string? GroupId, string? UserId, RecipientInfo? UserRecipient)>();

                foreach (var item in pending)
                {
                    if (!resultMap.TryGetValue(item.User, out var groupMap) || !groupMap.TryGetValue(groupEmail, out var statuses))
                    {
                        continue;
                    }

                    var isPresent = await TestGroupRolePresenceAsync(item.GroupKind, item.Role, groupEmail, item.GroupId, item.UserId, item.UserRecipient, item.User);
                    var isValidated = operationType == PermissionOperationType.Add ? isPresent : !isPresent;

                    if (isValidated)
                    {
                        var suffix = operationType == PermissionOperationType.Add ? "Added" : "Removed";
                        var verb = operationType == PermissionOperationType.Add ? "added" : "removed";
                        Logger.Log($"    SUCCESS: {item.Role} {verb} and verified for '{item.User}'.", LogType.Success);
                        ReplaceStatus(statuses, $"{item.Role} (Pending Validation)", $"{item.Role} ({suffix})");
                    }
                    else
                    {
                        remaining.Add(item);
                    }
                }

                pending = remaining;
            }

            foreach (var item in pending)
            {
                if (!resultMap.TryGetValue(item.User, out var groupMap) || !groupMap.TryGetValue(groupEmail, out var statuses))
                {
                    continue;
                }

                var suffix = operationType == PermissionOperationType.Add ? "Add - Verify Failed" : "Remove - Verify Failed";
                var opWord = operationType == PermissionOperationType.Add ? "add" : "remove";
                var stillOrNo = operationType == PermissionOperationType.Add ? "verification found no matching role" : "role still present";
                Logger.Log($"    WARNING: {item.Role} {opWord} command succeeded but {stillOrNo} for '{item.User}' within 10 seconds.", LogType.Warning);
                ReplaceStatus(statuses, $"{item.Role} (Pending Validation)", $"{item.Role} ({suffix})");
            }
        }

        private static bool GetOrCreateResultList(Dictionary<string, Dictionary<string, List<string>>> resultMap, string user, string group, out List<string> statuses)
        {
            var created = false;
            if (!resultMap.TryGetValue(user, out var groupMap))
            {
                groupMap = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
                resultMap[user] = groupMap;
            }
            if (!groupMap.TryGetValue(group, out var list))
            {
                list = new List<string>();
                groupMap[group] = list;
                created = true;
            }
            statuses = list;
            return created;
        }

        private static void ReplaceStatus(List<string> statuses, string oldValue, string newValue)
        {
            for (var i = 0; i < statuses.Count; i++)
            {
                if (string.Equals(statuses[i], oldValue, StringComparison.OrdinalIgnoreCase))
                {
                    statuses[i] = newValue;
                }
            }
        }

        private static bool IsNotFoundError(string message) =>
            message.Contains("isn't a member", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("wasn't found", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("doesn't exist", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("not found", StringComparison.OrdinalIgnoreCase);
    }
}

using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace EXOKit.Services
{
    /// <summary>
    /// Result of attempting to restore one item from a snapshot.
    /// </summary>
    public class RestoreItemResult
    {
        public string User { get; set; } = string.Empty;
        public string Role { get; set; } = string.Empty;
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
    }

    /// <summary>
    /// Replays a saved <see cref="SnapshotRecord"/> back through the existing EXO/Graph add-role and
    /// add-permission calls, effectively undoing a prior removal captured by <see cref="SnapshotService"/>.
    /// </summary>
    public class SnapshotRestoreService
    {
        private readonly ExoPowerShellService _exo;
        private readonly GraphService _graph;

        public SnapshotRestoreService(ExoPowerShellService exo, GraphService graph)
        {
            _exo = exo;
            _graph = graph;
        }

        public async Task<List<RestoreItemResult>> RestoreAsync(SnapshotRecord record)
        {
            if (!_exo.IsConnected) throw new InvalidOperationException("Connect Exchange Online before restoring.");
            SnapshotService.ValidateForRestore(record, _exo.ConnectedTenantId);
            var results = new List<RestoreItemResult>();

            switch (record.OperationType)
            {
                case "GroupMembershipRemoval":
                    await RestoreGroupMembershipAsync(record, results);
                    break;
                case "MailboxPermissionRemoval":
                case "GroupDelegateRemoval":
                    await RestoreMailboxPermissionsAsync(record, results);
                    break;
                default:
                    throw new InvalidOperationException($"Snapshot type '{record.OperationType}' is not supported for restore.");
            }

            return results;
        }

        private async Task RestoreGroupMembershipAsync(SnapshotRecord record, List<RestoreItemResult> results)
        {
            var groupEmail = record.Target;
            var isM365 = record.Metadata.TryGetValue("GroupKind", out var kind) && string.Equals(kind, GroupKind.M365.ToString(), StringComparison.OrdinalIgnoreCase);
            record.Metadata.TryGetValue("M365GroupId", out var m365GroupId);
            if (isM365 && (!_graph.IsConnected || !string.Equals(_graph.ConnectedTenantId, _exo.ConnectedTenantId, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("Connect Microsoft Graph to the snapshot tenant before restoring group roles.");
            if (kind != GroupKind.M365.ToString() && kind != GroupKind.DistributionGroup.ToString())
                throw new InvalidOperationException("Unknown snapshot group kind.");

            foreach (var item in record.Items)
            {
                try
                {
                    if (isM365)
                    {
                        if (item.Role != "Member" && item.Role != "Owner") throw new InvalidOperationException($"Unknown role '{item.Role}'.");
                        if (string.IsNullOrEmpty(m365GroupId) || string.IsNullOrEmpty(item.UserId))
                        {
                            throw new InvalidOperationException("Missing M365 group id or user id required to restore.");
                        }

                        if (string.Equals(item.Role, GroupRole.Member.ToString(), StringComparison.OrdinalIgnoreCase))
                        {
                            await PermissionVerification.ApplyAsync(() => _graph.AddGroupMemberAsync(m365GroupId, item.UserId), () => _graph.IsGroupMemberAsync(m365GroupId, item.UserId), true);
                        }
                        else
                        {
                            await PermissionVerification.ApplyAsync(() => _graph.AddGroupOwnerAsync(m365GroupId, item.UserId), () => _graph.IsGroupOwnerAsync(m365GroupId, item.UserId), true);
                        }
                    }
                    else
                    {
                        if (item.Role != "Member" && item.Role != "Owner") throw new InvalidOperationException($"Unknown role '{item.Role}'.");
                        var recipient = await _exo.GetRecipientAsync(item.User) ?? throw new InvalidOperationException("Restore recipient could not be resolved.");
                        if (string.Equals(item.Role, GroupRole.Member.ToString(), StringComparison.OrdinalIgnoreCase))
                        {
                            await PermissionVerification.ApplyAsync(() => _exo.AddDistributionGroupMemberAsync(groupEmail, item.User), () => _exo.IsDistributionGroupMemberAsync(groupEmail, recipient, item.User), true);
                        }
                        else
                        {
                            await PermissionVerification.ApplyAsync(() => _exo.AddDistributionGroupOwnerAsync(groupEmail, item.User), () => _exo.IsDistributionGroupOwnerAsync(groupEmail, recipient, item.User), true);
                        }
                    }

                    Logger.Log($"Restored {item.Role} '{item.User}' to group '{groupEmail}'.", LogType.Success);
                    results.Add(new RestoreItemResult { User = item.User, Role = item.Role, Success = true, Message = "Restored" });
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    Logger.Log($"Failed to restore {item.Role} '{item.User}' to group '{groupEmail}'. DETAILS: {ex.Message}", LogType.Error);
                    results.Add(new RestoreItemResult { User = item.User, Role = item.Role, Success = false, Message = ex.Message });
                }
            }
        }

        private async Task RestoreMailboxPermissionsAsync(SnapshotRecord record, List<RestoreItemResult> results)
        {
            var targetIdentity = record.Target;

            foreach (var item in record.Items)
            {
                try
                {
                    await ApplyRestorePermissionAsync(targetIdentity, item.User, item.Role, record.OperationType == "GroupDelegateRemoval");
                    Logger.Log($"Restored {item.Role} for '{item.User}' on '{targetIdentity}'.", LogType.Success);
                    results.Add(new RestoreItemResult { User = item.User, Role = item.Role, Success = true, Message = "Restored" });
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    Logger.Log($"Failed to restore {item.Role} for '{item.User}' on '{targetIdentity}'. DETAILS: {ex.Message}", LogType.Error);
                    results.Add(new RestoreItemResult { User = item.User, Role = item.Role, Success = false, Message = ex.Message });
                }
            }
        }

        private async Task ApplyRestorePermissionAsync(string targetIdentity, string user, string role, bool isGroup)
        {
            if (isGroup && role != "Send As" && role != "Send on Behalf") throw new InvalidOperationException($"Unsupported group delegate role '{role}'.");
            switch (role)
            {
                case "Full Access":
                    await _exo.AddFullAccessAsync(targetIdentity, user);
                    break;
                case "Send As":
                    await _exo.AddSendAsAsync(targetIdentity, user);
                    break;
                case "Send on Behalf":
                    if (isGroup)
                        await PermissionVerification.ApplyAsync(() => _exo.AddGroupSendOnBehalfAsync(targetIdentity, user), () => _exo.HasGroupSendOnBehalfAsync(targetIdentity, user), true);
                    else
                    {
                        var recipient = await _exo.GetRecipientAsync(user) ?? throw new InvalidOperationException("Restore recipient could not be resolved.");
                        await PermissionVerification.ApplyAsync(() => _exo.AddSendOnBehalfAsync(targetIdentity, user), () => _exo.HasSendOnBehalfAsync(targetIdentity, recipient), true);
                    }
                    break;
                default:
                    throw new InvalidOperationException($"Unknown permission role '{role}'.");
            }
        }
    }
}

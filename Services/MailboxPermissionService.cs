using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EXOKit.Services
{
    public enum PermissionTargetType
    {
        Mailbox,
        Resource
    }

    public enum PermissionOperationType
    {
        Add,
        Remove
    }

    public class PermissionSelections
    {
        public bool FullAccess { get; set; }
        public bool SendAs { get; set; }
        public bool SendOnBehalf { get; set; }
    }

    /// <summary>
    /// Result for a single (User, Target) permission processing pass, keyed the same way as the
    /// script's $results[$user][$targetIdentity] list of status strings (e.g. "Full Access (Added)",
    /// "Send As (Already Exists)", "Send on Behalf (Pending Validation)").
    /// </summary>
    public class PermissionResult
    {
        public string User { get; set; } = string.Empty;
        public string Target { get; set; } = string.Empty;
        public List<string> Statuses { get; } = new();
    }

    /// <summary>
    /// Ports Invoke-PermissionOperation (and its Add/Remove-MailboxPermissions / Add/Remove-ResourcePermissions
    /// wrappers) from the toolkit script: for each target mailbox/resource and each user, adds or
    /// removes Full Access, Send As, and Send on Behalf permissions, batching Send on Behalf changes
    /// into a single Set-Mailbox call per target, then polls once per second for up to 10 seconds to
    /// verify the change actually applied before reporting final status.
    /// </summary>
    public class MailboxPermissionService
    {
        private readonly ExoPowerShellService _exo;
        private readonly SnapshotService _snapshots;

        public MailboxPermissionService(ExoPowerShellService exo, SnapshotService? snapshots = null)
        {
            _exo = exo;
            _snapshots = snapshots ?? new SnapshotService();
        }

        public async Task<List<PermissionResult>> InvokePermissionOperationAsync(
            PermissionOperationType operationType,
            PermissionTargetType targetType,
            IEnumerable<string> targets,
            IEnumerable<string> users,
            PermissionSelections permissions)
        {
            var targetList = targets.ToList();
            var userList = users.ToList();
            var resultMap = new Dictionary<string, Dictionary<string, List<string>>>(StringComparer.OrdinalIgnoreCase);

            var actionVerb = operationType == PermissionOperationType.Add ? "Assignment" : "Removal";
            Logger.Log($"--- Starting {targetType} Permission {actionVerb} ---");

            foreach (var targetIdentity in targetList)
            {
                Logger.Log($"Processing {targetType}: {targetIdentity}");

                bool mailboxExists;
                try { mailboxExists = await _exo.MailboxExistsAsync(targetIdentity, resourceOnly: targetType == PermissionTargetType.Resource); }
                catch (Exception exception) when (exception is not OperationCanceledException)
                {
                    Logger.Log($"Target '{targetIdentity}' lookup failed: {exception.Message}", LogType.Error);
                    foreach (var user in userList)
                    {
                        GetOrCreateResultList(resultMap, user, targetIdentity, out var statuses);
                        statuses.Add("Target Lookup Error");
                    }
                    continue;
                }
                if (!mailboxExists)
                {
                    Logger.Log($"{targetType} '{targetIdentity}' not found. Skipping.", LogType.Error);
                    foreach (var user in userList)
                    {
                        GetOrCreateResultList(resultMap, user, targetIdentity, out var statuses);
                        statuses.Add("Target Not Found");
                    }
                    continue;
                }

                var sendOnBehalfList = new List<string>();
                var userObjectMap = new Dictionary<string, RecipientInfo>(StringComparer.OrdinalIgnoreCase);
                var validationQueue = new List<(string User, string Permission)>();
                var snapshotItems = new List<SnapshotItem>();

                foreach (var user in userList)
                {
                    Logger.Log($"Processing User: {user} for {targetType}: {targetIdentity}");

                    if (!GetOrCreateResultList(resultMap, user, targetIdentity, out var statuses))
                    {
                        // just created; continue
                    }

                    RecipientInfo? userObject;
                    try { userObject = await _exo.GetRecipientAsync(user); }
                    catch (Exception exception) when (exception is not OperationCanceledException)
                    {
                        Logger.Log($"User '{user}' lookup failed: {exception.Message}", LogType.Error);
                        statuses.Add("User Lookup Error");
                        continue;
                    }

                    if (userObject == null)
                    {
                        Logger.Log($"User '{user}' not found. Skipping.", LogType.Error);
                        statuses.Add("User Not Found");
                        continue;
                    }
                    userObjectMap[user] = userObject;

                    if (permissions.FullAccess)
                    {
                        await ProcessFullAccessAsync(operationType, targetIdentity, user, userObject, statuses, validationQueue, snapshotItems);
                    }

                    if (permissions.SendAs)
                    {
                        await ProcessSendAsAsync(operationType, targetIdentity, user, statuses, validationQueue, snapshotItems);
                    }

                    if (permissions.SendOnBehalf)
                    {
                        ProcessSendOnBehalfMark(operationType, user, sendOnBehalfList, statuses);
                    }
                }

                if (permissions.SendOnBehalf && operationType == PermissionOperationType.Remove)
                {
                    await CaptureSendOnBehalfBeforeStateAsync(targetIdentity, sendOnBehalfList, snapshotItems, resultMap);
                }

                if (sendOnBehalfList.Count > 0)
                {
                    await ApplySendOnBehalfBatchAsync(operationType, targetIdentity, sendOnBehalfList, resultMap, validationQueue);
                }

                if (validationQueue.Count > 0)
                {
                    await ValidatePendingChangesAsync(operationType, targetIdentity, validationQueue, userObjectMap, resultMap);
                }

            }

            var summaryResults = resultMap.ToDictionary(
                userEntry => userEntry.Key,
                userEntry => (object)userEntry.Value,
                StringComparer.OrdinalIgnoreCase);
            Logger.WriteSummary($"{targetType} Permission {actionVerb}", summaryResults);

            return resultMap.SelectMany(userEntry => userEntry.Value.Select(targetEntry => new PermissionResult
            {
                User = userEntry.Key,
                Target = targetEntry.Key,
                Statuses = { }
            }.Also(r => r.Statuses.AddRange(targetEntry.Value)))).ToList();
        }

        private static bool GetOrCreateResultList(Dictionary<string, Dictionary<string, List<string>>> resultMap, string user, string target, out List<string> statuses)
        {
            var created = false;
            if (!resultMap.TryGetValue(user, out var targetMap))
            {
                targetMap = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
                resultMap[user] = targetMap;
            }
            if (!targetMap.TryGetValue(target, out var list))
            {
                list = new List<string>();
                targetMap[target] = list;
                created = true;
            }
            statuses = list;
            return created;
        }

        private async Task ProcessFullAccessAsync(
            PermissionOperationType operationType, string targetIdentity, string user, RecipientInfo userObject,
            List<string> statuses, List<(string User, string Permission)> validationQueue, List<SnapshotItem> snapshotItems)
        {
            Logger.Log($"Attempting {operationType} Full Access...");

            // Per Microsoft's own documentation (Manage permissions for recipients in Exchange Online),
            // Full Access is only supported on User mailboxes, Resource mailboxes, Shared mailboxes, and
            // Discovery mailboxes. Microsoft 365 Group mailboxes are NOT a supported target for
            // Add-MailboxPermission / Remove-MailboxPermission at all; against those targets the cmdlets
            // don't just intermittently fail, they reliably throw the "Object reference not set" error
            // because the operation is unsupported, not because of a transient server bug. Check the
            // target's recipient type upfront and fail fast with an actionable message instead of
            // retrying an operation that can never succeed.
            string recipientTypeDetails;
            try { recipientTypeDetails = await _exo.GetRecipientTypeAsync(targetIdentity); }
            catch (Exception exception) when (exception is not OperationCanceledException)
            {
                Logger.Log($"Full Access target lookup failed: {exception.Message}", LogType.Error);
                statuses.Add("Full Access (Error - Target Lookup Failed)");
                return;
            }
            if (string.Equals(recipientTypeDetails, "GroupMailbox", StringComparison.OrdinalIgnoreCase))
            {
                Logger.Log(
                    $"Full Access is not supported on Microsoft 365 Group mailboxes (target '{targetIdentity}' is Group-backed). " +
                    "Use Send As or Send on Behalf instead, or manage this via the group's owners/members.",
                    LogType.Error);
                statuses.Add("Full Access (Not Supported on M365 Group)");
                return;
            }

            try
            {
                var hasFullAccess = await _exo.HasFullAccessAsync(targetIdentity, user, userObject);

                if (operationType == PermissionOperationType.Add)
                {
                    if (hasFullAccess)
                    {
                        Logger.Log("Already granted Full Access.", LogType.Warning);
                        statuses.Add("Full Access (Already Exists)");
                    }
                    else
                    {
                        await _exo.AddFullAccessAsync(targetIdentity, user);
                        Logger.Log("Full Access add command submitted. Pending validation.");
                        statuses.Add("Full Access (Pending Validation)");
                        validationQueue.Add((user, "Full Access"));
                    }
                }
                else
                {
                    if (!hasFullAccess)
                    {
                        statuses.Add("Full Access (Not Found - No Change)");
                        return;
                    }
                    if (hasFullAccess)
                    {
                        snapshotItems.Add(new SnapshotItem { User = user, Role = "Full Access" });
                        SaveBeforeRemoval(targetIdentity, user, "Full Access");
                    }
                    await _exo.RemoveFullAccessAsync(targetIdentity, user);
                    Logger.Log("Full Access remove command submitted. Pending validation.");
                    statuses.Add("Full Access (Pending Validation)");
                    validationQueue.Add((user, "Full Access"));
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                if (IsNullReferenceServerError(ex.Message))
                {
                    Logger.Log($"Full Access {operationType} could not establish permission state after retries. Verify manually.", LogType.Warning);
                    statuses.Add("Full Access (Unconfirmed - Read Failed)");
                }
                else
                {
                    Logger.Log($"Failed to {operationType} Full Access. DETAILS: {ex.Message}", LogType.Error);
                    statuses.Add("Full Access (Error)");
                }
            }
        }

        private static bool IsNullReferenceServerError(string message) =>
            message.Contains("Object reference not set to an instance of an object", StringComparison.OrdinalIgnoreCase);

        private async Task ProcessSendAsAsync(
            PermissionOperationType operationType, string targetIdentity, string user,
            List<string> statuses, List<(string User, string Permission)> validationQueue, List<SnapshotItem> snapshotItems)
        {
            Logger.Log($"Attempting {operationType} Send As...");
            try
            {
                var existing = await _exo.HasSendAsAsync(targetIdentity, user);

                if (operationType == PermissionOperationType.Add)
                {
                    if (existing)
                    {
                        Logger.Log("Already granted Send As.", LogType.Warning);
                        statuses.Add("Send As (Already Exists)");
                    }
                    else
                    {
                        await _exo.AddSendAsAsync(targetIdentity, user);
                        Logger.Log("Send As add command submitted. Pending validation.");
                        statuses.Add("Send As (Pending Validation)");
                        validationQueue.Add((user, "Send As"));
                    }
                }
                else
                {
                    if (!existing)
                    {
                        statuses.Add("Send As (Not Found - No Change)");
                        return;
                    }
                    snapshotItems.Add(new SnapshotItem { User = user, Role = "Send As" });
                    SaveBeforeRemoval(targetIdentity, user, "Send As");

                    await _exo.RemoveSendAsAsync(targetIdentity, user);
                    Logger.Log("Send As remove command submitted. Pending validation.");
                    statuses.Add("Send As (Pending Validation)");
                    validationQueue.Add((user, "Send As"));
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                if (IsNullReferenceServerError(ex.Message))
                {
                    Logger.Log($"Send As {operationType} could not establish permission state after retries. Verify manually.", LogType.Warning);
                    statuses.Add("Send As (Unconfirmed - Read Failed)");
                }
                else
                {
                    Logger.Log($"Failed to {operationType} Send As. DETAILS: {ex.Message}", LogType.Error);
                    statuses.Add("Send As (Error)");
                    snapshotItems.RemoveAll(si => si.Role == "Send As" && string.Equals(si.User, user, StringComparison.OrdinalIgnoreCase));
                }
            }
        }

        private static void ProcessSendOnBehalfMark(
            PermissionOperationType operationType, string user, List<string> sendOnBehalfList, List<string> statuses)
        {
            if (operationType == PermissionOperationType.Add)
            {
                Logger.Log("Checking Send on Behalf...");
                // Existence check happens in ApplySendOnBehalfBatchAsync's caller via HasSendOnBehalfAsync
                // is intentionally deferred here to keep the batching identical to the script, so we just mark.
                if (!sendOnBehalfList.Contains(user, StringComparer.OrdinalIgnoreCase))
                {
                    sendOnBehalfList.Add(user);
                }
                statuses.Add("Send on Behalf (Pending)");
            }
            else
            {
                Logger.Log($"Marking '{user}' for Remove Send on Behalf.");
                if (!sendOnBehalfList.Contains(user, StringComparer.OrdinalIgnoreCase))
                {
                    sendOnBehalfList.Add(user);
                }
                statuses.Add("Send on Behalf (Pending Removal)");
            }
        }

        private async Task CaptureSendOnBehalfBeforeStateAsync(string targetIdentity, List<string> sendOnBehalfList, List<SnapshotItem> snapshotItems,
            Dictionary<string, Dictionary<string, List<string>>> resultMap)
        {
            try
            {
                var existingDelegates = await _exo.GetAllSendOnBehalfDelegatesAsync(targetIdentity);
                var captured = new List<string>();
                foreach (var requestedUser in sendOnBehalfList)
                {
                    var recipient = await _exo.GetRecipientAsync(requestedUser)
                        ?? throw new InvalidOperationException($"Recipient '{requestedUser}' could not be resolved for snapshot.");
                    if (existingDelegates.Contains(recipient.PrimarySmtpAddress ?? requestedUser, StringComparer.OrdinalIgnoreCase))
                    {
                        snapshotItems.Add(new SnapshotItem { User = requestedUser, Role = "Send on Behalf" });
                        SaveBeforeRemoval(targetIdentity, requestedUser, "Send on Behalf");
                        captured.Add(requestedUser);
                    }
                    else
                    {
                        ReplaceStatus(resultMap[requestedUser][targetIdentity], "Send on Behalf (Pending Removal)", "Send on Behalf (Not Found - No Change)");
                    }
                }
                sendOnBehalfList.RemoveAll(user => !captured.Contains(user, StringComparer.OrdinalIgnoreCase));
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                Logger.Log($"Failed to capture Send on Behalf snapshot for '{targetIdentity}'. DETAILS: {ex.Message}", LogType.Warning);
                foreach (var user in sendOnBehalfList)
                    ReplaceStatus(resultMap[user][targetIdentity], "Send on Behalf (Pending Removal)", "Send on Behalf (Error - Before-State Unconfirmed)");
                sendOnBehalfList.Clear();
            }
        }

        private async Task ApplySendOnBehalfBatchAsync(
            PermissionOperationType operationType, string targetIdentity, List<string> sendOnBehalfList,
            Dictionary<string, Dictionary<string, List<string>>> resultMap,
            List<(string User, string Permission)> validationQueue)
        {
            var usersStr = string.Join(", ", sendOnBehalfList);
            Logger.Log($"Applying Send on Behalf {operationType} for '{targetIdentity}' (Users: {usersStr})...");

            try
            {
                if (operationType == PermissionOperationType.Add)
                {
                    await _exo.AddSendOnBehalfBatchAsync(targetIdentity, sendOnBehalfList.ToArray());
                }
                else
                {
                    await _exo.RemoveSendOnBehalfBatchAsync(targetIdentity, sendOnBehalfList.ToArray());
                }

                Logger.Log("Send on Behalf command submitted. Pending validation.");

                foreach (var processedUser in sendOnBehalfList)
                {
                    if (resultMap.TryGetValue(processedUser, out var targetMap) && targetMap.TryGetValue(targetIdentity, out var statuses))
                    {
                        validationQueue.Add((processedUser, "Send on Behalf"));
                        ReplaceStatus(statuses, operationType == PermissionOperationType.Add ? "Send on Behalf (Pending)" : "Send on Behalf (Pending Removal)", "Send on Behalf (Pending Validation)");
                    }
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                Logger.Log($"Failed Send on Behalf {operationType} for '{targetIdentity}'. DETAILS: {ex.Message}", LogType.Error);
                foreach (var failedUser in sendOnBehalfList)
                {
                    if (resultMap.TryGetValue(failedUser, out var targetMap) && targetMap.TryGetValue(targetIdentity, out var statuses))
                    {
                        ReplaceStatus(statuses, operationType == PermissionOperationType.Add ? "Send on Behalf (Pending)" : "Send on Behalf (Pending Removal)", "Send on Behalf (Error)");
                    }
                }
            }
        }

        private async Task ValidatePendingChangesAsync(
            PermissionOperationType operationType, string targetIdentity,
            List<(string User, string Permission)> validationQueue,
            Dictionary<string, RecipientInfo> userObjectMap,
            Dictionary<string, Dictionary<string, List<string>>> resultMap)
        {
            // Exchange Online's Get-* cmdlets (Get-EXOMailboxPermission, Get-RecipientPermission,
            // Get-Mailbox) read from an eventually-consistent cache that can lag noticeably behind the
            // Add/Remove/Set cmdlets that actually mutate the permission, especially against Microsoft
            // 365 Group-backed shared mailboxes. Polling once per second for 10 seconds was too
            // aggressive and produced false "Verify Failed"/"Not Found" results even though the change
            // had landed. Poll less frequently (every 2 seconds) for longer (up to 30 seconds total) to
            // give replication more time to catch up before giving up.
            const int maxAttempts = 15;
            const int delayMs = 2000;
            Logger.Log($"Validating applied permission changes for '{targetIdentity}' every {delayMs / 1000} seconds for up to {maxAttempts * delayMs / 1000} seconds...");
            var pending = new List<(string User, string Permission)>(validationQueue);

            for (var attempt = 1; attempt <= maxAttempts && pending.Count > 0; attempt++)
            {
                await Task.Delay(delayMs, _exo.OperationCancellationToken);
                var remaining = new List<(string User, string Permission)>();

                foreach (var item in pending)
                {
                    if (!resultMap.TryGetValue(item.User, out var targetMap) || !targetMap.TryGetValue(targetIdentity, out var statuses))
                    {
                        continue;
                    }

                    bool isPresent;
                    try
                    {
                        isPresent = item.Permission switch
                        {
                        "Full Access" => await _exo.HasFullAccessAsync(targetIdentity, item.User, userObjectMap.GetValueOrDefault(item.User) ?? new RecipientInfo()),
                        "Send As" => await _exo.HasSendAsAsync(targetIdentity, item.User),
                        "Send on Behalf" => userObjectMap.TryGetValue(item.User, out var uo) && await _exo.HasSendOnBehalfAsync(targetIdentity, uo),
                        _ => false
                        };
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception exception)
                    {
                        Logger.Log($"Verification unavailable for '{item.User}': {exception.Message}", LogType.Warning);
                        remaining.Add(item);
                        continue;
                    }

                    var isValidated = operationType == PermissionOperationType.Add ? isPresent : !isPresent;
                    if (isValidated)
                    {
                        var verb = operationType == PermissionOperationType.Add ? "added" : "removed";
                        var suffix = operationType == PermissionOperationType.Add ? "Added" : "Removed";
                        Logger.Log($"SUCCESS: {item.Permission} {verb} and verified for '{item.User}'.", LogType.Success);
                        ReplaceStatus(statuses, $"{item.Permission} (Pending Validation)", $"{item.Permission} ({suffix})");
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
                if (!resultMap.TryGetValue(item.User, out var targetMap) || !targetMap.TryGetValue(targetIdentity, out var statuses))
                {
                    continue;
                }

                var opWord = operationType == PermissionOperationType.Add ? "add" : "remove";
                var stillOrNo = operationType == PermissionOperationType.Add ? "verification found no matching permission" : "permission still present";
                var timeoutSeconds = maxAttempts * delayMs / 1000;
                // The command was already confirmed to have been submitted successfully (no exception was
                // thrown, or the retry/verify path already handled the known server-side error). Reaching
                // this point only means EXO's read-side cache hadn't caught up within the polling window,
                // not that the change failed. Report it as "Unconfirmed" rather than "Verify Failed" so the
                // ticket note doesn't read as a hard failure when manual verification will likely show the
                // change actually applied.
                Logger.Log($"{item.Permission} {opWord} command succeeded but could not be confirmed for '{item.User}' within {timeoutSeconds} seconds. This is often an Exchange Online replication delay rather than an actual failure - please verify manually before assuming it did not apply.", LogType.Warning);
                var suffix = operationType == PermissionOperationType.Add ? "Add - Unconfirmed, Verify Manually" : "Remove - Unconfirmed, Verify Manually";
                ReplaceStatus(statuses, $"{item.Permission} (Pending Validation)", $"{item.Permission} ({suffix})");
            }
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

        private void SaveBeforeRemoval(string target, string user, string role)
        {
            _snapshots.SaveSnapshot(new SnapshotRecord
            {
                OperationType = "MailboxPermissionRemoval",
                Target = target,
                Description = $"Before removing {role} for '{user}' from '{target}'",
                Items = { new SnapshotItem { User = user, Role = role } }
            });
        }

    }

    internal static class FluentExtensions
    {
        public static T Also<T>(this T self, Action<T> action)
        {
            action(self);
            return self;
        }
    }
}

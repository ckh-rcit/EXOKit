using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace EXOKit.Services
{
    /// <summary>
    /// Microsoft Graph operations for M365 Group membership/ownership and Graph user resolution,
    /// ported from the WinForms script's Get-MgUser / Get-MgGroupMember / New-MgGroupMember /
    /// New-MgGroupOwnerByRef / Remove-MgGroupMemberByRef / Remove-MgGroupOwnerByRef calls, plus the
    /// Bookings license group membership check/add used by Enable-BookingsAccessGUI.
    /// </summary>
    public class GraphService
    {
        private readonly AuthService _authService;
        private readonly string[] _scopes;
        private GraphServiceClient? _graphClient;
        public bool IsConnected => _authService.IsGraphConnected;
        public string? ConnectedTenantId => _authService.ConnectedTenantId;

        public GraphService(AuthService authService, string[] scopes)
        {
            _authService = authService;
            _scopes = scopes;
        }

        internal GraphService(AuthService authService, GraphServiceClient client) : this(authService, Array.Empty<string>()) =>
            _graphClient = client;

        public void InitializeGraphClient()
        {
            var authProvider = new TokenAuthenticationProvider(_authService, _scopes);
            _graphClient = new GraphServiceClient(authProvider);
        }

        private GraphServiceClient Client => _graphClient ?? throw new InvalidOperationException("Graph client not initialized. Connect to Microsoft Graph first.");

        public async Task<List<UserSearchResult>> SearchUsersAsync(string query, bool includeDeleted = false, Action<string>? onProgress = null)
        {
            var parsed = SearchQueryParser.Parse(query) ?? throw new ArgumentException("Enter a name, employee ID, UPN, email, or SAM account.", nameof(query));
            var cancellationToken = _authService.OperationCancellationToken;
            cancellationToken.ThrowIfCancellationRequested();
            var client = Client;
            var select = new[] { "id", "displayName", "userPrincipalName", "mail", "employeeId", "onPremisesSamAccountName",
                "jobTitle", "department", "accountEnabled", "createdDateTime", "businessPhones" };

            async Task<List<UserSearchResult>> ReadUsersAsync(string filter, bool deleted)
            {
                var users = new List<UserSearchResult>();
                var page = deleted
                    ? await client.Directory.DeletedItems.GraphUser.GetAsync(config =>
                    {
                        config.QueryParameters.Filter = filter;
                        config.QueryParameters.Select = select;
                        config.QueryParameters.Top = 50;
                        config.QueryParameters.Count = true;
                        config.Headers.Add("ConsistencyLevel", "eventual");
                    }, cancellationToken)
                    : await client.Users.GetAsync(config =>
                    {
                        config.QueryParameters.Filter = filter;
                        config.QueryParameters.Select = select;
                        config.QueryParameters.Top = 50;
                        config.QueryParameters.Count = true;
                        config.Headers.Add("ConsistencyLevel", "eventual");
                    }, cancellationToken);

                while (true)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (page?.Value == null) throw new InvalidOperationException("Graph returned an unreadable user search response. Search is incomplete.");
                    foreach (var user in page.Value)
                    {
                        if (string.IsNullOrWhiteSpace(user.Id)) throw new InvalidOperationException("Graph returned a user without an object ID. Search is incomplete.");
                        users.Add(new UserSearchResult
                        {
                            ObjectId = user.Id,
                            DisplayName = user.DisplayName ?? string.Empty,
                            UserPrincipalName = user.UserPrincipalName ?? string.Empty,
                            Mail = user.Mail ?? string.Empty,
                            EmployeeId = user.EmployeeId ?? string.Empty,
                            SamAccountName = user.OnPremisesSamAccountName ?? string.Empty,
                            JobTitle = user.JobTitle ?? string.Empty,
                            Department = user.Department ?? string.Empty,
                            Enabled = user.AccountEnabled,
                            CreatedDateTime = user.CreatedDateTime,
                            BusinessPhone = string.Join("; ", user.BusinessPhones ?? new List<string>()),
                            IsDeleted = deleted
                        });
                    }
                    onProgress?.Invoke($"Searching {(deleted ? "deleted" : "active")} users: {users.Count} matches read...");
                    if (string.IsNullOrEmpty(page.OdataNextLink)) return users;
                    page = deleted
                        ? await client.Directory.DeletedItems.GraphUser.WithUrl(page.OdataNextLink).GetAsync(config => config.Headers.Add("ConsistencyLevel", "eventual"), cancellationToken)
                        : await client.Users.WithUrl(page.OdataNextLink).GetAsync(config => config.Headers.Add("ConsistencyLevel", "eventual"), cancellationToken);
                }
            }

            async Task<List<UserSearchResult>> SearchScopeAsync(bool deleted)
            {
                var users = await ReadUsersAsync(parsed.Filter, deleted);
                if (users.Count == 0 && parsed.FallbackFilter != null)
                    users = await ReadUsersAsync(parsed.FallbackFilter, deleted);
                return users;
            }

            Logger.Log($"User search: {parsed.Description}{(includeDeleted ? ", including deleted users" : "") }.");
            var results = await SearchScopeAsync(false);
            if (includeDeleted) results.AddRange(await SearchScopeAsync(true));
            cancellationToken.ThrowIfCancellationRequested();
            return results.DistinctBy(user => user.ObjectId).OrderBy(user => user.DisplayName, StringComparer.OrdinalIgnoreCase).ToList();
        }

        public async Task VerifyConnectionAsync()
        {
            var response = await Client.Users.GetAsync(config =>
            {
                config.QueryParameters.Select = new[] { "id" };
                config.QueryParameters.Top = 1;
            }, _authService.OperationCancellationToken);
            if (response?.Value == null) throw new InvalidOperationException("Graph returned no directory response.");
        }

        public async Task<string?> GetUserIdAsync(string userIdentity)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(userIdentity);
            try
            {
                if (Guid.TryParse(userIdentity, out _))
                {
                    var user = await Client.Users[userIdentity].GetAsync(config =>
                        config.QueryParameters.Select = new[] { "id" }, _authService.OperationCancellationToken);
                    return user?.Id ?? throw new InvalidOperationException("Graph returned no user identity.");
                }
                var escaped = userIdentity.Replace("'", "''");
                var response = await Client.Users.GetAsync(config =>
                {
                    config.QueryParameters.Filter = $"userPrincipalName eq '{escaped}' or mail eq '{escaped}' or proxyAddresses/any(address:address eq 'smtp:{escaped}')";
                    config.QueryParameters.Select = new[] { "id" };
                    config.QueryParameters.Top = 2;
                }, _authService.OperationCancellationToken);
                if (response?.Value == null) throw new InvalidOperationException("Graph returned no user lookup response.");
                if (response.OdataNextLink != null || response.Value.Count > 1)
                    throw new InvalidOperationException("User identity is ambiguous; use the directory object ID.");
                if (response.Value.Count == 0) return null;
                return response.Value[0].Id ?? throw new InvalidOperationException("Graph returned no user identity.");
            }
            catch (ODataError odataEx)
            {
                Logger.Log($"Graph user lookup failed for '{userIdentity}': {odataEx.Error?.Code} - {odataEx.Error?.Message}", LogType.Error);
                if (odataEx.ResponseStatusCode == 404) return null;
                throw;
            }
            catch (Exception ex)
            {
                Logger.Log($"Graph user lookup failed for '{userIdentity}': {ex.Message}", LogType.Error);
                throw;
            }
        }

        public async Task<bool> IsGroupMemberAsync(string groupId, string userId)
        {
            try
            {
                var response = await Client.Groups[groupId].Members.GetAsync(config => config.QueryParameters.Select = new[] { "id" }, _authService.OperationCancellationToken);
                while (response != null)
                {
                    if (response.Value == null) throw new InvalidOperationException("Graph returned an unreadable membership list.");
                    if (response.Value.Any(member => member.Id == userId)) return true;
                    if (string.IsNullOrEmpty(response.OdataNextLink)) return false;
                    response = await Client.Groups[groupId].Members.WithUrl(response.OdataNextLink).GetAsync(cancellationToken: _authService.OperationCancellationToken);
                }
                throw new InvalidOperationException("Graph returned no membership response.");
            }
            catch (ODataError odataEx)
            {
                Logger.Log($"Graph group member check failed: {odataEx.Error?.Code} - {odataEx.Error?.Message}", LogType.Error);
                throw;
            }
            catch (Exception ex)
            {
                Logger.Log($"Graph group member check failed: {ex.Message}", LogType.Error);
                throw;
            }
        }

        public async Task<bool> IsGroupOwnerAsync(string groupId, string userId)
        {
            try
            {
                var response = await Client.Groups[groupId].Owners.GetAsync(config => config.QueryParameters.Select = new[] { "id" }, _authService.OperationCancellationToken);
                while (response != null)
                {
                    if (response.Value == null) throw new InvalidOperationException("Graph returned an unreadable ownership list.");
                    if (response.Value.Any(owner => owner.Id == userId)) return true;
                    if (string.IsNullOrEmpty(response.OdataNextLink)) return false;
                    response = await Client.Groups[groupId].Owners.WithUrl(response.OdataNextLink).GetAsync(cancellationToken: _authService.OperationCancellationToken);
                }
                throw new InvalidOperationException("Graph returned no ownership response.");
            }
            catch (ODataError odataEx)
            {
                Logger.Log($"Graph group owner check failed: {odataEx.Error?.Code} - {odataEx.Error?.Message}", LogType.Error);
                throw;
            }
            catch (Exception ex)
            {
                Logger.Log($"Graph group owner check failed: {ex.Message}", LogType.Error);
                throw;
            }
        }

        public async Task AddGroupMemberAsync(string groupId, string userId)
        {
            var requestBody = new ReferenceCreate
            {
                OdataId = $"https://graph.microsoft.com/v1.0/directoryObjects/{userId}"
            };
            await Client.Groups[groupId].Members.Ref.PostAsync(requestBody, cancellationToken: _authService.OperationCancellationToken);
        }

        public async Task RemoveGroupMemberAsync(string groupId, string userId)
        {
            await Client.Groups[groupId].Members[userId].Ref.DeleteAsync(cancellationToken: _authService.OperationCancellationToken);
        }

        public async Task AddGroupOwnerAsync(string groupId, string userId)
        {
            var requestBody = new ReferenceCreate
            {
                OdataId = $"https://graph.microsoft.com/v1.0/directoryObjects/{userId}"
            };
            await Client.Groups[groupId].Owners.Ref.PostAsync(requestBody, cancellationToken: _authService.OperationCancellationToken);
        }

        public async Task RemoveGroupOwnerAsync(string groupId, string userId)
        {
            await Client.Groups[groupId].Owners[userId].Ref.DeleteAsync(cancellationToken: _authService.OperationCancellationToken);
        }

        /// <summary>
        /// Resolves an M365 (Unified) group's id and RecipientTypeDetails-equivalent classification
        /// by mail nickname / mail / id, mirroring Resolve-GroupOperationContext's M365 branch.
        /// Returns null if not found or not a Unified (M365) group.
        /// </summary>
        public async Task<string?> ResolveM365GroupIdAsync(string groupIdentity)
        {
            try
            {
                var escaped = groupIdentity.Replace("'", "''");
                var isGuid = Guid.TryParse(groupIdentity, out _);

                var filter = isGuid
                    ? $"id eq '{escaped}' or mail eq '{escaped}' or mailNickname eq '{escaped}'"
                    : $"mail eq '{escaped}' or mailNickname eq '{escaped}'";

                var response = await Client.Groups.GetAsync(config =>
                {
                    config.QueryParameters.Filter = filter;
                    config.QueryParameters.Select = new[] { "id", "groupTypes", "mail" };
                    config.Headers.Add("ConsistencyLevel", "eventual");
                    config.QueryParameters.Count = true;
                }, _authService.OperationCancellationToken);

                if (response?.Value == null) throw new InvalidOperationException("Graph returned no group lookup response.");
                if (response?.OdataNextLink != null || response?.Value?.Count > 1) throw new InvalidOperationException("Group identity is ambiguous; use its object ID.");
                var group = response?.Value?.SingleOrDefault(g => g.GroupTypes != null && g.GroupTypes.Contains("Unified"));
                return group?.Id;
            }
            catch (ODataError odataEx)
            {
                Logger.Log($"Graph group lookup failed for '{groupIdentity}': {odataEx.Error?.Code} - {odataEx.Error?.Message}", LogType.Error);
                throw;
            }
            catch (Exception ex)
            {
                Logger.Log($"Graph group lookup failed for '{groupIdentity}': {ex.Message}", LogType.Error);
                throw;
            }
        }

        /// <summary>
        /// Finds the Bookings license group by name (Settings.LicenseGroups.Bookings.GroupName)
        /// and returns its group id, used by BookingsService to add users to the license group.
        /// </summary>
        public async Task<string?> GetGroupIdByNameAsync(string groupName)
        {
            try
            {
                var escaped = groupName.Replace("'", "''");
                var response = await Client.Groups.GetAsync(config =>
                {
                    config.QueryParameters.Filter = $"displayName eq '{escaped}'";
                    config.QueryParameters.Select = new[] { "id", "displayName" };
                }, _authService.OperationCancellationToken);
                if (response?.Value == null) throw new InvalidOperationException("Graph returned no group lookup response.");
                if (response?.OdataNextLink != null || response?.Value?.Count > 1)
                    throw new InvalidOperationException("Multiple groups share this display name. Configure the Bookings GroupId instead.");
                return response?.Value?.SingleOrDefault()?.Id;
            }
            catch (ODataError odataEx)
            {
                Logger.Log($"Graph group lookup failed for '{groupName}': {odataEx.Error?.Code} - {odataEx.Error?.Message}", LogType.Error);
                throw;
            }
            catch (Exception ex)
            {
                Logger.Log($"Graph group lookup failed for '{groupName}': {ex.Message}", LogType.Error);
                throw;
            }
        }

        /// <summary>
        /// Provisions a Microsoft Teams team on top of an existing Microsoft 365 Group, mirroring the
        /// "Create a team for this group" option in the Microsoft 365 admin center's group creation
        /// wizard (PUT /groups/{id}/team). Newly created groups may require a few seconds for Graph
        /// directory replication before the team can be provisioned, so 404/NotFound responses are
        /// retried with linear backoff.
        /// </summary>
        public async Task CreateTeamFromGroupAsync(string groupId)
        {
            try
            {
                if (await Client.Groups[groupId].Team.GetAsync(cancellationToken: _authService.OperationCancellationToken) != null) return;
            }
            catch (ODataError exception) when (exception.ResponseStatusCode == 404) { }
            const int maxRetries = 3;
            var retryDelaysMs = new[] { 10000, 10000, 10000 };

            for (var attempt = 0; attempt <= maxRetries; attempt++)
            {
                try
                {
                    var team = await Client.Groups[groupId].Team.PutAsync(new Team(), cancellationToken: _authService.OperationCancellationToken);
                    if (team == null) throw new InvalidOperationException("Team creation returned no confirmation; verify before retrying.");
                    return;
                }
                catch (ODataError odataEx) when (attempt < maxRetries && IsTransientGroupNotFoundError(odataEx))
                {
                    var delay = retryDelaysMs[attempt];
                    Logger.Log($"Group '{groupId}' is still replicating in Microsoft Graph ({odataEx.Error?.Code}). Retrying Team creation in {delay / 1000}s (attempt {attempt + 1}/{maxRetries})...", LogType.Warning);
                    await Task.Delay(delay, _authService.OperationCancellationToken);
                }
                catch (ODataError odataEx)
                {
                    Logger.Log($"Failed to create Team for group '{groupId}': {odataEx.Error?.Code} - {odataEx.Error?.Message}. If the group is new, wait at least 15 minutes after creation and use Resume Team Provisioning for this group ID.", LogType.Error);
                    throw;
                }
            }
        }

        private static bool IsTransientGroupNotFoundError(ODataError error)
        {
            var code = error.Error?.Code;
            var message = error.Error?.Message;
            return string.Equals(code, "NotFound", StringComparison.OrdinalIgnoreCase)
                || string.Equals(code, "ResourceNotFound", StringComparison.OrdinalIgnoreCase)
                || string.Equals(code, "Request_ResourceNotFound", StringComparison.OrdinalIgnoreCase)
                || (message != null && message.Contains("not found", StringComparison.OrdinalIgnoreCase));
        }
    }
}

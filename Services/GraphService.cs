using System;
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

        public void InitializeGraphClient()
        {
            var authProvider = new TokenAuthenticationProvider(_authService, _scopes);
            _graphClient = new GraphServiceClient(authProvider);
        }

        private GraphServiceClient Client => _graphClient ?? throw new InvalidOperationException("Graph client not initialized. Connect to Microsoft Graph first.");

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
            try
            {
                var user = await Client.Users[userIdentity].GetAsync(config =>
                {
                    config.QueryParameters.Select = new[] { "id" };
                });
                return user?.Id;
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
                var response = await Client.Groups[groupId].Members.GetAsync(config => config.QueryParameters.Select = new[] { "id" });
                while (response != null)
                {
                    if (response.Value?.Any(member => member.Id == userId) == true) return true;
                    if (string.IsNullOrEmpty(response.OdataNextLink)) return false;
                    response = await Client.Groups[groupId].Members.WithUrl(response.OdataNextLink).GetAsync();
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
                var response = await Client.Groups[groupId].Owners.GetAsync(config => config.QueryParameters.Select = new[] { "id" });
                while (response != null)
                {
                    if (response.Value?.Any(owner => owner.Id == userId) == true) return true;
                    if (string.IsNullOrEmpty(response.OdataNextLink)) return false;
                    response = await Client.Groups[groupId].Owners.WithUrl(response.OdataNextLink).GetAsync();
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
            await Client.Groups[groupId].Members.Ref.PostAsync(requestBody);
        }

        public async Task RemoveGroupMemberAsync(string groupId, string userId)
        {
            await Client.Groups[groupId].Members[userId].Ref.DeleteAsync();
        }

        public async Task AddGroupOwnerAsync(string groupId, string userId)
        {
            var requestBody = new ReferenceCreate
            {
                OdataId = $"https://graph.microsoft.com/v1.0/directoryObjects/{userId}"
            };
            await Client.Groups[groupId].Owners.Ref.PostAsync(requestBody);
        }

        public async Task RemoveGroupOwnerAsync(string groupId, string userId)
        {
            await Client.Groups[groupId].Owners[userId].Ref.DeleteAsync();
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
                });

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
                });
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
                if (await Client.Groups[groupId].Team.GetAsync() != null) return;
            }
            catch (ODataError exception) when (exception.ResponseStatusCode == 404) { }
            const int maxRetries = 3;
            var retryDelaysMs = new[] { 10000, 10000, 10000 };

            for (var attempt = 0; attempt <= maxRetries; attempt++)
            {
                try
                {
                    await Client.Groups[groupId].Team.PutAsync(new Team());
                    return;
                }
                catch (ODataError odataEx) when (attempt < maxRetries && IsTransientGroupNotFoundError(odataEx))
                {
                    var delay = retryDelaysMs[attempt];
                    Logger.Log($"Group '{groupId}' is still replicating in Microsoft Graph ({odataEx.Error?.Code}). Retrying Team creation in {delay / 1000}s (attempt {attempt + 1}/{maxRetries})...", LogType.Warning);
                    await Task.Delay(delay);
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

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
                return null;
            }
        }

        public async Task<bool> IsGroupMemberAsync(string groupId, string userId)
        {
            try
            {
                var response = await Client.Groups[groupId].Members.GraphUser.GetAsync(config =>
                {
                    config.QueryParameters.Filter = $"id eq '{userId}'";
                    config.QueryParameters.Count = true;
                    config.Headers.Add("ConsistencyLevel", "eventual");
                });
                return (response?.OdataCount ?? 0) > 0 || (response?.Value?.Count ?? 0) > 0;
            }
            catch (ODataError odataEx)
            {
                Logger.Log($"Graph group member check failed: {odataEx.Error?.Code} - {odataEx.Error?.Message}", LogType.Error);
                return false;
            }
        }

        public async Task<bool> IsGroupOwnerAsync(string groupId, string userId)
        {
            try
            {
                var response = await Client.Groups[groupId].Owners.GraphUser.GetAsync(config =>
                {
                    config.QueryParameters.Filter = $"id eq '{userId}'";
                    config.QueryParameters.Count = true;
                    config.Headers.Add("ConsistencyLevel", "eventual");
                });
                return (response?.OdataCount ?? 0) > 0 || (response?.Value?.Count ?? 0) > 0;
            }
            catch (ODataError odataEx)
            {
                Logger.Log($"Graph group owner check failed: {odataEx.Error?.Code} - {odataEx.Error?.Message}", LogType.Error);
                return false;
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
                var response = await Client.Groups.GetAsync(config =>
                {
                    config.QueryParameters.Filter = $"mail eq '{groupIdentity}' or mailNickname eq '{groupIdentity}' or id eq '{groupIdentity}'";
                    config.QueryParameters.Select = new[] { "id", "groupTypes", "mail" };
                    config.Headers.Add("ConsistencyLevel", "eventual");
                    config.QueryParameters.Count = true;
                });

                var group = response?.Value?.FirstOrDefault(g => g.GroupTypes != null && g.GroupTypes.Contains("Unified"));
                return group?.Id;
            }
            catch (ODataError odataEx)
            {
                Logger.Log($"Graph group lookup failed for '{groupIdentity}': {odataEx.Error?.Code} - {odataEx.Error?.Message}", LogType.Error);
                return null;
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
                var response = await Client.Groups.GetAsync(config =>
                {
                    config.QueryParameters.Filter = $"displayName eq '{groupName}'";
                    config.QueryParameters.Select = new[] { "id", "displayName" };
                });
                return response?.Value?.FirstOrDefault()?.Id;
            }
            catch (ODataError odataEx)
            {
                Logger.Log($"Graph group lookup failed for '{groupName}': {odataEx.Error?.Code} - {odataEx.Error?.Message}", LogType.Error);
                return null;
            }
        }

        /// <summary>
        /// Provisions a Microsoft Teams team on top of an existing Microsoft 365 Group, mirroring the
        /// "Create a team for this group" option in the Microsoft 365 admin center's group creation
        /// wizard (PUT /groups/{id}/team). The group must already exist before this call.
        /// </summary>
        public async Task CreateTeamFromGroupAsync(string groupId)
        {
            try
            {
                await Client.Groups[groupId].Team.PutAsync(new Team());
            }
            catch (ODataError odataEx)
            {
                Logger.Log($"Failed to create Team for group '{groupId}': {odataEx.Error?.Code} - {odataEx.Error?.Message}", LogType.Error);
                throw;
            }
        }
    }
}

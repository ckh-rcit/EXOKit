using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

namespace EXOKit.Services
{
    /// <summary>
    /// Wraps MSAL.NET interactive public-client authentication for Microsoft Graph, ported from
    /// Entra Scout's AuthService. The toolkit connects to Exchange Online first (see
    /// ExoPowerShellService), then Microsoft Graph using the scopes configured in config.json
    /// (Settings.GraphApi.Scopes), matching the script's enforced EXO-then-Graph connection order.
    /// </summary>
    public class AuthService
    {
        private readonly string[] _graphScopes;
        private readonly IPublicClientApplication? _app;
        private readonly string _tenantId;
        public string? ConnectedTenantId { get; private set; }
        public System.Threading.CancellationToken OperationCancellationToken { get; set; }
        private readonly Func<IntPtr> _parentWindowHandleProvider;

        public bool IsGraphConnected { get; private set; }
        public string? ConnectedUser { get; private set; }

        public AuthService(string[] graphScopes, Func<IntPtr> parentWindowHandleProvider, string clientId, string tenantId)
        {
            _tenantId = tenantId;
            _graphScopes = graphScopes;
            _parentWindowHandleProvider = parentWindowHandleProvider;
            if (!Guid.TryParse(clientId, out _) || !Guid.TryParse(tenantId, out _)) return;
            _app = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority($"https://login.microsoftonline.com/{tenantId}")
                .WithDefaultRedirectUri()
                .Build();
        }

        public async Task<bool> ConnectGraphAsync()
        {
            try
            {
                Logger.Log("Connecting to Microsoft Graph...");
                var result = await AcquireTokenAsync(_graphScopes);
                IsGraphConnected = result != null;
                if (result != null)
                {
                    if (!string.Equals(result.TenantId, _tenantId, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("Graph token tenant does not match the configured tenant.");
                    ConnectedTenantId = result.TenantId;
                    ConnectedUser = result.Account?.Username;
                    Logger.Log($"Connected to Microsoft Graph as {ConnectedUser}.", LogType.Success);
                }
                return IsGraphConnected;
            }
            catch (Exception ex)
            {
                IsGraphConnected = false;
                Logger.Log($"Graph connection failed: {ex.Message}", LogType.Error);
                return false;
            }
        }

        public async Task DisconnectGraphAsync()
        {
            try
            {
                var accounts = _app == null ? Array.Empty<IAccount>() : await _app.GetAccountsAsync();
                foreach (var account in accounts)
                {
                    await _app!.RemoveAsync(account);
                }
                IsGraphConnected = false;
                ConnectedUser = null;
                ConnectedTenantId = null;
                Logger.Log("Disconnected from Microsoft Graph.");
            }
            catch (Exception ex)
            {
                Logger.Log($"Graph disconnect error: {ex.Message}", LogType.Error);
            }
            finally
            {
                IsGraphConnected = false;
                ConnectedUser = null;
                ConnectedTenantId = null;
            }
        }

        public async Task<string?> GetAccessTokenAsync(string[]? scopes = null, string? claims = null, System.Threading.CancellationToken cancellationToken = default)
        {
            OperationCancellationToken.ThrowIfCancellationRequested();
            if (!IsGraphConnected) throw new InvalidOperationException("Connect Microsoft Graph first.");
            var result = await AcquireTokenAsync(scopes ?? _graphScopes, claims, cancellationToken);
            if (!string.Equals(result?.TenantId, _tenantId, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Graph token tenant mismatch.");
            return result?.AccessToken;
        }

        private async Task<AuthenticationResult?> AcquireTokenAsync(string[] scopes, string? claims = null, System.Threading.CancellationToken cancellationToken = default)
        {
            using var timeout = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(OperationCancellationToken, cancellationToken);
            timeout.CancelAfter(TimeSpan.FromMinutes(5));
            if (_app == null) throw new InvalidOperationException("Configure an EXOKit-owned Graph ClientId and TenantId in Settings first.");
            var accounts = await _app.GetAccountsAsync();
            var account = accounts.FirstOrDefault();

            if (account != null && string.IsNullOrWhiteSpace(claims))
            {
                try
                {
                    return await _app.AcquireTokenSilent(scopes, account).ExecuteAsync(timeout.Token);
                }
                catch (MsalUiRequiredException exception)
                {
                    claims = exception.Claims;
                }
            }

            return await _app.AcquireTokenInteractive(scopes)
                .WithClaims(claims)
                .WithParentActivityOrWindow(_parentWindowHandleProvider())
                .ExecuteAsync(timeout.Token);
        }
    }
}

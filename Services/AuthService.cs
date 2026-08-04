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
        // Microsoft Graph PowerShell well-known public client ID (multi-tenant, pre-registered redirect URIs).
        private const string ClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
        private const string Authority = "https://login.microsoftonline.com/organizations";

        private readonly string[] _graphScopes;
        private readonly IPublicClientApplication _app;
        private readonly Func<IntPtr> _parentWindowHandleProvider;

        public bool IsGraphConnected { get; private set; }
        public string? ConnectedUser { get; private set; }

        public AuthService(string[] graphScopes, Func<IntPtr> parentWindowHandleProvider)
        {
            _graphScopes = graphScopes;
            _parentWindowHandleProvider = parentWindowHandleProvider;
            _app = PublicClientApplicationBuilder.Create(ClientId)
                .WithAuthority(Authority)
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
                    ConnectedUser = result.Account?.Username;
                    Logger.Log($"Connected to Microsoft Graph as {ConnectedUser}.", LogType.Success);
                }
                return IsGraphConnected;
            }
            catch (Exception ex)
            {
                Logger.Log($"Graph connection failed: {ex.Message}", LogType.Error);
                return false;
            }
        }

        public async Task DisconnectGraphAsync()
        {
            try
            {
                var accounts = await _app.GetAccountsAsync();
                foreach (var account in accounts)
                {
                    await _app.RemoveAsync(account);
                }
                IsGraphConnected = false;
                ConnectedUser = null;
                Logger.Log("Disconnected from Microsoft Graph.");
            }
            catch (Exception ex)
            {
                Logger.Log($"Graph disconnect error: {ex.Message}", LogType.Error);
            }
        }

        public async Task<string?> GetAccessTokenAsync(string[]? scopes = null)
        {
            var result = await AcquireTokenAsync(scopes ?? _graphScopes);
            return result?.AccessToken;
        }

        private async Task<AuthenticationResult?> AcquireTokenAsync(string[] scopes)
        {
            var accounts = await _app.GetAccountsAsync();
            var account = accounts.FirstOrDefault();

            if (account != null)
            {
                try
                {
                    return await _app.AcquireTokenSilent(scopes, account).ExecuteAsync();
                }
                catch (MsalUiRequiredException)
                {
                    // Fall through to interactive - additional scope consent likely required.
                }
            }

            return await _app.AcquireTokenInteractive(scopes)
                .WithParentActivityOrWindow(_parentWindowHandleProvider())
                .ExecuteAsync();
        }
    }
}

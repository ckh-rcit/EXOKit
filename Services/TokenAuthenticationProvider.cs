using System.Threading;
using System.Threading.Tasks;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Authentication;

namespace EXOKit.Services
{
    /// <summary>
    /// Kiota authentication provider that fetches an MSAL access token (silent-then-interactive)
    /// for a fixed set of scopes and attaches it to outgoing Graph requests. Ported from Entra Scout.
    /// </summary>
    public class TokenAuthenticationProvider : IAuthenticationProvider
    {
        private readonly AuthService _authService;
        private readonly string[] _scopes;

        public TokenAuthenticationProvider(AuthService authService, string[] scopes)
        {
            _authService = authService;
            _scopes = scopes;
        }

        public async Task AuthenticateRequestAsync(
            RequestInformation request,
            System.Collections.Generic.Dictionary<string, object>? additionalAuthenticationContext = null,
            CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var uri = request.URI;
            if (uri.Scheme != "https" || !string.Equals(uri.Host, "graph.microsoft.com", System.StringComparison.OrdinalIgnoreCase))
                throw new System.InvalidOperationException("Refusing to send a Graph token to an untrusted endpoint.");
            string? claims = null;
            if (additionalAuthenticationContext?.TryGetValue("claims", out var claimValue) == true && claimValue is string encodedClaims)
            {
                claims = System.Text.Encoding.UTF8.GetString(System.Convert.FromBase64String(encodedClaims));
                using var document = System.Text.Json.JsonDocument.Parse(claims);
                if (document.RootElement.ValueKind != System.Text.Json.JsonValueKind.Object) throw new System.InvalidOperationException("Invalid authentication claims challenge.");
            }
            var token = await _authService.GetAccessTokenAsync(_scopes, claims, cancellationToken);
            if (string.IsNullOrEmpty(token)) throw new System.InvalidOperationException("Graph authentication did not return a token.");
            if (!string.IsNullOrEmpty(token))
            {
                request.Headers.Remove("Authorization");
                request.Headers.Add("Authorization", $"Bearer {token}");
            }
        }
    }
}

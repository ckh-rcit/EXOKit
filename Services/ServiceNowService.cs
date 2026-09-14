using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Linq;
using System.Threading.Tasks;
using Azure.Identity;
using Azure.Security.KeyVault.Secrets;

namespace EXOKit.Services
{
    public class ServiceNowResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
    }

    /// <summary>
    /// Ports Get-KeyVaultSecret and Close-ServiceNowTask: retrieves the ServiceNow service-account
    /// credentials from Azure Key Vault, looks up a ticket by number, and updates/closes it with
    /// work notes and additional comments. Uses interactive Azure AD auth (InteractiveBrowserCredential)
    /// for Key Vault access instead of the script's Connect-AzAccount/Az.Accounts dependency.
    /// </summary>
    public class ServiceNowService
    {
        private readonly ServiceNowConfig _config;
        private static readonly HttpClient _httpClient = new(new HttpClientHandler { AllowAutoRedirect = false }) { Timeout = TimeSpan.FromSeconds(60) };

        private readonly InteractiveBrowserCredential? _keyVaultCredential;

        public ServiceNowService(ServiceNowConfig config, GraphApiConfig authentication)
        {
            _config = config;
            if (Guid.TryParse(authentication.ClientId, out _) && Guid.TryParse(authentication.TenantId, out _))
                _keyVaultCredential = new InteractiveBrowserCredential(new InteractiveBrowserCredentialOptions
                {
                    ClientId = authentication.ClientId,
                    TenantId = authentication.TenantId,
                    RedirectUri = new Uri("http://localhost"),
                    TokenCachePersistenceOptions = new TokenCachePersistenceOptions { Name = $"EXOKit.KeyVault.{authentication.TenantId}.{authentication.ClientId}" }
                });
        }

        private async Task<string> GetKeyVaultSecretAsync(string vaultUrl, string secretName)
        {
            if (_keyVaultCredential == null) throw new InvalidOperationException("Configure the EXOKit app registration and tenant before using Key Vault.");
            var client = new SecretClient(new Uri(vaultUrl), _keyVaultCredential);
            using var timeout = new System.Threading.CancellationTokenSource(TimeSpan.FromMinutes(5));
            var secret = await client.GetSecretAsync(secretName, cancellationToken: timeout.Token);
            return secret.Value.Value;
        }

        /// <summary>
        /// Validates that the configured ServiceNow instance URL and Key Vault secrets are reachable
        /// and correct, without making any destructive changes. Used by the Settings page to give the
        /// user immediate feedback when saving ServiceNow configuration.
        /// </summary>
        public async Task<ServiceNowResult> TestConnectionAsync()
        {
            try { ValidateConfiguration(_config); }
            catch (Exception exception) { return new ServiceNowResult { Message = exception.Message }; }
            if (string.IsNullOrWhiteSpace(_config.InstanceUrl) || string.IsNullOrWhiteSpace(_config.KeyVaultUrl) ||
                string.IsNullOrWhiteSpace(_config.UsernameSecretName) || string.IsNullOrWhiteSpace(_config.PasswordSecretName))
            {
                return new ServiceNowResult { Success = false, Message = "Instance URL, Key Vault URL, Username Secret Name, and Password Secret Name are all required." };
            }

            if (!Uri.TryCreate(_config.InstanceUrl, UriKind.Absolute, out var instanceUri) || instanceUri.Scheme != Uri.UriSchemeHttps)
            {
                return new ServiceNowResult { Success = false, Message = $"Instance URL '{_config.InstanceUrl}' is not a valid absolute http(s) URL." };
            }

            if (!Uri.TryCreate(_config.KeyVaultUrl, UriKind.Absolute, out var keyVaultUri) || keyVaultUri.Scheme != Uri.UriSchemeHttps)
            {
                return new ServiceNowResult { Success = false, Message = $"Key Vault URL '{_config.KeyVaultUrl}' is not a valid absolute http(s) URL." };
            }

            string snUser, snPass;
            try
            {
                snUser = await GetKeyVaultSecretAsync(_config.KeyVaultUrl, _config.UsernameSecretName);
                snPass = await GetKeyVaultSecretAsync(_config.KeyVaultUrl, _config.PasswordSecretName);
            }
            catch (Exception ex)
            {
                return new ServiceNowResult { Success = false, Message = $"Failed to retrieve ServiceNow credentials from Key Vault '{_config.KeyVaultUrl}': {ex.Message}" };
            }

            if (string.IsNullOrEmpty(snUser) || string.IsNullOrEmpty(snPass))
            {
                return new ServiceNowResult { Success = false, Message = "Retrieved ServiceNow username or password secret is empty." };
            }

            var encodedCreds = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{snUser}:{snPass}"));
            var instanceUrl = _config.InstanceUrl.TrimEnd('/');
            var table = string.IsNullOrWhiteSpace(_config.Table) ? "task" : _config.Table;

            try
            {
                var testUrl = $"{instanceUrl}/api/now/table/{table}?sysparm_limit=1";
                using var request = new HttpRequestMessage(HttpMethod.Get, testUrl);
                request.Headers.Authorization = new AuthenticationHeaderValue("Basic", encodedCreds);
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                using var response = await _httpClient.SendAsync(request);

                if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized || response.StatusCode == System.Net.HttpStatusCode.Forbidden)
                {
                    return new ServiceNowResult { Success = false, Message = $"ServiceNow rejected the credentials retrieved from Key Vault ({(int)response.StatusCode} {response.StatusCode})." };
                }

                if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
                {
                    return new ServiceNowResult { Success = false, Message = $"ServiceNow table '{table}' or instance URL '{instanceUrl}' could not be found (404)." };
                }

                if (!response.IsSuccessStatusCode)
                {
                    return new ServiceNowResult { Success = false, Message = $"ServiceNow returned an unexpected status code: {(int)response.StatusCode} {response.StatusCode}." };
                }

                return new ServiceNowResult { Success = true, Message = "Successfully connected to ServiceNow and validated credentials." };
            }
            catch (HttpRequestException ex)
            {
                return new ServiceNowResult { Success = false, Message = $"Could not reach ServiceNow instance '{instanceUrl}': {ex.Message}" };
            }
            catch (Exception ex)
            {
                return new ServiceNowResult { Success = false, Message = $"ServiceNow connection test failed: {ex.Message}" };
            }
        }

        public async Task<ServiceNowResult> CloseTaskAsync(string ticketNumber, string workNotes = "", string additionalComments = "")
        {
            try
            {
                ValidateConfiguration(_config);
                if (!Regex.IsMatch(ticketNumber, @"\A[A-Za-z0-9_-]{1,64}\z")) throw new InvalidOperationException("Invalid ticket number.");
            }
            catch (Exception exception) { return new ServiceNowResult { Message = exception.Message }; }
            if (!_config.Enabled)
            {
                return new ServiceNowResult { Success = false, Message = "ServiceNow integration is not enabled in config." };
            }

            if (string.IsNullOrWhiteSpace(_config.InstanceUrl) || string.IsNullOrWhiteSpace(_config.KeyVaultUrl) ||
                string.IsNullOrWhiteSpace(_config.UsernameSecretName) || string.IsNullOrWhiteSpace(_config.PasswordSecretName))
            {
                return new ServiceNowResult { Success = false, Message = "ServiceNow config field is missing or empty." };
            }

            string snUser, snPass;
            try
            {
                snUser = await GetKeyVaultSecretAsync(_config.KeyVaultUrl, _config.UsernameSecretName);
                snPass = await GetKeyVaultSecretAsync(_config.KeyVaultUrl, _config.PasswordSecretName);
            }
            catch (Exception ex)
            {
                return new ServiceNowResult { Success = false, Message = $"Failed to retrieve ServiceNow credentials from Key Vault: {ex.Message}" };
            }

            var encodedCreds = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{snUser}:{snPass}"));
            var instanceUrl = _config.InstanceUrl.TrimEnd('/');
            var table = string.IsNullOrWhiteSpace(_config.Table) ? "task" : _config.Table;
            var numField = string.IsNullOrWhiteSpace(_config.TicketNumberField) ? "number" : _config.TicketNumberField;

            string sysId, sysClass;
            try
            {
                var query = Uri.EscapeDataString($"{numField}={ticketNumber}");
                var lookupUrl = $"{instanceUrl}/api/now/table/{table}?sysparm_query={query}&sysparm_fields=sys_id,{numField},sys_class_name&sysparm_limit=2";
                using var lookupRequest = new HttpRequestMessage(HttpMethod.Get, lookupUrl);
                lookupRequest.Headers.Authorization = new AuthenticationHeaderValue("Basic", encodedCreds);
                lookupRequest.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                using var lookupResponse = await _httpClient.SendAsync(lookupRequest);
                lookupResponse.EnsureSuccessStatusCode();
                var lookupJson = await lookupResponse.Content.ReadAsStringAsync();
                using var lookupDoc = JsonDocument.Parse(lookupJson);

                if (!lookupDoc.RootElement.TryGetProperty("result", out var resultArray) || resultArray.GetArrayLength() == 0)
                {
                    return new ServiceNowResult { Success = false, Message = $"Ticket '{ticketNumber}' not found in ServiceNow." };
                }

                var record = resultArray[0];
                if (resultArray.GetArrayLength() != 1 || !string.Equals(GetFieldValue(record, numField), ticketNumber, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Ticket lookup was ambiguous or did not match the requested number.");
                sysId = GetFieldValue(record, "sys_id");
                sysClass = GetFieldValue(record, "sys_class_name");
            }
            catch (Exception ex)
            {
                return new ServiceNowResult { Success = false, Message = $"Lookup failed for '{ticketNumber}': {ex.Message}" };
            }

            var updateTable = !string.IsNullOrWhiteSpace(_config.UpdateTable) ? _config.UpdateTable : (!string.IsNullOrWhiteSpace(sysClass) ? sysClass : table);
            var stateField = string.IsNullOrWhiteSpace(_config.CloseStateField) ? "state" : _config.CloseStateField;
            var stateValue = string.IsNullOrWhiteSpace(_config.CloseStateValue) ? "3" : _config.CloseStateValue;
            var workNotesField = string.IsNullOrWhiteSpace(_config.WorkNotesField) ? "work_notes" : _config.WorkNotesField;
            var commentsField = string.IsNullOrWhiteSpace(_config.AdditionalCommentsField) ? "comments" : _config.AdditionalCommentsField;

            try
            {
                if (!Regex.IsMatch(sysId, @"\A[a-fA-F0-9]{32}\z") || !IsIdentifier(updateTable)) throw new InvalidOperationException("Invalid ServiceNow update target.");
                var updatePayload = new System.Collections.Generic.Dictionary<string, object?>
                {
                    [stateField] = stateValue,
                    [workNotesField] = workNotes,
                    [commentsField] = additionalComments
                };

                if (_config.AdditionalFields != null)
                {
                    foreach (var field in _config.AdditionalFields)
                    {
                        if (updatePayload.ContainsKey(field.Key)) throw new InvalidOperationException("AdditionalFields cannot override closure state or notes.");
                        updatePayload[field.Key] = field.Value;
                    }
                }

                var updateUrl = $"{instanceUrl}/api/now/table/{updateTable}/{sysId}?sysparm_display_value=false&sysparm_fields=sys_id,{stateField}";
                using var updateRequest = new HttpRequestMessage(new HttpMethod("PATCH"), updateUrl);
                updateRequest.Headers.Authorization = new AuthenticationHeaderValue("Basic", encodedCreds);
                updateRequest.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                updateRequest.Content = new StringContent(JsonSerializer.Serialize(updatePayload), Encoding.UTF8, "application/json");

                using var updateResponse = await _httpClient.SendAsync(updateRequest);
                updateResponse.EnsureSuccessStatusCode();
                using var updated = JsonDocument.Parse(await updateResponse.Content.ReadAsStringAsync());
                if (!updated.RootElement.TryGetProperty("result", out var updatedRecord) || GetFieldValue(updatedRecord, stateField) != stateValue)
                    return new ServiceNowResult { Message = $"Ticket '{ticketNumber}' update was submitted, but closure state is unconfirmed. Verify it in ServiceNow." };

                return new ServiceNowResult { Success = true, Message = $"Ticket '{ticketNumber}' updated and closed successfully." };
            }
            catch (Exception ex)
            {
                return new ServiceNowResult { Success = false, Message = $"Failed to update ticket '{ticketNumber}': {ex.Message}" };
            }
        }

        private static string GetFieldValue(JsonElement record, string fieldName)
        {
            if (!record.TryGetProperty(fieldName, out var field)) return string.Empty;
            if (field.ValueKind == JsonValueKind.Object && field.TryGetProperty("value", out var valueProp))
            {
                return valueProp.GetString() ?? string.Empty;
            }
            return field.GetString() ?? string.Empty;
        }

        private static bool IsIdentifier(string value) => Regex.IsMatch(value, @"\A[a-zA-Z_][a-zA-Z0-9_]*\z");

        public static void ValidateConfiguration(ServiceNowConfig config)
        {
            foreach (var endpoint in new[] { config.InstanceUrl, config.KeyVaultUrl })
            {
                if (!Uri.TryCreate(endpoint, UriKind.Absolute, out var uri) || uri.Scheme != "https" || !string.IsNullOrEmpty(uri.UserInfo)
                    || !string.IsNullOrEmpty(uri.Query) || !string.IsNullOrEmpty(uri.Fragment) || uri.AbsolutePath != "/")
                    throw new InvalidOperationException("ServiceNow and Key Vault require HTTPS origin URLs without credentials, paths or queries.");
            }
            if (!new Uri(config.KeyVaultUrl).Host.EndsWith(".vault.azure.net", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Key Vault must use an Azure public-cloud vault endpoint.");
            var fields = new[] { config.Table, config.TicketNumberField, config.CloseStateField, config.WorkNotesField, config.AdditionalCommentsField };
            if (fields.Any(field => !IsIdentifier(field)) || (!string.IsNullOrEmpty(config.UpdateTable) && !IsIdentifier(config.UpdateTable))
                || config.AdditionalFields?.Keys.Any(field => !IsIdentifier(field)) == true)
                throw new InvalidOperationException("ServiceNow table and field names must be valid API identifiers.");
        }
    }
}

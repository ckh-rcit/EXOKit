using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
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
        private static readonly HttpClient _httpClient = new();

        public ServiceNowService(ServiceNowConfig config)
        {
            _config = config;
        }

        private async Task<string> GetKeyVaultSecretAsync(string vaultUrl, string secretName)
        {
            var credential = new InteractiveBrowserCredential();
            var client = new SecretClient(new Uri(vaultUrl), credential);
            var secret = await client.GetSecretAsync(secretName);
            return secret.Value.Value;
        }

        public async Task<ServiceNowResult> CloseTaskAsync(string ticketNumber, string workNotes = "", string additionalComments = "")
        {
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
                var lookupUrl = $"{instanceUrl}/api/now/table/{table}?sysparm_query={numField}={ticketNumber}&sysparm_fields=sys_id,number,sys_class_name&sysparm_limit=1";
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
                        updatePayload[field.Key] = field.Value;
                    }
                }

                var updateUrl = $"{instanceUrl}/api/now/table/{updateTable}/{sysId}";
                using var updateRequest = new HttpRequestMessage(new HttpMethod("PATCH"), updateUrl);
                updateRequest.Headers.Authorization = new AuthenticationHeaderValue("Basic", encodedCreds);
                updateRequest.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                updateRequest.Content = new StringContent(JsonSerializer.Serialize(updatePayload), Encoding.UTF8, "application/json");

                using var updateResponse = await _httpClient.SendAsync(updateRequest);
                updateResponse.EnsureSuccessStatusCode();

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
    }
}

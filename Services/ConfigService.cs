using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace EXOKit.Services
{
    public class LicenseGroupConfig
    {
        [JsonPropertyName("GroupName")]
        public string GroupName { get; set; } = string.Empty;

        [JsonPropertyName("GroupId")]
        public string? GroupId { get; set; }
    }

    public class LicenseGroupsConfig
    {
        [JsonPropertyName("Bookings")]
        public LicenseGroupConfig Bookings { get; set; } = new();
    }

    public class OwaPoliciesConfig
    {
        [JsonPropertyName("BookingsCreators")]
        public string BookingsCreators { get; set; } = string.Empty;
    }

    public class GraphApiConfig
    {
        public string ClientId { get; set; } = string.Empty;
        public string TenantId { get; set; } = string.Empty;
        [JsonPropertyName("Scopes")]
        public List<string> Scopes { get; set; } = new();
    }

    public class ServiceNowConfig
    {
        [JsonPropertyName("Enabled")]
        public bool Enabled { get; set; }

        [JsonPropertyName("InstanceUrl")]
        public string InstanceUrl { get; set; } = string.Empty;

        [JsonPropertyName("KeyVaultUrl")]
        public string KeyVaultUrl { get; set; } = string.Empty;

        [JsonPropertyName("SubscriptionId")]
        public string SubscriptionId { get; set; } = string.Empty;

        [JsonPropertyName("UsernameSecretName")]
        public string UsernameSecretName { get; set; } = string.Empty;

        [JsonPropertyName("PasswordSecretName")]
        public string PasswordSecretName { get; set; } = string.Empty;

        [JsonPropertyName("Table")]
        public string Table { get; set; } = "task";

        [JsonPropertyName("TicketNumberField")]
        public string TicketNumberField { get; set; } = "number";

        [JsonPropertyName("CloseStateField")]
        public string CloseStateField { get; set; } = "state";

        [JsonPropertyName("CloseStateValue")]
        public string CloseStateValue { get; set; } = "3";

        [JsonPropertyName("WorkNotesField")]
        public string WorkNotesField { get; set; } = "work_notes";

        [JsonPropertyName("AdditionalCommentsField")]
        public string AdditionalCommentsField { get; set; } = "comments";

        [JsonPropertyName("UpdateTable")]
        public string UpdateTable { get; set; } = string.Empty;

        [JsonPropertyName("AdditionalFields")]
        public Dictionary<string, object>? AdditionalFields { get; set; }
    }

    public class SettingsConfig
    {
        public bool ExoUseBrowserSignIn { get; set; }
        public string UpdateFeedUrl { get; set; } = string.Empty;
        [JsonPropertyName("LicenseGroups")]
        public LicenseGroupsConfig LicenseGroups { get; set; } = new();

        [JsonPropertyName("OwaPolicies")]
        public OwaPoliciesConfig OwaPolicies { get; set; } = new();

        [JsonPropertyName("GraphApi")]
        public GraphApiConfig GraphApi { get; set; } = new();

        [JsonPropertyName("ServiceNow")]
        public ServiceNowConfig? ServiceNow { get; set; }
    }

    public class ToolConfig
    {
        [JsonPropertyName("Settings")]
        public SettingsConfig Settings { get; set; } = new();
    }

    /// <summary>
    /// Loads and validates config.json, ported from Test-ConfigStructure. Required fields for the
    /// ported (non-Provisioning, non-Calendar) tabs are LicenseGroups.Bookings.GroupName,
    /// OwaPolicies.BookingsCreators, and a non-empty GraphApi.Scopes array. ServiceNow is optional
    /// (disabled ticket auto-close is a valid configuration).
    /// </summary>
    public static class ConfigService
    {
        public static ToolConfig CreateDefault() => new()
        {
            Settings = new SettingsConfig
            {
                LicenseGroups = new LicenseGroupsConfig { Bookings = new LicenseGroupConfig { GroupName = "License_M365_Bookings" } },
                OwaPolicies = new OwaPoliciesConfig { BookingsCreators = "BookingsCreators" },
                GraphApi = new GraphApiConfig { Scopes = new List<string> { "User.Read.All", "Group.ReadWrite.All" } }
            }
        };
        /// <summary>
        /// Returns a writable directory for config.json. MSIX-packaged installs run from a read-only
        /// location under Program Files\WindowsApps, so runtime edits (Settings page Save) must be
        /// persisted under the per-user LocalAppData folder instead of the app's install directory.
        /// </summary>
        private static string GetWritableConfigDirectory()
        {
            var directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "EXOKit");
            Directory.CreateDirectory(directory);
            return directory;
        }

        public static ToolConfig Load(string? configDirectory = null)
        {
            var directory = configDirectory ?? GetWritableConfigDirectory();
            var configPath = Path.Combine(directory, "config.json");

            if (!File.Exists(configPath))
            {
                // First run (or no prior save): seed the writable location from the bundled config.json
                // shipped alongside the app, if one exists.
                var bundledPath = Path.Combine(AppContext.BaseDirectory, "config.json");
                if (configDirectory == null && File.Exists(bundledPath))
                {
                    File.Copy(bundledPath, configPath);
                }
            }

            if (!File.Exists(configPath))
            {
                throw new InvalidOperationException($"Configuration file 'config.json' not found in the application directory: {directory}");
            }

            ToolConfig? config;
            try
            {
                var json = File.ReadAllText(configPath);
                config = JsonSerializer.Deserialize<ToolConfig>(json);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to parse 'config.json': {ex.Message}", ex);
            }

            if (config == null)
            {
                throw new InvalidOperationException("Failed to load 'config.json': file is empty or invalid.");
            }

            Validate(config);
            return config;
        }

        /// <summary>
        /// Validates and writes the given config back to config.json, so the in-app Settings
        /// section can persist edits made at runtime.
        /// </summary>
        public static void Save(ToolConfig config, string? configDirectory = null)
        {
            Validate(config);

            var directory = configDirectory ?? GetWritableConfigDirectory();
            var configPath = Path.Combine(directory, "config.json");

            var json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
            var temporaryPath = configPath + ".tmp";
            File.WriteAllText(temporaryPath, json);
            if (File.Exists(configPath)) File.Copy(configPath, configPath + ".bak", true);
            File.Move(temporaryPath, configPath, true);
        }

        private static void Validate(ToolConfig config)
        {
            var missingFields = new List<string>();

            if (string.IsNullOrWhiteSpace(config.Settings?.LicenseGroups?.Bookings?.GroupName)
                && string.IsNullOrWhiteSpace(config.Settings?.LicenseGroups?.Bookings?.GroupId))
            {
                missingFields.Add("Settings.LicenseGroups.Bookings.GroupName");
            }
            if (string.IsNullOrWhiteSpace(config.Settings?.OwaPolicies?.BookingsCreators))
            {
                missingFields.Add("Settings.OwaPolicies.BookingsCreators");
            }
            if (config.Settings?.GraphApi?.Scopes == null || config.Settings.GraphApi.Scopes.Count == 0)
            {
                missingFields.Add("Settings.GraphApi.Scopes");
            }

            if (missingFields.Count > 0)
            {
                throw new InvalidOperationException("Configuration validation failed. Missing or null required fields:" + Environment.NewLine + string.Join(Environment.NewLine, missingFields));
            }

            if (config.Settings!.GraphApi.Scopes.Exists(string.IsNullOrWhiteSpace))
            {
                throw new InvalidOperationException("Settings.GraphApi.Scopes contains one or more empty values.");
            }
        }
    }
}

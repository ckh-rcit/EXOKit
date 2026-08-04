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
        public static ToolConfig Load(string? configDirectory = null)
        {
            var directory = configDirectory ?? AppContext.BaseDirectory;
            var configPath = Path.Combine(directory, "config.json");

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

            var directory = configDirectory ?? AppContext.BaseDirectory;
            var configPath = Path.Combine(directory, "config.json");

            var json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(configPath, json);
        }

        private static void Validate(ToolConfig config)
        {
            var missingFields = new List<string>();

            if (string.IsNullOrWhiteSpace(config.Settings?.LicenseGroups?.Bookings?.GroupName))
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

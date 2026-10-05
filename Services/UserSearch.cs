using System;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;

namespace EXOKit.Services;

public sealed class UserSearchResult : INotifyPropertyChanged
{
    public string DisplayName { get; init; } = string.Empty;
    public string UserPrincipalName { get; init; } = string.Empty;
    public string Mail { get; init; } = string.Empty;
    public string EmployeeId { get; init; } = string.Empty;
    public string SamAccountName { get; init; } = string.Empty;
    public string JobTitle { get; init; } = string.Empty;
    public string Department { get; init; } = string.Empty;
    public bool? Enabled { get; init; }
    public string ObjectId { get; init; } = string.Empty;
    public DateTimeOffset? CreatedDateTime { get; init; }
    public string BusinessPhone { get; init; } = string.Empty;
    public bool IsDeleted { get; init; }
    public string AccountStatus => IsDeleted ? "Deleted" : Enabled switch { true => "Enabled", false => "Disabled", _ => "Unknown" };
    private string _exoMailbox = "Not checked";
    public string ExoMailbox
    {
        get => _exoMailbox;
        private set
        {
            _exoMailbox = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ExoMailbox)));
        }
    }
    public string MailboxDiagnostic { get; private set; } = string.Empty;
    public void SetMailboxStatus(string status, string diagnostic = "")
    {
        MailboxDiagnostic = diagnostic;
        ExoMailbox = status;
    }
    public event PropertyChangedEventHandler? PropertyChanged;

    public string Details => string.Join(Environment.NewLine, new[]
    {
        $"Display name: {DisplayName}", $"UPN: {UserPrincipalName}", $"Email: {Mail}",
        $"Employee ID: {EmployeeId}", $"SAM account: {SamAccountName}", $"Job title: {JobTitle}",
        $"Department: {Department}", $"Account: {AccountStatus}", $"EXO mailbox: {ExoMailbox}",
        $"Object ID: {ObjectId}", $"Created (UTC): {CreatedDateTime?.UtcDateTime.ToString("u") ?? "Unknown"}",
        $"Business phone: {BusinessPhone}", $"Deleted: {IsDeleted}", $"Mailbox details: {MailboxDiagnostic}"
    });
}

public sealed record ParsedSearchQuery(string Description, string Filter, string? FallbackFilter = null);

public static class SearchQueryParser
{
    public static ParsedSearchQuery? Parse(string rawQuery)
    {
        if (string.IsNullOrWhiteSpace(rawQuery)) return null;
        var query = Regex.Replace(rawQuery.Trim(), @"\s+", " ");
        if (query.Contains('@'))
            return new("UPN / email", $"userPrincipalName eq '{Escape(query)}' or mail eq '{Escape(query)}'");

        if (query.Contains(','))
        {
            var parts = query.Split(',', 2);
            var last = parts[0].Trim();
            var first = parts[1].Trim();
            if (last.Length == 0 || first.Length == 0)
                throw new ArgumentException("Enter both names for a Last, First search.", nameof(rawQuery));
            return new("Last, First", NameFilter($"{first} {last}", $"{last}, {first}", $"{last} {first}"));
        }

        if (Regex.IsMatch(query, @"^[Ee]?\d+$"))
        {
            var numeric = query.All(char.IsDigit);
            var normalized = numeric ? "E" + query : query.ToUpperInvariant();
            return new("Employee ID", $"employeeId eq '{Escape(normalized)}'", numeric ? $"employeeId eq '{Escape(query)}'" : null);
        }

        if (query.Contains(' '))
        {
            var parts = query.Split(' ');
            return new("Full name", parts.Length == 2
                ? NameFilter(query, $"{parts[1]}, {parts[0]}", $"{parts[1]} {parts[0]}")
                : NameFilter(query));
        }

        var escaped = Escape(query);
        return new("SAM account / name", $"onPremisesSamAccountName eq '{escaped}' or startswith(displayName,'{escaped}') or startswith(userPrincipalName,'{escaped}')");
    }

    private static string NameFilter(params string[] names) =>
        string.Join(" or ", names.Distinct(StringComparer.OrdinalIgnoreCase).Select(name => $"startswith(displayName,'{Escape(name)}')"));

    private static string Escape(string value) => value.Replace("'", "''");
}
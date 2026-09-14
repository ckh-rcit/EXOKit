using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text.RegularExpressions;

namespace EXOKit.Services
{
    /// <summary>
    /// Ports ConvertTo-InputList, Test-IsEmailLikeValue, and Get-InvalidEmailLikeValues from the
    /// WinForms script. Input boxes accept comma/semicolon/newline/space separated values, which
    /// are trimmed, stripped of surrounding quote/angle-bracket characters, and deduplicated
    /// case-insensitively while preserving first-seen order.
    /// </summary>
    public static class InputParsingHelpers
    {
        private static readonly Regex SplitPattern = new(@"[,;\r\n\s]+", RegexOptions.Compiled);
        private static readonly Regex TrimCharsPattern = new(@"^[""'<]+|[""'>]+$", RegexOptions.Compiled);

        public static string[] ConvertToInputList(string? input)
        {
            if (string.IsNullOrWhiteSpace(input)) return Array.Empty<string>();

            var parsedValues = new List<string>();
            var seenValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var rawValue in SplitPattern.Split(input))
            {
                var cleanValue = TrimCharsPattern.Replace(rawValue.Trim(), string.Empty);
                if (string.IsNullOrWhiteSpace(cleanValue)) continue;
                if (seenValues.Add(cleanValue))
                {
                    parsedValues.Add(cleanValue);
                }
            }

            return parsedValues.ToArray();
        }

        public static bool IsEmailLikeValue(string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return false;
            if (value.Any(char.IsWhiteSpace)) return false;

            try
            {
                var mailAddress = new MailAddress(value);
                return string.Equals(mailAddress.Address, value, StringComparison.OrdinalIgnoreCase);
            }
            catch (FormatException)
            {
                return false;
            }
        }

        public static string[] ConvertToRecipientList(string? input)
        {
            return (input ?? string.Empty).Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(value => value.Trim()).Where(value => value.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        }

        public static string[] ParseRecipientFile(string content, bool isCsv)
        {
            if (!isCsv) return ConvertToRecipientList(content);

            using var reader = new System.IO.StringReader(content);
            using var parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(reader);
            parser.SetDelimiters(",", ";");
            parser.HasFieldsEnclosedInQuotes = true;
            var firstRow = parser.ReadFields();
            if (firstRow == null) return Array.Empty<string>();

            var headers = new[] { "Identity", "Recipient", "Email", "EmailAddress", "PrimarySmtpAddress", "UserPrincipalName", "UPN" };
            var column = Array.FindIndex(firstRow, value => headers.Contains(value.Trim(), StringComparer.OrdinalIgnoreCase));
            var hasHeader = column >= 0;
            if (!hasHeader && firstRow.Length != 1)
                throw new FormatException("A multi-column CSV must have an Identity, Recipient, Email, EmailAddress, PrimarySmtpAddress, UserPrincipalName, or UPN header.");
            if (!hasHeader) column = 0;

            var values = new List<string>();
            void AddRow(string[] row)
            {
                if (row.Length != firstRow.Length)
                    throw new FormatException("Each CSV row must have the same number of columns as the first row.");
                var value = row[column].Trim();
                if (value.IndexOfAny(new[] { '\r', '\n' }) >= 0)
                    throw new FormatException("A recipient identity cannot span multiple lines.");
                if (value.Length > 0) values.Add(value);
            }

            if (!hasHeader) AddRow(firstRow);
            while (!parser.EndOfData) AddRow(parser.ReadFields() ?? Array.Empty<string>());
            return values.Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        }

        public static string[] GetInvalidEmailLikeValues(IEnumerable<string> values)
        {
            return values.Where(v => !IsEmailLikeValue(v)).ToArray();
        }
    }
}

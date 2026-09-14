using System;
using System.Collections;
using System.Linq;
using System.Management.Automation;
using System.Threading;
using System.Threading.Tasks;

namespace EXOKit.Services
{
    public static class PermissionVerification
    {
        public static string[] ReadStrings(PSObject row, string propertyName)
        {
            var property = row.Properties[propertyName]
                ?? throw new InvalidOperationException($"Exchange result is missing {propertyName}; state is unconfirmed.");
            return ReadStrings(property.Value);
        }

        public static string[] ReadStrings(object? value)
        {
            if (value is null) return Array.Empty<string>();
            if (value is PSObject wrapped)
            {
                if (wrapped.BaseObject is PSCustomObject)
                {
                    var serializedText = wrapped.ToString();
                    if (string.IsNullOrWhiteSpace(serializedText) || serializedText.StartsWith("@{", StringComparison.Ordinal))
                        throw new InvalidOperationException("Exchange returned an unreadable identity; state is unconfirmed.");
                    return new[] { serializedText };
                }
                return ReadStrings(wrapped.BaseObject);
            }
            if (value is IDictionary || value is PSCustomObject)
                throw new InvalidOperationException("Exchange returned an unreadable collection; state is unconfirmed.");
            if (value is IEnumerable values && value is not string)
                return values.Cast<object?>().SelectMany(ReadStrings).ToArray();
            var text = value.ToString();
            if (string.IsNullOrWhiteSpace(text)) return Array.Empty<string>();
            if (value.GetType() == typeof(object))
                throw new InvalidOperationException("Exchange returned an unreadable value; state is unconfirmed.");
            return new[] { text };
        }

        public static bool HasAccessRight(object? accessRights, string expectedRight)
        {
            return ReadStrings(accessRights).Any(right => string.Equals(right, expectedRight, StringComparison.OrdinalIgnoreCase));
        }

        public static bool IsTransient(Exception exception) =>
            exception.Message.Contains("Object reference not set", StringComparison.OrdinalIgnoreCase);

        public static async Task<T> ReadAsync<T>(Func<Task<T>> read, Func<Task>? delay = null, CancellationToken cancellationToken = default)
        {
            for (var attempt = 0; ; attempt++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                try { return await read(); }
                catch (Exception exception) when (exception is not OperationCanceledException && attempt < 2 && IsTransient(exception))
                {
                    await (delay?.Invoke() ?? Task.Delay(TimeSpan.FromSeconds(2), cancellationToken));
                }
            }
        }

        public static async Task ApplyAsync(Func<Task> apply, Func<Task<bool>> present, bool expected, Func<Task>? delay = null, CancellationToken cancellationToken = default)
        {
            if (await ReadAsync(present, delay, cancellationToken) == expected) return;
            Exception? lastError = null;
            for (var attempt = 0; attempt < 2; attempt++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var mutationThrew = false;
                try { await apply(); }
                catch (Exception exception) when (exception is not OperationCanceledException && IsTransient(exception)) { lastError = exception; mutationThrew = true; }

                for (var poll = 0; poll < 15; poll++)
                {
                    await (delay?.Invoke() ?? Task.Delay(TimeSpan.FromSeconds(2), cancellationToken));
                    try
                    {
                        if (await ReadAsync(present, delay, cancellationToken) == expected) return;
                    }
                    catch (Exception exception) when (exception is not OperationCanceledException && IsTransient(exception)) { lastError = exception; }
                }
                if (!mutationThrew) break;
            }
            throw new InvalidOperationException("Permission change is unconfirmed. Verify the target before retrying or closing the ticket.", lastError);
        }
    }
}
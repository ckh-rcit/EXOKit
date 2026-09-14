using System;
using System.Threading.Tasks;

namespace EXOKit.Services
{
    public static class PermissionVerification
    {
        public static bool IsTransient(Exception exception) =>
            exception.Message.Contains("Object reference not set", StringComparison.OrdinalIgnoreCase);

        public static async Task<T> ReadAsync<T>(Func<Task<T>> read, Func<Task>? delay = null)
        {
            for (var attempt = 0; ; attempt++)
            {
                try { return await read(); }
                catch (Exception exception) when (attempt < 2 && IsTransient(exception))
                {
                    await (delay?.Invoke() ?? Task.Delay(TimeSpan.FromSeconds(2)));
                }
            }
        }

        public static async Task ApplyAsync(Func<Task> apply, Func<Task<bool>> present, bool expected, Func<Task>? delay = null)
        {
            if (await ReadAsync(present, delay) == expected) return;
            Exception? lastError = null;
            for (var attempt = 0; attempt < 2; attempt++)
            {
                var mutationThrew = false;
                try { await apply(); }
                catch (Exception exception) when (IsTransient(exception)) { lastError = exception; mutationThrew = true; }

                for (var poll = 0; poll < 15; poll++)
                {
                    await (delay?.Invoke() ?? Task.Delay(TimeSpan.FromSeconds(2)));
                    try
                    {
                        if (await ReadAsync(present, delay) == expected) return;
                    }
                    catch (Exception exception) when (IsTransient(exception)) { lastError = exception; }
                }
                if (!mutationThrew) break;
            }
            throw new InvalidOperationException("Permission change is unconfirmed. Verify the target before retrying or closing the ticket.", lastError);
        }
    }
}
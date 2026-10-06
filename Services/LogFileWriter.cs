using System;
using System.IO;

namespace EXOKit.Services;

/// <summary>Appends log lines to a primary file and switches to a fallback file when the primary cannot be written.</summary>
internal sealed class LogFileWriter(Func<string> primaryPath, Func<string> fallbackPath)
{
    private bool _usingFallback;

    internal bool Writable { get; private set; } = true;
    internal string CurrentPath => _usingFallback ? fallbackPath() : primaryPath();

    /// <summary>Returns a notice to show the user when this call moved the log to the fallback location.</summary>
    internal string? Append(string line)
    {
        string? notice = null;
        if (!_usingFallback)
        {
            var primary = primaryPath();
            var failure = TryAppend(primary, line);
            if (failure == null) return null;
            _usingFallback = true;
            notice = $"Cannot write the log to '{primary}' ({failure}). Using '{fallbackPath()}' instead.";
            TryAppend(fallbackPath(), notice);
        }
        Writable = TryAppend(fallbackPath(), line) == null;
        return notice;
    }

    private static string? TryAppend(string path, string line)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.AppendAllText(path, line + Environment.NewLine);
            return null;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or NotSupportedException or System.Security.SecurityException)
        {
            return exception.Message;
        }
    }
}

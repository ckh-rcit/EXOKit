using System;
using System.Runtime.InteropServices;

namespace EXOKit.Services;

internal static class ErrorReporter
{
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int MessageBoxW(IntPtr window, string text, string caption, uint type);

    internal static void Report(string operation, Exception error, bool fatal = false)
    {
        try { Logger.Log($"{operation}: {error}", LogType.Error); } catch { }
        try
        {
            MessageBoxW(IntPtr.Zero, $"{operation}\n\n{error.Message}\n\n{(fatal ? "The application must close. " : "")}Review applied changes before retrying.\nLog: {Logger.LogPath}", "EXOKit", 0x10);
        }
        catch { }
    }
}
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace EXOKit.Services;

internal static class PackagedPathResolver
{
    private const int ErrorInsufficientBuffer = 122;

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetCurrentPackageFamilyName(ref uint length, StringBuilder? name);

    internal static string? CurrentPackageFamilyName()
    {
        try
        {
            uint length = 0;
            if (GetCurrentPackageFamilyName(ref length, null) != ErrorInsufficientBuffer) return null;
            var name = new StringBuilder((int)length);
            return GetCurrentPackageFamilyName(ref length, name) == 0 ? name.ToString() : null;
        }
        catch (EntryPointNotFoundException) { return null; }
    }

    /// <summary>
    /// Packaged apps see a new folder under AppData\Local as the real one, but Windows stores it under
    /// Packages\&lt;family&gt;\LocalCache\Local. Returns the location a person can browse to.
    /// </summary>
    internal static string ToBrowsablePath(string path, string? localAppData = null, string? familyName = null)
    {
        localAppData ??= Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        familyName ??= CurrentPackageFamilyName();
        if (string.IsNullOrEmpty(familyName) || string.IsNullOrEmpty(localAppData)) return path;

        var root = localAppData.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        if (!path.StartsWith(root, StringComparison.OrdinalIgnoreCase)) return path;

        var relative = path[root.Length..];
        var packaged = Path.Combine(root, "Packages", familyName, "LocalCache", "Local", relative);
        if (File.Exists(packaged) || Directory.Exists(packaged)) return packaged;

        // Windows opens the package copy first and uses the real folder only when it already existed.
        var topFolder = relative.Split(Path.DirectorySeparatorChar)[0];
        return Directory.Exists(Path.Combine(root, topFolder)) || File.Exists(path) ? path : packaged;
    }
}

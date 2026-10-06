using EXOKit.Services;
using Xunit;

namespace EXOKit.Tests;

public sealed class LogLocationTests : IDisposable
{
    private const string Family = "Example_73b7v5sxh7xgp";
    private readonly string _root = Path.Combine(Path.GetTempPath(), "EXOKit-logloc-" + Guid.NewGuid().ToString("N"));

    public LogLocationTests() => Directory.CreateDirectory(_root);
    public void Dispose() => Directory.Delete(_root, true);

    [Fact]
    public void NewUserLogPathPointsAtThePackageCopyWhereWindowsStoresIt()
    {
        var path = Path.Combine(_root, "EXOKit", "Logs", "today.log");
        var expected = Path.Combine(_root, "Packages", Family, "LocalCache", "Local", "EXOKit", "Logs", "today.log");
        Assert.Equal(expected, PackagedPathResolver.ToBrowsablePath(path, _root, Family));
    }

    [Fact]
    public void ExistingRealFolderKeepsTheRealPath()
    {
        Directory.CreateDirectory(Path.Combine(_root, "EXOKit"));
        var path = Path.Combine(_root, "EXOKit", "Logs", "today.log");
        Assert.Equal(path, PackagedPathResolver.ToBrowsablePath(path, _root, Family));
    }

    [Fact]
    public void PackageCopyWinsWhenItExists()
    {
        Directory.CreateDirectory(Path.Combine(_root, "EXOKit"));
        var packaged = Path.Combine(_root, "Packages", Family, "LocalCache", "Local", "EXOKit", "Snapshots");
        Directory.CreateDirectory(packaged);
        Assert.Equal(packaged, PackagedPathResolver.ToBrowsablePath(Path.Combine(_root, "EXOKit", "Snapshots"), _root, Family));
    }

    [Fact]
    public void UnpackagedAndExternalPathsAreUnchanged()
    {
        var path = Path.Combine(_root, "EXOKit", "Logs", "today.log");
        Assert.Equal(path, PackagedPathResolver.ToBrowsablePath(path, _root, ""));
        var elsewhere = Path.Combine(Path.GetTempPath(), "elsewhere", "today.log");
        Assert.Equal(elsewhere, PackagedPathResolver.ToBrowsablePath(elsewhere, _root, Family));
        Assert.Null(PackagedPathResolver.CurrentPackageFamilyName());
    }

    [Fact]
    public void LogWriterReportsAndFallsBackWhenPrimaryCannotBeWritten()
    {
        var blocker = Path.Combine(_root, "blocker");
        File.WriteAllText(blocker, "a file where a folder is required");
        var fallback = Path.Combine(_root, "temp", "log.txt");
        var writer = new LogFileWriter(() => Path.Combine(blocker, "Logs", "log.txt"), () => fallback);

        var notice = writer.Append("first");
        Assert.NotNull(notice);
        Assert.Contains("Cannot write the log", notice);
        Assert.Contains(fallback, notice);
        Assert.Null(writer.Append("second"));
        Assert.Equal(fallback, writer.CurrentPath);
        Assert.True(writer.Writable);
        Assert.Equal(new[] { notice, "first", "second" }, File.ReadAllLines(fallback));
    }

    [Fact]
    public void LogWriterUsesPrimaryAndStaysQuietWhenItWorks()
    {
        var primary = Path.Combine(_root, "Logs", "log.txt");
        var writer = new LogFileWriter(() => primary, () => Path.Combine(_root, "unused.txt"));
        Assert.Null(writer.Append("one"));
        Assert.Null(writer.Append("two"));
        Assert.Equal(primary, writer.CurrentPath);
        Assert.Equal(new[] { "one", "two" }, File.ReadAllLines(primary));
        Assert.False(File.Exists(Path.Combine(_root, "unused.txt")));
    }

    [Fact]
    public void LogWriterReportsNotWritableWhenNothingWorks()
    {
        var blocker = Path.Combine(_root, "blocker");
        File.WriteAllText(blocker, "x");
        var writer = new LogFileWriter(() => Path.Combine(blocker, "a", "log.txt"), () => Path.Combine(blocker, "b", "log.txt"));
        Assert.NotNull(writer.Append("lost"));
        Assert.False(writer.Writable);
    }
}

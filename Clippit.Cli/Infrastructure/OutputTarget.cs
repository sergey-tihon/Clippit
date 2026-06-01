namespace Clippit.Cli.Infrastructure;

internal sealed class OutputTarget
{
    public const string StdoutToken = "-";

    private readonly string? _path;

    private OutputTarget(string displayPath, string? path, bool isStdout)
    {
        DisplayPath = displayPath;
        _path = path;
        IsStdout = isStdout;
    }

    public bool IsStdout { get; }

    public string DisplayPath { get; }

    public static OutputTarget FromOption(string? outputOption, Func<string> defaultPath)
    {
        if (outputOption == StdoutToken)
            return Stdout();

        var path = outputOption ?? defaultPath();
        return File(path);
    }

    public static OutputTarget File(string path)
    {
        var fullPath = Path.IsPathRooted(path) ? path : Path.GetFullPath(path);
        return new OutputTarget(fullPath, fullPath, isStdout: false);
    }

    public static OutputTarget Stdout() => new("<stdout>", null, isStdout: true);

    public void EnsureDirectoryExists()
    {
        if (_path is null)
            return;

        var directory = Path.GetDirectoryName(_path);
        if (directory is { Length: > 0 })
            Directory.CreateDirectory(directory);
    }

    public void EnsureCanWrite(bool force, string itemName)
    {
        if (_path is null || force || !System.IO.File.Exists(_path))
            return;

        throw CliException.OutputError($"{itemName} already exists: {_path}. Pass --force to overwrite.");
    }

    public Stream OpenWrite(out string? tempPath)
    {
        if (_path is null)
        {
            tempPath = null;
            return new MemoryStream();
        }

        var directory = Path.GetDirectoryName(_path) ?? Directory.GetCurrentDirectory();
        var fileName = Path.GetFileName(_path);
        tempPath = Path.Combine(directory, $".{fileName}.{Guid.NewGuid():N}.tmp");
        return new FileStream(tempPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
    }

    public void Commit(string? tempPath)
    {
        if (_path is null || tempPath is null)
            return;

        System.IO.File.Move(tempPath, _path, overwrite: true);
    }

    public static void DeleteTemp(string? tempPath)
    {
        if (tempPath is not null && System.IO.File.Exists(tempPath))
            System.IO.File.Delete(tempPath);
    }

    public void Flush(Stream stream)
    {
        if (!IsStdout)
            return;

        stream.Position = 0;
        using var stdout = Console.OpenStandardOutput();
        stream.CopyTo(stdout);
    }
}

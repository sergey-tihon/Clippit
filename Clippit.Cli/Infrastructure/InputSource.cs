using System.CommandLine;
using System.CommandLine.Parsing;

namespace Clippit.Cli.Infrastructure;

internal sealed class InputSource
{
    public const string StdinToken = "-";

    private readonly string _path;

    private InputSource(string path, string displayName, string logicalName, bool isStdin)
    {
        _path = path;
        DisplayName = displayName;
        LogicalName = logicalName;
        IsStdin = isStdin;
    }

    public bool IsStdin { get; }

    public string DisplayName { get; }

    public string LogicalName { get; }

    public static Argument<string> BuildArgument(string name, string description, string fileDescription = "File")
    {
        var argument = new Argument<string>(name) { Description = description };
        argument.Validators.Add(result => ValidateExistsOrStdin(result, fileDescription));
        return argument;
    }

    public static InputSource From(string value, string stdinLogicalName)
    {
        if (value == StdinToken)
            return new InputSource(value, "<stdin>", stdinLogicalName, isStdin: true);

        var fileInfo = new FileInfo(value);
        return new InputSource(fileInfo.FullName, fileInfo.FullName, fileInfo.Name, isStdin: false);
    }

    public Stream OpenSeekable() =>
        IsStdin ? ReadStdinToMemory() : new FileStream(_path, FileMode.Open, FileAccess.Read, FileShare.Read);

    public static void ValidateExistsOrStdin(ArgumentResult result, string fileDescription = "File")
    {
        var value = result.GetValueOrDefault<string?>();
        if (string.IsNullOrEmpty(value) || value == StdinToken)
            return;

        if (!File.Exists(value))
            result.AddError($"{fileDescription} not found: {value}");
    }

    private static MemoryStream ReadStdinToMemory()
    {
        var buffer = new MemoryStream();
        using (var stdin = Console.OpenStandardInput())
            stdin.CopyTo(buffer);
        buffer.Position = 0;
        return buffer;
    }
}

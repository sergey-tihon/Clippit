using System.CommandLine;
using System.Reflection;
using Clippit.Cli.Infrastructure;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Cli.Commands.Version;

internal static class VersionCommand
{
    public static Command Build()
    {
        var cmd = new Command("version", "Print the current version of Clippit CLI.");
        var (formatOption, quietOption) = cmd.AddOutputOptions();

        cmd.SetAction(parseResult =>
        {
            var format = parseResult.GetValue(formatOption);
            var quiet = parseResult.GetValue(quietOption);
            var writer = new OutputWriter(format, quiet);
            writer.WriteResult(BuildResult(), CliJsonContext.Default.VersionResult, VersionResult.WriteText);
            return 0;
        });
        return cmd;
    }

    /// <summary>
    /// Used by the <c>--version</c> root-level intercept so the structured
    /// payload is identical to <c>clippit version</c>.
    /// </summary>
    public static VersionResult BuildResult() =>
        new() { Version = GetCleanVersion(), OpenXmlSdkVersion = GetCleanVersion(typeof(OpenXmlPackage).Assembly) };

    public static string GetCleanVersion() => GetCleanVersion(typeof(VersionCommand).Assembly);

    private static string GetCleanVersion(Assembly assembly)
    {
        var version =
            assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
            ?? assembly.GetName().Version?.ToString()
            ?? "unknown";

        // Strip commit hash suffix if present (e.g. "1.2.3+abc1234" -> "1.2.3").
        var index = version.IndexOf('+', StringComparison.Ordinal);
        return index >= 0 ? version[..index] : version;
    }
}

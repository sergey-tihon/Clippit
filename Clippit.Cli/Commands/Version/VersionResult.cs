namespace Clippit.Cli.Commands.Version;

/// <summary>
/// Result of the "version" command.
/// </summary>
internal sealed record VersionResult
{
    public required string Version { get; init; }

    public required string OpenXmlSdkVersion { get; init; }

    public static void WriteText(VersionResult result, TextWriter writer)
    {
        writer.WriteLine($"clippit         {result.Version}");
        writer.WriteLine($"OpenXml SDK     {result.OpenXmlSdkVersion}");
    }
}

namespace Clippit.Cli.Commands.Install;

internal sealed class InstallResult
{
    public required IReadOnlyList<InstalledSkillResult> Installed { get; init; }
    public required string Version { get; init; }

    public static void WriteText(InstallResult result, TextWriter writer)
    {
        if (result.Installed.Count == 0)
        {
            writer.WriteLine("No Clippit skills installed.");
            return;
        }

        writer.WriteLine(result.Installed.Count == 1 ? "Installed Clippit skill:" : "Installed Clippit skills:");
        foreach (var installed in result.Installed)
            writer.WriteLine($"  {installed.Path}");
    }
}

internal sealed class InstalledSkillResult
{
    public required string Target { get; init; }
    public required string Path { get; init; }
}

internal sealed class InstallPlanResult
{
    public required IReadOnlyList<string> Paths { get; init; }

    public static void WriteText(InstallPlanResult result, TextWriter writer)
    {
        if (result.Paths.Count == 0)
        {
            writer.WriteLine("No Clippit skills planned.");
            return;
        }

        foreach (var path in result.Paths)
            writer.WriteLine(path);
    }
}

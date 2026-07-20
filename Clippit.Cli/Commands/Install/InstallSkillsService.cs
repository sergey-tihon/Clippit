using System.Reflection;
using Clippit.Cli.Commands.Version;
using Clippit.Cli.Infrastructure;

namespace Clippit.Cli.Commands.Install;

internal static class InstallSkillsService
{
    public const string DefaultTarget = "agents";

    private static readonly string[] s_skillFiles =
    [
        "SKILL.md",
        "references/workflows.md",
        "references/manifests.md",
        "references/output.md",
    ];

    private static readonly Assembly s_assembly = typeof(InstallSkillsService).Assembly;

    public static InstallResult Install(string? targetValue)
    {
        var targets = ResolveTargets(targetValue);
        var files = ReadBundledSkillFiles();
        var installed = new List<InstalledSkillResult>();

        foreach (var target in targets)
        {
            var targetDirectory = ResolveTargetDirectory(target);
            var skillFile = new FileInfo(Path.Combine(targetDirectory.FullName, "SKILL.md"));

            if (targetDirectory.Exists)
                targetDirectory.Delete(recursive: true);

            targetDirectory.Create();
            WriteSkillFiles(targetDirectory, files);

            installed.Add(
                new InstalledSkillResult { Target = target, Path = RelativeToCurrentDirectory(skillFile.FullName) }
            );
        }

        return new InstallResult { Installed = installed, Version = VersionCommand.GetCleanVersion() };
    }

    public static InstallPlanResult Plan(string? targetValue) =>
        new()
        {
            Paths = ResolveTargets(targetValue)
                .Select(target =>
                    RelativeToCurrentDirectory(Path.Combine(ResolveTargetDirectory(target).FullName, "SKILL.md"))
                )
                .ToList(),
        };

    private static IReadOnlyList<string> ResolveTargets(string? targetValue)
    {
        var normalized = string.IsNullOrWhiteSpace(targetValue) ? "AGENTS" : targetValue.Trim().ToUpperInvariant();
        return normalized switch
        {
            "CLAUDE" => ["claude"],
            "AGENTS" => ["agents"],
            "ALL" => ["agents", "claude"],
            _ => throw CliException.InvalidArguments(
                $"Invalid value for --skills: '{targetValue}'. Allowed values are: claude, agents, all."
            ),
        };
    }

    private static DirectoryInfo ResolveTargetDirectory(string target) =>
        new(Path.Combine(Directory.GetCurrentDirectory(), $".{target}", "skills", "clippit"));

    private static Dictionary<string, string> ReadBundledSkillFiles()
    {
        var files = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var relativePath in s_skillFiles)
        {
            var resourceName = FindResourceName(relativePath);
            using var stream = s_assembly.GetManifestResourceStream(resourceName);
            if (stream is null)
            {
                throw new FileNotFoundException($"Bundled Clippit skill resource could not be loaded: {resourceName}");
            }

            using var reader = new StreamReader(stream);
            var content = reader.ReadToEnd();
            files[relativePath] =
                relativePath == "SKILL.md"
                    ? content.Replace(
                        "<!-- clippit-skill-version: bundled -->",
                        $"<!-- clippit-skill-version: {VersionCommand.GetCleanVersion()} -->",
                        StringComparison.Ordinal
                    )
                    : content;
        }

        return files;
    }

    private static string FindResourceName(string relativePath)
    {
        var suffix = relativePath.Replace('/', '.');
        var resourceName = s_assembly
            .GetManifestResourceNames()
            .FirstOrDefault(name => name.EndsWith(suffix, StringComparison.Ordinal));

        if (resourceName is null)
        {
            throw new FileNotFoundException(
                $"Bundled Clippit skill resource was not found for '{relativePath}'. "
                    + $"Available resources: {string.Join(", ", s_assembly.GetManifestResourceNames())}"
            );
        }

        return resourceName;
    }

    private static void WriteSkillFiles(DirectoryInfo targetDirectory, Dictionary<string, string> files)
    {
        foreach (var (relativePath, content) in files)
        {
            var destination = new FileInfo(Path.Combine(targetDirectory.FullName, relativePath));
            destination.Directory?.Create();
            File.WriteAllText(destination.FullName, content);
        }
    }

    private static string RelativeToCurrentDirectory(string path) =>
        Path.GetRelativePath(Directory.GetCurrentDirectory(), path).Replace(Path.DirectorySeparatorChar, '/');
}

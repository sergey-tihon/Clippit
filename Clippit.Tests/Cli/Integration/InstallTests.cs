namespace Clippit.Tests.Cli.Integration;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
internal sealed class InstallTests : CliIntegrationTestBase
{
    [Test]
    public async Task CLI194_InstallSkills_InstallsAgentsSkillByDefault()
    {
        var directory = CliTestRunner.CreateTempDirectory(nameof(CLI194_InstallSkills_InstallsAgentsSkillByDefault));

        var result = await CliTestRunner
            .RunManagedAsync(directory, "install", "--skills", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert
            .That(File.Exists(Path.Combine(directory.FullName, ".agents", "skills", "clippit", "SKILL.md")))
            .IsTrue();
        await Assert
            .That(
                File.Exists(
                    Path.Combine(directory.FullName, ".agents", "skills", "clippit", "references", "workflows.md")
                )
            )
            .IsTrue();
        await Assert.That(Directory.Exists(Path.Combine(directory.FullName, ".claude"))).IsFalse();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("installed").GetArrayLength()).IsEqualTo(1);
        var installed = json.RootElement.GetProperty("installed")[0];
        await Assert.That(installed.GetProperty("target").GetString()).IsEqualTo("agents");
        await Assert.That(installed.GetProperty("path").GetString()).IsEqualTo(".agents/skills/clippit/SKILL.md");
        await Assert.That(json.RootElement.GetProperty("version").GetString()).IsNotNull();
    }

    [Test]
    public async Task CLI195_InstallSkillsAgents_InstallsAgentsSkill()
    {
        var directory = CliTestRunner.CreateTempDirectory(nameof(CLI195_InstallSkillsAgents_InstallsAgentsSkill));

        var result = await CliTestRunner.RunManagedAsync(directory, "install", "--skills=agents").ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert
            .That(File.Exists(Path.Combine(directory.FullName, ".agents", "skills", "clippit", "SKILL.md")))
            .IsTrue();
        await Assert.That(result.StandardOutput).Contains(".agents/skills/clippit/SKILL.md");
    }

    [Test]
    public async Task CLI196_InstallSkillsAll_InstallsBothTargets()
    {
        var directory = CliTestRunner.CreateTempDirectory(nameof(CLI196_InstallSkillsAll_InstallsBothTargets));

        var result = await CliTestRunner
            .RunManagedAsync(directory, "install", "--skills=all", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert
            .That(File.Exists(Path.Combine(directory.FullName, ".claude", "skills", "clippit", "SKILL.md")))
            .IsTrue();
        await Assert
            .That(File.Exists(Path.Combine(directory.FullName, ".agents", "skills", "clippit", "SKILL.md")))
            .IsTrue();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("installed").GetArrayLength()).IsEqualTo(2);
    }

    [Test]
    public async Task CLI197_InstallSkills_IsIdempotent()
    {
        var directory = CliTestRunner.CreateTempDirectory(nameof(CLI197_InstallSkills_IsIdempotent));

        var first = await CliTestRunner.RunManagedAsync(directory, "install", "--skills").ConfigureAwait(false);
        var second = await CliTestRunner.RunManagedAsync(directory, "install", "--skills").ConfigureAwait(false);

        await Assert.That(first.ExitCode).IsEqualTo(0);
        await Assert.That(second.ExitCode).IsEqualTo(0);
        await Assert.That(second.StandardError).IsEmpty();
        await Assert
            .That(File.Exists(Path.Combine(directory.FullName, ".agents", "skills", "clippit", "SKILL.md")))
            .IsTrue();
    }

    [Test]
    public async Task CLI198_InstallSkillsDryRun_DoesNotWriteFiles()
    {
        var directory = CliTestRunner.CreateTempDirectory(nameof(CLI198_InstallSkillsDryRun_DoesNotWriteFiles));

        var result = await CliTestRunner
            .RunManagedAsync(directory, "install", "--skills=all", "--dry-run", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("paths").GetArrayLength()).IsEqualTo(2);
        var paths = json.RootElement.GetProperty("paths").EnumerateArray().Select(path => path.GetString()).ToList();
        await Assert.That(paths).Contains(".agents/skills/clippit/SKILL.md");
        await Assert.That(paths).Contains(".claude/skills/clippit/SKILL.md");
        await Assert.That(Directory.Exists(Path.Combine(directory.FullName, ".claude"))).IsFalse();
        await Assert.That(Directory.Exists(Path.Combine(directory.FullName, ".agents"))).IsFalse();
    }

    [Test]
    public async Task CLI199_InstallSkills_OverwritesModifiedFilesAndRemovesExtraFiles()
    {
        var directory = CliTestRunner.CreateTempDirectory(
            nameof(CLI199_InstallSkills_OverwritesModifiedFilesAndRemovesExtraFiles)
        );

        var first = await CliTestRunner.RunManagedAsync(directory, "install", "--skills").ConfigureAwait(false);
        await Assert.That(first.ExitCode).IsEqualTo(0);

        var skillPath = Path.Combine(directory.FullName, ".agents", "skills", "clippit", "SKILL.md");
        await File.WriteAllTextAsync(skillPath, "modified skill content").ConfigureAwait(false);

        var extraPath = Path.Combine(directory.FullName, ".agents", "skills", "clippit", "references", "legacy.md");
        await File.WriteAllTextAsync(extraPath, "old reference").ConfigureAwait(false);

        var second = await CliTestRunner.RunManagedAsync(directory, "install", "--skills").ConfigureAwait(false);
        await Assert.That(second.ExitCode).IsEqualTo(0);

        var content = await File.ReadAllTextAsync(skillPath).ConfigureAwait(false);
        await Assert.That(content).DoesNotContain("modified skill content");
        await Assert.That(content).Contains("<!-- clippit-skill-version:");
        await Assert.That(File.Exists(extraPath)).IsFalse();
    }
}
#pragma warning restore CA1707

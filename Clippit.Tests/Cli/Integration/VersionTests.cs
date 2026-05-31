namespace Clippit.Tests.Cli.Integration;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
internal sealed class VersionTests : CliIntegrationTestBase
{
    [Test]
    public async Task CLI001_Version_ReturnsJson()
    {
        var result = await CliTestRunner.RunManagedAsync("version").ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("version").GetString()).IsNotNull();
        await Assert.That(json.RootElement.GetProperty("openXmlSdkVersion").GetString()).IsNotNull();
    }

    [Test]
    public async Task CLI001a_RootVersion_MatchesVersionSubcommand()
    {
        var rootResult = await CliTestRunner.RunManagedAsync("--version").ConfigureAwait(false);
        var versionResult = await CliTestRunner.RunManagedAsync("version").ConfigureAwait(false);

        await Assert.That(rootResult.ExitCode).IsEqualTo(0);
        await Assert.That(rootResult.StandardError).IsEmpty();
        await Assert.That(versionResult.ExitCode).IsEqualTo(0);
        await Assert.That(versionResult.StandardError).IsEmpty();

        // --version and `version` must produce byte-identical structured JSON
        // so downstream tooling can rely on either entry point.
        await Assert.That(rootResult.StandardOutput.Trim()).IsEqualTo(versionResult.StandardOutput.Trim());

        using var rootJson = rootResult.ReadStdoutJson();
        await Assert.That(rootJson.RootElement.GetProperty("version").GetString()).IsNotNull();
        // Informational version suffix (commit hash) must be stripped.
        await Assert.That(rootResult.StandardOutput).DoesNotContain("+");
    }

    [Test]
    public async Task CLI001b_RootHelp_ShowsCommands()
    {
        var result = await CliTestRunner.RunManagedAsync("--help").ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(result.StandardOutput).Contains("Usage:");
        await Assert.That(result.StandardOutput).Contains("pptx");
        await Assert.That(result.StandardOutput).Contains("version");
    }

    [Test]
    public async Task CLI001c_HelpAlias_ShowsNestedCommandHelp()
    {
        var result = await CliTestRunner.RunManagedAsync("help", "pptx", "split").ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(result.StandardOutput).Contains("--slides");
        await Assert.That(result.StandardOutput).Contains("Examples:");
        await Assert.That(result.StandardOutput).Contains("clippit pptx split deck.pptx --slides 1,3,6-9");
    }
}
#pragma warning restore CA1707

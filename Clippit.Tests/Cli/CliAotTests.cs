using System.Text.Json;
using TUnit.Core;

namespace Clippit.Tests.Cli;

[NotInParallel]
#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
internal sealed class CliAotTests
{
    [Test]
    [SkipUnlessAotTestsEnabled]
    public async Task CLIAOT001_NativeAotSmokeTests()
    {
        var executable = await PublishNativeAotAsync().ConfigureAwait(false);

        var version = await RunAotAsync(executable, "version").ConfigureAwait(false);
        await Assert.That(version.ExitCode).IsEqualTo(0);
        await Assert.That(version.StandardError).IsEmpty();
        using (var json = version.ReadStdoutJson())
        {
            await Assert.That(json.RootElement.GetProperty("version").GetString()).IsNotNull();
        }

        var directory = CliTestRunner.CreateTempDirectory("aot-smoke");
        var manifest = new FileInfo(Path.Combine(directory.FullName, "deck.json"));
        var init = await RunAotAsync(
                executable,
                "pptx",
                "build",
                "init",
                "--output",
                manifest.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(init.ExitCode).IsEqualTo(0);
        await Assert.That(init.StandardError).IsEmpty();
        await Assert.That(manifest.Exists).IsTrue();

        WriteSmokeManifest(manifest, CliTestRunner.TestFile("PB001-Input1.pptx"));
        var run = await RunAotAsync(executable, "pptx", "build", "run", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(run.ExitCode).IsEqualTo(0);
        await Assert.That(run.StandardError).IsEmpty();
        using (var json = run.ReadStdoutJson())
        {
            await Assert.That(json.RootElement.GetProperty("totalSlides").GetInt32()).IsGreaterThan(0);
            await Assert.That(json.RootElement.GetProperty("entries").GetArrayLength()).IsGreaterThan(0);
        }
        await Assert.That(File.Exists(Path.Combine(directory.FullName, "aot-smoke.pptx"))).IsTrue();

        var splitOutput = CliTestRunner.CreateTempDirectory("aot-split");
        var split = await RunAotAsync(
                executable,
                "pptx",
                "split",
                CliTestRunner.TestFile("PB001-Input1.pptx").FullName,
                "--output",
                splitOutput.FullName,
                "--slides",
                "1",
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(split.ExitCode).IsEqualTo(0);
        await Assert.That(split.StandardError).IsEmpty();
        using (var json = split.ReadStdoutJson())
        {
            await Assert.That(json.RootElement.GetProperty("count").GetInt32()).IsEqualTo(1);
        }
    }

    private static async Task<FileInfo> PublishNativeAotAsync()
    {
        var outputDirectory = CliTestRunner.CreateTempDirectory("aot-publish");
        var project = Path.Combine(CliTestRunner.RepositoryRoot.FullName, "Clippit.Cli", "Clippit.Cli.csproj");
        var result = await CliTestRunner
            .RunAsync(
                "dotnet",
                [
                    "publish",
                    project,
                    "-c",
                    "Release",
                    "--self-contained",
                    "-p:NativeAot=true",
                    "-o",
                    outputDirectory.FullName,
                ],
                CliTestRunner.RepositoryRoot.FullName,
                TimeSpan.FromMinutes(10)
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);

        var executableNames = OperatingSystem.IsWindows()
            ? new[] { "clippit.exe", "Clippit.Cli.exe" }
            : ["clippit", "Clippit.Cli"];
        var executable = executableNames
            .Select(name => new FileInfo(Path.Combine(outputDirectory.FullName, name)))
            .FirstOrDefault(file => file.Exists);

        if (executable is null)
        {
            throw new FileNotFoundException(
                "Native AOT executable was not found in publish output: "
                    + string.Join(", ", outputDirectory.GetFiles().Select(file => file.Name))
            );
        }

        return executable;
    }

    private static Task<CliResult> RunAotAsync(FileInfo executable, params string[] arguments) =>
        CliTestRunner.RunAsync(
            executable.FullName,
            arguments,
            CliTestRunner.RepositoryRoot.FullName,
            TimeSpan.FromMinutes(2)
        );

    private static void WriteSmokeManifest(FileInfo manifest, FileInfo source)
    {
        var json = JsonSerializer.Serialize(
            new
            {
                title = "Native AOT Smoke Test",
                output = "aot-smoke.pptx",
                deck = new[] { "[Smoke]", source.FullName },
            }
        );
        File.WriteAllText(manifest.FullName, json);
    }
}
#pragma warning restore CA1707

internal sealed class SkipUnlessAotTestsEnabledAttribute()
    : SkipAttribute("Set CLIPPIT_TEST_AOT=1 to run Native AOT CLI smoke tests.")
{
    public override Task<bool> ShouldSkip(TestRegisteredContext context) =>
        Task.FromResult(
            !string.Equals(Environment.GetEnvironmentVariable("CLIPPIT_TEST_AOT"), "1", StringComparison.Ordinal)
        );
}

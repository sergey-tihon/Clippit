using System.Runtime.InteropServices;
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
        var executable = StageNpmPackageFiles(await PublishNativeAotAsync().ConfigureAwait(false));

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
        var htmlOutput = new FileInfo(Path.Combine(directory.FullName, "tabs.html"));
        var toHtml = await RunAotAsync(
                executable,
                "word",
                "to-html",
                CliTestRunner.TestFile("HC024-Tabs-01.docx").FullName,
                "--output",
                htmlOutput.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(toHtml.ExitCode).IsEqualTo(0);
        await Assert.That(toHtml.StandardError).IsEmpty();
        await Assert.That(htmlOutput.Exists).IsTrue();
    }

    private static async Task<DirectoryInfo> PublishNativeAotAsync()
    {
        var outputDirectory = CliTestRunner.CreateTempDirectory("aot-publish");
        var project = Path.Combine(CliTestRunner.RepositoryRoot.FullName, "Clippit.Cli", "Clippit.Cli.csproj");
        var result = await CliTestRunner
            .RunAsync(
                "dotnet",
                [
                    "publish",
                    project,
                    "-f",
                    "net11.0",
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
        return outputDirectory;
    }

    private static FileInfo StageNpmPackageFiles(DirectoryInfo publishDirectory)
    {
        var os =
            OperatingSystem.IsWindows() ? "win32"
            : OperatingSystem.IsMacOS() ? "darwin"
            : OperatingSystem.IsLinux() ? "linux"
            : throw new PlatformNotSupportedException();
        var arch = RuntimeInformation.OSArchitecture switch
        {
            Architecture.X64 => "x64",
            Architecture.Arm64 => "arm64",
            _ => throw new PlatformNotSupportedException(
                $"Unsupported architecture: {RuntimeInformation.OSArchitecture}"
            ),
        };
        var packageDirectory = Path.Combine(CliTestRunner.RepositoryRoot.FullName, "npm", $"clippit-{os}-{arch}");
        using var packageJson = JsonDocument.Parse(File.ReadAllText(Path.Combine(packageDirectory, "package.json")));
        var stagedDirectory = CliTestRunner.CreateTempDirectory("aot-npm");

        foreach (var file in packageJson.RootElement.GetProperty("files").EnumerateArray())
        {
            var name = file.GetString() ?? throw new InvalidDataException("npm package file entries must be strings.");
            var source = new FileInfo(Path.Combine(publishDirectory.FullName, name));

            if (!source.Exists)
            {
                throw new FileNotFoundException(
                    $"npm package file '{name}' was not found in publish output: {string.Join(", ", publishDirectory.GetFiles().Select(file => file.Name))}"
                );
            }

            source.CopyTo(Path.Combine(stagedDirectory.FullName, name));
        }

        var executableName = OperatingSystem.IsWindows() ? "clippit.exe" : "clippit";
        return new FileInfo(Path.Combine(stagedDirectory.FullName, executableName));
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

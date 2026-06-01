using System.Text.Json;

namespace Clippit.Tests.Cli.Integration.Pptx;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
internal sealed class PptxBuildTests : CliIntegrationTestBase
{
    private const string DeckManifestSchema = "https://sergey-tihon.github.io/Clippit/schemas/deck-manifest.v1.json";

    [Test]
    public async Task CLI002_PptxBuildInit_CreatesManifest()
    {
        var directory = CliTestRunner.CreateTempDirectory("build-init");
        var manifest = new FileInfo(Path.Combine(directory.FullName, "deck.json"));

        var result = await CliTestRunner
            .RunManagedAsync("pptx", "build", "init", "--output", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(manifest.Exists).IsTrue();

        using var outputJson = result.ReadStdoutJson();
        await Assert.That(outputJson.RootElement.GetProperty("manifest").GetString()).IsEqualTo(manifest.FullName);

        var text = await File.ReadAllTextAsync(manifest.FullName).ConfigureAwait(false);
        using var manifestJson = JsonDocument.Parse(text);
        await Assert.That(manifestJson.RootElement.GetProperty("$schema").GetString()).IsEqualTo(DeckManifestSchema);
        await Assert.That(manifestJson.RootElement.GetProperty("title").GetString()).IsEqualTo("My Presentation");
        await Assert.That(manifestJson.RootElement.GetProperty("deck").GetArrayLength()).IsEqualTo(2);
    }

    [Test]
    public async Task CLI003_PptxBuildRun_WritesPresentationAndReportsCounts()
    {
        var directory = CliTestRunner.CreateTempDirectory("build-run");
        var outputName = "built.pptx";
        var manifest = await WriteManifestAsync(directory, outputName, CliTestRunner.TestFile("PB001-Input1.pptx"))
            .ConfigureAwait(false);
        var expectedSlides = CountSlides(CliTestRunner.TestFile("PB001-Input1.pptx"));
        var expectedOutput = Path.Combine(directory.FullName, outputName);

        var result = await CliTestRunner
            .RunManagedAsync("pptx", "build", "run", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(File.Exists(expectedOutput)).IsTrue();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("output").GetString()).IsEqualTo(expectedOutput);
        await Assert.That(json.RootElement.GetProperty("totalSlides").GetInt32()).IsEqualTo(expectedSlides);

        var entries = json.RootElement.GetProperty("entries");
        await Assert.That(entries.GetArrayLength()).IsEqualTo(2);
        await Assert.That(entries[0].GetProperty("section").GetString()).IsEqualTo("CLI Section");
        await Assert.That(entries[1].GetProperty("slides").GetInt32()).IsEqualTo(expectedSlides);

        await Assert.That(CountSlides(new FileInfo(expectedOutput))).IsEqualTo(expectedSlides);
    }

    [Test]
    public async Task CLI008_PptxBuildRunWithMissingSource_ReturnsJsonError()
    {
        var directory = CliTestRunner.CreateTempDirectory("build-missing-source");
        var manifest = await WriteManifestAsync(
                directory,
                "missing-source.pptx",
                new FileInfo(Path.Combine(directory.FullName, "missing.pptx"))
            )
            .ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync("pptx", "build", "run", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(3);
        await Assert.That(result.StandardOutput).IsEmpty();

        using var json = result.ReadStderrJson();
        await Assert.That(json.RootElement.GetProperty("code").GetString()).IsEqualTo("FILE_NOT_FOUND");
    }

    [Test]
    public async Task CLI009_InvalidFormat_ReturnsParserError()
    {
        var directory = CliTestRunner.CreateTempDirectory("invalid-format");
        var manifest = new FileInfo(Path.Combine(directory.FullName, "deck.json"));

        var result = await CliTestRunner
            .RunManagedAsync("pptx", "build", "init", "--output", manifest.FullName, "--format", "xml")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(1);
        await Assert
            .That(result.StandardError)
            .Contains("Invalid value for --format: 'xml'. Allowed values are: json, text.");
        await Assert.That(result.StandardOutput).Contains("Usage:");
        await Assert.That(manifest.Exists).IsFalse();
    }

    [Test]
    public async Task CLI012_PptxBuildInit_FailsOnExistingFile_UnlessForce()
    {
        var directory = CliTestRunner.CreateTempDirectory("init-collision");
        var manifest = new FileInfo(Path.Combine(directory.FullName, "deck.json"));

        var first = await CliTestRunner
            .RunManagedAsync("pptx", "build", "init", "--output", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);
        await Assert.That(first.ExitCode).IsEqualTo(0);

        var second = await CliTestRunner
            .RunManagedAsync("pptx", "build", "init", "--output", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);
        await Assert.That(second.ExitCode).IsEqualTo(5);
        using (var json = second.ReadStderrJson())
            await Assert.That(json.RootElement.GetProperty("code").GetString()).IsEqualTo("OUTPUT_ERROR");

        var forced = await CliTestRunner
            .RunManagedAsync("pptx", "build", "init", "--output", manifest.FullName, "--force", "--format", "json")
            .ConfigureAwait(false);
        await Assert.That(forced.ExitCode).IsEqualTo(0);
    }

    [Test]
    public async Task CLI013_PptxBuildInit_OutputDash_WritesManifestToStdout()
    {
        var result = await CliTestRunner
            .RunManagedAsync("pptx", "build", "init", "--output", "-")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        // The manifest is the only thing on stdout; success summary is suppressed when output goes to stdout.
        using var manifest = JsonDocument.Parse(result.StandardOutput);
        await Assert.That(manifest.RootElement.GetProperty("$schema").GetString()).IsEqualTo(DeckManifestSchema);
        await Assert.That(manifest.RootElement.GetProperty("deck").GetArrayLength()).IsEqualTo(2);
    }

    [Test]
    public async Task CLI015_PptxBuildRun_StdinManifest_StdoutBinary()
    {
        var directory = CliTestRunner.CreateTempDirectory("build-pipe");
        var source = CliTestRunner.TestFile("PB001-Input1.pptx");
        var manifest = await WriteManifestAsync(directory, "ignored.pptx", source).ConfigureAwait(false);
        var manifestBytes = await File.ReadAllBytesAsync(manifest.FullName).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedWithStdinAsync(manifestBytes, "pptx", "build", "run", "-", "--output", "-")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        // PPTX is a ZIP container — verify the magic header rather than try to open from memory.
        await Assert.That(result.StandardOutput.Length).IsGreaterThan(4);
        await Assert.That(result.StandardOutput[0]).IsEqualTo((byte)'P');
        await Assert.That(result.StandardOutput[1]).IsEqualTo((byte)'K');
    }

    [Test]
    public async Task CLI028_PptxBuildRun_BlankFileEntry_ReturnsInvalidArguments()
    {
        var directory = CliTestRunner.CreateTempDirectory("build-blank-file");
        var manifest = new FileInfo(Path.Combine(directory.FullName, "deck.json"));
        var json = JsonSerializer.Serialize(
            new
            {
                title = "Invalid",
                output = "out.pptx",
                deck = new[] { " " },
            }
        );
        await File.WriteAllTextAsync(manifest.FullName, json).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync("pptx", "build", "run", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(2);
        using var error = result.ReadStderrJson();
        await Assert.That(error.RootElement.GetProperty("code").GetString()).IsEqualTo("INVALID_ARGUMENTS");
    }

    [Test]
    public async Task CLI029_PptxBuildRun_KeepSectionsAfterManifestSection_ReturnsInvalidArguments()
    {
        var directory = CliTestRunner.CreateTempDirectory("build-keepsections-after-section");
        var manifest = new FileInfo(Path.Combine(directory.FullName, "deck.json"));
        var source = CliTestRunner.TestFile("PB001-Input1.pptx");
        var json = JsonSerializer.Serialize(
            new
            {
                title = "Invalid",
                output = "out.pptx",
                deck = new object[] { "[Section]", new { file = source.FullName, keepSections = true } },
            }
        );
        await File.WriteAllTextAsync(manifest.FullName, json).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync("pptx", "build", "run", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(2);
        using var error = result.ReadStderrJson();
        await Assert.That(error.RootElement.GetProperty("code").GetString()).IsEqualTo("INVALID_ARGUMENTS");
    }

    [Test]
    public async Task CLI030_PptxBuildRun_OutputCannotOverwriteSource()
    {
        var directory = CliTestRunner.CreateTempDirectory("build-output-over-source");
        var source = new FileInfo(Path.Combine(directory.FullName, "source.pptx"));
        File.Copy(CliTestRunner.TestFile("PB001-Input1.pptx").FullName, source.FullName);
        var sourceLength = source.Length;
        var manifest = new FileInfo(Path.Combine(directory.FullName, "deck.json"));
        var json = JsonSerializer.Serialize(
            new
            {
                title = "Invalid",
                output = source.Name,
                deck = new[] { source.FullName },
            }
        );
        await File.WriteAllTextAsync(manifest.FullName, json).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync("pptx", "build", "run", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);

        source.Refresh();
        await Assert.That(result.ExitCode).IsEqualTo(5);
        await Assert.That(source.Length).IsEqualTo(sourceLength);
        using var error = result.ReadStderrJson();
        await Assert.That(error.RootElement.GetProperty("code").GetString()).IsEqualTo("OUTPUT_ERROR");
    }

    [Test]
    public async Task CLI030a_PptxBuildRun_FailsOnExistingOutputUnlessForce()
    {
        var directory = CliTestRunner.CreateTempDirectory("build-output-collision");
        var outputName = "built.pptx";
        var source = CliTestRunner.TestFile("PB001-Input1.pptx");
        var manifest = await WriteManifestAsync(directory, outputName, source).ConfigureAwait(false);
        var output = new FileInfo(Path.Combine(directory.FullName, outputName));
        await File.WriteAllTextAsync(output.FullName, "existing").ConfigureAwait(false);

        var blocked = await CliTestRunner
            .RunManagedAsync("pptx", "build", "run", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(blocked.ExitCode).IsEqualTo(5);
        using (var error = blocked.ReadStderrJson())
            await Assert.That(error.RootElement.GetProperty("code").GetString()).IsEqualTo("OUTPUT_ERROR");
        await Assert.That(await File.ReadAllTextAsync(output.FullName).ConfigureAwait(false)).IsEqualTo("existing");

        var forced = await CliTestRunner
            .RunManagedAsync("pptx", "build", "run", manifest.FullName, "--force", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(forced.ExitCode).IsEqualTo(0);
        await Assert.That(CountSlides(output)).IsGreaterThan(0);
    }

    [Test]
    public async Task CLI030b_PptxBuildRun_FailedBuildKeepsExistingOutput()
    {
        var directory = CliTestRunner.CreateTempDirectory("build-output-atomic");
        var outputName = "built.pptx";
        var output = new FileInfo(Path.Combine(directory.FullName, outputName));
        var invalidSource = new FileInfo(Path.Combine(directory.FullName, "invalid.pptx"));
        var manifest = new FileInfo(Path.Combine(directory.FullName, "deck.json"));

        await File.WriteAllTextAsync(output.FullName, "existing output").ConfigureAwait(false);
        await File.WriteAllTextAsync(invalidSource.FullName, "not a presentation").ConfigureAwait(false);
        var json = JsonSerializer.Serialize(
            new
            {
                title = "Invalid",
                output = outputName,
                deck = new[] { invalidSource.FullName },
            }
        );
        await File.WriteAllTextAsync(manifest.FullName, json).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync("pptx", "build", "run", manifest.FullName, "--force", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(4);
        await Assert
            .That(await File.ReadAllTextAsync(output.FullName).ConfigureAwait(false))
            .IsEqualTo("existing output");
        await Assert.That(directory.EnumerateFiles("*.tmp")).IsEmpty();
        using var error = result.ReadStderrJson();
        await Assert.That(error.RootElement.GetProperty("code").GetString()).IsEqualTo("INVALID_FORMAT");
    }

    [Test]
    public async Task CLI031_PptxBuildRun_MultipleKeepSectionsSources_AppendsSectionsInOrder()
    {
        var directory = CliTestRunner.CreateTempDirectory("build-multiple-keepsections");
        var source = CliTestRunner.TestFile("PB001-Input1.pptx");

        var firstManifest = new FileInfo(Path.Combine(directory.FullName, "first.json"));
        var firstJson = JsonSerializer.Serialize(
            new
            {
                title = "First",
                output = "first.pptx",
                deck = new object[] { "[First A]", source.FullName, "[First B]", source.FullName },
            }
        );
        await File.WriteAllTextAsync(firstManifest.FullName, firstJson).ConfigureAwait(false);
        var firstBuild = await CliTestRunner
            .RunManagedAsync("pptx", "build", "run", firstManifest.FullName, "--format", "json")
            .ConfigureAwait(false);
        await Assert.That(firstBuild.ExitCode).IsEqualTo(0);

        var secondManifest = new FileInfo(Path.Combine(directory.FullName, "second.json"));
        var secondJson = JsonSerializer.Serialize(
            new
            {
                title = "Second",
                output = "second.pptx",
                deck = new object[] { "[Second A]", source.FullName, "[Second B]", source.FullName },
            }
        );
        await File.WriteAllTextAsync(secondManifest.FullName, secondJson).ConfigureAwait(false);
        var secondBuild = await CliTestRunner
            .RunManagedAsync("pptx", "build", "run", secondManifest.FullName, "--format", "json")
            .ConfigureAwait(false);
        await Assert.That(secondBuild.ExitCode).IsEqualTo(0);

        var combinedManifest = new FileInfo(Path.Combine(directory.FullName, "combined.json"));
        var combinedJson = JsonSerializer.Serialize(
            new
            {
                title = "Combined",
                output = "combined.pptx",
                deck = new object[]
                {
                    new { file = "first.pptx", keepSections = true },
                    new { file = "second.pptx", keepSections = true },
                },
            }
        );
        await File.WriteAllTextAsync(combinedManifest.FullName, combinedJson).ConfigureAwait(false);

        var combinedBuild = await CliTestRunner
            .RunManagedAsync("pptx", "build", "run", combinedManifest.FullName, "--format", "json")
            .ConfigureAwait(false);

        var combined = new FileInfo(Path.Combine(directory.FullName, "combined.pptx"));
        await Assert.That(combinedBuild.ExitCode).IsEqualTo(0);
        await Assert.That(GetSectionNames(combined)).IsEquivalentTo(["First A", "First B", "Second A", "Second B"]);
        await Assert.That(SectionListUsesNumericSlideIds(combined)).IsTrue();
    }
}
#pragma warning restore CA1707

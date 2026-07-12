using System.Text.Json;
using Clippit.Tests.Word;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Cli.Integration.Word;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
internal sealed class WordBuildTests : CliIntegrationTestBase
{
    private const string WordBuildManifestSchema =
        "https://sergey-tihon.github.io/Clippit/schemas/word-build-manifest.v1.json";

    // Known schema validation issues in DocumentBuilder output (attribute not declared for newer schema versions,
    // duplicate IDs that DocumentBuilder doesn't deduplicate, extension elements from Word 2010/2012, etc.).
    private static readonly List<string> s_expectedErrors =
    [
        .. WmlComparerTests.ExpectedErrors,
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:evenHBand' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:evenVBand' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRowFirstColumn' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRowLastColumn' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRowFirstColumn' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRowLastColumn' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:oddHBand' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:oddVBand' attribute is not declared.",
        "The element has unexpected child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:updateFields'.",
        "The attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:name' has invalid value 'useWord2013TrackBottomHyphenation'. The Enumeration constraint failed.",
        "The 'http://schemas.microsoft.com/office/word/2012/wordml:restartNumberingAfterBreak' attribute is not declared.",
        "Attribute 'id' should have unique value. Its current value '",
        "The 'urn:schemas-microsoft-com:mac:vml:blur' attribute is not declared.",
        "Attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:id' should have unique value. Its current value '",
        "The element has unexpected child element 'http://schemas.microsoft.com/office/word/2012/wordml:",
        "The element has invalid child element 'http://schemas.microsoft.com/office/word/2012/wordml:",
        "The 'urn:schemas-microsoft-com:mac:vml:complextextbox' attribute is not declared.",
        "http://schemas.microsoft.com/office/word/2010/wordml:",
        "http://schemas.microsoft.com/office/word/2008/9/12/wordml:",
    ];

    private static FileInfo Source1 => CliTestRunner.TestFile("DB006-Source1.docx");
    private static FileInfo Source2 => CliTestRunner.TestFile("DB007-WhitePaper.docx");
    private static FileInfo Source3 => CliTestRunner.TestFile("DB007-Spec.docx");

    [Test]
    public async Task CLI180_WordBuildInit_CreatesManifest()
    {
        var directory = CliTestRunner.CreateTempDirectory("word-build-init");
        var manifest = new FileInfo(Path.Combine(directory.FullName, "word-build.json"));

        var result = await CliTestRunner
            .RunManagedAsync("word", "build", "init", "--output", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(manifest.Exists).IsTrue();

        using var outputJson = result.ReadStdoutJson();
        await Assert.That(outputJson.RootElement.GetProperty("manifest").GetString()).IsEqualTo(manifest.FullName);

        var text = await File.ReadAllTextAsync(manifest.FullName).ConfigureAwait(false);
        using var manifestJson = JsonDocument.Parse(text);
        await Assert
            .That(manifestJson.RootElement.GetProperty("$schema").GetString())
            .IsEqualTo(WordBuildManifestSchema);
        await Assert.That(manifestJson.RootElement.GetProperty("output").GetString()).IsEqualTo("merged.docx");
        await Assert.That(manifestJson.RootElement.GetProperty("entries").GetArrayLength()).IsEqualTo(2);
    }

    [Test]
    public async Task CLI181_WordBuildRun_SingleSource_WritesDocxAndReportsCounts()
    {
        var directory = CliTestRunner.CreateTempDirectory("word-build-single");
        var outputName = "merged.docx";
        var manifest = await WriteWordManifestAsync(directory, outputName, Source1).ConfigureAwait(false);
        var expectedOutput = new FileInfo(Path.Combine(directory.FullName, outputName));

        var result = await CliTestRunner
            .RunManagedAsync("word", "build", "run", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(expectedOutput.Exists).IsTrue();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("output").GetString()).IsEqualTo(expectedOutput.FullName);
        await Assert.That(json.RootElement.GetProperty("outputSize").GetInt64()).IsGreaterThan(0);
        await Assert.That(json.RootElement.GetProperty("entryCount").GetInt32()).IsEqualTo(2);

        var entries = json.RootElement.GetProperty("entries");
        await Assert.That(entries.GetArrayLength()).IsEqualTo(2);
        await Assert.That(entries[0].GetProperty("section").GetString()).IsEqualTo("CLI Section");
        await Assert.That(entries[1].GetProperty("elements").GetInt32()).IsGreaterThan(0);

        using var doc = WordprocessingDocument.Open(expectedOutput.FullName, false);
        await Validate(doc, s_expectedErrors);
    }

    [Test]
    public async Task CLI182_WordBuildRun_MultipleSources_MergesAllDocuments()
    {
        var directory = CliTestRunner.CreateTempDirectory("word-build-multi");
        var manifest = new FileInfo(Path.Combine(directory.FullName, "word-build.json"));
        var outputName = "merged.docx";
        var json = JsonSerializer.Serialize(
            new
            {
                output = outputName,
                entries = new object[]
                {
                    new { section = "Part 1" },
                    Source1.FullName,
                    new { section = "Part 2" },
                    Source2.FullName,
                },
            }
        );
        await File.WriteAllTextAsync(manifest.FullName, json).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync("word", "build", "run", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var resultJson = result.ReadStdoutJson();
        await Assert.That(resultJson.RootElement.GetProperty("entryCount").GetInt32()).IsEqualTo(4);

        var entries = resultJson.RootElement.GetProperty("entries");
        await Assert.That(entries.GetArrayLength()).IsEqualTo(4);
        await Assert.That(entries[0].GetProperty("section").GetString()).IsEqualTo("Part 1");
        await Assert.That(entries[2].GetProperty("section").GetString()).IsEqualTo("Part 2");
        await Assert.That(entries[1].GetProperty("elements").GetInt32()).IsGreaterThan(0);
        await Assert.That(entries[3].GetProperty("elements").GetInt32()).IsGreaterThan(0);

        var output = new FileInfo(Path.Combine(directory.FullName, outputName));
        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await Validate(doc, s_expectedErrors);
    }

    [Test]
    public async Task CLI183_WordBuildRun_StartAndCount_ExcerptsElements()
    {
        var directory = CliTestRunner.CreateTempDirectory("word-build-excerpt");
        var manifest = new FileInfo(Path.Combine(directory.FullName, "word-build.json"));
        var outputName = "excerpt.docx";
        var json = JsonSerializer.Serialize(
            new
            {
                output = outputName,
                entries = new object[]
                {
                    new
                    {
                        file = Source1.FullName,
                        start = 0,
                        count = 5,
                    },
                },
            }
        );
        await File.WriteAllTextAsync(manifest.FullName, json).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync("word", "build", "run", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var resultJson = result.ReadStdoutJson();
        var entries = resultJson.RootElement.GetProperty("entries");
        await Assert.That(entries[0].GetProperty("elements").GetInt32()).IsEqualTo(5);

        var output = new FileInfo(Path.Combine(directory.FullName, outputName));
        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await Validate(doc, s_expectedErrors);
    }

    [Test]
    public async Task CLI184_WordBuildInit_OutputDash_WritesManifestToStdout()
    {
        var result = await CliTestRunner
            .RunManagedAsync("word", "build", "init", "--output", "-")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        using var manifest = JsonDocument.Parse(result.StandardOutput);
        await Assert.That(manifest.RootElement.GetProperty("$schema").GetString()).IsEqualTo(WordBuildManifestSchema);
        await Assert.That(manifest.RootElement.GetProperty("entries").GetArrayLength()).IsEqualTo(2);
    }

    [Test]
    public async Task CLI185_WordBuildInit_FailsOnExistingFile_UnlessForce()
    {
        var directory = CliTestRunner.CreateTempDirectory("word-build-init-force");
        var manifest = new FileInfo(Path.Combine(directory.FullName, "word-build.json"));

        var first = await CliTestRunner
            .RunManagedAsync("word", "build", "init", "--output", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);
        await Assert.That(first.ExitCode).IsEqualTo(0);

        var second = await CliTestRunner
            .RunManagedAsync("word", "build", "init", "--output", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);
        await Assert.That(second.ExitCode).IsEqualTo(5);
        using (var json = second.ReadStderrJson())
            await Assert.That(json.RootElement.GetProperty("code").GetString()).IsEqualTo("OUTPUT_ERROR");

        var forced = await CliTestRunner
            .RunManagedAsync("word", "build", "init", "--output", manifest.FullName, "--force", "--format", "json")
            .ConfigureAwait(false);
        await Assert.That(forced.ExitCode).IsEqualTo(0);
    }

    [Test]
    public async Task CLI186_WordBuildRun_MissingSourceFile_ReturnsFileNotFound()
    {
        var directory = CliTestRunner.CreateTempDirectory("word-build-missing");
        var manifest = new FileInfo(Path.Combine(directory.FullName, "word-build.json"));
        var json = JsonSerializer.Serialize(
            new { output = "out.docx", entries = new[] { Path.Combine(directory.FullName, "missing.docx") } }
        );
        await File.WriteAllTextAsync(manifest.FullName, json).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync("word", "build", "run", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(3);
        await Assert.That(result.StandardOutput).IsEmpty();
        using var error = result.ReadStderrJson();
        await Assert.That(error.RootElement.GetProperty("code").GetString()).IsEqualTo("FILE_NOT_FOUND");
    }

    [Test]
    public async Task CLI187_WordBuildRun_OutputCannotOverwriteSource()
    {
        var directory = CliTestRunner.CreateTempDirectory("word-build-overwrite");
        var source = new FileInfo(Path.Combine(directory.FullName, "source.docx"));
        File.Copy(Source1.FullName, source.FullName);
        var manifest = new FileInfo(Path.Combine(directory.FullName, "word-build.json"));
        var json = JsonSerializer.Serialize(new { output = source.Name, entries = new[] { source.FullName } });
        await File.WriteAllTextAsync(manifest.FullName, json).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync("word", "build", "run", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(5);
        using var error = result.ReadStderrJson();
        await Assert.That(error.RootElement.GetProperty("code").GetString()).IsEqualTo("OUTPUT_ERROR");
    }

    [Test]
    public async Task CLI188_WordBuildRun_FailsOnExistingOutputUnlessForce()
    {
        var directory = CliTestRunner.CreateTempDirectory("word-build-collision");
        var manifest = await WriteWordManifestAsync(directory, "merged.docx", Source1).ConfigureAwait(false);
        var output = new FileInfo(Path.Combine(directory.FullName, "merged.docx"));
        await File.WriteAllTextAsync(output.FullName, "existing").ConfigureAwait(false);

        var blocked = await CliTestRunner
            .RunManagedAsync("word", "build", "run", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);
        await Assert.That(blocked.ExitCode).IsEqualTo(5);
        using (var error = blocked.ReadStderrJson())
            await Assert.That(error.RootElement.GetProperty("code").GetString()).IsEqualTo("OUTPUT_ERROR");
        await Assert.That(await File.ReadAllTextAsync(output.FullName).ConfigureAwait(false)).IsEqualTo("existing");

        var forced = await CliTestRunner
            .RunManagedAsync("word", "build", "run", manifest.FullName, "--force", "--format", "json")
            .ConfigureAwait(false);
        await Assert.That(forced.ExitCode).IsEqualTo(0);
        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await Validate(doc, s_expectedErrors);
    }

    [Test]
    public async Task CLI189_WordBuildRun_InvalidManifestJson_ReturnsInvalidFormat()
    {
        var directory = CliTestRunner.CreateTempDirectory("word-build-invalid-json");
        var manifest = new FileInfo(Path.Combine(directory.FullName, "word-build.json"));
        await File.WriteAllTextAsync(manifest.FullName, "not valid json").ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync("word", "build", "run", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(4);
        using var error = result.ReadStderrJson();
        await Assert.That(error.RootElement.GetProperty("code").GetString()).IsEqualTo("INVALID_FORMAT");
    }

    [Test]
    public async Task CLI190_WordBuildRun_EmptyDeck_ReturnsInvalidArguments()
    {
        var directory = CliTestRunner.CreateTempDirectory("word-build-empty-deck");
        var manifest = new FileInfo(Path.Combine(directory.FullName, "word-build.json"));
        var json = JsonSerializer.Serialize(new { output = "out.docx", entries = Array.Empty<string>() });
        await File.WriteAllTextAsync(manifest.FullName, json).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync("word", "build", "run", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(2);
        using var error = result.ReadStderrJson();
        await Assert.That(error.RootElement.GetProperty("code").GetString()).IsEqualTo("INVALID_ARGUMENTS");
    }

    [Test]
    public async Task CLI191_WordBuildRun_OnlySectionEntries_ReturnsInvalidArguments()
    {
        var directory = CliTestRunner.CreateTempDirectory("word-build-sections-only");
        var manifest = new FileInfo(Path.Combine(directory.FullName, "word-build.json"));
        var json = JsonSerializer.Serialize(
            new { output = "out.docx", entries = new[] { "[Section A]", "[Section B]" } }
        );
        await File.WriteAllTextAsync(manifest.FullName, json).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync("word", "build", "run", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(2);
        using var error = result.ReadStderrJson();
        await Assert.That(error.RootElement.GetProperty("code").GetString()).IsEqualTo("INVALID_ARGUMENTS");
    }

    [Test]
    public async Task CLI192_WordBuildRun_StdinManifest_StdoutBinary()
    {
        var directory = CliTestRunner.CreateTempDirectory("word-build-pipe");
        var manifest = await WriteWordManifestAsync(directory, "ignored.docx", Source1).ConfigureAwait(false);
        var manifestBytes = await File.ReadAllBytesAsync(manifest.FullName).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedWithStdinAsync(manifestBytes, "word", "build", "run", "-", "--output", "-")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        // DOCX is a ZIP container — verify the magic header (PK).
        await Assert.That(result.StandardOutput.Length).IsGreaterThan(4);
        await Assert.That(result.StandardOutput[0]).IsEqualTo((byte)'P');
        await Assert.That(result.StandardOutput[1]).IsEqualTo((byte)'K');
    }

    [Test]
    public async Task CLI193_WordBuildRun_KeepSections_PreservesSourceSectionStructure()
    {
        var directory = CliTestRunner.CreateTempDirectory("word-build-keepsections");
        var manifest = new FileInfo(Path.Combine(directory.FullName, "word-build.json"));
        var json = JsonSerializer.Serialize(
            new
            {
                output = "merged.docx",
                entries = new object[]
                {
                    new { file = Source1.FullName, keepSections = true },
                    new { file = Source2.FullName, keepSections = true },
                },
            }
        );
        await File.WriteAllTextAsync(manifest.FullName, json).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync("word", "build", "run", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        var output = new FileInfo(Path.Combine(directory.FullName, "merged.docx"));
        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await Validate(doc, s_expectedErrors);
    }

    private static async Task<FileInfo> WriteWordManifestAsync(DirectoryInfo directory, string output, FileInfo source)
    {
        var manifest = new FileInfo(Path.Combine(directory.FullName, "word-build.json"));
        var json = JsonSerializer.Serialize(
            new { output, entries = new object[] { new { section = "CLI Section" }, source.FullName } }
        );
        await File.WriteAllTextAsync(manifest.FullName, json).ConfigureAwait(false);
        return manifest;
    }
}
#pragma warning restore CA1707

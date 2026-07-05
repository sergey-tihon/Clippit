using System.IO.Compression;
using System.Text;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Cli.Integration.Word;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
internal sealed class WordSimplifyMarkupTests : CliIntegrationTestBase
{
    [Test]
    public async Task CLI118_WordSimplifyMarkup_NoFlags_ReturnsInvalidArguments()
    {
        var input = CliTestRunner.TestFile("DB007-Spec.docx");
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-no-flags");
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync("word", "simplify-markup", input.FullName, "--output", output.FullName)
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(2);
        await Assert.That(result.StandardOutput).IsEmpty();
        await Assert.That(result.StandardError).Contains("INVALID_ARGUMENTS");
    }

    [Test]
    public async Task CLI119_WordSimplifyMarkup_AllFlag_ProducesValidDocx()
    {
        var input = CliTestRunner.TestFile("DB007-Spec.docx");
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-all");
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--all",
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("input").GetString()).IsEqualTo(input.FullName);
        await Assert.That(json.RootElement.GetProperty("output").GetString()).IsEqualTo(output.FullName);
        await Assert.That(json.RootElement.GetProperty("outputSize").GetInt64()).IsGreaterThan(0);
        await Assert.That(output.Exists).IsTrue();

        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await ValidateRelationships(doc);
        await Validate(doc);
    }

    [Test]
    public async Task CLI120_WordSimplifyMarkup_AcceptRevisions_RemovesTrackedChanges()
    {
        var input = CliTestRunner.TestFile("RA001-Tracked-Revisions-01.docx");
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-accept-revisions");
        var output = new FileInfo(Path.Combine(tempDir.FullName, "clean.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--accept-revisions",
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(output.Exists).IsTrue();

        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await Assert.That(RevisionAccepter.HasTrackedRevisions(doc)).IsFalse();
        await ValidateRelationships(doc);
    }

    [Test]
    public async Task CLI121_WordSimplifyMarkup_RemoveRsidInfo_EliminatesRsidAttributes()
    {
        var input = CliTestRunner.TestFile("DB007-Spec.docx");
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-rsid");
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--remove-rsid-info",
                "--output",
                output.FullName
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(output.Exists).IsTrue();

        // Verify no rsid attributes remain in document.xml
        using var zip = ZipFile.OpenRead(output.FullName);
        var docEntry = zip.GetEntry("word/document.xml");
        await Assert.That(docEntry).IsNotNull();

        using var stream = docEntry!.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        var xmlContent = await reader.ReadToEndAsync().ConfigureAwait(false);
        await Assert.That(xmlContent).DoesNotContain(":rsid");
    }

    [Test]
    public async Task CLI122_WordSimplifyMarkup_DefaultOutputPath_DerivedFromInput()
    {
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-default-output");
        var source = CliTestRunner.TestFile("DB007-Spec.docx");
        var localInput = new FileInfo(Path.Combine(tempDir.FullName, "DB007-Spec.docx"));
        File.Copy(source.FullName, localInput.FullName);

        var expectedOutput = new FileInfo(Path.Combine(tempDir.FullName, "DB007-Spec-simplified.docx"));

        var result = await CliTestRunner
            .RunManagedAsync("word", "simplify-markup", localInput.FullName, "--remove-rsid-info", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("output").GetString()).IsEqualTo(expectedOutput.FullName);
        await Assert.That(expectedOutput.Exists).IsTrue();
    }

    [Test]
    public async Task CLI123_WordSimplifyMarkup_QuietSuppressesStdout()
    {
        var input = CliTestRunner.TestFile("DB007-Spec.docx");
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-quiet");
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--remove-rsid-info",
                "--output",
                output.FullName,
                "--quiet"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardOutput).IsEmpty();
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(output.Exists).IsTrue();
    }

    [Test]
    public async Task CLI124_WordSimplifyMarkup_InvalidInput_ReturnsInvalidFormat()
    {
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-invalid-input");
        var badInput = new FileInfo(Path.Combine(tempDir.FullName, "not-a-docx.docx"));
        await File.WriteAllTextAsync(badInput.FullName, "not a valid DOCX").ConfigureAwait(false);
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                badInput.FullName,
                "--all",
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(4);
        await Assert.That(result.StandardOutput).IsEmpty();
        await Assert.That(result.StandardError).Contains("INVALID_FORMAT");
    }

    [Test]
    public async Task CLI125_WordSimplifyMarkup_OutputToStdout_StreamsBinaryDocx()
    {
        var input = CliTestRunner.TestFile("DB007-Spec.docx");
        var inputBytes = await File.ReadAllBytesAsync(input.FullName).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedWithStdinAsync(inputBytes, "word", "simplify-markup", "-", "--all", "--output", "-")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        // Output should be a DOCX (ZIP/PK header)
        await Assert.That(result.StandardOutput.Length).IsGreaterThan(0);
        await Assert.That(result.StandardOutput[0]).IsEqualTo((byte)'P');
        await Assert.That(result.StandardOutput[1]).IsEqualTo((byte)'K');
    }
}
#pragma warning restore CA1707

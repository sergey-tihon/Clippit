using System.Text;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Cli.Integration.Word;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
internal sealed class WordAcceptRevisionsTests : CliIntegrationTestBase
{
    [Test]
    public async Task CLI097_WordAcceptRevisions_DocumentWithRevisions_AcceptsAndWritesOutput()
    {
        var input = CliTestRunner.TestFile("RA001-Tracked-Revisions-01.docx");
        var tempDir = CliTestRunner.CreateTempDirectory("accept-revisions-basic");
        var output = new FileInfo(Path.Combine(tempDir.FullName, "clean.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "accept-revisions",
                input.FullName,
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

        // Verify all tracked revisions are gone
        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await Assert.That(RevisionAccepter.HasTrackedRevisions(doc)).IsFalse();
    }

    [Test]
    public async Task CLI098_WordAcceptRevisions_DefaultOutputPathDerivedFromInput()
    {
        var tempDir = CliTestRunner.CreateTempDirectory("accept-revisions-default-output");
        var source = CliTestRunner.TestFile("RA001-Tracked-Revisions-01.docx");
        var localInput = new FileInfo(Path.Combine(tempDir.FullName, "RA001-Tracked-Revisions-01.docx"));
        File.Copy(source.FullName, localInput.FullName);

        var expectedOutput = new FileInfo(Path.Combine(tempDir.FullName, "RA001-Tracked-Revisions-01-accepted.docx"));

        var result = await CliTestRunner
            .RunManagedAsync("word", "accept-revisions", localInput.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("output").GetString()).IsEqualTo(expectedOutput.FullName);
        await Assert.That(expectedOutput.Exists).IsTrue();
    }

    [Test]
    public async Task CLI099_WordAcceptRevisions_DocumentAlreadyClean_CopiesUnchanged()
    {
        // Use a document that has no tracked revisions — the output should still be a valid docx.
        var input = CliTestRunner.TestFile("Blank-wml.docx");
        var tempDir = CliTestRunner.CreateTempDirectory("accept-revisions-clean");
        var output = new FileInfo(Path.Combine(tempDir.FullName, "clean.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "accept-revisions",
                input.FullName,
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
        await Validate(doc);
    }

    [Test]
    public async Task CLI100_WordAcceptRevisions_ReadsFromStdin_ProducesOutput()
    {
        var input = CliTestRunner.TestFile("RA001-Tracked-Revisions-01.docx");
        var tempDir = CliTestRunner.CreateTempDirectory("accept-revisions-stdin");
        var output = new FileInfo(Path.Combine(tempDir.FullName, "clean.docx"));
        var inputBytes = await File.ReadAllBytesAsync(input.FullName).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedWithStdinAsync(
                inputBytes,
                "word",
                "accept-revisions",
                "-",
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        var stdout = Encoding.UTF8.GetString(result.StandardOutput);
        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = System.Text.Json.JsonDocument.Parse(stdout);
        await Assert.That(json.RootElement.GetProperty("input").GetString()).IsEqualTo("<stdin>");
        await Assert.That(output.Exists).IsTrue();
    }

    [Test]
    public async Task CLI101_WordAcceptRevisions_QuietSuppressesStdout()
    {
        var input = CliTestRunner.TestFile("RA001-Tracked-Revisions-01.docx");
        var tempDir = CliTestRunner.CreateTempDirectory("accept-revisions-quiet");
        var output = new FileInfo(Path.Combine(tempDir.FullName, "clean.docx"));

        var result = await CliTestRunner
            .RunManagedAsync("word", "accept-revisions", input.FullName, "--output", output.FullName, "--quiet")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardOutput).IsEmpty();
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(output.Exists).IsTrue();
    }

    [Test]
    public async Task CLI102_WordAcceptRevisions_ExistingOutputWithoutForce_ReturnsOutputError()
    {
        var input = CliTestRunner.TestFile("RA001-Tracked-Revisions-01.docx");
        var tempDir = CliTestRunner.CreateTempDirectory("accept-revisions-no-force");
        var output = new FileInfo(Path.Combine(tempDir.FullName, "clean.docx"));
        await File.WriteAllTextAsync(output.FullName, "placeholder").ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "accept-revisions",
                input.FullName,
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(5);
        await Assert.That(result.StandardOutput).IsEmpty();
        await Assert.That(result.StandardError).Contains("OUTPUT_ERROR");
    }
}
#pragma warning restore CA1707

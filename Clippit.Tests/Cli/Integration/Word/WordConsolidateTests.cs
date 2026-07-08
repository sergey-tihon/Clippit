using System.Text;
using Clippit.Tests.Word;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Cli.Integration.Word;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
internal sealed class WordConsolidateTests : CliIntegrationTestBase
{
    private static readonly string s_original = CliTestRunner.TestFile("WC/WC027-Twenty-Paras-Before.docx").FullName;
    private static readonly string s_rev1 = CliTestRunner.TestFile("WC/WC027-Twenty-Paras-After-1.docx").FullName;
    private static readonly string s_rev2 = CliTestRunner.TestFile("WC/WC027-Twenty-Paras-After-2.docx").FullName;
    private static readonly string s_rev3 = CliTestRunner.TestFile("WC/WC027-Twenty-Paras-After-3.docx").FullName;

    // WmlComparer.Consolidate produces tblLook attributes that reference conditional
    // table-style attributes; these are pre-existing schema warnings from the comparer.
    private static readonly List<string> s_expectedErrors = WmlComparerTests.ExpectedErrors.ToList();

    [Test]
    public async Task CLI164_WordConsolidate_SingleRevision_ProducesValidOutput()
    {
        var output = new FileInfo(Path.Combine(TempDir, "consolidated-1rev.docx"));

        var result = await CliTestRunner
            .RunManagedAsync("word", "consolidate", s_original, s_rev1, "--output", output.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("original").GetString()).IsEqualTo(s_original);
        await Assert.That(json.RootElement.GetProperty("output").GetString()).IsEqualTo(output.FullName);
        await Assert.That(json.RootElement.GetProperty("outputSize").GetInt64()).IsGreaterThan(0);
        var revisions = json.RootElement.GetProperty("revisions");
        await Assert.That(revisions.GetArrayLength()).IsEqualTo(1);
        await Assert.That(revisions[0].GetProperty("file").GetString()).IsEqualTo(s_rev1);

        await Assert.That(output.Exists).IsTrue();
        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await Validate(doc, s_expectedErrors);
    }

    [Test]
    public async Task CLI165_WordConsolidate_MultipleRevisions_ProducesValidOutput()
    {
        var output = new FileInfo(Path.Combine(TempDir, "consolidated-3rev.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "consolidate",
                s_original,
                s_rev1,
                s_rev2,
                s_rev3,
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        var revisions = json.RootElement.GetProperty("revisions");
        await Assert.That(revisions.GetArrayLength()).IsEqualTo(3);
        await Assert.That(output.Exists).IsTrue();
        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await Validate(doc, s_expectedErrors);
    }

    [Test]
    public async Task CLI166_WordConsolidate_CustomRevisorAndColor_AppliedToResult()
    {
        var output = new FileInfo(Path.Combine(TempDir, "consolidated-revisor.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "consolidate",
                s_original,
                s_rev1,
                "--revisor",
                "Alice",
                "--color",
                "#FF0000",
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        var revisions = json.RootElement.GetProperty("revisions");
        await Assert.That(revisions[0].GetProperty("revisor").GetString()).IsEqualTo("Alice");
        await Assert.That(revisions[0].GetProperty("color").GetString()).IsEqualTo("#FF0000");
        await Assert.That(output.Exists).IsTrue();
        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await Validate(doc, s_expectedErrors);
    }

    [Test]
    public async Task CLI167_WordConsolidate_NoTableConsolidation_ProducesValidOutput()
    {
        var output = new FileInfo(Path.Combine(TempDir, "consolidated-notable.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "consolidate",
                s_original,
                s_rev1,
                "--no-table-consolidation",
                "--output",
                output.FullName
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(output.Exists).IsTrue();
        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await Validate(doc, s_expectedErrors);
    }

    [Test]
    public async Task CLI168_WordConsolidate_CaseInsensitive_ProducesValidOutput()
    {
        var output = new FileInfo(Path.Combine(TempDir, "consolidated-ci.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "consolidate",
                s_original,
                s_rev1,
                "--case-insensitive",
                "--output",
                output.FullName
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(output.Exists).IsTrue();
        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await Validate(doc, s_expectedErrors);
    }

    [Test]
    public async Task CLI169_WordConsolidate_MismatchedRevisorCount_ReturnsInvalidArguments()
    {
        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "consolidate",
                s_original,
                s_rev1,
                s_rev2,
                "--revisor",
                "OnlyOne",
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(2);
        await Assert.That(result.StandardOutput).IsEmpty();
        await Assert.That(result.StandardError).Contains("--revisor");
    }

    [Test]
    public async Task CLI170_WordConsolidate_MismatchedColorCount_ReturnsInvalidArguments()
    {
        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "consolidate",
                s_original,
                s_rev1,
                s_rev2,
                "--color",
                "#FF0000",
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(2);
        await Assert.That(result.StandardOutput).IsEmpty();
        await Assert.That(result.StandardError).Contains("--color");
    }

    [Test]
    public async Task CLI171_WordConsolidate_NoRevisions_ReturnsInvalidArguments()
    {
        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "consolidate",
                s_original,
                "--output",
                Path.Combine(TempDir, "out.docx"),
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(2);
        await Assert.That(result.StandardOutput).IsEmpty();
    }

    [Test]
    public async Task CLI172_WordConsolidate_DefaultOutputNaming_UsesConsolidatedSuffix()
    {
        var originalCopy = new FileInfo(Path.Combine(TempDir, "source.docx"));
        File.Copy(s_original, originalCopy.FullName);
        var expectedOutput = new FileInfo(Path.Combine(TempDir, "source-consolidated.docx"));

        var result = await CliTestRunner
            .RunManagedAsync("word", "consolidate", originalCopy.FullName, s_rev1)
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(expectedOutput.Exists).IsTrue();
        using var doc = WordprocessingDocument.Open(expectedOutput.FullName, false);
        await Validate(doc, s_expectedErrors);
    }

    [Test]
    public async Task CLI173_WordConsolidate_QuietSuppressesStdout()
    {
        var output = new FileInfo(Path.Combine(TempDir, "consolidated-quiet.docx"));

        var result = await CliTestRunner
            .RunManagedAsync("word", "consolidate", s_original, s_rev1, "--output", output.FullName, "--quiet")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardOutput).IsEmpty();
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(output.Exists).IsTrue();
    }

    [Test]
    public async Task CLI174_WordConsolidate_NonExistentRevision_ReturnsFileNotFound()
    {
        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "consolidate",
                s_original,
                "nonexistent-rev.docx",
                "--output",
                Path.Combine(TempDir, "out.docx")
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsNotEqualTo(0);
    }

    [Test]
    public async Task CLI175_WordConsolidate_OriginalFromStdin_ProducesValidOutput()
    {
        var originalBytes = await File.ReadAllBytesAsync(s_original).ConfigureAwait(false);
        var output = new FileInfo(Path.Combine(TempDir, "consolidated-stdin.docx"));

        var binaryResult = await CliTestRunner
            .RunManagedWithStdinAsync(
                originalBytes,
                "word",
                "consolidate",
                "-",
                s_rev1,
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(binaryResult.ExitCode).IsEqualTo(0);
        var stdout = Encoding.UTF8.GetString(binaryResult.StandardOutput);
        using var json = System.Text.Json.JsonDocument.Parse(stdout);
        await Assert.That(json.RootElement.GetProperty("original").GetString()).IsEqualTo("<stdin>");
        await Assert.That(output.Exists).IsTrue();
        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await Validate(doc, s_expectedErrors);
    }

    [Test]
    public async Task CLI176_WordConsolidate_OutputToStdout_ProducesValidDocx()
    {
        var binaryResult = await CliTestRunner
            .RunManagedWithStdinAsync([], "word", "consolidate", s_original, s_rev1, "--output", "-")
            .ConfigureAwait(false);

        await Assert.That(binaryResult.ExitCode).IsEqualTo(0);
        await Assert.That(binaryResult.StandardOutput.Length).IsGreaterThan(0);

        using var stream = new MemoryStream(binaryResult.StandardOutput);
        using var doc = WordprocessingDocument.Open(stream, false);
        await Validate(doc, s_expectedErrors);
    }

    [Test]
    public async Task CLI177_WordConsolidate_TwoRevisorsAndColors_ReflectedInResult()
    {
        var output = new FileInfo(Path.Combine(TempDir, "consolidated-2rv.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "consolidate",
                s_original,
                s_rev1,
                s_rev2,
                "--revisor",
                "Alice",
                "--revisor",
                "Bob",
                "--color",
                "#FF0000",
                "--color",
                "#0000FF",
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        var revisions = json.RootElement.GetProperty("revisions");
        await Assert.That(revisions.GetArrayLength()).IsEqualTo(2);
        await Assert.That(revisions[0].GetProperty("revisor").GetString()).IsEqualTo("Alice");
        await Assert.That(revisions[0].GetProperty("color").GetString()).IsEqualTo("#FF0000");
        await Assert.That(revisions[1].GetProperty("revisor").GetString()).IsEqualTo("Bob");
        await Assert.That(revisions[1].GetProperty("color").GetString()).IsEqualTo("#0000FF");
        await Assert.That(output.Exists).IsTrue();
        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await Validate(doc, s_expectedErrors);
    }

    [Test]
    public async Task CLI178_WordConsolidate_StdinForRevision_ReturnsInvalidArguments()
    {
        var result = await CliTestRunner
            .RunManagedAsync("word", "consolidate", s_original, "-", "--output", Path.Combine(TempDir, "out.docx"))
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsNotEqualTo(0);
        await Assert.That(result.StandardError).Contains("stdin");
    }
}
#pragma warning restore CA1707

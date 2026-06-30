using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Cli.Integration.Word;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
internal sealed class WordCompareTests : CliIntegrationTestBase
{
    [Test]
    public async Task CLI046_WordCompare_ValidDocuments_ReturnsResultAndWritesOutput()
    {
        var source = CliTestRunner.TestFile("WC/WC011-Before.docx");
        var revised = CliTestRunner.TestFile("WC/WC011-After.docx");
        var output = new FileInfo(Path.Combine(TempDir, "wc011-compared.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "compare",
                source.FullName,
                revised.FullName,
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("source").GetString()).IsEqualTo(source.FullName);
        await Assert.That(json.RootElement.GetProperty("revised").GetString()).IsEqualTo(revised.FullName);
        await Assert.That(json.RootElement.GetProperty("output").GetString()).IsEqualTo(output.FullName);
        await Assert.That(json.RootElement.GetProperty("outputSize").GetInt64()).IsGreaterThan(0);
        await Assert.That(json.RootElement.GetProperty("revisions").GetInt32()).IsGreaterThan(0);
        await Assert.That(output.Exists).IsTrue();

        using var compared = WordprocessingDocument.Open(output.FullName, false);
        await Validate(compared);
    }

    [Test]
    public async Task CLI047_WordCompare_ReadsSourceFromStdin()
    {
        var source = CliTestRunner.TestFile("WC/WC011-Before.docx");
        var revised = CliTestRunner.TestFile("WC/WC011-After.docx");
        var output = new FileInfo(Path.Combine(TempDir, "wc011-stdin-compared.docx"));
        var sourceBytes = await File.ReadAllBytesAsync(source.FullName).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedWithStdinAsync(
                sourceBytes,
                "word",
                "compare",
                "-",
                revised.FullName,
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        var stdout = Encoding.UTF8.GetString(result.StandardOutput);
        using var json = System.Text.Json.JsonDocument.Parse(stdout);
        await Assert.That(json.RootElement.GetProperty("source").GetString()).IsEqualTo("<stdin>");
        await Assert.That(json.RootElement.GetProperty("revised").GetString()).IsEqualTo(revised.FullName);
        await Assert.That(output.Exists).IsTrue();
    }

    [Test]
    public async Task CLI048_WordCompare_BothInputsFromStdin_ReturnsError()
    {
        var source = CliTestRunner.TestFile("WC/WC011-Before.docx");
        var sourceBytes = await File.ReadAllBytesAsync(source.FullName).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedWithStdinAsync(sourceBytes, "word", "compare", "-", "-", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(2);
        await Assert.That(result.StandardOutput).IsEmpty();
        await Assert.That(result.StandardError).Contains("Only one input can be read from stdin.");
    }

    [Test]
    public async Task CLI049_WordCompare_QuietSuppressesStdoutButPreservesExitCode()
    {
        var source = CliTestRunner.TestFile("WC/WC011-Before.docx");
        var revised = CliTestRunner.TestFile("WC/WC011-After.docx");
        var output = new FileInfo(Path.Combine(TempDir, "wc011-quiet-compared.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "compare",
                source.FullName,
                revised.FullName,
                "--output",
                output.FullName,
                "--quiet",
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardOutput).IsEmpty();
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(output.Exists).IsTrue();
    }
}
#pragma warning restore CA1707

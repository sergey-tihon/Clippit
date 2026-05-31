using System.Text.Json;

namespace Clippit.Tests.Cli.Integration.Pptx;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
internal sealed class PptxVerifyTests : CliIntegrationTestBase
{
    [Test]
    public async Task CLI021_PptxVerify_ValidDeck_ReturnsValidJson()
    {
        var input = CliTestRunner.TestFile("PB001-Input1.pptx");

        var result = await CliTestRunner
            .RunManagedAsync("pptx", "verify", input.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("input").GetString()).IsEqualTo(input.FullName);
        await Assert.That(json.RootElement.GetProperty("officeVersion").GetString()).IsEqualTo("Microsoft365");
        await Assert.That(json.RootElement.GetProperty("valid").GetBoolean()).IsTrue();
        await Assert.That(json.RootElement.TryGetProperty("slideCount", out _)).IsFalse();
        await Assert.That(json.RootElement.TryGetProperty("errors", out _)).IsFalse();
        await Assert.That(json.RootElement.TryGetProperty("errorCount", out _)).IsFalse();
        await Assert.That(json.RootElement.TryGetProperty("truncated", out _)).IsFalse();
        await Assert.That(json.RootElement.GetProperty("diagnostics").GetArrayLength()).IsEqualTo(0);
    }

    [Test]
    public async Task CLI022_PptxVerify_NonPptx_ReturnsInvalidResultOnStdout()
    {
        var directory = CliTestRunner.CreateTempDirectory("verify-invalid-package");
        var input = new FileInfo(Path.Combine(directory.FullName, "not-a-deck.pptx"));
        await File.WriteAllTextAsync(input.FullName, "not a zip package").ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync("pptx", "verify", input.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(4);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("officeVersion").GetString()).IsEqualTo("Microsoft365");
        await Assert.That(json.RootElement.GetProperty("valid").GetBoolean()).IsFalse();
        await Assert
            .That(json.RootElement.GetProperty("diagnostics")[0].GetProperty("kind").GetString())
            .IsEqualTo("package");
        await Assert.That(json.RootElement.GetProperty("diagnostics")[0].TryGetProperty("severity", out _)).IsFalse();
    }

    [Test]
    public async Task CLI023_PptxVerify_DanglingRelationship_ReturnsRelationshipDiagnostic()
    {
        var directory = CliTestRunner.CreateTempDirectory("verify-dangling-rel");
        var input = new FileInfo(Path.Combine(directory.FullName, "dangling.pptx"));
        File.Copy(CliTestRunner.TestFile("PB001-Input1.pptx").FullName, input.FullName);
        InjectDanglingOleObjRelId(input, "rId_cli_verify_dangling");

        var result = await CliTestRunner
            .RunManagedAsync("pptx", "verify", input.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(4);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("valid").GetBoolean()).IsFalse();

        var diagnostics = json.RootElement.GetProperty("diagnostics");
        var relationshipErrorFound = false;
        for (var i = 0; i < diagnostics.GetArrayLength(); i++)
        {
            var diagnostic = diagnostics[i];
            if (
                diagnostic.GetProperty("kind").GetString() == "relationship"
                && diagnostic.GetProperty("relationshipId").GetString() == "rId_cli_verify_dangling"
            )
            {
                relationshipErrorFound = true;
            }
        }
        await Assert.That(relationshipErrorFound).IsTrue();
    }

    [Test]
    public async Task CLI024_PptxVerify_ReadsFromStdin()
    {
        var input = CliTestRunner.TestFile("PB001-Input1.pptx");
        var bytes = await File.ReadAllBytesAsync(input.FullName).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedWithStdinAsync(bytes, "pptx", "verify", "-", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        var stdout = System.Text.Encoding.UTF8.GetString(result.StandardOutput);
        using var json = JsonDocument.Parse(stdout);
        await Assert.That(json.RootElement.GetProperty("input").GetString()).IsEqualTo("<stdin>");
        await Assert.That(json.RootElement.GetProperty("officeVersion").GetString()).IsEqualTo("Microsoft365");
        await Assert.That(json.RootElement.GetProperty("valid").GetBoolean()).IsTrue();
    }

    [Test]
    public async Task CLI024a_PptxVerify_OfficeVersionOverride_IsReported()
    {
        var input = CliTestRunner.TestFile("PB001-Input1.pptx");

        var result = await CliTestRunner
            .RunManagedAsync("pptx", "verify", input.FullName, "--office-version", "Office2021", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("officeVersion").GetString()).IsEqualTo("Office2021");
        await Assert.That(json.RootElement.GetProperty("valid").GetBoolean()).IsTrue();
    }

    [Test]
    public async Task CLI024b_PptxVerify_InvalidOfficeVersion_ReturnsParserError()
    {
        var input = CliTestRunner.TestFile("PB001-Input1.pptx");

        var result = await CliTestRunner
            .RunManagedAsync("pptx", "verify", input.FullName, "--office-version", "Office2099", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(1);
        await Assert.That(result.StandardOutput).Contains("Usage:");
        await Assert.That(result.StandardError).Contains("Invalid value for --office-version: 'Office2099'.");
    }

    [Test]
    public async Task CLI024c_PptxVerify_StrictOption_IsUnknownOption()
    {
        var input = CliTestRunner.TestFile("PB001-Input1.pptx");

        var result = await CliTestRunner
            .RunManagedAsync("pptx", "verify", input.FullName, "--strict", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(1);
        await Assert.That(result.StandardOutput).Contains("Usage:");
        await Assert.That(result.StandardError).Contains("Unrecognized command or argument '--strict'");
    }

    [Test]
    public async Task CLI025_PptxVerify_QuietSuppressesStdoutButPreservesExitCode()
    {
        var input = CliTestRunner.TestFile("PB001-Input1.pptx");

        var result = await CliTestRunner
            .RunManagedAsync("pptx", "verify", input.FullName, "--quiet", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardOutput).IsEmpty();
        await Assert.That(result.StandardError).IsEmpty();
    }

    [Test]
    public async Task CLI026_PptxVerify_MaxErrors_IsUnknownOption()
    {
        var input = CliTestRunner.TestFile("PB001-Input1.pptx");

        var result = await CliTestRunner
            .RunManagedAsync("pptx", "verify", input.FullName, "--max-errors", "1", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(1);
        await Assert.That(result.StandardOutput).Contains("Usage:");
        await Assert.That(result.StandardError).Contains("Unrecognized command or argument '--max-errors'");
    }
}
#pragma warning restore CA1707

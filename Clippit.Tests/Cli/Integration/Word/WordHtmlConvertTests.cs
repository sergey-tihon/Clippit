using System.Text;
using System.Text.Json;

namespace Clippit.Tests.Cli.Integration.Word;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
internal sealed class WordHtmlConvertTests : CliIntegrationTestBase
{
    private static FileInfo SimpleDocx => CliTestRunner.TestFile("Blank-wml.docx");
    private static FileInfo DocxWithImage => CliTestRunner.TestFile("HC042-Image-Png.docx");
    private static FileInfo HtmlWithStyle => CliTestRunner.TestFile("T0870.html");

    // ── word to-html ──────────────────────────────────────────────────────

    [Test]
    public async Task CLI050_ToHtml_SimpleDocument_ReturnsValidJson()
    {
        var output = new FileInfo(Path.Combine(TempDir, "simple.html"));
        var result = await CliTestRunner
            .RunManagedAsync("word", "to-html", SimpleDocx.FullName, "--output", output.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("input").GetString()).IsEqualTo(SimpleDocx.FullName);
        await Assert.That(json.RootElement.GetProperty("output").GetString()).IsEqualTo(output.FullName);
        await Assert.That(json.RootElement.GetProperty("outputSize").GetInt64()).IsGreaterThan(0);
        await Assert.That(output.Exists).IsTrue();
        await Assert.That(output.Length).IsGreaterThan(0);

        // Verify the output is valid HTML
        var html = await File.ReadAllTextAsync(output.FullName).ConfigureAwait(false);
        await Assert.That(html).StartsWith("<!DOCTYPE html >");
        await Assert.That(html).Contains("<html");
        await Assert.That(html).Contains("</html>");
    }

    [Test]
    public async Task CLI051_ToHtml_InlineImages_EmbedsBase64()
    {
        var output = new FileInfo(Path.Combine(TempDir, "with-image.html"));
        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "to-html",
                DocxWithImage.FullName,
                "--output",
                output.FullName,
                "--inline-images",
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("outputSize").GetInt64()).IsGreaterThan(0);

        var html = await File.ReadAllTextAsync(output.FullName).ConfigureAwait(false);
        // Should contain a base64 data URI for the embedded image
        await Assert.That(html).Contains("data:image/");
        await Assert.That(html).Contains("<img");
    }

    [Test]
    public async Task CLI052_ToHtml_AdditionalCss_IsInjected()
    {
        var output = new FileInfo(Path.Combine(TempDir, "with-css.html"));
        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "to-html",
                SimpleDocx.FullName,
                "--output",
                output.FullName,
                "--additional-css",
                "body { max-width: 800px; }",
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        var html = await File.ReadAllTextAsync(output.FullName).ConfigureAwait(false);
        await Assert.That(html).Contains("max-width: 800px");
    }

    [Test]
    public async Task CLI053_ToHtml_PageTitle_IsSet()
    {
        var output = new FileInfo(Path.Combine(TempDir, "with-title.html"));
        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "to-html",
                SimpleDocx.FullName,
                "--output",
                output.FullName,
                "--page-title",
                "My Custom Title",
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        var html = await File.ReadAllTextAsync(output.FullName).ConfigureAwait(false);
        await Assert.That(html).Contains("<title>My Custom Title</title>");
    }

    [Test]
    public async Task CLI054_ToHtml_NoFabricateCss_UsesInlineStyles()
    {
        var output = new FileInfo(Path.Combine(TempDir, "no-fab.html"));
        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "to-html",
                SimpleDocx.FullName,
                "--output",
                output.FullName,
                "--no-fabricate-css",
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        var html = await File.ReadAllTextAsync(output.FullName).ConfigureAwait(false);
        // Without CSS class fabrication, styles should be inline
        await Assert.That(html).Contains("style=\"");
        // Should NOT have pt-Normal or pt- prefixed classes
        await Assert.That(html).DoesNotContain("class=\"pt-");
    }

    [Test]
    public async Task CLI055_ToHtml_CustomCssPrefix_IsUsed()
    {
        var output = new FileInfo(Path.Combine(TempDir, "custom-prefix.html"));
        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "to-html",
                SimpleDocx.FullName,
                "--output",
                output.FullName,
                "--css-prefix",
                "my-",
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        var html = await File.ReadAllTextAsync(output.FullName).ConfigureAwait(false);
        await Assert.That(html).Contains("class=\"my-");
    }

    [Test]
    public async Task CLI056_ToHtml_DefaultOutputPath_DerivedFromInput()
    {
        // When no --output is given, the output should be <input>.html
        var sourceCopy = new FileInfo(Path.Combine(TempDir, "source-test.docx"));
        File.Copy(SimpleDocx.FullName, sourceCopy.FullName, overwrite: true);

        var result = await CliTestRunner
            .RunManagedAsync("word", "to-html", sourceCopy.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        var expectedHtml = new FileInfo(Path.ChangeExtension(sourceCopy.FullName, ".html"));
        // The output path in JSON should match the default
        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("output").GetString()).IsEqualTo(expectedHtml.FullName);
        await Assert.That(expectedHtml.Exists).IsTrue();
    }

    [Test]
    public async Task CLI057_ToHtml_ReadsFromStdin()
    {
        var docxBytes = await File.ReadAllBytesAsync(SimpleDocx.FullName).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedWithStdinAsync(docxBytes, "word", "to-html", "-", "--output", "-", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        // When --output - is used with stdin, no success summary is written (binary/HTML stream)
        // The output should contain the HTML content directly
        var stdout = Encoding.UTF8.GetString(result.StandardOutput);
        await Assert.That(stdout).StartsWith("<!DOCTYPE html >");
        await Assert.That(stdout).Contains("</html>");
    }

    [Test]
    public async Task CLI058_ToHtml_OutputToStdout_NoSuccessSummary()
    {
        var result = await CliTestRunner
            .RunManagedAsync("word", "to-html", SimpleDocx.FullName, "--output", "-")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        // When --output -, stdout is the HTML content, not a summary
        await Assert.That(result.StandardOutput).StartsWith("<!DOCTYPE html >");
    }

    [Test]
    public async Task CLI059_ToHtml_QuietSuppressesStdout()
    {
        var output = new FileInfo(Path.Combine(TempDir, "quiet.html"));
        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "to-html",
                SimpleDocx.FullName,
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
        // Output file should still be written
        await Assert.That(output.Exists).IsTrue();
        await Assert.That(output.Length).IsGreaterThan(0);
    }

    [Test]
    public async Task CLI059a_ToHtml_MissingInput_ReturnsParserError()
    {
        var missingFile = new FileInfo(Path.Combine(TempDir, "nonexistent.docx"));
        var result = await CliTestRunner
            .RunManagedAsync("word", "to-html", missingFile.FullName, "--format", "json")
            .ConfigureAwait(false);

        // Parser-level validation returns exit code 1
        await Assert.That(result.ExitCode).IsEqualTo(1);
        await Assert.That(result.StandardOutput).Contains("Usage:");
        await Assert.That(result.StandardError).Contains("not found");
    }

    // ── word from-html ────────────────────────────────────────────────────

    [Test]
    public async Task CLI060_FromHtml_SimpleHtml_ReturnsValidJson()
    {
        var output = new FileInfo(Path.Combine(TempDir, "simple.docx"));
        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "from-html",
                HtmlWithStyle.FullName,
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("input").GetString()).IsEqualTo(HtmlWithStyle.FullName);
        await Assert.That(json.RootElement.GetProperty("output").GetString()).IsEqualTo(output.FullName);
        await Assert.That(json.RootElement.GetProperty("outputSize").GetInt64()).IsGreaterThan(0);
        await Assert.That(output.Exists).IsTrue();
        await Assert.That(output.Length).IsGreaterThan(0);

        // Verify the output is a valid DOCX (starts with PK zip signature)
        var docxBytes = await File.ReadAllBytesAsync(output.FullName).ConfigureAwait(false);
        await Assert.That(docxBytes[0]).IsEqualTo((byte)'P');
        await Assert.That(docxBytes[1]).IsEqualTo((byte)'K');
    }

    [Test]
    public async Task CLI061_FromHtml_WithExternalCss_UsesFile()
    {
        var output = new FileInfo(Path.Combine(TempDir, "external-css.docx"));
        var cssFile = new FileInfo(Path.Combine(TempDir, "custom.css"));
        await File.WriteAllTextAsync(cssFile.FullName, "p { color: red; }").ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "from-html",
                HtmlWithStyle.FullName,
                "--output",
                output.FullName,
                "--css",
                cssFile.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(output.Exists).IsTrue();
    }

    [Test]
    public async Task CLI062_FromHtml_WithCustomFonts_UsesFontSettings()
    {
        var output = new FileInfo(Path.Combine(TempDir, "font-test.docx"));
        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "from-html",
                HtmlWithStyle.FullName,
                "--output",
                output.FullName,
                "--major-font",
                "Arial",
                "--minor-font",
                "Georgia",
                "--font-size",
                "11",
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(output.Exists).IsTrue();
    }

    [Test]
    public async Task CLI063_FromHtml_WithUserCss_AppliesOverrides()
    {
        var output = new FileInfo(Path.Combine(TempDir, "user-css.docx"));
        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "from-html",
                HtmlWithStyle.FullName,
                "--output",
                output.FullName,
                "--user-css",
                "p { margin: 0; }",
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(output.Exists).IsTrue();
    }

    [Test]
    public async Task CLI064_FromHtml_WithBaseUri_ResolvesImages()
    {
        var output = new FileInfo(Path.Combine(TempDir, "base-uri.docx"));
        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "from-html",
                HtmlWithStyle.FullName,
                "--output",
                output.FullName,
                "--base-uri",
                "https://example.com/images/",
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(output.Exists).IsTrue();
    }

    [Test]
    public async Task CLI065_FromHtml_DefaultOutputPath_DerivedFromInput()
    {
        var sourceCopy = new FileInfo(Path.Combine(TempDir, "source-test.html"));
        File.Copy(HtmlWithStyle.FullName, sourceCopy.FullName, overwrite: true);

        var result = await CliTestRunner
            .RunManagedAsync("word", "from-html", sourceCopy.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        var expectedDocx = new FileInfo(Path.ChangeExtension(sourceCopy.FullName, ".docx"));
        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("output").GetString()).IsEqualTo(expectedDocx.FullName);
        await Assert.That(expectedDocx.Exists).IsTrue();
    }

    [Test]
    public async Task CLI066_FromHtml_ReadsFromStdin()
    {
        var htmlBytes = await File.ReadAllBytesAsync(HtmlWithStyle.FullName).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedWithStdinAsync(htmlBytes, "word", "from-html", "-", "--output", "-")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        // When --output -, stdout is the binary DOCX content
        await Assert.That(result.StandardOutput.Length).IsGreaterThan(0);
        // Verify it looks like a ZIP/DOCX
        await Assert.That(result.StandardOutput[0]).IsEqualTo((byte)'P');
        await Assert.That(result.StandardOutput[1]).IsEqualTo((byte)'K');
    }

    [Test]
    public async Task CLI067_FromHtml_QuietSuppressesStdout()
    {
        var output = new FileInfo(Path.Combine(TempDir, "quiet.docx"));
        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "from-html",
                HtmlWithStyle.FullName,
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
        await Assert.That(output.Length).IsGreaterThan(0);
    }

    [Test]
    public async Task CLI068_FromHtml_MissingInput_ReturnsParserError()
    {
        var missingFile = new FileInfo(Path.Combine(TempDir, "nonexistent.html"));
        var result = await CliTestRunner
            .RunManagedAsync("word", "from-html", missingFile.FullName, "--format", "json")
            .ConfigureAwait(false);

        // Parser-level validation returns exit code 1
        await Assert.That(result.ExitCode).IsEqualTo(1);
        await Assert.That(result.StandardOutput).Contains("Usage:");
        await Assert.That(result.StandardError).Contains("not found");
    }

    [Test]
    public async Task CLI069_FromHtml_InvalidNonXmlHtml_ReturnsError()
    {
        var badHtml = new FileInfo(Path.Combine(TempDir, "bad.html"));
        await File.WriteAllTextAsync(badHtml.FullName, "<html><body><p unclosed>Hello</html>").ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync("word", "from-html", badHtml.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(4);
        await Assert.That(result.StandardError).Contains("not well-formed XML");
    }

    [Test]
    public async Task CLI069a_FromHtml_MissingCssFile_ReturnsError()
    {
        var missingCss = new FileInfo(Path.Combine(TempDir, "missing.css"));
        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "from-html",
                HtmlWithStyle.FullName,
                "--css",
                missingCss.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(3);
        await Assert.That(result.StandardError).Contains("not found");
    }

    // ── word to-html → from-html roundtrip ────────────────────────────────

    [Test]
    public async Task CLI070_Roundtrip_ToHtmlThenFromHtml_ProducesValidDocx()
    {
        var htmlOutput = new FileInfo(Path.Combine(TempDir, "roundtrip.html"));
        var docxOutput = new FileInfo(Path.Combine(TempDir, "roundtrip.docx"));

        // Step 1: docx → html
        var step1 = await CliTestRunner
            .RunManagedAsync(
                "word",
                "to-html",
                SimpleDocx.FullName,
                "--output",
                htmlOutput.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(step1.ExitCode).IsEqualTo(0);

        // Step 2: html → docx
        var step2 = await CliTestRunner
            .RunManagedAsync(
                "word",
                "from-html",
                htmlOutput.FullName,
                "--output",
                docxOutput.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(step2.ExitCode).IsEqualTo(0);
        await Assert.That(docxOutput.Exists).IsTrue();
        await Assert.That(docxOutput.Length).IsGreaterThan(0);

        // Verify it's a valid ZIP (DOCX)
        var bytes = await File.ReadAllBytesAsync(docxOutput.FullName).ConfigureAwait(false);
        await Assert.That(bytes[0]).IsEqualTo((byte)'P');
        await Assert.That(bytes[1]).IsEqualTo((byte)'K');
    }

    // ── word to-html text output ──────────────────────────────────────────

    [Test]
    public async Task CLI071_ToHtml_TextOutput_HumanReadable()
    {
        var output = new FileInfo(Path.Combine(TempDir, "text-out.html"));
        var result = await CliTestRunner
            .RunManagedAsync("word", "to-html", SimpleDocx.FullName, "--output", output.FullName, "--format", "text")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        // Text output should contain the input → output arrow
        await Assert.That(result.StandardOutput).Contains(SimpleDocx.Name);
        await Assert.That(result.StandardOutput).Contains("Output size:");
    }

    [Test]
    public async Task CLI072_FromHtml_TextOutput_HumanReadable()
    {
        var output = new FileInfo(Path.Combine(TempDir, "text-out.docx"));
        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "from-html",
                HtmlWithStyle.FullName,
                "--output",
                output.FullName,
                "--format",
                "text"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(result.StandardOutput).Contains(HtmlWithStyle.Name);
        await Assert.That(result.StandardOutput).Contains("Output size:");
    }

    // ── word to-html with external image files ────────────────────────────

    [Test]
    public async Task CLI073_ToHtml_ExternalImages_CreatesFilesDirectory()
    {
        var output = new FileInfo(Path.Combine(TempDir, "extimg", "output.html"));
        var result = await CliTestRunner
            .RunManagedAsync("word", "to-html", DocxWithImage.FullName, "--output", output.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        // Without --inline-images, images should be written to a _files directory
        var filesDir = new DirectoryInfo(
            Path.Combine(output.DirectoryName!, Path.GetFileNameWithoutExtension(output.Name) + "_files")
        );
        await Assert.That(filesDir.Exists).IsTrue();
        await Assert.That(filesDir.GetFiles().Length).IsGreaterThan(0);

        // The HTML should reference external images (not data URIs)
        var html = await File.ReadAllTextAsync(output.FullName).ConfigureAwait(false);
        await Assert.That(html).DoesNotContain("data:image/");
        await Assert.That(html).Contains("<img");
    }

    [Test]
    public async Task CLI074_ToHtml_NestedOutputDirectory_CreatedForDocumentWithoutImages()
    {
        var output = new FileInfo(Path.Combine(TempDir, "nested", "no-image", "output.html"));
        var result = await CliTestRunner
            .RunManagedAsync("word", "to-html", SimpleDocx.FullName, "--output", output.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(output.Exists).IsTrue();
    }

    [Test]
    public async Task CLI075_ToHtml_StdoutOutput_SuppressesSuccessSummary()
    {
        var result = await CliTestRunner
            .RunManagedAsync("word", "to-html", SimpleDocx.FullName, "--output", "-", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        var stdout = result.StandardOutput;
        await Assert.That(stdout).StartsWith("<!DOCTYPE html >");
        // Success summary must not be mixed into the binary/HTML stream.
        await Assert.That(stdout).DoesNotContain("\"input\"");
        await Assert.That(stdout).DoesNotContain("Output size:");
    }

    [Test]
    public async Task CLI076_FromHtml_NestedOutputDirectory_Created()
    {
        var output = new FileInfo(Path.Combine(TempDir, "nested", "from-html", "output.docx"));
        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "from-html",
                HtmlWithStyle.FullName,
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(output.Exists).IsTrue();
    }
}
#pragma warning restore CA1707

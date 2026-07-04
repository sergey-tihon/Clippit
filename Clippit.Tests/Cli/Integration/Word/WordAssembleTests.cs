using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Cli.Integration.Word;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
internal sealed class WordAssembleTests : CliIntegrationTestBase
{
    private static FileInfo Template => CliTestRunner.TestFile("DA/DA001-TemplateDocument.docx");
    private static FileInfo Data => CliTestRunner.TestFile("DA/DA-Data.xml");
    private static FileInfo TemplateWithErrors => CliTestRunner.TestFile("DA/DA003-Select-XPathFindsNoData.docx");

    [Test]
    public async Task CLI110_WordAssemble_ValidTemplateAndData_ReturnsResultAndWritesOutput()
    {
        var output = new FileInfo(Path.Combine(TempDir, "da001-assembled.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "assemble",
                Template.FullName,
                Data.FullName,
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("template").GetString()).IsEqualTo(Template.FullName);
        await Assert.That(json.RootElement.GetProperty("data").GetString()).IsEqualTo(Data.FullName);
        await Assert.That(json.RootElement.GetProperty("output").GetString()).IsEqualTo(output.FullName);
        await Assert.That(json.RootElement.GetProperty("outputSize").GetInt64()).IsGreaterThan(0);
        await Assert.That(json.RootElement.GetProperty("templateError").GetBoolean()).IsEqualTo(false);
        await Assert.That(output.Exists).IsTrue();

        // Document-level validation is covered by DocumentAssemblerTests; here we just check the file is valid docx.
        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await Assert.That(doc.MainDocumentPart).IsNotNull();
    }

    [Test]
    public async Task CLI111_WordAssemble_DefaultOutputPath_WritesAdjacentFile()
    {
        var tempDir = CliTestRunner.CreateTempDirectory("da001-default-output");
        var templateCopy = new FileInfo(Path.Combine(tempDir.FullName, "DA001-TemplateDocument.docx"));
        File.Copy(Template.FullName, templateCopy.FullName, overwrite: true);
        var expectedOutput = new FileInfo(Path.Combine(tempDir.FullName, "DA001-TemplateDocument-assembled.docx"));

        var result = await CliTestRunner
            .RunManagedAsync("word", "assemble", templateCopy.FullName, Data.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(expectedOutput.Exists).IsTrue();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("output").GetString()).IsEqualTo(expectedOutput.FullName);
    }

    [Test]
    public async Task CLI112_WordAssemble_ReadsDataFromStdin()
    {
        var output = new FileInfo(Path.Combine(TempDir, "da001-stdin-data.docx"));
        var dataBytes = await File.ReadAllBytesAsync(Data.FullName).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedWithStdinAsync(
                dataBytes,
                "word",
                "assemble",
                Template.FullName,
                "-",
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        var stdout = System.Text.Encoding.UTF8.GetString(result.StandardOutput);
        using var json = System.Text.Json.JsonDocument.Parse(stdout);
        await Assert.That(json.RootElement.GetProperty("data").GetString()).IsEqualTo("<stdin>");
        await Assert.That(output.Exists).IsTrue();
    }

    [Test]
    public async Task CLI113_WordAssemble_BothInputsFromStdin_ReturnsError()
    {
        var templateBytes = await File.ReadAllBytesAsync(Template.FullName).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedWithStdinAsync(templateBytes, "word", "assemble", "-", "-", "--output", "out.docx")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsNotEqualTo(0);
        await Assert.That(result.StandardError).Contains("stdin");
    }

    [Test]
    public async Task CLI114_WordAssemble_MalformedXmlData_ReturnsInvalidFormat()
    {
        var output = new FileInfo(Path.Combine(TempDir, "da001-bad-xml.docx"));
        var badXmlBytes = System.Text.Encoding.UTF8.GetBytes("<NotXml>missing close");

        var result = await CliTestRunner
            .RunManagedWithStdinAsync(
                badXmlBytes,
                "word",
                "assemble",
                Template.FullName,
                "-",
                "--output",
                output.FullName
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(4); // INVALID_FORMAT
        await Assert.That(result.StandardError).Contains("XML");
    }

    [Test]
    public async Task CLI115_WordAssemble_TemplateError_ReportsInResult()
    {
        var output = new FileInfo(Path.Combine(TempDir, "da003-assembled.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "assemble",
                TemplateWithErrors.FullName,
                Data.FullName,
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(output.Exists).IsTrue();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("templateError").GetBoolean()).IsEqualTo(true);
    }

    [Test]
    public async Task CLI116_WordAssemble_QuietMode_SuppressesOutput()
    {
        var output = new FileInfo(Path.Combine(TempDir, "da001-quiet.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "assemble",
                Template.FullName,
                Data.FullName,
                "--output",
                output.FullName,
                "--quiet"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardOutput).IsEmpty();
        await Assert.That(output.Exists).IsTrue();
    }

    [Test]
    public async Task CLI117_WordAssemble_OutputToStdout_WritesDocxBytes()
    {
        var result = await CliTestRunner
            .RunManagedWithStdinAsync(
                Array.Empty<byte>(),
                "word",
                "assemble",
                Template.FullName,
                Data.FullName,
                "--output",
                "-"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardOutput.Length).IsGreaterThan(0);
        // ZIP magic bytes: PK (0x50 0x4B)
        await Assert.That(result.StandardOutput[0]).IsEqualTo((byte)0x50);
        await Assert.That(result.StandardOutput[1]).IsEqualTo((byte)0x4B);
    }
}

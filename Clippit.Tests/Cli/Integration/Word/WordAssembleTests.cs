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
        await Assert.That(json.RootElement.GetProperty("templateError").GetBoolean()).IsFalse();
        await Assert.That(output.Exists).IsTrue();

        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await Validate(doc, s_expectedErrors);
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
        await Assert.That(json.RootElement.GetProperty("templateError").GetBoolean()).IsTrue();
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

    [Test]
    [Arguments("template")]
    [Arguments("data")]
    public async Task CLI118_WordAssemble_OutputPathMatchesInput_ReturnsOutputError(string conflictingInput)
    {
        var tempDir = CliTestRunner.CreateTempDirectory("da001-overwrite-guard");
        var templateCopy = new FileInfo(Path.Combine(tempDir.FullName, "template.docx"));
        var dataCopy = new FileInfo(Path.Combine(tempDir.FullName, "data.xml"));
        File.Copy(Template.FullName, templateCopy.FullName, overwrite: true);
        File.Copy(Data.FullName, dataCopy.FullName, overwrite: true);
        var outputPath = conflictingInput == "template" ? templateCopy.FullName : dataCopy.FullName;

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "assemble",
                templateCopy.FullName,
                dataCopy.FullName,
                "--output",
                outputPath,
                "--force",
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(5);
        await Assert.That(result.StandardOutput).IsEmpty();
        await Assert.That(result.StandardError).Contains("must not overwrite");
    }

    [Test]
    public async Task CLI119_WordAssemble_DtdInXmlData_ReturnsInvalidFormat()
    {
        var output = new FileInfo(Path.Combine(TempDir, "da001-dtd.docx"));
        var dtdXmlBytes = System.Text.Encoding.UTF8.GetBytes(
            """
            <?xml version="1.0" encoding="utf-8"?>
            <!DOCTYPE Customer [<!ENTITY injected "blocked">]>
            <Customer>
                <Name>&injected;</Name>
            </Customer>
            """
        );

        var result = await CliTestRunner
            .RunManagedWithStdinAsync(
                dtdXmlBytes,
                "word",
                "assemble",
                Template.FullName,
                "-",
                "--output",
                output.FullName
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(4);
        await Assert.That(result.StandardOutput).IsEmpty();
        await Assert.That(result.StandardError).Contains("DTD");
    }

    private static readonly List<string> s_expectedErrors =
    [
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:evenHBand' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:evenVBand' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstColumn' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRow' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRowFirstColumn' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRowLastColumn' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastColumn' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRow' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRowFirstColumn' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRowLastColumn' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noHBand' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noVBand' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:oddHBand' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:oddVBand' attribute is not declared.",
        "The 'http://schemas.microsoft.com/office/word/2012/wordml:restartNumberingAfterBreak' attribute is not declared.",
        "The 'http://schemas.microsoft.com/office/word/2016/wordml/cid:durableId' attribute is not declared.",
        "Attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:val' should have unique value. Its current value",
        "The element has unexpected child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:bCs'.",
        "The element has unexpected child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:kern'.",
        "The element has unexpected child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:rFonts'.",
    ];
}

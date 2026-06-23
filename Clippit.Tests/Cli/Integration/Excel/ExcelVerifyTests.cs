using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;

namespace Clippit.Tests.Cli.Integration.Excel;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
internal sealed class ExcelVerifyTests : CliIntegrationTestBase
{
    [Test]
    public async Task CLI078_ExcelVerify_ValidSpreadsheet_ReturnsValidJson()
    {
        var input = CliTestRunner.TestFile("SH001-Table.xlsx");

        var result = await CliTestRunner
            .RunManagedAsync("excel", "verify", input.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("input").GetString()).IsEqualTo(input.FullName);
        await Assert.That(json.RootElement.GetProperty("officeVersion").GetString()).IsEqualTo("Microsoft365");
        await Assert.That(json.RootElement.GetProperty("valid").GetBoolean()).IsTrue();
        await Assert.That(json.RootElement.GetProperty("diagnostics").GetArrayLength()).IsEqualTo(0);
    }

    [Test]
    public async Task CLI079_ExcelVerify_NonXlsx_ReturnsInvalidResultOnStdout()
    {
        var directory = CliTestRunner.CreateTempDirectory("verify-invalid-package");
        var input = new FileInfo(Path.Combine(directory.FullName, "not-a-spreadsheet.xlsx"));
        await File.WriteAllTextAsync(input.FullName, "not a zip package").ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync("excel", "verify", input.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(4);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("officeVersion").GetString()).IsEqualTo("Microsoft365");
        await Assert.That(json.RootElement.GetProperty("valid").GetBoolean()).IsFalse();
        await Assert
            .That(json.RootElement.GetProperty("diagnostics")[0].GetProperty("kind").GetString())
            .IsEqualTo("package");
    }

    [Test]
    public async Task CLI080_ExcelVerify_DanglingRelationship_ReturnsRelationshipDiagnostic()
    {
        var directory = CliTestRunner.CreateTempDirectory("verify-dangling-rel");
        var input = new FileInfo(Path.Combine(directory.FullName, "dangling.xlsx"));
        File.Copy(CliTestRunner.TestFile("SH001-Table.xlsx").FullName, input.FullName);
        InjectDanglingRelationship(input, "rId_cli_verify_dangling");

        var result = await CliTestRunner
            .RunManagedAsync("excel", "verify", input.FullName, "--format", "json")
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
    public async Task CLI081_ExcelVerify_ReadsFromStdin()
    {
        var input = CliTestRunner.TestFile("SH001-Table.xlsx");
        var bytes = await File.ReadAllBytesAsync(input.FullName).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedWithStdinAsync(bytes, "excel", "verify", "-", "--format", "json")
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
    public async Task CLI081a_ExcelVerify_OfficeVersionOverride_IsReported()
    {
        var input = CliTestRunner.TestFile("SH001-Table.xlsx");

        var result = await CliTestRunner
            .RunManagedAsync("excel", "verify", input.FullName, "--office-version", "Office2021", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("officeVersion").GetString()).IsEqualTo("Office2021");
        await Assert.That(json.RootElement.GetProperty("valid").GetBoolean()).IsTrue();
    }

    [Test]
    public async Task CLI081b_ExcelVerify_InvalidOfficeVersion_ReturnsParserError()
    {
        var input = CliTestRunner.TestFile("SH001-Table.xlsx");

        var result = await CliTestRunner
            .RunManagedAsync("excel", "verify", input.FullName, "--office-version", "Office2099", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(1);
        await Assert.That(result.StandardOutput).Contains("Usage:");
        await Assert.That(result.StandardError).Contains("Invalid value for --office-version: 'Office2099'.");
    }

    [Test]
    public async Task CLI081c_ExcelVerify_StrictOption_IsUnknownOption()
    {
        var input = CliTestRunner.TestFile("SH001-Table.xlsx");

        var result = await CliTestRunner
            .RunManagedAsync("excel", "verify", input.FullName, "--strict", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(1);
        await Assert.That(result.StandardOutput).Contains("Usage:");
        await Assert.That(result.StandardError).Contains("Unrecognized command or argument '--strict'");
    }

    [Test]
    public async Task CLI082_ExcelVerify_QuietSuppressesStdoutButPreservesExitCode()
    {
        var input = CliTestRunner.TestFile("SH001-Table.xlsx");

        var result = await CliTestRunner
            .RunManagedAsync("excel", "verify", input.FullName, "--quiet", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardOutput).IsEmpty();
        await Assert.That(result.StandardError).IsEmpty();
    }

    private static void InjectDanglingRelationship(FileInfo xlsx, params string[] danglingIds)
    {
        XNamespace sNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        using var zip = ZipFile.Open(xlsx.FullName, ZipArchiveMode.Update);
        var sheetEntry = zip.GetEntry("xl/worksheets/sheet1.xml")!;

        XDocument xDoc;
        using (var stream = sheetEntry.Open())
            xDoc = XDocument.Load(stream);

        var target = xDoc.Descendants(sNs + "sheetData").FirstOrDefault() ?? xDoc.Root;
        foreach (var danglingId in danglingIds)
            target?.Add(new XElement(sNs + "row", new XAttribute(rNs + "id", danglingId)));

        var fullName = sheetEntry.FullName;
        sheetEntry.Delete();
        var newEntry = zip.CreateEntry(fullName);
        using var writer = new StreamWriter(newEntry.Open());
        using var xmlWriter = System.Xml.XmlWriter.Create(writer);
        xDoc.WriteTo(xmlWriter);
    }
}
#pragma warning restore CA1707

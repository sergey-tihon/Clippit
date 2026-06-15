using System.Text.Json;
using Clippit.Tests.Cli.Integration;

namespace Clippit.Tests.Cli.Integration.Excel;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
internal sealed class ExcelToHtmlTests : CliIntegrationTestBase
{
    [Test]
    public async Task CLI061_ExcelToHtml_ConvertTable_ProducesValidHtmlAndJsonResult()
    {
        var input = CliTestRunner.TestFile("SH001-Table.xlsx");
        var tempDir = CliTestRunner.CreateTempDirectory("excel-to-html-table");
        var output = Path.Combine(tempDir.FullName, "output.html");

        var result = await CliTestRunner
            .RunManagedAsync(
                "excel",
                "to-html",
                input.FullName,
                "--output",
                output,
                "--table",
                "MyTable",
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("input").GetString()).IsEqualTo(input.FullName);
        await Assert.That(json.RootElement.GetProperty("output").GetString()).IsEqualTo(output);
        await Assert.That(json.RootElement.GetProperty("outputSize").GetInt64()).IsGreaterThan(0);

        await Assert.That(File.Exists(output)).IsTrue();
        var html = await File.ReadAllTextAsync(output).ConfigureAwait(false);
        await Assert.That(html).Contains("<table");
        await Assert.That(html).Contains("MyTable");
    }

    [Test]
    public async Task CLI062_ExcelToHtml_ConvertWholeSheet_ProducesValidHtml()
    {
        var input = CliTestRunner.TestFile("SH101-SimpleFormats.xlsx");
        var tempDir = CliTestRunner.CreateTempDirectory("excel-to-html-sheet");
        var output = Path.Combine(tempDir.FullName, "output.html");

        var result = await CliTestRunner
            .RunManagedAsync(
                "excel",
                "to-html",
                input.FullName,
                "--output",
                output,
                "--sheet",
                "Sheet1",
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        await Assert.That(File.Exists(output)).IsTrue();
        var html = await File.ReadAllTextAsync(output).ConfigureAwait(false);
        await Assert.That(html).Contains("<table");
    }

    [Test]
    public async Task CLI063_ExcelToHtml_ConvertRange_ProducesValidHtml()
    {
        var input = CliTestRunner.TestFile("SH101-SimpleFormats.xlsx");
        var tempDir = CliTestRunner.CreateTempDirectory("excel-to-html-range");
        var output = Path.Combine(tempDir.FullName, "output.html");

        var result = await CliTestRunner
            .RunManagedAsync(
                "excel",
                "to-html",
                input.FullName,
                "--output",
                output,
                "--sheet",
                "Sheet1",
                "--range",
                "A1:B10",
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        await Assert.That(File.Exists(output)).IsTrue();
        var html = await File.ReadAllTextAsync(output).ConfigureAwait(false);
        await Assert.That(html).Contains("<table");
    }

    [Test]
    public async Task CLI064_ExcelToHtml_InvalidTableCombination_ReturnsError()
    {
        var input = CliTestRunner.TestFile("SH001-Table.xlsx");

        var result = await CliTestRunner
            .RunManagedAsync(
                "excel",
                "to-html",
                input.FullName,
                "--table",
                "MyTable",
                "--sheet",
                "Sheet1",
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(2);
        await Assert.That(result.StandardOutput).IsEmpty();

        using var json = result.ReadStderrJson();
        await Assert.That(json.RootElement.GetProperty("code").GetString()).IsEqualTo("INVALID_ARGUMENTS");
        await Assert.That(json.RootElement.GetProperty("error").GetString()).Contains("cannot be combined with");
    }

    [Test]
    public async Task CLI065_ExcelToHtml_RangeWithoutSheet_ReturnsError()
    {
        var input = CliTestRunner.TestFile("SH101-SimpleFormats.xlsx");

        var result = await CliTestRunner
            .RunManagedAsync("excel", "to-html", input.FullName, "--range", "A1:B10", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(2);
        await Assert.That(result.StandardOutput).IsEmpty();

        using var json = result.ReadStderrJson();
        await Assert.That(json.RootElement.GetProperty("code").GetString()).IsEqualTo("INVALID_ARGUMENTS");
        await Assert.That(json.RootElement.GetProperty("error").GetString()).Contains("requires a --sheet");
    }

    [Test]
    public async Task CLI066_ExcelToHtml_TableNotFound_ReturnsError()
    {
        var input = CliTestRunner.TestFile("SH001-Table.xlsx");

        var result = await CliTestRunner
            .RunManagedAsync("excel", "to-html", input.FullName, "--table", "NonExistentTable", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(2);
        await Assert.That(result.StandardOutput).IsEmpty();

        using var json = result.ReadStderrJson();
        await Assert.That(json.RootElement.GetProperty("code").GetString()).IsEqualTo("INVALID_ARGUMENTS");
        await Assert.That(json.RootElement.GetProperty("error").GetString()).Contains("was not found");
    }
}
#pragma warning restore CA1707

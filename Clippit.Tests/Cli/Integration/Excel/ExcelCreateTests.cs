using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Clippit.Tests.Cli.Integration.Excel;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
internal sealed class ExcelCreateTests : CliIntegrationTestBase
{
    private static string SimpleWorkbook(string sheetName = "Sheet1") =>
        $$"""
            {
              "worksheets": [
                {
                  "name": "{{sheetName}}",
                  "rows": [
                    { "cells": [{ "value": "Hello" }, { "value": 42, "cellDataType": "Number" }] },
                    { "cells": [{ "value": "World" }, { "value": 3.14, "cellDataType": "Number" }] }
                  ]
                }
              ]
            }
            """;

    private static string WorkbookWithTable() =>
        """
            {
              "worksheets": [
                {
                  "name": "Sales",
                  "tableName": "SalesTable",
                  "columnHeadings": [
                    { "value": "Product", "bold": true },
                    { "value": "Revenue", "bold": true }
                  ],
                  "rows": [
                    { "cells": [{ "value": "Widget" }, { "value": 1234.5, "cellDataType": "Number", "formatCode": "#,##0.00", "horizontalCellAlignment": "Right" }] },
                    { "cells": [{ "value": "Gadget" }, { "value": 567.8, "cellDataType": "Number", "formatCode": "#,##0.00", "horizontalCellAlignment": "Right" }] }
                  ]
                }
              ]
            }
            """;

    private static async Task<FileInfo> WriteWorkbookJsonAsync(
        DirectoryInfo dir,
        string content,
        string name = "workbook.json"
    )
    {
        var file = new FileInfo(Path.Combine(dir.FullName, name));
        await File.WriteAllTextAsync(file.FullName, content).ConfigureAwait(false);
        return file;
    }

    private async Task ValidateXlsxAsync(FileInfo xlsx)
    {
        using var doc = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(xlsx.FullName, false);
        await Validate(doc, []);
    }

    [Test]
    public async Task CLI179_ExcelCreate_ValidDefinition_ProducesValidXlsx()
    {
        var dir = CliTestRunner.CreateTempDirectory("excel-create-basic");
        var input = await WriteWorkbookJsonAsync(dir, SimpleWorkbook()).ConfigureAwait(false);
        var output = new FileInfo(Path.Combine(dir.FullName, "output.xlsx"));

        var result = await CliTestRunner
            .RunManagedAsync("excel", "create", input.FullName, "--output", output.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(output.Exists).IsTrue();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("input").GetString()).IsEqualTo(input.FullName);
        await Assert.That(json.RootElement.GetProperty("output").GetString()).IsEqualTo(output.FullName);
        await Assert.That(json.RootElement.GetProperty("outputSize").GetInt64()).IsGreaterThan(0);
        await Assert.That(json.RootElement.GetProperty("worksheetCount").GetInt32()).IsEqualTo(1);
        await ValidateXlsxAsync(output).ConfigureAwait(false);
    }

    [Test]
    public async Task CLI180_ExcelCreate_MultipleWorksheets_AllWritten()
    {
        var dir = CliTestRunner.CreateTempDirectory("excel-create-multi");
        var content = """
            {
              "worksheets": [
                {
                  "name": "Alpha",
                  "rows": [{ "cells": [{ "value": "A1" }] }]
                },
                {
                  "name": "Beta",
                  "rows": [{ "cells": [{ "value": "B1" }] }]
                },
                {
                  "name": "Gamma",
                  "rows": [{ "cells": [{ "value": "C1" }] }]
                }
              ]
            }
            """;
        var input = await WriteWorkbookJsonAsync(dir, content).ConfigureAwait(false);
        var output = new FileInfo(Path.Combine(dir.FullName, "multi.xlsx"));

        var result = await CliTestRunner
            .RunManagedAsync("excel", "create", input.FullName, "--output", output.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("worksheetCount").GetInt32()).IsEqualTo(3);

        using var doc = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(output.FullName, false);
        var sheets = doc.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().Select(s => s.Name?.Value).ToList();
        await Assert.That(sheets).Contains("Alpha");
        await Assert.That(sheets).Contains("Beta");
        await Assert.That(sheets).Contains("Gamma");
    }

    [Test]
    public async Task CLI181_ExcelCreate_WithTableAndColumnHeadings_ProducesValidXlsx()
    {
        var dir = CliTestRunner.CreateTempDirectory("excel-create-table");
        var input = await WriteWorkbookJsonAsync(dir, WorkbookWithTable()).ConfigureAwait(false);
        var output = new FileInfo(Path.Combine(dir.FullName, "table.xlsx"));

        var result = await CliTestRunner
            .RunManagedAsync("excel", "create", input.FullName, "--output", output.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await ValidateXlsxAsync(output).ConfigureAwait(false);
    }

    [Test]
    public async Task CLI182_ExcelCreate_DateCellType_ParsesIso8601()
    {
        var dir = CliTestRunner.CreateTempDirectory("excel-create-date");
        var content = """
            {
              "worksheets": [
                {
                  "name": "Dates",
                  "rows": [
                    { "cells": [
                      { "value": "2026-07-09", "cellDataType": "Date", "formatCode": "yyyy-mm-dd" }
                    ]}
                  ]
                }
              ]
            }
            """;
        var input = await WriteWorkbookJsonAsync(dir, content).ConfigureAwait(false);
        var output = new FileInfo(Path.Combine(dir.FullName, "dates.xlsx"));

        var result = await CliTestRunner
            .RunManagedAsync("excel", "create", input.FullName, "--output", output.FullName)
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await ValidateXlsxAsync(output).ConfigureAwait(false);
    }

    [Test]
    public async Task CLI183_ExcelCreate_BooleanCell_Accepted()
    {
        var dir = CliTestRunner.CreateTempDirectory("excel-create-bool");
        var content = """
            {
              "worksheets": [
                {
                  "name": "Flags",
                  "rows": [
                    { "cells": [
                      { "value": true, "cellDataType": "Boolean" },
                      { "value": false, "cellDataType": "Boolean" }
                    ]}
                  ]
                }
              ]
            }
            """;
        var input = await WriteWorkbookJsonAsync(dir, content).ConfigureAwait(false);
        var output = new FileInfo(Path.Combine(dir.FullName, "bools.xlsx"));

        var result = await CliTestRunner
            .RunManagedAsync("excel", "create", input.FullName, "--output", output.FullName)
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await ValidateXlsxAsync(output).ConfigureAwait(false);
    }

    [Test]
    public async Task CLI184_ExcelCreate_NullCellValue_TreatedAsEmpty()
    {
        var dir = CliTestRunner.CreateTempDirectory("excel-create-null");
        var content = """
            {
              "worksheets": [
                {
                  "name": "NullsSheet",
                  "rows": [
                    { "cells": [{ "value": null }, { "value": "present" }] }
                  ]
                }
              ]
            }
            """;
        var input = await WriteWorkbookJsonAsync(dir, content).ConfigureAwait(false);
        var output = new FileInfo(Path.Combine(dir.FullName, "nulls.xlsx"));

        var result = await CliTestRunner
            .RunManagedAsync("excel", "create", input.FullName, "--output", output.FullName)
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await ValidateXlsxAsync(output).ConfigureAwait(false);
    }

    [Test]
    public async Task CLI185_ExcelCreate_DefaultOutputPath_AdjacentToInput()
    {
        var dir = CliTestRunner.CreateTempDirectory("excel-create-default-out");
        var input = await WriteWorkbookJsonAsync(dir, SimpleWorkbook(), "report.json").ConfigureAwait(false);
        var expectedOutput = new FileInfo(Path.Combine(dir.FullName, "report.xlsx"));

        var result = await CliTestRunner
            .RunManagedAsync("excel", "create", input.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(expectedOutput.Exists).IsTrue();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("output").GetString()).IsEqualTo(expectedOutput.FullName);
    }

    [Test]
    public async Task CLI186_ExcelCreate_StdinInput_DefaultsToWorkbookXlsx()
    {
        var dir = CliTestRunner.CreateTempDirectory("excel-create-stdin");
        var outputFile = new FileInfo(Path.Combine(dir.FullName, "from-stdin.xlsx"));
        var jsonBytes = Encoding.UTF8.GetBytes(SimpleWorkbook());

        var result = await CliTestRunner
            .RunManagedWithStdinAsync(
                jsonBytes,
                "excel",
                "create",
                "-",
                "--output",
                outputFile.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(outputFile.Exists).IsTrue();

        using var json = JsonDocument.Parse(result.StandardOutput);
        await Assert.That(json.RootElement.GetProperty("input").GetString()).IsEqualTo("<stdin>");
    }

    [Test]
    public async Task CLI187_ExcelCreate_StdoutOutput_WritesBinaryToStdout()
    {
        var dir = CliTestRunner.CreateTempDirectory("excel-create-stdout");
        var input = await WriteWorkbookJsonAsync(dir, SimpleWorkbook()).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedWithStdinAsync([], "excel", "create", input.FullName, "--output", "-")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        // xlsx files start with PK magic bytes (ZIP)
        await Assert.That(result.StandardOutput.Length).IsGreaterThan(0);
        await Assert.That(result.StandardOutput[0]).IsEqualTo((byte)'P');
        await Assert.That(result.StandardOutput[1]).IsEqualTo((byte)'K');
    }

    [Test]
    public async Task CLI188_ExcelCreate_OutputExists_FailsWithoutForce()
    {
        var dir = CliTestRunner.CreateTempDirectory("excel-create-overwrite");
        var input = await WriteWorkbookJsonAsync(dir, SimpleWorkbook()).ConfigureAwait(false);
        var output = new FileInfo(Path.Combine(dir.FullName, "existing.xlsx"));
        await File.WriteAllTextAsync(output.FullName, "placeholder").ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync("excel", "create", input.FullName, "--output", output.FullName)
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsNotEqualTo(0);
        await Assert.That(result.StandardError).Contains("already exists");
    }

    [Test]
    public async Task CLI189_ExcelCreate_OutputExistsWithForce_Overwrites()
    {
        var dir = CliTestRunner.CreateTempDirectory("excel-create-force");
        var input = await WriteWorkbookJsonAsync(dir, SimpleWorkbook()).ConfigureAwait(false);
        var output = new FileInfo(Path.Combine(dir.FullName, "overwrite.xlsx"));
        await File.WriteAllTextAsync(output.FullName, "placeholder").ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync("excel", "create", input.FullName, "--output", output.FullName, "--force")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await ValidateXlsxAsync(output).ConfigureAwait(false);
    }

    [Test]
    public async Task CLI190_ExcelCreate_MalformedJson_ReturnsInvalidFormat()
    {
        var dir = CliTestRunner.CreateTempDirectory("excel-create-invalid-json");
        var input = new FileInfo(Path.Combine(dir.FullName, "bad.json"));
        await File.WriteAllTextAsync(input.FullName, "{ not valid json }").ConfigureAwait(false);

        var result = await CliTestRunner.RunManagedAsync("excel", "create", input.FullName).ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(4);
        using var json = result.ReadStderrJson();
        await Assert.That(json.RootElement.GetProperty("code").GetString()).IsEqualTo("INVALID_FORMAT");
    }

    [Test]
    public async Task CLI191_ExcelCreate_EmptyWorksheets_ReturnsInvalidArguments()
    {
        var dir = CliTestRunner.CreateTempDirectory("excel-create-empty");
        var content = """{ "worksheets": [] }""";
        var input = await WriteWorkbookJsonAsync(dir, content).ConfigureAwait(false);

        var result = await CliTestRunner.RunManagedAsync("excel", "create", input.FullName).ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(2);
        using var json = result.ReadStderrJson();
        await Assert.That(json.RootElement.GetProperty("code").GetString()).IsEqualTo("INVALID_ARGUMENTS");
    }

    [Test]
    public async Task CLI192_ExcelCreate_QuietMode_SuppressesPayload()
    {
        var dir = CliTestRunner.CreateTempDirectory("excel-create-quiet");
        var input = await WriteWorkbookJsonAsync(dir, SimpleWorkbook()).ConfigureAwait(false);
        var output = new FileInfo(Path.Combine(dir.FullName, "quiet.xlsx"));

        var result = await CliTestRunner
            .RunManagedAsync("excel", "create", input.FullName, "--output", output.FullName, "--quiet")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardOutput).IsEmpty();
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(output.Exists).IsTrue();
    }

    [Test]
    public async Task CLI193_ExcelCreate_InvalidSheetName_ReturnsInvalidArguments()
    {
        var dir = CliTestRunner.CreateTempDirectory("excel-create-invalid-sheet");
        var content = """
            {
              "worksheets": [
                {
                  "name": "Bad/Name",
                  "rows": [{ "cells": [{ "value": "x" }] }]
                }
              ]
            }
            """;
        var input = await WriteWorkbookJsonAsync(dir, content).ConfigureAwait(false);

        var result = await CliTestRunner.RunManagedAsync("excel", "create", input.FullName).ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(2);
        using var json = result.ReadStderrJson();
        await Assert.That(json.RootElement.GetProperty("code").GetString()).IsEqualTo("INVALID_ARGUMENTS");
    }
}
#pragma warning restore CA1707

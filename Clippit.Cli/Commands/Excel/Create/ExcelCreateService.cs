using System.Text.Json;
using Clippit.Cli.Infrastructure;
using Clippit.Excel;

namespace Clippit.Cli.Commands.Excel.Create;

internal static class ExcelCreateService
{
    public static CreateResult Execute(InputSource input, OutputTarget output, bool force)
    {
        WorkbookDefinition definition;
        try
        {
            using var stream = input.OpenSeekable();
            definition =
                JsonSerializer.Deserialize(stream, CliJsonContext.Default.WorkbookDefinition)
                ?? throw CliException.InvalidFormat("Workbook definition JSON is null.");
        }
        catch (JsonException ex)
        {
            throw CliException.InvalidFormat($"Invalid workbook definition JSON: {ex.Message}");
        }

        if (definition.Worksheets is null || definition.Worksheets.Count == 0)
            throw CliException.InvalidArguments("Workbook definition must contain at least one worksheet.");

        var workbook = MapWorkbook(definition);

        var memStream = new MemoryStream();
        try
        {
            workbook.WriteTo(memStream);
        }
        catch (Exception ex) when (ex is InvalidSheetNameException or WorksheetAlreadyExistsException)
        {
            throw CliException.InvalidArguments(ex.Message);
        }
        catch (Exception ex) when (ex is not CliException)
        {
            throw CliException.InvalidFormat($"Could not generate workbook: {ex.Message}");
        }

        var outputSize = memStream.Length;
        memStream.Position = 0;

        string? tempPath = null;
        try
        {
            output.EnsureCanWrite(force, "output");
            output.EnsureDirectoryExists();
            using (var outStream = output.OpenWrite(out tempPath))
            {
                memStream.CopyTo(outStream);
                outStream.Flush();

                if (output.IsStdout)
                    output.Flush(outStream);
            }

            if (!output.IsStdout)
                output.Commit(tempPath);
        }
        catch (Exception ex) when (ex is not CliException)
        {
            throw CliException.OutputError($"Could not write output: {ex.Message}");
        }
        finally
        {
            OutputTarget.DeleteTemp(tempPath);
        }

        return new CreateResult
        {
            Input = input.DisplayName,
            Output = output.DisplayPath,
            OutputSize = outputSize,
            WorksheetCount = definition.Worksheets.Count,
        };
    }

    private static WorkbookDfn MapWorkbook(WorkbookDefinition definition) =>
        new() { Worksheets = definition.Worksheets.Select(MapWorksheet).ToList() };

    private static WorksheetDfn MapWorksheet(WorksheetDefinition ws)
    {
        var dfn = new WorksheetDfn
        {
            Name = ws.Name,
            Rows = (ws.Rows ?? [])
                .Select(r => new RowDfn { Cells = (r.Cells ?? []).Select(MapCell).ToList() })
                .ToList(),
        };
        // TableName and ColumnHeadings are optional in WorksheetDfn (legacy API accepts null)
        if (ws.TableName is not null)
            dfn.TableName = ws.TableName;
        if (ws.ColumnHeadings is { Count: > 0 })
            dfn.ColumnHeadings = ws.ColumnHeadings.Select(MapCell).ToList();
        return dfn;
    }

    private static CellDfn MapCell(CellDefinition cell)
    {
        var dataType = MapCellDataType(cell.CellDataType);
        var dfn = new CellDfn
        {
            // CellDfn is a legacy class; null assignments are safe at runtime
            Value = ConvertValue(cell.Value, dataType),
            CellDataType = dataType,
            HorizontalCellAlignment = MapAlignment(cell.HorizontalCellAlignment),
            Bold = cell.Bold,
            Italic = cell.Italic,
        };
        if (cell.FormatCode is not null)
            dfn.FormatCode = cell.FormatCode;
        return dfn;
    }

    private static object ConvertValue(System.Text.Json.JsonElement? element, CellDataType? dataType)
    {
        if (element is null)
            return string.Empty;

        return element.Value.ValueKind switch
        {
            System.Text.Json.JsonValueKind.Null => string.Empty,
            System.Text.Json.JsonValueKind.True => true,
            System.Text.Json.JsonValueKind.False => false,
            System.Text.Json.JsonValueKind.Number => element.Value.GetDouble(),
            System.Text.Json.JsonValueKind.String => ParseStringValue(
                element.Value.GetString() ?? string.Empty,
                dataType
            ),
            _ => string.Empty,
        };
    }

    private static object ParseStringValue(string value, CellDataType? dataType)
    {
        if (
            dataType == CellDataType.Date
            && DateTime.TryParse(
                value,
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.RoundtripKind,
                out var dt
            )
        )
            return dt;

        return value;
    }

    private static CellDataType? MapCellDataType(CellDataTypeJson? json) =>
        json switch
        {
            CellDataTypeJson.Boolean => CellDataType.Boolean,
            CellDataTypeJson.Date => CellDataType.Date,
            CellDataTypeJson.Number => CellDataType.Number,
            CellDataTypeJson.String => CellDataType.String,
            _ => null,
        };

    private static HorizontalCellAlignment? MapAlignment(HorizontalAlignmentJson? json) =>
        json switch
        {
            HorizontalAlignmentJson.Left => HorizontalCellAlignment.Left,
            HorizontalAlignmentJson.Center => HorizontalCellAlignment.Center,
            HorizontalAlignmentJson.Right => HorizontalCellAlignment.Right,
            _ => null,
        };
}

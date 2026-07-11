using System.Text.Json;
using System.Text.Json.Serialization;

namespace Clippit.Cli.Commands.Excel.Create;

internal sealed class WorkbookDefinition
{
    public List<WorksheetDefinition> Worksheets { get; set; } = [];
}

internal sealed class WorksheetDefinition
{
    public required string Name { get; set; }
    public string? TableName { get; set; }
    public List<CellDefinition>? ColumnHeadings { get; set; }
    public List<RowDefinition> Rows { get; set; } = [];
}

internal sealed class RowDefinition
{
    public List<CellDefinition> Cells { get; set; } = [];
}

/// <summary>
/// JSON representation of a spreadsheet cell. <see cref="Value"/> is a raw
/// <see cref="JsonElement"/> so it can hold string, number, boolean, or null
/// without a custom AOT converter.
/// </summary>
internal sealed class CellDefinition
{
    public JsonElement? Value { get; set; }

    [JsonConverter(typeof(JsonStringEnumConverter<CellDataTypeJson>))]
    public CellDataTypeJson? CellDataType { get; set; }

    public string? FormatCode { get; set; }

    [JsonConverter(typeof(JsonStringEnumConverter<HorizontalAlignmentJson>))]
    public HorizontalAlignmentJson? HorizontalCellAlignment { get; set; }

    public bool? Bold { get; set; }
    public bool? Italic { get; set; }
}

internal enum CellDataTypeJson
{
    Boolean,
    Date,
    Number,
    String,
}

internal enum HorizontalAlignmentJson
{
    Left,
    Center,
    Right,
}

namespace Clippit.Cli.Commands.Common.Verify;

/// <summary>
/// Result of the "verify" command for Word, Excel, or PowerPoint documents.
/// </summary>
internal sealed record VerifyResult
{
    public required string Input { get; init; }
    public required string OfficeVersion { get; init; }
    public required bool Valid { get; init; }
    public required IReadOnlyList<VerifyDiagnostic> Diagnostics { get; init; }

    public static void WriteText(VerifyResult result, TextWriter writer)
    {
        writer.WriteLine($"{(result.Valid ? "Valid" : "Invalid")} document: {result.Input}");
        writer.WriteLine($"Office version: {result.OfficeVersion}");

        if (result.Valid)
            return;

        writer.WriteLine($"Diagnostics: {result.Diagnostics.Count}");
        foreach (var diagnostic in result.Diagnostics)
        {
            var code = diagnostic.Code is not null ? $"/{diagnostic.Code}" : string.Empty;
            var location = diagnostic.Part is not null ? $" {diagnostic.Part}:" : string.Empty;
            writer.WriteLine($"  [{diagnostic.Kind}{code}]{location} {diagnostic.Description}");
        }
    }
}

internal sealed record VerifyDiagnostic
{
    public required string Kind { get; init; }
    public string? Code { get; init; }
    public required string Description { get; init; }
    public string? Part { get; init; }
    public string? Path { get; init; }
    public string? Element { get; init; }
    public string? Attribute { get; init; }
    public string? RelationshipId { get; init; }
}

using System.Text.Json.Serialization;
using Clippit.Cli.Infrastructure;

namespace Clippit.Cli.Commands.Pptx.Build;

/// <summary>
/// Top-level deck manifest — serialized as clippit-deck.json.
/// All file paths resolve relative to the manifest file's directory.
/// </summary>
internal sealed class BuildManifest
{
    [JsonPropertyName("$schema")]
    public string? Schema { get; set; }

    /// <summary>Written to CoreProperties.Title of the output PPTX.</summary>
    public required string Title { get; set; }

    /// <summary>Output file path (relative to manifest dir, or absolute).</summary>
    public required string Output { get; set; }

    /// <summary>
    /// Ordered list of deck entries defining the final presentation.
    /// Each entry is either a plain string (section divider or file path)
    /// or an object with explicit options.
    ///
    /// String shorthand:
    ///   "[Section Name]"    → named PPTX section divider
    ///   "path/to/file.pptx" → append all slides from this file
    ///
    /// Object form:
    ///   { "section": "Section Name" }
    ///   { "file": "path/to/file.pptx", "masters": true, "slides": false, "keepSections": true }
    ///
    /// Object options:
    ///   masters      → copy all slide masters/layouts from this file instead of only those used by copied slides
    ///   slides       → copy slides from this file; set false for a masters-only template entry
    ///   keepSections → preserve this file's own internal PPTX section structure
    /// </summary>
    [JsonConverter(typeof(DeckEntryListConverter))]
    public required IList<DeckEntry> Deck { get; set; }
}

/// <summary>
/// A single entry in the deck list. Exactly one of <see cref="Section"/> or
/// <see cref="File"/> must be non-null.
/// </summary>
internal sealed class DeckEntry
{
    /// <summary>
    /// Non-null → this entry is a named PPTX section divider.
    /// Corresponds to the string shorthand "[Name]".
    /// </summary>
    public string? Section { get; init; }

    /// <summary>
    /// Non-null → this entry is a source .pptx file.
    /// Path is relative to manifest dir or absolute.
    /// </summary>
    public string? File { get; init; }

    /// <summary>
    /// How to handle slide masters and layouts from this file.
    ///   null / false → lazy: copy only the masters/layouts used by the slides
    ///                  being added (default — produces the smallest output).
    ///   true         → eager: copy all masters and all their layouts from this
    ///                  file (use for a dedicated brand template file).
    /// </summary>
    public bool? Masters { get; init; }

    /// <summary>
    /// Whether to copy slides from this file. Default: true.
    /// Set false combined with masters=true for a masters-only template entry
    /// that contributes no slides.
    /// </summary>
    public bool? Slides { get; init; }

    /// <summary>
    /// When true, the file's own internal PPTX section structure is imported
    /// and preserved in the output instead of placing slides into the current
    /// manifest section. Default: false.
    /// </summary>
    public bool? KeepSections { get; init; }

    // -------------------------------------------------------------------------
    // Resolved defaults
    // -------------------------------------------------------------------------

    [JsonIgnore]
    public bool IsSection => Section is not null;

    [JsonIgnore]
    public bool IsFile => File is not null;

    [JsonIgnore]
    public bool ShouldCopySlides => Slides ?? true;

    [JsonIgnore]
    public bool ShouldCopyAllMasters => Masters ?? false;

    [JsonIgnore]
    public bool ShouldKeepSections => KeepSections ?? false;

    // -------------------------------------------------------------------------
    // String shorthand factory
    // -------------------------------------------------------------------------

    /// <summary>
    /// Parses a plain string entry:
    ///   "[Name]"    → section divider
    ///   "path.pptx" → file entry (all slides, lazy masters)
    /// </summary>
    public static DeckEntry FromString(string value)
    {
        var trimmed = value.Trim();
        if (trimmed.StartsWith('[') && trimmed.EndsWith(']'))
            return new DeckEntry { Section = trimmed[1..^1].Trim() };
        return new DeckEntry { File = trimmed };
    }

    public string? Validate()
    {
        if (Section is null && File is null)
            return "Entry has neither 'section' nor 'file' — one is required.";
        if (Section is not null && File is not null)
            return "Entry has both 'section' and 'file' — only one is allowed.";
        if (Section is { } section && string.IsNullOrWhiteSpace(section))
            return "Entry has an empty 'section' value.";
        if (File is { } file && string.IsNullOrWhiteSpace(file))
            return "Entry has an empty 'file' value.";
        return null;
    }
}

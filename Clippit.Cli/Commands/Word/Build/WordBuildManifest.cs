using System.Text.Json.Serialization;

namespace Clippit.Cli.Commands.Word.Build;

/// <summary>
/// Top-level Word build manifest — serialized as clippit-word-build.json.
/// All file paths resolve relative to the manifest file's directory.
/// </summary>
internal sealed class WordBuildManifest
{
    [JsonPropertyName("$schema")]
    public string? Schema { get; set; }

    /// <summary>Output file path (relative to manifest dir, or absolute).</summary>
    public required string Output { get; set; }

    /// <summary>
    /// Ordered list of deck entries defining the merged document.
    /// Each entry is either a plain string (section label or file path)
    /// or an object with explicit options.
    ///
    /// String shorthand:
    ///   "[Section Name]"     → named logical section label (no content added)
    ///   "path/to/file.docx"  → append all body elements from this file
    ///
    /// Object form:
    ///   { "section": "Section Name" }
    ///   { "file": "path/to/file.docx", "start": 0, "count": 50,
    ///     "keepSections": true, "discardHeadersAndFootersInKeptSections": false }
    /// </summary>
    [JsonConverter(typeof(WordDeckEntryListConverter))]
    public required IList<WordDeckEntry> Deck { get; set; }
}

/// <summary>
/// A single entry in the Word build deck. Exactly one of <see cref="Section"/>
/// or <see cref="File"/> must be non-null.
/// </summary>
internal sealed class WordDeckEntry
{
    /// <summary>Non-null → logical section label; no content is added to the document.</summary>
    public string? Section { get; init; }

    /// <summary>Non-null → source .docx path (relative to manifest dir or absolute).</summary>
    public string? File { get; init; }

    /// <summary>0-based index of the first body element to copy. Defaults to 0.</summary>
    public int? Start { get; init; }

    /// <summary>Maximum number of body elements to copy. Defaults to all remaining.</summary>
    public int? Count { get; init; }

    /// <summary>When true, preserve this source document's section structure. Defaults to false.</summary>
    public bool? KeepSections { get; init; }

    /// <summary>When KeepSections is true, discard the source's headers/footers. Defaults to false.</summary>
    public bool? DiscardHeadersAndFootersInKeptSections { get; init; }

    [JsonIgnore]
    public bool IsSection => Section is not null;

    [JsonIgnore]
    public bool IsFile => File is not null;

    /// <summary>
    /// Parses a plain string entry:
    ///   "[Name]"    → section label
    ///   "path.docx" → file entry (all elements, default settings)
    /// </summary>
    public static WordDeckEntry FromString(string value)
    {
        var trimmed = value.Trim();
        if (trimmed.StartsWith('[') && trimmed.EndsWith(']'))
            return new WordDeckEntry { Section = trimmed[1..^1].Trim() };
        return new WordDeckEntry { File = trimmed };
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
        if (Start is < 0)
            return "Entry 'start' must be a non-negative integer.";
        if (Count is < 1)
            return "Entry 'count' must be a positive integer.";
        return null;
    }
}

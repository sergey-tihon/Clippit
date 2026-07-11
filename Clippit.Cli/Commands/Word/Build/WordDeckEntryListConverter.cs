using System.Text.Json;
using System.Text.Json.Serialization;

namespace Clippit.Cli.Commands.Word.Build;

/// <summary>
/// Deserializes the "deck" array where each element is either:
///   - a JSON string  → parsed via WordDeckEntry.FromString()
///                        "[Name]"      → section label
///                        "file.docx"   → file entry
///   - a JSON object  → standard WordDeckEntry deserialization (full options)
///
/// Serializes using string shorthand when only <c>section</c> or <c>file</c> is set,
/// and falls back to objects for entries with additional options.
/// Fully AOT-safe: uses CliJsonContext source-gen for the object branch.
/// </summary>
internal sealed class WordDeckEntryListConverter : JsonConverter<IList<WordDeckEntry>>
{
    public override IList<WordDeckEntry> Read(
        ref Utf8JsonReader reader,
        Type typeToConvert,
        JsonSerializerOptions options
    )
    {
        if (reader.TokenType != JsonTokenType.StartArray)
            throw new JsonException("Expected start of array for 'deck'.");

        var list = new List<WordDeckEntry>();
        while (reader.Read() && reader.TokenType != JsonTokenType.EndArray)
        {
            switch (reader.TokenType)
            {
                case JsonTokenType.String:
                    list.Add(WordDeckEntry.FromString(reader.GetString()!));
                    break;
                case JsonTokenType.StartObject:
                {
                    var entry = JsonSerializer.Deserialize(ref reader, CliJsonContext.Default.WordDeckEntry);
                    if (entry is not null)
                        list.Add(entry);
                    break;
                }
                default:
                    throw new JsonException(
                        $"Unexpected token '{reader.TokenType}' in 'deck' array. "
                            + "Each entry must be a string or an object."
                    );
            }
        }

        return list;
    }

    public override void Write(Utf8JsonWriter writer, IList<WordDeckEntry> value, JsonSerializerOptions options)
    {
        writer.WriteStartArray();
        foreach (var entry in value)
        {
            if (CanUseStringShorthand(entry))
            {
                if (entry.Section is not null)
                    writer.WriteStringValue($"[{entry.Section}]");
                else
                    writer.WriteStringValue(entry.File);
            }
            else
            {
                JsonSerializer.Serialize(writer, entry, CliJsonContext.Default.WordDeckEntry);
            }
        }
        writer.WriteEndArray();
    }

    private static bool CanUseStringShorthand(WordDeckEntry entry) =>
        (entry.Section is null) != (entry.File is null)
        && entry is { Start: null, Count: null, KeepSections: null, DiscardHeadersAndFootersInKeptSections: null };
}

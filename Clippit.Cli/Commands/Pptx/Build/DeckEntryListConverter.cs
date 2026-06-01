using System.Text.Json;
using System.Text.Json.Serialization;

namespace Clippit.Cli.Commands.Pptx.Build;

/// <summary>
/// Deserializes the "deck" array where each element is either:
///   - a JSON string  → parsed via DeckEntry.FromString()
///                        "[Name]"      → section divider
///                        "file.pptx"   → file entry
///   - a JSON object  → standard DeckEntry deserialization (full options)
///
/// Serializes always as objects for a stable round-trip.
/// Fully AOT-safe: uses CliJsonContext source-gen for the object branch.
/// </summary>
internal sealed class DeckEntryListConverter : JsonConverter<IList<DeckEntry>>
{
    public override IList<DeckEntry> Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        if (reader.TokenType != JsonTokenType.StartArray)
            throw new JsonException("Expected start of array for 'deck'.");

        var list = new List<DeckEntry>();
        while (reader.Read() && reader.TokenType != JsonTokenType.EndArray)
        {
            switch (reader.TokenType)
            {
                case JsonTokenType.String:
                    list.Add(DeckEntry.FromString(reader.GetString()!));
                    break;
                case JsonTokenType.StartObject:
                {
                    var entry = JsonSerializer.Deserialize(ref reader, CliJsonContext.Default.DeckEntry);
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

    public override void Write(Utf8JsonWriter writer, IList<DeckEntry> value, JsonSerializerOptions options)
    {
        writer.WriteStartArray();
        foreach (var entry in value)
        {
            // Emit the compact string shorthand when no extra options are set;
            // this keeps round-tripped manifests human-friendly.
            if (CanUseStringShorthand(entry))
            {
                if (entry.Section is not null)
                    writer.WriteStringValue($"[{entry.Section}]");
                else
                    writer.WriteStringValue(entry.File);
            }
            else
            {
                JsonSerializer.Serialize(writer, entry, CliJsonContext.Default.DeckEntry);
            }
        }
        writer.WriteEndArray();
    }

    private static bool CanUseStringShorthand(DeckEntry entry) =>
        (entry.Section is null) != (entry.File is null) && entry is { Masters: null, Slides: null, KeepSections: null };
}

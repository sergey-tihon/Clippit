---
uid: Tutorial.Common.RelationshipValidator
---

# RelationshipValidator

Namespace: `Clippit.Core`

Detects **dangling relationship references** in any OpenXml package (DOCX, PPTX, XLSX).

```csharp
public static class RelationshipValidator
{
    // Returns one error per unresolvable relationship attribute value.
    public static IEnumerable<RelationshipValidationError> Validate(OpenXmlPackage package);

    // Convenience: true when no dangling references are found.
    public static bool IsValid(OpenXmlPackage package);
}
```

## Background

The built-in [`OpenXmlValidator`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.validation.openxmlvalidator)
checks element-level schema conformance but does **not** verify relationship-reference integrity.
A dangling `r:id` — an attribute whose value does not correspond to any relationship registered
with that part — silently passes schema validation yet throws `KeyNotFoundException` at runtime
when code tries to follow it (e.g., during slide copying or publishing).

`RelationshipValidator` fills this gap by scanning every XML part of the package and reporting
any attribute from the relationship namespace (`r:id`, `r:embed`, `r:link`, `r:cs`, `r:dm`,
`r:lo`, `r:qs`, `r:href`, `r:pict`, `r:blip`) whose value is not registered on that part.

## RelationshipValidationError

Each error is a `record` with the following properties:

| Property | Type | Description |
|---|---|---|
| `PartUri` | `Uri` | URI of the part that owns the problematic attribute |
| `ElementName` | `XName` | XML element carrying the dangling relationship ID |
| `AttributeName` | `XName` | XML attribute holding the unresolved ID |
| `RelationshipId` | `string` | The relationship ID value that could not be resolved |
| `Description` | `string` | Human-readable diagnostic message |

## Usage

### Check whether a package is valid

```csharp
using Clippit.Core;
using DocumentFormat.OpenXml.Packaging;

using var doc = PresentationDocument.Open("deck.pptx", false);
if (!RelationshipValidator.IsValid(doc))
    Console.WriteLine("Package contains dangling relationship references!");
```

### Enumerate and log all errors

```csharp
using Clippit.Core;
using DocumentFormat.OpenXml.Packaging;

using var doc = WordprocessingDocument.Open("document.docx", false);
foreach (var error in RelationshipValidator.Validate(doc))
{
    Console.WriteLine(error.Description);
    // e.g.: Part '/word/document.xml': element 'blip' attribute 'embed'
    //       references relationship ID 'rId999' which is not registered on this part.
}
```

### Guard before processing

Running the validator before slide-copy or publish operations lets callers detect the
problem early and skip or repair the affected element instead of crashing:

```csharp
using Clippit.Core;
using Clippit.PowerPoint;
using DocumentFormat.OpenXml.Packaging;

using var src = PresentationDocument.Open("source.pptx", false);
var danglingParts = RelationshipValidator
    .Validate(src)
    .Select(e => e.PartUri)
    .ToHashSet();

if (danglingParts.Count > 0)
    Console.WriteLine($"Warning: {danglingParts.Count} part(s) have dangling references.");
```

## Related

- [`OpenXmlValidator`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.validation.openxmlvalidator) — schema-level validation (complements `RelationshipValidator`)
- [Issue #155](https://github.com/sergey-tihon/Clippit/issues/155) — root-cause issue that motivated this utility

# Clippit — Fresh PowerTools for OpenXml

![NuGet Version](https://badgen.net/nuget/v/Clippit) ![NuGet Downloads](https://badgen.net/nuget/dt/Clippit)

<img style="float: right;" src="images/logo.jpeg">

Clippit is a .NET library for programmatically creating, modifying, and converting
Word (DOCX), Excel (XLSX), and PowerPoint (PPTX) documents. Built on top of the
[Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK), it provides high-level
APIs that handle the complexity of the Open XML format so you can focus on your content.

## Getting Started

Install from NuGet:

```bash
dotnet add package Clippit
```

Split a PowerPoint presentation into individual slides:

```csharp
using Clippit.PowerPoint;

var presentation = new PmlDocument("conference-deck.pptx");
var slides = PresentationBuilder.PublishSlides(presentation);

foreach (var slide in slides)
{
    slide.SaveAs(Path.Combine("output", slide.FileName));
}
```

## Features

Clippit covers a broad range of document processing scenarios across all three
Office formats. Every feature listed below has a dedicated tutorial with API
signatures and code samples.

### Word

| Feature                                                                     | Description                                                                                                                                                                                      |
| --------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| [DocumentAssembler](xref:Tutorial.Word.DocumentAssembler.DocumentTemplates) | Populate DOCX templates with data from XML, including [images](xref:Tutorial.Word.DocumentAssembler.ImagesSupport) and [inline HTML](xref:Tutorial.Word.DocumentAssembler.InlineHtmlSupport)     |
| [DocumentBuilder](xref:Tutorial.Word.DocumentBuilder.ISource)               | Merge, split, and reorganize DOCX files with an [extensible ISource model](xref:Tutorial.Word.DocumentBuilder.ISource) and [TableCellSource](xref:Tutorial.Word.DocumentBuilder.TableCellSource) |
| [WmlComparer](xref:Tutorial.Word.WmlComparer)                               | Compare two DOCX files and produce a diff with revision tracking markup                                                                                                                          |
| [WmlToHtmlConverter](xref:Tutorial.Word.WmlToHtmlConverter)                 | High-fidelity conversion from DOCX to HTML/CSS                                                                                                                                                   |
| [HtmlToWmlConverter](xref:Tutorial.Word.HtmlToWmlConverter)                 | Convert HTML/CSS back into a properly structured DOCX                                                                                                                                            |
| [RevisionProcessor](xref:Tutorial.Word.RevisionProcessor)                   | Accept or reject tracked revisions programmatically                                                                                                                                              |
| [MarkupSimplifier](xref:Tutorial.Word.MarkupSimplifier)                     | Clean up and normalize DOCX markup for easier processing                                                                                                                                         |

### Excel

| Feature                                                    | Description                                                                                                                   |
| ---------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------- |
| [SpreadsheetWriter](xref:Tutorial.Excel.SpreadsheetWriter) | Generate multi-sheet XLSX files with formatted tables, streaming support for millions of rows, and a concise Cell Builder API |
| [SmlDataRetriever](xref:Tutorial.Excel.SmlDataRetriever)   | Extract data and formatting from existing spreadsheets as structured XML                                                      |

### PowerPoint

| Feature                                                             | Description                                                                                                                                                                                                                        |
| ------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| [PresentationBuilder](xref:Tutorial.PowerPoint.PresentationBuilder) | Merge and split PPTX files, with a [Fluent API](xref:Tutorial.PowerPoint.BuildPresentation.FluentApi) for ergonomic slide composition and optimized [slide publishing](xref:Tutorial.PowerPoint.PresentationBuilder.PublishSlides) |

### Common

| Feature                                             | Description                                                             |
| --------------------------------------------------- | ----------------------------------------------------------------------- |
| [OpenXmlRegex](xref:Tutorial.Common.OpenXmlRegex)   | Search and replace content across DOCX/PPTX using regular expressions   |
| [MetricsGetter](xref:Tutorial.Common.MetricsGetter) | Retrieve document metrics — style hierarchy, languages, fonts, and more |

## Compatibility

- **Targets:** `net8.0` and `net10.0`
- **Dependency:** [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml) (Open XML SDK)
- **Platforms:** Windows and Linux (continuously tested on both)
- **Side-by-side:** Can coexist with the original Open-Xml-PowerTools assembly

## Heritage

Clippit originated as a fork of [Open-Xml-PowerTools](https://github.com/EricWhiteDev/Open-Xml-PowerTools)
and has since evolved into an independently maintained library with new features,
performance improvements, and modern .NET support. See the
[Changelog](api/CHANGELOG.md) for the full release history.

## Questions and Contributing

Have a question or idea? Start a [GitHub Discussion](https://github.com/sergey-tihon/Clippit/discussions).

Found a bug or want to request a feature? Open an [Issue](https://github.com/sergey-tihon/Clippit/issues).

```
Copyright (c) Microsoft Corporation 2012-2017
Portions Copyright (c) Eric White Inc 2018-2019
Portions Copyright (c) Sergey Tihon 2019-2026
Licensed under the MIT License.
```

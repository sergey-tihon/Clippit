# Clippit — Fresh PowerTools for OpenXml

[![NuGet Version](https://badgen.net/nuget/v/Clippit)](https://www.nuget.org/packages/Clippit)
[![NuGet Downloads](https://badgen.net/nuget/dt/Clippit)](https://www.nuget.org/packages/Clippit)
[![Build and Test](https://github.com/sergey-tihon/Clippit/actions/workflows/main.yml/badge.svg)](https://github.com/sergey-tihon/Clippit/actions/workflows/main.yml)
[![License: MIT](https://badgen.net/badge/license/MIT/blue)](LICENSE)

Clippit is a .NET library for programmatically creating, modifying, and converting
Word (DOCX), Excel (XLSX), and PowerPoint (PPTX) documents. Built on top of the
[Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK), it provides high-level
APIs that handle the complexity of the Open XML format so you can focus on your content.

📖 **[Full documentation and tutorials →](https://sergey-tihon.github.io/Clippit/)**

## Installation

```bash
dotnet add package Clippit
```

## Quick Start

```csharp
using Clippit.PowerPoint;

// Split a presentation into individual slides
var presentation = new PmlDocument("deck.pptx");
var slides = PresentationBuilder.PublishSlides(presentation);
foreach (var slide in slides)
    slide.SaveAs(Path.Combine("output", slide.FileName));
```

## Features

| Area        | Capabilities                                                                 |
| ----------- | ---------------------------------------------------------------------------- |
| **Word**    | Template assembler, merge/split docs, track-change diff, DOCX↔HTML, regex   |
| **Excel**   | Generate XLSX with formatted tables, streaming writes, data extraction       |
| **PowerPoint** | Merge/split PPTX, fluent slide builder, publish slides                   |
| **Common**  | OpenXml regex search/replace, document metrics                               |

See the [tutorials](https://sergey-tihon.github.io/Clippit/tutorials/) for API signatures and full code samples.

## Compatibility

- **Targets:** `net8.0` and `net10.0`
- **Platforms:** Windows and Linux (continuously tested)
- **Dependency:** [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml) (Open XML SDK)

## Contributing

Questions and ideas → [GitHub Discussions](https://github.com/sergey-tihon/Clippit/discussions)  
Bugs and feature requests → [GitHub Issues](https://github.com/sergey-tihon/Clippit/issues)

To build locally:

```bash
./build.sh   # Unix
build.cmd    # Windows
```

To preview the docs site locally:

```bash
dotnet tool restore
dotnet docfx docs/docfx.json --serve
```

## Heritage

Clippit originated as a fork of [Open-Xml-PowerTools](https://github.com/EricWhiteDev/Open-Xml-PowerTools)
and has since evolved into an independently maintained library. See the [Changelog](https://sergey-tihon.github.io/Clippit/api/CHANGELOG.html) for the full release history.

```
Copyright (c) Microsoft Corporation 2012-2017
Portions Copyright (c) Eric White Inc 2018-2019
Portions Copyright (c) Sergey Tihon 2019-2026
Licensed under the MIT License.
```

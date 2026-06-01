<style>
  .clippit-landing {
    --clippit-bg: #11171f;
    --clippit-panel: rgba(18, 31, 42, 0.72);
    --clippit-border: rgba(160, 255, 237, 0.2);
    --clippit-text: #eaf4fb;
    --clippit-muted: #aebcc8;
    --clippit-cyan: #76fff0;
    --clippit-mint: #62ff9d;
    --clippit-blue: #7b9cff;
    margin: -0.5rem 0 2.5rem;
  }

  .clippit-hero {
    display: grid;
    grid-template-columns: minmax(0, 1.05fr) minmax(320px, 0.95fr);
    gap: clamp(2rem, 5vw, 5rem);
    align-items: center;
    position: relative;
    overflow: hidden;
    padding: clamp(2rem, 5vw, 4.8rem);
    border: 1px solid var(--clippit-border);
    border-radius: 28px;
    background:
      radial-gradient(circle at 18% 18%, rgba(98, 255, 157, 0.13), transparent 32%),
      radial-gradient(circle at 70% 30%, rgba(118, 255, 240, 0.2), transparent 38%),
      linear-gradient(135deg, rgba(16, 22, 31, 0.96), rgba(20, 32, 43, 0.9) 48%, rgba(12, 16, 24, 0.98));
    box-shadow: 0 28px 90px rgba(0, 0, 0, 0.32);
  }

  .clippit-hero::before {
    content: "";
    position: absolute;
    inset: -1px;
    background:
      linear-gradient(90deg, rgba(118, 255, 240, 0.14) 1px, transparent 1px),
      linear-gradient(0deg, rgba(118, 255, 240, 0.08) 1px, transparent 1px);
    background-size: 44px 44px;
    mask-image: radial-gradient(circle at 70% 35%, black, transparent 68%);
    pointer-events: none;
  }

  .clippit-hero-copy,
  .clippit-hero-art {
    position: relative;
    z-index: 1;
  }

  .clippit-eyebrow {
    display: inline-flex;
    gap: 0.55rem;
    align-items: center;
    margin-bottom: 1.2rem;
    color: var(--clippit-cyan);
    font-size: 0.78rem;
    font-weight: 700;
    letter-spacing: 0.16em;
    text-transform: uppercase;
  }

  .clippit-hero h1 {
    margin: 0;
    color: var(--clippit-text);
    font-size: clamp(3.5rem, 9vw, 7rem);
    line-height: 0.9;
    letter-spacing: -0.07em;
  }

  .clippit-hero h1 span {
    display: block;
    margin-top: 0.5rem;
    color: transparent;
    font-size: clamp(1.7rem, 3.5vw, 3.1rem);
    line-height: 1.05;
    letter-spacing: -0.045em;
    background: linear-gradient(90deg, var(--clippit-cyan), var(--clippit-mint), var(--clippit-blue));
    -webkit-background-clip: text;
    background-clip: text;
  }

  .clippit-hero p {
    max-width: 64ch;
    margin: 1.5rem 0 0;
    color: var(--clippit-muted);
    font-size: clamp(1rem, 1.35vw, 1.18rem);
    line-height: 1.75;
  }

  .clippit-badges {
    display: flex;
    flex-wrap: wrap;
    gap: 0.45rem;
    margin: 1.4rem 0 0;
  }

  .clippit-actions {
    display: flex;
    flex-wrap: wrap;
    gap: 0.8rem;
    margin-top: 2rem;
  }

  .clippit-button {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    min-height: 44px;
    padding: 0 1.05rem;
    border: 1px solid rgba(255, 255, 255, 0.13);
    border-radius: 999px;
    color: var(--clippit-text) !important;
    text-decoration: none !important;
    background: rgba(255, 255, 255, 0.07);
    backdrop-filter: blur(16px);
  }

  .clippit-button-primary {
    border-color: rgba(98, 255, 157, 0.42);
    color: #071812 !important;
    font-weight: 750;
    background: linear-gradient(135deg, #74ffbd, #76fff0);
    box-shadow: 0 18px 38px rgba(98, 255, 157, 0.18);
  }

  .clippit-hero-art {
    display: grid;
    place-items: center;
  }

  .clippit-hero-art img {
    width: min(100%, 560px);
    border-radius: 32px;
    filter: drop-shadow(0 30px 60px rgba(0, 0, 0, 0.34));
  }

  .clippit-install-grid {
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 1rem;
    margin: 1.2rem 0 2.5rem;
  }

  .clippit-install-card {
    padding: 1.1rem;
    border: 1px solid rgba(130, 155, 175, 0.28);
    border-radius: 18px;
    background: linear-gradient(180deg, rgba(255, 255, 255, 0.035), rgba(255, 255, 255, 0.015));
  }

  .clippit-install-card strong {
    display: block;
    margin-bottom: 0.55rem;
    color: var(--clippit-text);
  }

  .clippit-install-card pre {
    margin: 0;
  }

  .clippit-install-card code {
    white-space: pre-wrap;
  }

  @media (max-width: 980px) {
    .clippit-hero,
    .clippit-install-grid {
      grid-template-columns: 1fr;
    }

    .clippit-hero-art {
      order: -1;
    }
  }
</style>

<div class="clippit-landing">
  <section class="clippit-hero">
    <div class="clippit-hero-copy">
      <div class="clippit-eyebrow">.NET 10 • Open XML • Scriptable CLI</div>
      <h1>Clippit <span>Fresh PowerTools for OpenXml</span></h1>
      <div class="clippit-badges">
        <img alt="NuGet Version" src="https://badgen.net/nuget/v/Clippit" />
        <img alt="NuGet Downloads" src="https://badgen.net/nuget/dt/Clippit" />
        <img alt="CLI NuGet Version" src="https://badgen.net/nuget/v/Clippit.Cli" />
        <img alt="CLI npm Version" src="https://badgen.net/npm/v/clippit" />
        <img alt="CLI npm Downloads" src="https://badgen.net/npm/dt/clippit" />
      </div>
      <p>
        Create, transform, compare, split, and validate Office documents with a modern .NET library
        and a native CLI built for automation. Clippit wraps the Open XML SDK with high-level APIs
        for Word, Excel, and PowerPoint workflows.
      </p>
      <div class="clippit-actions">
        <a class="clippit-button clippit-button-primary" href="#getting-started">Get Started</a>
        <a class="clippit-button" href="cli.md">CLI Docs</a>
        <a class="clippit-button" href="api/CHANGELOG.md">API Reference</a>
      </div>
    </div>
    <div class="clippit-hero-art">
      <img src="images/hero.jpg" alt="Friendly Clippit agent reviewing Open XML documents" />
    </div>
  </section>
</div>

# Clippit — Fresh PowerTools for OpenXml

Clippit is a .NET library for programmatically creating, modifying, and converting Word (DOCX), Excel (XLSX), and PowerPoint (PPTX) documents. Built on top of the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK), it provides high-level APIs that handle the complexity of the Open XML format so you can focus on your content. It also includes a [scriptable CLI](cli.md) for PowerPoint split, build, and validation workflows.

## Getting Started

<div class="clippit-install-grid">
  <div class="clippit-install-card">
    <strong>Library from NuGet</strong>
    <pre><code>dotnet add package Clippit</code></pre>
  </div>
  <div class="clippit-install-card">
    <strong>CLI from NuGet</strong>
    <pre><code>dotnet tool install -g Clippit.Cli</code></pre>
  </div>
  <div class="clippit-install-card">
    <strong>CLI from npm</strong>
    <pre><code>npm install -g clippit</code></pre>
  </div>
</div>

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

### CLI

| Feature | Description |
| ------- | ----------- |
| [Clippit CLI](cli.md) | Scriptable PPTX split/build/verify commands with human-readable text output, stable JSON output, stdin/stdout support, and JSON schemas for automation |

### Common

| Feature                                             | Description                                                             |
| --------------------------------------------------- | ----------------------------------------------------------------------- |
| [OpenXmlRegex](xref:Tutorial.Common.OpenXmlRegex)   | Search and replace content across DOCX/PPTX using regular expressions   |
| [MetricsGetter](xref:Tutorial.Common.MetricsGetter) | Retrieve document metrics — style hierarchy, languages, fonts, and more |

## Compatibility

- **Target:** `net10.0`
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

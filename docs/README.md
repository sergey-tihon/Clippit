## Why Clippit?

Clippit is a fork of [Open-Xml-PowerTools](https://github.com/EricWhiteDev/Open-Xml-PowerTools) (currently owned by Eric White) with new features, fixes and performance optimizations.

Key highlights:
- Shipped as new [NuGet package](https://www.nuget.org/packages/Clippit) published from latest `master`.
- Target `netstandard2.0` and uses latest `C# 8.0` features.
- Continuously tested on Windows, macOS and Linux.
- Can be used side-by-side with any existing Open-Xml-PowerTools assembly.
- Provides optimized [slide publishing API](/PowerPoint/PublishSlides) <Badge type="success">New</Badge> and improved [PresentationBuilder](/PowerPoint/BuildPresentation) 


Most of existing content about Open-Xml-PowerTools is still relevant:
- [DocumentBuilder Resource Center](http://www.ericwhite.com/blog/documentbuilder-developer-center/)
- [PresentationBuilder Resource Center](http://www.ericwhite.com/blog/presentationbuilder-developer-center/)
- [WmlToHtmlConverter Resource Center](http://www.ericwhite.com/blog/wmltohtmlconverter-developer-center/)
- [DocumentAssembler Resource Center](http://www.ericwhite.com/blog/documentassembler-developer-center/)

## About Open-XML-PowerTools

The Open XML PowerTools provides guidance and example code for programming with Open XML
Documents (DOCX, XLSX, and PPTX).  It is based on, and extends the functionality
of the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK).

It supports scenarios such as:
- Splitting DOCX/PPTX files into multiple files.
- Combining multiple DOCX/PPTX files into a single file.
- Populating content in template DOCX files with data from XML.
- High-fidelity conversion of DOCX to HTML/CSS.
- High-fidelity conversion of HTML/CSS to DOCX.
- Searching and replacing content in DOCX/PPTX using regular expressions.
- Managing tracked-revisions, including detecting tracked revisions, and accepting tracked revisions.
- Updating Charts in DOCX/PPTX files, including updating cached data, as well as the embedded XLSX.
- Comparing two DOCX files, producing a DOCX with revision tracking markup, and enabling retrieving a list of revisions.
- Retrieving metrics from DOCX files, including the hierarchy of styles used, the languages used, and the fonts used.
- Writing XLSX files using far simpler code than directly writing the markup, including a streaming approach that
  enables writing XLSX files with millions of rows.
- Extracting data (along with formatting) from spreadsheets.

```
Copyright (c) Microsoft Corporation 2012-2017
Portions Copyright (c) Eric White Inc 2018-2019
Licensed under the MIT License.
```
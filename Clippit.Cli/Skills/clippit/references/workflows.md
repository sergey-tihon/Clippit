# Clippit workflow recipes

Use these recipes when a user asks for a document transformation. Write outputs to new paths unless the user explicitly asks to overwrite.

## PPTX: split, edit manifest, rebuild

```bash
clippit pptx split source.pptx --output slides --manifest
# edit generated slides/source.manifest.json if needed
clippit pptx build run slides/source.manifest.json --output rebuilt.pptx
clippit pptx verify rebuilt.pptx --format json
```

Use `--slides` for subsets:

```bash
clippit pptx split source.pptx --slides 1,3,6-9 --output subset
```

## PPTX: scaffold a deck manifest

```bash
clippit pptx build init --output deck.json
# edit deck.json to list source decks, slide files, and section markers
clippit pptx build run deck.json --output deck.pptx
clippit pptx verify deck.pptx --format json
```

## DOCX: merge documents

```bash
clippit word build init --output word-build.json
# edit word-build.json to list source documents
clippit word build run word-build.json --output merged.docx
clippit word verify merged.docx --format json
```

## DOCX: compare two versions

Use when the user wants tracked revisions between two files.

```bash
clippit word compare before.docx after.docx --output compared.docx
clippit word verify compared.docx --format json
```

## DOCX: consolidate reviewer files

Use when there is one original and multiple independently edited copies.

```bash
clippit word consolidate original.docx reviewer-a.docx reviewer-b.docx --output consolidated.docx
clippit word verify consolidated.docx --format json
```

## DOCX: template assembly

Use when a `.docx` template contains DocumentAssembler markup and data is XML.

```bash
clippit word assemble template.docx data.xml --output assembled.docx
clippit word verify assembled.docx --format json
```

## DOCX: cleanup

Accept revisions only:

```bash
clippit word accept-revisions draft.docx --output accepted.docx
```

Remove selected non-content markup:

```bash
clippit word simplify-markup document.docx --accept-revisions --remove-comments --output simplified.docx
clippit word verify simplified.docx --format json
```

Check `clippit word simplify-markup --help` before using broad cleanup presets such as `--all`.

## HTML conversions

```bash
clippit word to-html document.docx --output document.html
clippit word from-html article.html --css styles.css --output article.docx
clippit word verify article.docx --format json
```

## XLSX: create from JSON

```bash
clippit excel create workbook.json --output report.xlsx
clippit excel verify report.xlsx --format json
```

## XLSX: export a sheet/range/table to HTML

```bash
clippit excel to-html workbook.xlsx --sheet "Sheet1" --output sheet.html
clippit excel to-html workbook.xlsx --sheet "Sheet1" --range "A1:D20" --output range.html
clippit excel to-html workbook.xlsx --table "SalesTable" --output table.html
```

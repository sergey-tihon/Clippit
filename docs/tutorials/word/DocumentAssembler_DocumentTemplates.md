---
uid: Tutorial.Word.DocumentAssembler.DocumentTemplates
---

# Inserting Documents or Document Templates into Word Files using Document Assembler

## Introduction

`DocumentAssembler` has added the `Document` and `DocumentTemplate` elements for inserting entire Word documents into the output of an assembled document.

This enables some interesting use cases and also allows developers using `DocumentAssembler` to componentise large templates with re-usable blocks.

## Document Element Usage

The `Document` element will take either a File path or a `base64` encoded string that represents a Word document:

```xml
<Document Path="C:\Temp\My Document.docx" />
```

Or

```xml
<Document Data="base64encodedstring" />
```

The `Document` element can only be used at the block level and not within a run.

## DocumentTemplate Element Usage

The `DocumentTemplate` element builds upon the Document element by taking a `Select` attribute which is the XML data you want to pass to the template.

Note that the `Select` element changes the context of the XML node. This can make for easier to read templates.

For example say we had the following XML document:

```
<xml>
  <invoice>
    <line-items>
      <line-item>
       <description>Mens Jeans</description>
       <amount>£40.00</amount>
      </line-item>
      <line-item>
        <description>T-Shirt<description>
        <amount>£15.00</amount>
      <line-item>
    </line-items>
    <total>
      <amount>£55.00</amount>
    </total>
  </invoice>
</xml>
```

The in a single file template we could address that as follows:

```xml
<Repeat Select="invoice/line-items/line-item" />
  <Content Select="description" /> - <Content Select="amount" />
</EndRepeat>
Total: <Content Select="invoice/total/amount" />
```

Using the DocumentTemplate element we could re-write the top-level template as:

```xml
<DocumentTemplate Path="C:\Templates\Invoice-Lines.docx" Select="invoice/line-items" />
Total: <Content Select="invoice/total/amount" />
```

Because we are passing in the context node then our `Invoice-Lines.docx` would then look like so:

```xml
<Repeat Select="line-item">
  <Content Select="description" /> - <Content Select="amount" />
</Repeat>
```

This is a contrived example, but in large templates with re-usable blocks then it can be very useful.

Note that a template that is called by a `DocumentTemplate` element can have another `DocumentTemplate` within it. How your structure and re-use the templates is up to you.

## How Does it Work?

To make this work we have introduced a dependency on `DocumentBuilder` in `DocumentAssembler`. We then perform the operation in two passes:

### 1. Inline all Document and DocumentTemplate Elements

Each `Document` and `DocumentTemplate` element is processed, and the data inlined as a `Document` element with a `Data` attribute.

e.g. `<Document Data="base64encodedstring" />`

### 2. DocumentBuilder combines all Document elements

Once `DocumentAssembler` has finished processing, we build a collection of `Sources` and pass that to `DocumentBuilder` which handles the complex job of merging the output document and all inlined `Document` elements into a single output document.

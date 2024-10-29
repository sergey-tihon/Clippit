---
uid: Tutorial.Word.DocumentAssembler.ImagesSupport
---

# Key highlights from [#31](https://github.com/sergey-tihon/Clippit/pull/31#issuecomment-874335292)

1. Images can be provided either by base64-encoded string or by specifying the filename. The Assembler will have to figure out the image type, based on either MIME or file extension. Here are both examples:

  - `<Logo>../../md-logo.png</Logo>`
  - `<Image>data:image/jpg;base64,/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAgGBgcGBQgHBwcetc…</Image>`

2. The Image content control will be surrounded by new type of select (similar to Repeat or Conditional) – `Image Select`. I found this approach easier than dealing just with the Image content control.

Examples below are taken from the Tests:

## Image Select

![image1](../../images/word/documentassembler/image1.jpg)

## Image Select within a Repeat

Note that when using a Repeat XPATH is aware of the Context it is in and should operate on the "current" Node.

![image2](../../images/word/documentassembler/image2.png)

## Image Select within a Table

Works in a very similar way to Repeat.

![image3](../../images/word/documentassembler/image3.png)

## XML Data used in above Examples

```xml
<?xml version="1.0" encoding="utf-8"?>
<Customer>
  <CustomerID>1</CustomerID>
  <Name>Cheryl</Name>
  <HighValueCustomer>True</HighValueCustomer>
  <CustomerLogo>../../../../TestFiles/img.png</CustomerLogo>
  <Header>../../../../TestFiles/T0936_files/image001.png</Header>
  <Orders>
    <Order>
      <ProductDescription>Unicycle</ProductDescription>
      <Quantity>3</Quantity>
      <OrderDate>9/5/2001</OrderDate>
	  <Thumbnail>../../../../TestFiles/img2.png</Thumbnail>
    </Order>
    <Order>
      <ProductDescription>Tricycle</ProductDescription>
      <Quantity>3</Quantity>
      <OrderDate>8/6/2000</OrderDate>
	  <Thumbnail>data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAMcAAACTCAYA...</Thumbnail>	  
    </Order>
  </Orders>
  <TotalQuantity>6</TotalQuantity>
  <Description><![CDATA[This
is a multiline
description that
contains details about
Cheryl.]]></Description>
</Customer>
```

## Further Reading

If you are interested in using the Image functionality in DocumentAssembler then you best bet is to look at `DocumentAssemblerTests.cs` and particularly the data files which can be found in the repository under `Test Files/DA`.

Changes merged in: [#31](https://github.com/sergey-tihon/Clippit/pull/31)
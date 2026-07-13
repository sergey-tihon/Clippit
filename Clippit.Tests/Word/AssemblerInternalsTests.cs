// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;
using System.Xml.XPath;
using Clippit.Word.Assembler;

namespace Clippit.Tests.Word;

/// <summary>
/// Unit tests for the internal Assembler helpers: XPathExtensions and ErrorHandler.
/// The main library exposes these to this project via InternalsVisibleTo.
/// </summary>
public class AssemblerInternalsTests
{
    // ── XPathExtensions.EvaluateXPath ──────────────────────────────────────

    [Test]
    public async Task EvaluateXPath_BlankXPath_ReturnsEmptyArray()
    {
        var element = new XElement("root", new XElement("child", "value"));
        var result = element.EvaluateXPath("   ", optional: false);
        await Assert.That(result).IsEmpty();
    }

    [Test]
    public async Task EvaluateXPath_NullXPath_ReturnsEmptyArray()
    {
        var element = new XElement("root");
        var result = element.EvaluateXPath(null!, optional: false);
        await Assert.That(result).IsEmpty();
    }

    [Test]
    public async Task EvaluateXPath_ValidElementXPath_ReturnsNodeValues()
    {
        var element = new XElement("root", new XElement("Name", "Alice"));
        var result = element.EvaluateXPath("Name", optional: false);
        await Assert.That(result).IsEquivalentTo(["Alice"]);
    }

    [Test]
    public async Task EvaluateXPath_XPathSelectsAttribute_ReturnsAttributeValue()
    {
        var element = new XElement("root", new XElement("child", new XAttribute("id", "42")));
        var result = element.EvaluateXPath("child/@id", optional: false);
        await Assert.That(result).IsEquivalentTo(["42"]);
    }

    [Test]
    public async Task EvaluateXPath_MultipleMatchingNodes_ReturnsAllValues()
    {
        var element = new XElement(
            "root",
            new XElement("Item", "first"),
            new XElement("Item", "second"),
            new XElement("Item", "third")
        );
        var result = element.EvaluateXPath("Item", optional: false);
        await Assert.That(result).IsEquivalentTo(["first", "second", "third"]);
    }

    [Test]
    public async Task EvaluateXPath_BooleanXPathExpression_ReturnsBooleanAsString()
    {
        var element = new XElement("root");
        // A boolean XPath expression returns a non-IEnumerable result.
        var result = element.EvaluateXPath("count(Item) > 0", optional: true);
        await Assert.That(result).Count().IsEqualTo(1);
        await Assert.That(result[0]).IsEqualTo("False");
    }

    [Test]
    public async Task EvaluateXPath_NumericXPathExpression_ReturnsNumberAsString()
    {
        var element = new XElement("root", new XElement("Item", "a"), new XElement("Item", "b"));
        var result = element.EvaluateXPath("count(Item)", optional: false);
        await Assert.That(result).Count().IsEqualTo(1);
        await Assert.That(result[0]).IsEqualTo("2");
    }

    [Test]
    public async Task EvaluateXPath_NoMatchNonOptional_ThrowsXPathException()
    {
        var element = new XElement("root");
        await Assert.That(() => element.EvaluateXPath("Missing", optional: false)).Throws<XPathException>();
    }

    [Test]
    public async Task EvaluateXPath_NoMatchOptional_ReturnsEmptyArray()
    {
        var element = new XElement("root");
        var result = element.EvaluateXPath("Missing", optional: true);
        await Assert.That(result).IsEmpty();
    }

    [Test]
    public async Task EvaluateXPath_InvalidXPath_ThrowsXPathException()
    {
        var element = new XElement("root");
        await Assert.That(() => element.EvaluateXPath("//[invalid", optional: false)).Throws<XPathException>();
    }

    // ── XPathExtensions.EvaluateXPathToString ─────────────────────────────

    [Test]
    public async Task EvaluateXPathToString_BlankXPathOptional_ReturnsEmptyString()
    {
        var element = new XElement("root", new XElement("Name", "Alice"));
        var result = element.EvaluateXPathToString("   ", optional: true);
        await Assert.That(result).IsEqualTo(string.Empty);
    }

    [Test]
    public async Task EvaluateXPathToString_BlankXPathNonOptional_ThrowsXPathException()
    {
        var element = new XElement("root", new XElement("Name", "Alice"));
        await Assert.That(() => element.EvaluateXPathToString("   ", optional: false)).Throws<XPathException>();
    }

    [Test]
    public async Task EvaluateXPathToString_ValidXPath_ReturnsSingleValue()
    {
        var element = new XElement("root", new XElement("Name", "Bob"));
        var result = element.EvaluateXPathToString("Name", optional: false);
        await Assert.That(result).IsEqualTo("Bob");
    }

    [Test]
    public async Task EvaluateXPathToString_NoMatchNonOptional_ThrowsXPathException()
    {
        var element = new XElement("root");
        await Assert.That(() => element.EvaluateXPathToString("Missing", optional: false)).Throws<XPathException>();
    }

    [Test]
    public async Task EvaluateXPathToString_NoMatchOptional_ReturnsEmptyString()
    {
        var element = new XElement("root");
        var result = element.EvaluateXPathToString("Missing", optional: true);
        await Assert.That(result).IsEqualTo(string.Empty);
    }

    [Test]
    public async Task EvaluateXPathToString_MultipleResults_ThrowsXPathException()
    {
        var element = new XElement("root", new XElement("Item", "first"), new XElement("Item", "second"));
        await Assert.That(() => element.EvaluateXPathToString("Item", optional: false)).Throws<XPathException>();
    }

    // ── ErrorHandler.CreateRunErrorMessage ────────────────────────────────

    [Test]
    public async Task CreateRunErrorMessage_SetsHasErrorAndReturnsRedHighlightedRun()
    {
        var templateError = new TemplateError();
        var run = ErrorHandler.CreateRunErrorMessage("Something went wrong", templateError);

        await Assert.That(templateError.HasError).IsTrue();
        await Assert.That(run.Name).IsEqualTo(W.r);

        var rPr = run.Element(W.rPr);
        await Assert.That(rPr).IsNotNull();
        var colorEl = rPr!.Element(W.color);
        await Assert.That(colorEl).IsNotNull();
        var colorValAttr = colorEl!.Attribute(W.val);
        await Assert.That(colorValAttr).IsNotNull();
        await Assert.That((string)colorValAttr!).IsEqualTo("FF0000");
        var highlightEl = rPr.Element(W.highlight);
        await Assert.That(highlightEl).IsNotNull();
        var highlightValAttr = highlightEl!.Attribute(W.val);
        await Assert.That(highlightValAttr).IsNotNull();
        await Assert.That((string)highlightValAttr!).IsEqualTo("yellow");

        var textEl = run.Element(W.t);
        await Assert.That(textEl).IsNotNull();
        await Assert.That(textEl!.Value).IsEqualTo("Something went wrong");
    }

    [Test]
    public async Task CreateParaErrorMessage_SetsHasErrorAndReturnsParaWithRun()
    {
        var templateError = new TemplateError();
        var para = ErrorHandler.CreateParaErrorMessage("Para error", templateError);

        await Assert.That(templateError.HasError).IsTrue();
        await Assert.That(para.Name).IsEqualTo(W.p);

        var run = para.Element(W.r);
        await Assert.That(run).IsNotNull();
        var textEl = run!.Element(W.t);
        await Assert.That(textEl).IsNotNull();
        await Assert.That(textEl!.Value).IsEqualTo("Para error");
    }

    [Test]
    public async Task CreateContextErrorMessage_WithDescendantParagraph_ReturnsParagraphWrapper()
    {
        var templateError = new TemplateError();
        var context = new XElement("sdt", new XElement(W.p, new XElement(W.r, new XElement(W.t, "existing"))));

        var result = context.CreateContextErrorMessage("Context error", templateError);

        await Assert.That(templateError.HasError).IsTrue();
        // When the context has a descendant w:p, the result is a new w:p containing the error run.
        var resultEl = result as XElement;
        await Assert.That(resultEl).IsNotNull();
        await Assert.That(resultEl!.Name).IsEqualTo(W.p);
    }

    [Test]
    public async Task CreateContextErrorMessage_WithNoDescendantParagraph_ReturnsRunDirectly()
    {
        var templateError = new TemplateError();
        // A context with a run but no paragraph.
        var context = new XElement("sdt", new XElement(W.r, new XElement(W.t, "existing")));

        var result = context.CreateContextErrorMessage("Run-only error", templateError);

        await Assert.That(templateError.HasError).IsTrue();
        var resultEl = result as XElement;
        await Assert.That(resultEl).IsNotNull();
        await Assert.That(resultEl!.Name).IsEqualTo(W.r);
    }
}

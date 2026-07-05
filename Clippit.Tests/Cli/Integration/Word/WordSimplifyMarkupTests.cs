using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Clippit.Word;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Cli.Integration.Word;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
internal sealed class WordSimplifyMarkupTests : CliIntegrationTestBase
{
    [Test]
    public async Task CLI118_WordSimplifyMarkup_NoFlags_ReturnsInvalidArguments()
    {
        var input = CliTestRunner.TestFile("DB007-Spec.docx");
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-no-flags");
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync("word", "simplify-markup", input.FullName, "--output", output.FullName)
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(2);
        await Assert.That(result.StandardOutput).IsEmpty();
        await Assert.That(result.StandardError).Contains("INVALID_ARGUMENTS");
    }

    [Test]
    public async Task CLI119_WordSimplifyMarkup_AllFlag_ProducesValidDocx()
    {
        var input = CliTestRunner.TestFile("DB007-Spec.docx");
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-all");
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--all",
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("input").GetString()).IsEqualTo(input.FullName);
        await Assert.That(json.RootElement.GetProperty("output").GetString()).IsEqualTo(output.FullName);
        await Assert.That(json.RootElement.GetProperty("outputSize").GetInt64()).IsGreaterThan(0);
        await Assert.That(output.Exists).IsTrue();

        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await ValidateRelationships(doc);
        await Validate(doc);
    }

    [Test]
    public async Task CLI120_WordSimplifyMarkup_AcceptRevisions_RemovesTrackedChanges()
    {
        var input = CliTestRunner.TestFile("RA001-Tracked-Revisions-01.docx");
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-accept-revisions");
        var output = new FileInfo(Path.Combine(tempDir.FullName, "clean.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--accept-revisions",
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(output.Exists).IsTrue();

        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await Assert.That(RevisionAccepter.HasTrackedRevisions(doc)).IsFalse();
        await ValidateRelationships(doc);
        // Full schema validation is intentionally omitted here: RA001-Tracked-Revisions-01.docx
        // contains pre-existing schema errors unrelated to revision acceptance.
    }

    [Test]
    public async Task CLI121_WordSimplifyMarkup_RemoveRsidInfo_EliminatesRsidAttributes()
    {
        var input = CliTestRunner.TestFile("DB007-Spec.docx");
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-rsid");
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--remove-rsid-info",
                "--output",
                output.FullName
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(output.Exists).IsTrue();

        // Verify no rsid attributes remain in document.xml
        using var zip = ZipFile.OpenRead(output.FullName);
        var docEntry = zip.GetEntry("word/document.xml");
        await Assert.That(docEntry).IsNotNull();

        using var stream = docEntry!.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        var xmlContent = await reader.ReadToEndAsync().ConfigureAwait(false);
        await Assert.That(xmlContent).DoesNotContain(":rsid");

        // Verify the w:rsids element is also removed from word/settings.xml
        var settingsEntry = zip.GetEntry("word/settings.xml");
        await Assert.That(settingsEntry).IsNotNull();
        using var settingsStream = settingsEntry!.Open();
        using var settingsReader = new StreamReader(settingsStream, Encoding.UTF8);
        var settingsContent = await settingsReader.ReadToEndAsync().ConfigureAwait(false);
        await Assert.That(settingsContent).DoesNotContain("w:rsids");
    }

    [Test]
    public async Task CLI122_WordSimplifyMarkup_DefaultOutputPath_DerivedFromInput()
    {
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-default-output");
        var source = CliTestRunner.TestFile("DB007-Spec.docx");
        var localInput = new FileInfo(Path.Combine(tempDir.FullName, "DB007-Spec.docx"));
        File.Copy(source.FullName, localInput.FullName);

        var expectedOutput = new FileInfo(Path.Combine(tempDir.FullName, "DB007-Spec-simplified.docx"));

        var result = await CliTestRunner
            .RunManagedAsync("word", "simplify-markup", localInput.FullName, "--remove-rsid-info", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("output").GetString()).IsEqualTo(expectedOutput.FullName);
        await Assert.That(expectedOutput.Exists).IsTrue();
    }

    [Test]
    public async Task CLI123_WordSimplifyMarkup_QuietSuppressesStdout()
    {
        var input = CliTestRunner.TestFile("DB007-Spec.docx");
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-quiet");
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--remove-rsid-info",
                "--output",
                output.FullName,
                "--quiet"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardOutput).IsEmpty();
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(output.Exists).IsTrue();
    }

    [Test]
    public async Task CLI124_WordSimplifyMarkup_InvalidInput_ReturnsInvalidFormat()
    {
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-invalid-input");
        var badInput = new FileInfo(Path.Combine(tempDir.FullName, "not-a-docx.docx"));
        await File.WriteAllTextAsync(badInput.FullName, "not a valid DOCX").ConfigureAwait(false);
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                badInput.FullName,
                "--all",
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(4);
        await Assert.That(result.StandardOutput).IsEmpty();
        await Assert.That(result.StandardError).Contains("INVALID_FORMAT");
    }

    [Test]
    public async Task CLI125_WordSimplifyMarkup_OutputToStdout_StreamsBinaryDocx()
    {
        var input = CliTestRunner.TestFile("DB007-Spec.docx");
        var inputBytes = await File.ReadAllBytesAsync(input.FullName).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedWithStdinAsync(inputBytes, "word", "simplify-markup", "-", "--all", "--output", "-")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        // Output should be a valid DOCX, not just a ZIP/PK header
        await Assert.That(result.StandardOutput.Length).IsGreaterThan(0);
        await Assert.That(result.StandardOutput[0]).IsEqualTo((byte)'P');
        await Assert.That(result.StandardOutput[1]).IsEqualTo((byte)'K');

        using var ms = new MemoryStream(result.StandardOutput);
        using var doc = WordprocessingDocument.Open(ms, false);
        await ValidateRelationships(doc);
        await Validate(doc);
    }

    [Test]
    public async Task CLI126_WordSimplifyMarkup_RemoveBookmarks_EliminatesBookmarkElements()
    {
        var input = CliTestRunner.TestFile("DB007-Spec.docx");
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-remove-bookmarks");
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--remove-bookmarks",
                "--output",
                output.FullName
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var zip = ZipFile.OpenRead(output.FullName);
        var docEntry = zip.GetEntry("word/document.xml");
        await Assert.That(docEntry).IsNotNull();
        using var stream = docEntry!.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        var xmlContent = await reader.ReadToEndAsync().ConfigureAwait(false);
        await Assert.That(xmlContent).DoesNotContain("w:bookmarkStart");

        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await ValidateRelationships(doc);
        await Validate(doc);
    }

    [Test]
    public async Task CLI127_WordSimplifyMarkup_RemoveContentControls_EliminatesSdtElements()
    {
        var input = CliTestRunner.TestFile("HC030-Content-Controls.docx");
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-remove-content-controls");
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--remove-content-controls",
                "--output",
                output.FullName
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var zip = ZipFile.OpenRead(output.FullName);
        var docEntry = zip.GetEntry("word/document.xml");
        await Assert.That(docEntry).IsNotNull();
        using var stream = docEntry!.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        var xmlContent = await reader.ReadToEndAsync().ConfigureAwait(false);
        await Assert.That(xmlContent).DoesNotContain("w:sdt");

        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await ValidateRelationships(doc);
        await Validate(doc);
    }

    [Test]
    public async Task CLI128_WordSimplifyMarkup_RemoveHyperlinks_EliminatesHyperlinkElements()
    {
        var input = CliTestRunner.TestFile("HC023-Hyperlink.docx");
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-remove-hyperlinks");
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--remove-hyperlinks",
                "--output",
                output.FullName
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var zip = ZipFile.OpenRead(output.FullName);
        var docEntry = zip.GetEntry("word/document.xml");
        await Assert.That(docEntry).IsNotNull();
        using var stream = docEntry!.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        var xmlContent = await reader.ReadToEndAsync().ConfigureAwait(false);
        await Assert.That(xmlContent).DoesNotContain("w:hyperlink");

        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await ValidateRelationships(doc);
        await Validate(doc);
    }

    [Test]
    public async Task CLI129_WordSimplifyMarkup_ReplaceTabsWithSpaces_RemovesTabCharacterElements()
    {
        var input = CliTestRunner.TestFile("HC024-Tabs-01.docx");
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-replace-tabs");
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--replace-tabs-with-spaces",
                "--output",
                output.FullName
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var zip = ZipFile.OpenRead(output.FullName);
        var docEntry = zip.GetEntry("word/document.xml");
        await Assert.That(docEntry).IsNotNull();
        using var stream = docEntry!.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        var xmlContent = await reader.ReadToEndAsync().ConfigureAwait(false);
        var docXml = XDocument.Parse(xmlContent);
        var runTabElements = docXml.Descendants(W.r).Elements(W.tab).ToList();
        await Assert.That(runTabElements).IsEmpty();

        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await ValidateRelationships(doc);
        await Validate(doc);
    }

    [Test]
    public async Task CLI130_WordSimplifyMarkup_RemoveFieldCodes_EliminatesInstrTextElements()
    {
        var input = CliTestRunner.TestFile("HC040-Hyperlink-Fieldcode-01.docx");
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-remove-field-codes");
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--remove-field-codes",
                "--output",
                output.FullName
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var zip = ZipFile.OpenRead(output.FullName);
        var docEntry = zip.GetEntry("word/document.xml");
        await Assert.That(docEntry).IsNotNull();
        using var stream = docEntry!.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        var xmlContent = await reader.ReadToEndAsync().ConfigureAwait(false);
        await Assert.That(xmlContent).DoesNotContain("w:instrText");

        using var doc = WordprocessingDocument.Open(output.FullName, false);
        await ValidateRelationships(doc);
        await Validate(doc);
    }

    [Test]
    public async Task CLI131_WordSimplifyMarkup_RemoveGoBackBookmark_RemovesOnlyGoBackBookmark()
    {
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-remove-go-back-bookmark");
        var input = await CreateDocxWithMainDocumentXmlAsync(
            tempDir,
            "in.docx",
            """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:bookmarkStart w:id="0" w:name="_GoBack"/>
                  <w:bookmarkEnd w:id="0"/>
                  <w:bookmarkStart w:id="1" w:name="KeepMe"/>
                  <w:r><w:t>Text</w:t></w:r>
                  <w:bookmarkEnd w:id="1"/>
                </w:p>
              </w:body>
            </w:document>
            """
        );
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--remove-go-back-bookmark",
                "--output",
                output.FullName
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        var docXml = await ReadPartXmlAsync(output, "word/document.xml");
        var bookmarkStarts = docXml.Descendants(W.bookmarkStart).Select(x => (string?)x.Attribute(W.name)).ToList();
        await Assert.That(bookmarkStarts).Contains("KeepMe");
        await Assert.That(bookmarkStarts).DoesNotContain("_GoBack");
    }

    [Test]
    public async Task CLI132_WordSimplifyMarkup_RemoveMarkupForDocumentComparison_RemovesRsidAttributes()
    {
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-remove-markup-for-document-comparison");
        var input = await CreateDocxWithMainDocumentXmlAsync(
            tempDir,
            "in.docx",
            """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p w:rsidR="001234AB" w:rsidRDefault="001234AB">
                  <w:r w:rsidRPr="00AB1234">
                    <w:t>Hello</w:t>
                  </w:r>
                </w:p>
              </w:body>
            </w:document>
            """
        );
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--remove-markup-for-document-comparison",
                "--output",
                output.FullName
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        var docXml = await ReadPartXmlAsync(output, "word/document.xml");
        var rsidAttributes = docXml
            .Descendants()
            .SelectMany(e => e.Attributes())
            .Where(a => a.Name.Namespace == W.rsidR.Namespace && a.Name.LocalName.StartsWith("rsid"))
            .ToList();
        await Assert.That(rsidAttributes).IsEmpty();
    }

    [Test]
    public async Task CLI133_WordSimplifyMarkup_RemoveComments_RemovesCommentMarkup()
    {
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-remove-comments");
        var input = await CreateDocxWithMainDocumentXmlAsync(
            tempDir,
            "in.docx",
            """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:commentRangeStart w:id="1"/>
                  <w:r><w:t>Hello</w:t></w:r>
                  <w:commentRangeEnd w:id="1"/>
                  <w:r>
                    <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
                    <w:commentReference w:id="1"/>
                  </w:r>
                </w:p>
              </w:body>
            </w:document>
            """,
            commentsXml: """
            <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:comment w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">
                <w:p><w:r><w:t>Comment</w:t></w:r></w:p>
              </w:comment>
            </w:comments>
            """
        );
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--remove-comments",
                "--output",
                output.FullName
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        var docXml = await ReadPartXmlAsync(output, "word/document.xml");
        await Assert.That(docXml.Descendants(W.commentRangeStart)).IsEmpty();
        await Assert.That(docXml.Descendants(W.commentRangeEnd)).IsEmpty();
        await Assert.That(docXml.Descendants(W.commentReference)).IsEmpty();
    }

    [Test]
    public async Task CLI134_WordSimplifyMarkup_RemoveEndAndFootnotes_RemovesNoteReferences()
    {
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-remove-end-and-footnotes");
        var input = await CreateDocxWithMainDocumentXmlAsync(
            tempDir,
            "in.docx",
            """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:r><w:t xml:space="preserve">Some text</w:t></w:r>
                  <w:r><w:footnoteReference w:id="1"/></w:r>
                  <w:r><w:endnoteReference w:id="1"/></w:r>
                </w:p>
              </w:body>
            </w:document>
            """
        );
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--remove-end-and-footnotes",
                "--output",
                output.FullName
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        var docXml = await ReadPartXmlAsync(output, "word/document.xml");
        await Assert.That(docXml.Descendants(W.footnoteReference)).IsEmpty();
        await Assert.That(docXml.Descendants(W.endnoteReference)).IsEmpty();
    }

    [Test]
    public async Task CLI135_WordSimplifyMarkup_RemoveLastRenderedPageBreak_RemovesElement()
    {
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-remove-last-rendered-page-break");
        var input = await CreateDocxWithMainDocumentXmlAsync(
            tempDir,
            "in.docx",
            """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:r><w:lastRenderedPageBreak/><w:t>Text</w:t></w:r>
                </w:p>
              </w:body>
            </w:document>
            """
        );
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--remove-last-rendered-page-break",
                "--output",
                output.FullName
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        var docXml = await ReadPartXmlAsync(output, "word/document.xml");
        await Assert.That(docXml.Descendants(W.lastRenderedPageBreak)).IsEmpty();
    }

    [Test]
    public async Task CLI136_WordSimplifyMarkup_RemovePermissions_RemovesPermissionMarkup()
    {
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-remove-permissions");
        var input = await CreateDocxWithMainDocumentXmlAsync(
            tempDir,
            "in.docx",
            """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:permStart w:id="1" w:edGrp="everyone"/>
                  <w:r><w:t>Protected text</w:t></w:r>
                  <w:permEnd w:id="1"/>
                </w:p>
              </w:body>
            </w:document>
            """
        );
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--remove-permissions",
                "--output",
                output.FullName
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        var docXml = await ReadPartXmlAsync(output, "word/document.xml");
        await Assert.That(docXml.Descendants(W.permStart)).IsEmpty();
        await Assert.That(docXml.Descendants(W.permEnd)).IsEmpty();
    }

    [Test]
    public async Task CLI137_WordSimplifyMarkup_RemoveProof_RemovesProofErrElements()
    {
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-remove-proof");
        var input = await CreateDocxWithMainDocumentXmlAsync(
            tempDir,
            "in.docx",
            """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:proofErr w:type="spellStart"/>
                  <w:r><w:t>mispelled</w:t></w:r>
                  <w:proofErr w:type="spellEnd"/>
                </w:p>
              </w:body>
            </w:document>
            """
        );
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync("word", "simplify-markup", input.FullName, "--remove-proof", "--output", output.FullName)
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        var docXml = await ReadPartXmlAsync(output, "word/document.xml");
        await Assert.That(docXml.Descendants(W.proofErr)).IsEmpty();
    }

    [Test]
    public async Task CLI138_WordSimplifyMarkup_RemoveSmartTags_RemovesSmartTagWrappers()
    {
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-remove-smart-tags");
        var input = await CreateDocxWithMainDocumentXmlAsync(
            tempDir,
            "in.docx",
            """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="country-region">
                    <w:r><w:t>Algeria</w:t></w:r>
                  </w:smartTag>
                </w:p>
              </w:body>
            </w:document>
            """
        );
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--remove-smart-tags",
                "--output",
                output.FullName
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        var docXml = await ReadPartXmlAsync(output, "word/document.xml");
        await Assert.That(docXml.Descendants(W.smartTag)).IsEmpty();
    }

    [Test]
    public async Task CLI139_WordSimplifyMarkup_RemoveSoftHyphens_RemovesSoftHyphenElements()
    {
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-remove-soft-hyphens");
        var input = await CreateDocxWithMainDocumentXmlAsync(
            tempDir,
            "in.docx",
            """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:r>
                    <w:t xml:space="preserve">extra</w:t>
                    <w:softHyphen/>
                    <w:t>ordinary</w:t>
                  </w:r>
                </w:p>
              </w:body>
            </w:document>
            """
        );
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--remove-soft-hyphens",
                "--output",
                output.FullName
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        var docXml = await ReadPartXmlAsync(output, "word/document.xml");
        await Assert.That(docXml.Descendants(W.softHyphen)).IsEmpty();
    }

    [Test]
    public async Task CLI140_WordSimplifyMarkup_RemoveWebHidden_RemovesWebHiddenElement()
    {
        var tempDir = CliTestRunner.CreateTempDirectory("simplify-remove-web-hidden");
        var input = await CreateDocxWithMainDocumentXmlAsync(
            tempDir,
            "in.docx",
            """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:r>
                    <w:rPr><w:webHidden/></w:rPr>
                    <w:t>Hidden on web</w:t>
                  </w:r>
                </w:p>
              </w:body>
            </w:document>
            """
        );
        var output = new FileInfo(Path.Combine(tempDir.FullName, "out.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "simplify-markup",
                input.FullName,
                "--remove-web-hidden",
                "--output",
                output.FullName
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        var docXml = await ReadPartXmlAsync(output, "word/document.xml");
        await Assert.That(docXml.Descendants(W.webHidden)).IsEmpty();
    }

    private static async Task<FileInfo> CreateDocxWithMainDocumentXmlAsync(
        DirectoryInfo directory,
        string fileName,
        string mainDocumentXml,
        string? commentsXml = null
    )
    {
        var path = Path.Combine(directory.FullName, fileName);

        await using (var file = File.Create(path))
        using (var doc = WordprocessingDocument.Create(file, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.PutXDocument(XDocument.Parse(mainDocumentXml));

            if (commentsXml is not null)
            {
                var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
                commentsPart.PutXDocument(XDocument.Parse(commentsXml));
            }
        }

        return new FileInfo(path);
    }

    private static async Task<XDocument> ReadPartXmlAsync(FileInfo docxFile, string partPath)
    {
        using var zip = ZipFile.OpenRead(docxFile.FullName);
        var partEntry = zip.GetEntry(partPath);
        await Assert.That(partEntry).IsNotNull();

        using var stream = partEntry!.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        var xmlContent = await reader.ReadToEndAsync().ConfigureAwait(false);
        return XDocument.Parse(xmlContent);
    }
}
#pragma warning restore CA1707

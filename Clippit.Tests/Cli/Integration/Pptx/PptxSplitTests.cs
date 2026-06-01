using System.Globalization;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Cli.Integration.Pptx;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
internal sealed class PptxSplitTests : CliIntegrationTestBase
{
    [Test]
    public async Task CLI004_PptxSplit_WritesSingleSlidePresentations()
    {
        var input = CliTestRunner.TestFile("PB001-Input1.pptx");
        var outputDirectory = CliTestRunner.CreateTempDirectory("split");
        var expectedSlides = CountSlides(input);

        var result = await CliTestRunner
            .RunManagedAsync("pptx", "split", input.FullName, "--output", outputDirectory.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("count").GetInt32()).IsEqualTo(expectedSlides);

        var files = outputDirectory.GetFiles("*.pptx").OrderBy(file => file.Name).ToList();
        await Assert.That(files).Count().IsEqualTo(expectedSlides);

        foreach (var file in files)
        {
            using var document = PresentationDocument.Open(file.FullName, false);
            var presentationPart = document.PresentationPart;
            ArgumentNullException.ThrowIfNull(presentationPart);
            var presentation = presentationPart.Presentation;
            ArgumentNullException.ThrowIfNull(presentation);
            var slideIdList = presentation.SlideIdList;
            ArgumentNullException.ThrowIfNull(slideIdList);
            await Assert.That(slideIdList.Count()).IsEqualTo(1);
            await ValidateRelationships(document).ConfigureAwait(false);
        }
    }

    [Test]
    public async Task CLI005_PptxSplitWithSlides_WritesSelectedSlides()
    {
        var input = CliTestRunner.TestFile("PB001-Input1.pptx");
        var outputDirectory = CliTestRunner.CreateTempDirectory("split-selected");

        var result = await CliTestRunner
            .RunManagedAsync(
                "pptx",
                "split",
                input.FullName,
                "--output",
                outputDirectory.FullName,
                "--slides",
                "1, 2-2, 1",
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("count").GetInt32()).IsEqualTo(2);
        var slideIndexes = new List<int>();
        var slides = json.RootElement.GetProperty("slides");
        for (var i = 0; i < slides.GetArrayLength(); i++)
            slideIndexes.Add(slides[i].GetProperty("index").GetInt32());
        await Assert.That(slideIndexes).IsEquivalentTo([1, 2]);

        var files = outputDirectory.GetFiles("*.pptx").OrderBy(file => file.Name).ToList();
        await Assert.That(files).Count().IsEqualTo(2);
        var fileNames = files.Select(file => file.Name).ToList();
        await Assert.That(fileNames).IsEquivalentTo(["PB001-Input1_001.pptx", "PB001-Input1_002.pptx"]);
    }

    [Test]
    public async Task CLI006_PptxSplitWithDescendingRange_ReturnsJsonError()
    {
        var input = CliTestRunner.TestFile("PB001-Input1.pptx");
        var outputDirectory = CliTestRunner.CreateTempDirectory("split-descending-range");

        var result = await CliTestRunner
            .RunManagedAsync(
                "pptx",
                "split",
                input.FullName,
                "--output",
                outputDirectory.FullName,
                "--slides",
                "3-1",
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(2);
        await Assert.That(result.StandardOutput).IsEmpty();

        using var json = result.ReadStderrJson();
        await Assert.That(json.RootElement.GetProperty("code").GetString()).IsEqualTo("INVALID_ARGUMENTS");
    }

    [Test]
    public async Task CLI007_PptxSplitWithOutOfRangeSlide_ReturnsJsonError()
    {
        var input = CliTestRunner.TestFile("PB001-Input1.pptx");
        var outputDirectory = CliTestRunner.CreateTempDirectory("split-out-of-range");
        var outOfRangeSlide = CountSlides(input) + 1;

        var result = await CliTestRunner
            .RunManagedAsync(
                "pptx",
                "split",
                input.FullName,
                "--output",
                outputDirectory.FullName,
                "--slides",
                outOfRangeSlide.ToString(CultureInfo.InvariantCulture),
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(2);
        await Assert.That(result.StandardOutput).IsEmpty();

        using var json = result.ReadStderrJson();
        await Assert.That(json.RootElement.GetProperty("code").GetString()).IsEqualTo("INVALID_ARGUMENTS");
    }

    [Test]
    public async Task CLI010_PptxSplit_FailsOnExistingFile_UnlessForce()
    {
        var input = CliTestRunner.TestFile("PB001-Input1.pptx");
        var outputDirectory = CliTestRunner.CreateTempDirectory("split-collision");

        var first = await CliTestRunner
            .RunManagedAsync(
                "pptx",
                "split",
                input.FullName,
                "--output",
                outputDirectory.FullName,
                "--slides",
                "1",
                "--format",
                "json"
            )
            .ConfigureAwait(false);
        await Assert.That(first.ExitCode).IsEqualTo(0);

        var second = await CliTestRunner
            .RunManagedAsync(
                "pptx",
                "split",
                input.FullName,
                "--output",
                outputDirectory.FullName,
                "--slides",
                "1",
                "--format",
                "json"
            )
            .ConfigureAwait(false);
        await Assert.That(second.ExitCode).IsEqualTo(5);
        await Assert.That(second.StandardOutput).IsEmpty();
        using (var json = second.ReadStderrJson())
            await Assert.That(json.RootElement.GetProperty("code").GetString()).IsEqualTo("OUTPUT_ERROR");

        var forced = await CliTestRunner
            .RunManagedAsync(
                "pptx",
                "split",
                input.FullName,
                "--output",
                outputDirectory.FullName,
                "--slides",
                "1",
                "--force",
                "--format",
                "json"
            )
            .ConfigureAwait(false);
        await Assert.That(forced.ExitCode).IsEqualTo(0);
        await Assert.That(forced.StandardError).IsEmpty();
    }

    [Test]
    public async Task CLI011_PptxSplit_QuietSuppressesStdout()
    {
        var input = CliTestRunner.TestFile("PB001-Input1.pptx");
        var outputDirectory = CliTestRunner.CreateTempDirectory("split-quiet");

        var result = await CliTestRunner
            .RunManagedAsync(
                "pptx",
                "split",
                input.FullName,
                "--output",
                outputDirectory.FullName,
                "--slides",
                "1",
                "--quiet"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardOutput).IsEmpty();
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(outputDirectory.GetFiles("*.pptx")).IsNotEmpty();
    }

    [Test]
    public async Task CLI014_PptxSplit_ReadsFromStdin()
    {
        var input = CliTestRunner.TestFile("PB001-Input1.pptx");
        var outputDirectory = CliTestRunner.CreateTempDirectory("split-stdin");
        var bytes = await File.ReadAllBytesAsync(input.FullName).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedWithStdinAsync(
                bytes,
                "pptx",
                "split",
                "-",
                "--output",
                outputDirectory.FullName,
                "--slides",
                "1",
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        var stdout = System.Text.Encoding.UTF8.GetString(result.StandardOutput);
        using var json = JsonDocument.Parse(stdout);
        await Assert.That(json.RootElement.GetProperty("input").GetString()).IsEqualTo("<stdin>");
        await Assert.That(json.RootElement.GetProperty("count").GetInt32()).IsEqualTo(1);
        await Assert.That(outputDirectory.GetFiles("stdin_*.pptx")).Count().IsEqualTo(1);
    }

    [Test]
    public async Task CLI016_PptxSplit_Manifest_WritesDeterministicPath()
    {
        var input = CliTestRunner.TestFile("PB001-Input1.pptx");
        var outputDirectory = CliTestRunner.CreateTempDirectory("split-manifest");
        var expectedManifest = Path.Combine(outputDirectory.FullName, "PB001-Input1.manifest.json");

        var result = await CliTestRunner
            .RunManagedAsync(
                "pptx",
                "split",
                input.FullName,
                "--output",
                outputDirectory.FullName,
                "--manifest",
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(File.Exists(expectedManifest)).IsTrue();

        using var json = result.ReadStdoutJson();
        await Assert.That(json.RootElement.GetProperty("manifest").GetString()).IsEqualTo(expectedManifest);

        // Manifest must round-trip through `pptx build run` and rebuild a valid deck.
        var rebuilt = Path.Combine(outputDirectory.FullName, "PB001-Input1.merged.pptx");
        var buildResult = await CliTestRunner
            .RunManagedAsync("pptx", "build", "run", expectedManifest, "--format", "json")
            .ConfigureAwait(false);
        await Assert.That(buildResult.ExitCode).IsEqualTo(0);
        await Assert.That(buildResult.StandardError).IsEmpty();
        await Assert.That(File.Exists(rebuilt)).IsTrue();

        var srcCount = CountSlides(input);
        await Assert.That(CountSlides(new FileInfo(rebuilt))).IsEqualTo(srcCount);

        var manifestText = await File.ReadAllTextAsync(expectedManifest).ConfigureAwait(false);
        using var manifestJson = JsonDocument.Parse(manifestText);
        await Assert.That(manifestJson.RootElement.GetProperty("title").GetString()).IsEqualTo("PB001-Input1");
        await Assert
            .That(manifestJson.RootElement.GetProperty("output").GetString())
            .IsEqualTo("PB001-Input1.merged.pptx");
        // No source sections in PB001-Input1.pptx → deck contains only file entries.
        var deck = manifestJson.RootElement.GetProperty("deck");
        await Assert.That(deck.GetArrayLength()).IsEqualTo(srcCount);
        for (var i = 0; i < deck.GetArrayLength(); i++)
        {
            var entry = deck[i].GetString();
            await Assert.That(entry).IsNotNull();
            await Assert.That(entry!.StartsWith('[')).IsFalse();
            await Assert.That(entry).EndsWith(".pptx");
        }
    }

    [Test]
    public async Task CLI017_PptxSplit_Manifest_FailsOnCollision_UnlessForce()
    {
        var input = CliTestRunner.TestFile("PB001-Input1.pptx");
        var outputDirectory = CliTestRunner.CreateTempDirectory("split-manifest-collision");

        var first = await CliTestRunner
            .RunManagedAsync(
                "pptx",
                "split",
                input.FullName,
                "--output",
                outputDirectory.FullName,
                "--slides",
                "1",
                "--manifest",
                "--format",
                "json"
            )
            .ConfigureAwait(false);
        await Assert.That(first.ExitCode).IsEqualTo(0);

        // Delete the slide file but keep the manifest → second run must still fail.
        foreach (var slide in outputDirectory.GetFiles("*.pptx"))
            slide.Delete();

        var second = await CliTestRunner
            .RunManagedAsync(
                "pptx",
                "split",
                input.FullName,
                "--output",
                outputDirectory.FullName,
                "--slides",
                "1",
                "--manifest",
                "--format",
                "json"
            )
            .ConfigureAwait(false);
        await Assert.That(second.ExitCode).IsEqualTo(5);
        using (var json = second.ReadStderrJson())
        {
            await Assert.That(json.RootElement.GetProperty("code").GetString()).IsEqualTo("OUTPUT_ERROR");
            await Assert.That(json.RootElement.GetProperty("error").GetString()).Contains("manifest.json");
        }

        var forced = await CliTestRunner
            .RunManagedAsync(
                "pptx",
                "split",
                input.FullName,
                "--output",
                outputDirectory.FullName,
                "--slides",
                "1",
                "--manifest",
                "--force",
                "--format",
                "json"
            )
            .ConfigureAwait(false);
        await Assert.That(forced.ExitCode).IsEqualTo(0);
    }

    [Test]
    public async Task CLI018_PptxSplit_Manifest_PreservesSourceSections()
    {
        // Build a sectioned deck via `pptx build run`, then split it and verify
        // the generated manifest preserves the section boundaries.
        var directory = CliTestRunner.CreateTempDirectory("split-manifest-sections");
        var source = CliTestRunner.TestFile("PB001-Input1.pptx");
        var sectionedPath = Path.Combine(directory.FullName, "sectioned.pptx");

        var buildManifest = new FileInfo(Path.Combine(directory.FullName, "build.json"));
        var buildJson = JsonSerializer.Serialize(
            new
            {
                title = "Sectioned",
                output = "sectioned.pptx",
                deck = new object[]
                {
                    "[Intro]",
                    new { file = source.FullName, slides = true },
                    "[Details]",
                    new { file = source.FullName, slides = true },
                },
            }
        );
        await File.WriteAllTextAsync(buildManifest.FullName, buildJson).ConfigureAwait(false);

        var buildResult = await CliTestRunner
            .RunManagedAsync("pptx", "build", "run", buildManifest.FullName, "--format", "json")
            .ConfigureAwait(false);
        await Assert.That(buildResult.ExitCode).IsEqualTo(0);
        await Assert.That(File.Exists(sectionedPath)).IsTrue();

        var srcSlideCount = CountSlides(source);
        var totalSlides = srcSlideCount * 2;

        var splitDir = CliTestRunner.CreateTempDirectory("split-manifest-sections-out");
        var splitResult = await CliTestRunner
            .RunManagedAsync(
                "pptx",
                "split",
                sectionedPath,
                "--output",
                splitDir.FullName,
                "--manifest",
                "--format",
                "json"
            )
            .ConfigureAwait(false);
        await Assert.That(splitResult.ExitCode).IsEqualTo(0);

        var manifestPath = Path.Combine(splitDir.FullName, "sectioned.manifest.json");
        await Assert.That(File.Exists(manifestPath)).IsTrue();

        var manifestText = await File.ReadAllTextAsync(manifestPath).ConfigureAwait(false);
        using var manifestJson = JsonDocument.Parse(manifestText);
        var deck = manifestJson.RootElement.GetProperty("deck");

        // Expected layout: [Intro], N slides, [Details], N slides → 2 + 2N entries.
        await Assert.That(deck.GetArrayLength()).IsEqualTo(2 + totalSlides);
        await Assert.That(deck[0].GetString()).IsEqualTo("[Intro]");
        await Assert.That(deck[1 + srcSlideCount].GetString()).IsEqualTo("[Details]");

        // Selecting only the first slide of each section must still emit both section markers.
        var selDir = CliTestRunner.CreateTempDirectory("split-manifest-sections-selected");
        var firstSelectedIndex = 1;
        var secondSelectedIndex = srcSlideCount + 1;
        var slidesExpr = $"{firstSelectedIndex},{secondSelectedIndex}";
        var selResult = await CliTestRunner
            .RunManagedAsync(
                "pptx",
                "split",
                sectionedPath,
                "--output",
                selDir.FullName,
                "--slides",
                slidesExpr,
                "--manifest",
                "--format",
                "json"
            )
            .ConfigureAwait(false);
        await Assert.That(selResult.ExitCode).IsEqualTo(0);

        var selManifestPath = Path.Combine(selDir.FullName, "sectioned.manifest.json");
        var selManifestText = await File.ReadAllTextAsync(selManifestPath).ConfigureAwait(false);
        using var selJson = JsonDocument.Parse(selManifestText);
        var selDeck = selJson.RootElement.GetProperty("deck");
        await Assert.That(selDeck.GetArrayLength()).IsEqualTo(4);
        await Assert.That(selDeck[0].GetString()).IsEqualTo("[Intro]");
        await Assert.That(selDeck[2].GetString()).IsEqualTo("[Details]");

        var rebuiltPath = Path.Combine(splitDir.FullName, "sectioned.merged.pptx");
        var roundTrip = await CliTestRunner
            .RunManagedAsync("pptx", "build", "run", manifestPath, "--format", "json")
            .ConfigureAwait(false);
        await Assert.That(roundTrip.ExitCode).IsEqualTo(0);
        await Assert.That(File.Exists(rebuiltPath)).IsTrue();
        await Assert.That(SectionListUsesNumericSlideIds(new FileInfo(rebuiltPath))).IsTrue();
    }

    [Test]
    public async Task CLI019_PptxSplit_Manifest_TextOutputIncludesManifestLine()
    {
        var input = CliTestRunner.TestFile("PB001-Input1.pptx");
        var outputDirectory = CliTestRunner.CreateTempDirectory("split-manifest-text");

        var result = await CliTestRunner
            .RunManagedAsync(
                "pptx",
                "split",
                input.FullName,
                "--output",
                outputDirectory.FullName,
                "--slides",
                "1",
                "--manifest",
                "--format",
                "text"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();
        await Assert.That(result.StandardOutput).Contains("Manifest:");
        await Assert.That(result.StandardOutput).Contains("PB001-Input1.manifest.json");
    }

    [Test]
    [Arguments("PB001-Input1.pptx")]
    [Arguments("PB001-Input2.pptx")]
    [Arguments("PB001-Input3.pptx")]
    [Arguments("PP006-Videos.pptx")]
    [Arguments("Presentation.pptx")]
    public async Task CLI020_PptxSplit_Then_BuildRun_RoundTrip_PassesValidation(string sourceName)
    {
        // End-to-end round-trip: split → use generated manifest → rebuild →
        // assert (1) same slide count and (2) the rebuilt deck passes OpenXml
        // schema + relationship validation.
        var source = CliTestRunner.TestFile(sourceName);
        var directory = CliTestRunner.CreateTempDirectory($"roundtrip-{Path.GetFileNameWithoutExtension(sourceName)}");
        var expectedSlides = CountSlides(source);

        var splitResult = await CliTestRunner
            .RunManagedAsync(
                "pptx",
                "split",
                source.FullName,
                "--output",
                directory.FullName,
                "--manifest",
                "--format",
                "json"
            )
            .ConfigureAwait(false);
        await Assert.That(splitResult.ExitCode).IsEqualTo(0);
        await Assert.That(splitResult.StandardError).IsEmpty();

        var baseName = Path.GetFileNameWithoutExtension(sourceName);
        var manifestPath = Path.Combine(directory.FullName, $"{baseName}.manifest.json");
        var rebuiltPath = Path.Combine(directory.FullName, $"{baseName}.merged.pptx");
        await Assert.That(File.Exists(manifestPath)).IsTrue();

        var buildResult = await CliTestRunner
            .RunManagedAsync("pptx", "build", "run", manifestPath, "--format", "json")
            .ConfigureAwait(false);
        await Assert.That(buildResult.ExitCode).IsEqualTo(0);
        await Assert.That(buildResult.StandardError).IsEmpty();
        await Assert.That(File.Exists(rebuiltPath)).IsTrue();

        // Slide-count parity.
        await Assert.That(CountSlides(new FileInfo(rebuiltPath))).IsEqualTo(expectedSlides);

        // OpenXml schema + relationship validation on the rebuilt deck.
        using var rebuilt = PresentationDocument.Open(rebuiltPath, false);
        await Validate(rebuilt).ConfigureAwait(false);

        // Flat generated manifests should not create a synthetic empty PowerPoint section.
        await Assert.That(HasSectionList(new FileInfo(rebuiltPath))).IsFalse();
    }

    [Test]
    public async Task CLI020a_PptxVerify_MalformedSectionList_ReturnsSectionDiagnostics()
    {
        var directory = CliTestRunner.CreateTempDirectory("verify-malformed-section-list");
        var source = CliTestRunner.TestFile("PB001-Input1.pptx");
        var manifest = await WriteManifestAsync(directory, "sectioned.pptx", source).ConfigureAwait(false);

        var buildResult = await CliTestRunner
            .RunManagedAsync("pptx", "build", "run", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);
        await Assert.That(buildResult.ExitCode).IsEqualTo(0);

        var output = new FileInfo(Path.Combine(directory.FullName, "sectioned.pptx"));
        CorruptSectionListToUseRelationshipIds(output);

        var verify = await CliTestRunner
            .RunManagedAsync("pptx", "verify", output.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(verify.ExitCode).IsEqualTo(4);
        using var json = verify.ReadStdoutJson();
        var diagnostics = json.RootElement.GetProperty("diagnostics");
        var hasSectionError = false;
        for (var i = 0; i < diagnostics.GetArrayLength(); i++)
        {
            if (diagnostics[i].GetProperty("kind").GetString() == "presentation.section")
                hasSectionError = true;
        }
        await Assert.That(hasSectionError).IsTrue();
    }

    [Test]
    public async Task CLI027_PptxSplit_InvalidPackage_ReturnsInvalidFormat()
    {
        var directory = CliTestRunner.CreateTempDirectory("split-invalid-package");
        var input = new FileInfo(Path.Combine(directory.FullName, "not-a-deck.pptx"));
        await File.WriteAllTextAsync(input.FullName, "not a zip package").ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync("pptx", "split", input.FullName, "--output", directory.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(4);
        await Assert.That(result.StandardOutput).IsEmpty();
        using var json = result.ReadStderrJson();
        await Assert.That(json.RootElement.GetProperty("code").GetString()).IsEqualTo("INVALID_FORMAT");
    }
}
#pragma warning restore CA1707

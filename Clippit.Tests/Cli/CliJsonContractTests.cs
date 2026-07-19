using System.Text.Json;
using Json.Schema;

namespace Clippit.Tests.Cli;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
internal sealed class CliJsonContractTests : TestsBase
{
    [Test]
    public async Task CLI100_PptxBuildRun_JsonResult_MatchesSchema()
    {
        var directory = CliTestRunner.CreateTempDirectory("contract-build-result");
        var source = CliTestRunner.TestFile("PB001-Input1.pptx");
        var manifest = new FileInfo(Path.Combine(directory.FullName, "deck.json"));
        var manifestJson = JsonSerializer.Serialize(
            new
            {
                title = "Contract Build",
                output = "built.pptx",
                deck = new[] { "[Section A]", source.FullName },
            }
        );
        await File.WriteAllTextAsync(manifest.FullName, manifestJson).ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync("pptx", "build", "run", manifest.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var payload = result.ReadStdoutJson();
        ValidateJsonAgainstSchema(payload.RootElement, "build-result.v1.json");
    }

    [Test]
    public async Task CLI101_PptxSplit_JsonResult_MatchesSchema()
    {
        var input = CliTestRunner.TestFile("PB001-Input1.pptx");
        var outputDirectory = CliTestRunner.CreateTempDirectory("contract-split-result");

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

        using var payload = result.ReadStdoutJson();
        ValidateJsonAgainstSchema(payload.RootElement, "split-result.v1.json");
    }

    [Test]
    public async Task CLI102_PptxVerify_JsonResult_MatchesSchema()
    {
        var directory = CliTestRunner.CreateTempDirectory("contract-verify-result");
        var input = new FileInfo(Path.Combine(directory.FullName, "not-a-deck.pptx"));
        await File.WriteAllTextAsync(input.FullName, "not a zip package").ConfigureAwait(false);

        var result = await CliTestRunner
            .RunManagedAsync("pptx", "verify", input.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(4);
        await Assert.That(result.StandardError).IsEmpty();

        using var payload = result.ReadStdoutJson();
        ValidateJsonAgainstSchema(payload.RootElement, "verify-result.v1.json");
    }

    [Test]
    public async Task CLI103_PptxBuildInit_Manifest_MatchesSchema()
    {
        var directory = CliTestRunner.CreateTempDirectory("contract-manifest");
        var outputManifest = new FileInfo(Path.Combine(directory.FullName, "deck.json"));

        var result = await CliTestRunner
            .RunManagedAsync("pptx", "build", "init", "--output", outputManifest.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var payload = JsonDocument.Parse(
            await File.ReadAllTextAsync(outputManifest.FullName).ConfigureAwait(false)
        );
        ValidateJsonAgainstSchema(payload.RootElement, "deck-manifest.v1.json");
    }

    [Test]
    public async Task CLI105_ToHtml_JsonResult_MatchesSchema()
    {
        var input = CliTestRunner.TestFile("Blank-wml.docx");
        var output = new FileInfo(Path.Combine(TempDir, "schema-test.html"));

        var result = await CliTestRunner
            .RunManagedAsync("word", "to-html", input.FullName, "--output", output.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var payload = result.ReadStdoutJson();
        ValidateJsonAgainstSchema(payload.RootElement, "convert-result.v1.json");
    }

    [Test]
    public async Task CLI106_FromHtml_JsonResult_MatchesSchema()
    {
        var input = CliTestRunner.TestFile("T0870.html");
        var output = new FileInfo(Path.Combine(TempDir, "schema-test.docx"));

        var result = await CliTestRunner
            .RunManagedAsync("word", "from-html", input.FullName, "--output", output.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var payload = result.ReadStdoutJson();
        ValidateJsonAgainstSchema(payload.RootElement, "convert-result.v1.json");
    }

    [Test]
    public async Task CLI108_WordCompare_JsonResult_MatchesSchema()
    {
        var directory = CliTestRunner.CreateTempDirectory("contract-compare-result");
        var source = CliTestRunner.TestFile("WC/WC011-Before.docx");
        var revised = CliTestRunner.TestFile("WC/WC011-After.docx");
        var output = new FileInfo(Path.Combine(directory.FullName, "compared.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "compare",
                source.FullName,
                revised.FullName,
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var payload = result.ReadStdoutJson();
        ValidateJsonAgainstSchema(payload.RootElement, "compare-result.v1.json");
    }

    [Test]
    public async Task CLI109_CompareResult_Schema_RejectsInvalidPayloads()
    {
        using var missingRevisions = JsonDocument.Parse(
            """
            {"source":"/tmp/before.docx","revised":"/tmp/after.docx","output":"/tmp/out.docx","outputSize":100,"authorForRevisions":"a","dateTimeForRevisions":"2026-01-01T00:00:00Z","caseInsensitive":false}
            """
        );
        await Assert.That(IsValid(missingRevisions.RootElement, "compare-result.v1.json")).IsFalse();

        using var wrongCaseInsensitiveType = JsonDocument.Parse(
            """
            {"source":"/tmp/before.docx","revised":"/tmp/after.docx","output":"/tmp/out.docx","outputSize":100,"revisions":1,"authorForRevisions":"a","dateTimeForRevisions":"2026-01-01T00:00:00Z","caseInsensitive":"no"}
            """
        );
        await Assert.That(IsValid(wrongCaseInsensitiveType.RootElement, "compare-result.v1.json")).IsFalse();
    }

    [Test]
    public async Task CLI107_ConvertResult_Schema_RejectsInvalidPayloads()
    {
        using var missingInput = JsonDocument.Parse(
            """
            {"output":"/tmp/out.html","outputSize":100}
            """
        );
        await Assert.That(IsValid(missingInput.RootElement, "convert-result.v1.json")).IsFalse();

        using var missingOutput = JsonDocument.Parse(
            """
            {"input":"/tmp/in.docx","outputSize":100}
            """
        );
        await Assert.That(IsValid(missingOutput.RootElement, "convert-result.v1.json")).IsFalse();

        using var missingOutputSize = JsonDocument.Parse(
            """
            {"input":"/tmp/in.docx","output":"/tmp/out.html"}
            """
        );
        await Assert.That(IsValid(missingOutputSize.RootElement, "convert-result.v1.json")).IsFalse();

        using var wrongType = JsonDocument.Parse(
            """
            {"input":"/tmp/in.docx","output":"/tmp/out.html","outputSize":"not-a-number"}
            """
        );
        await Assert.That(IsValid(wrongType.RootElement, "convert-result.v1.json")).IsFalse();
    }

    [Test]
    public async Task CLI110_WordAssemble_JsonResult_MatchesSchema()
    {
        var directory = CliTestRunner.CreateTempDirectory("contract-assemble-result");
        var template = CliTestRunner.TestFile("DA/DA001-TemplateDocument.docx");
        var data = CliTestRunner.TestFile("DA/DA-Data.xml");
        var output = new FileInfo(Path.Combine(directory.FullName, "assembled.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "assemble",
                template.FullName,
                data.FullName,
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var payload = result.ReadStdoutJson();
        ValidateJsonAgainstSchema(payload.RootElement, "assemble-result.v1.json");
    }

    [Test]
    public async Task CLI111_WordConsolidate_JsonResult_MatchesSchema()
    {
        var directory = CliTestRunner.CreateTempDirectory("contract-consolidate-result");
        var original = CliTestRunner.TestFile("WC/WC027-Twenty-Paras-Before.docx");
        var rev1 = CliTestRunner.TestFile("WC/WC027-Twenty-Paras-After-1.docx");
        var output = new FileInfo(Path.Combine(directory.FullName, "consolidated.docx"));

        var result = await CliTestRunner
            .RunManagedAsync(
                "word",
                "consolidate",
                original.FullName,
                rev1.FullName,
                "--output",
                output.FullName,
                "--format",
                "json"
            )
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var payload = result.ReadStdoutJson();
        ValidateJsonAgainstSchema(payload.RootElement, "consolidate-result.v1.json");
    }

    [Test]
    public async Task CLI112_ExcelCreate_JsonResult_MatchesSchema()
    {
        var directory = CliTestRunner.CreateTempDirectory("contract-excel-create-result");
        var input = new FileInfo(Path.Combine(directory.FullName, "workbook.json"));
        await File.WriteAllTextAsync(
                input.FullName,
                """
                {
                  "worksheets": [
                    {
                      "name": "Sheet1",
                      "rows": [
                        { "cells": [{ "value": "Hello" }, { "value": 1, "cellDataType": "Number" }] }
                      ]
                    }
                  ]
                }
                """
            )
            .ConfigureAwait(false);
        var output = new FileInfo(Path.Combine(directory.FullName, "result.xlsx"));

        var result = await CliTestRunner
            .RunManagedAsync("excel", "create", input.FullName, "--output", output.FullName, "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var payload = result.ReadStdoutJson();
        ValidateJsonAgainstSchema(payload.RootElement, "excel-create-result.v1.json");
    }

    [Test]
    public async Task CLI199_InstallSkills_JsonResult_MatchesSchema()
    {
        var directory = CliTestRunner.CreateTempDirectory("contract-install-result");

        var result = await CliTestRunner
            .RunManagedAsync(directory, "install", "--skills=all", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var payload = result.ReadStdoutJson();
        ValidateJsonAgainstSchema(payload.RootElement, "install-result.v1.json");
    }

    [Test]
    public async Task CLI200_InstallSkillsDryRun_JsonPlan_MatchesSchema()
    {
        var directory = CliTestRunner.CreateTempDirectory("contract-install-plan");

        var result = await CliTestRunner
            .RunManagedAsync(directory, "install", "--skills=all", "--dry-run", "--format", "json")
            .ConfigureAwait(false);

        await Assert.That(result.ExitCode).IsEqualTo(0);
        await Assert.That(result.StandardError).IsEmpty();

        using var payload = result.ReadStdoutJson();
        ValidateJsonAgainstSchema(payload.RootElement, "install-plan.v1.json");
    }

    [Test]
    public async Task CLI104_JsonSchema_RejectsInvalidPayloads()
    {
        using var negativeCount = JsonDocument.Parse(
            """
            {"input":"deck.pptx","outputDir":"out","slides":[],"count":-1}
            """
        );
        await Assert.That(IsValid(negativeCount.RootElement, "split-result.v1.json")).IsFalse();

        using var blankManifestTitle = JsonDocument.Parse(
            """
            {"title":"","output":"out.pptx","deck":["part.pptx"]}
            """
        );
        await Assert.That(IsValid(blankManifestTitle.RootElement, "deck-manifest.v1.json")).IsFalse();

        using var missingSlides = JsonDocument.Parse(
            """
            {"output":"out.pptx","totalSlides":1,"entries":[{"file":"part.pptx"}]}
            """
        );
        await Assert.That(IsValid(missingSlides.RootElement, "build-result.v1.json")).IsFalse();
    }

    private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, Lazy<JsonSchema>> s_schemaCache =
        new();

    private static void ValidateJsonAgainstSchema(JsonElement payload, string schemaFileName)
    {
        var result = Evaluate(payload, schemaFileName);
        if (!result.IsValid)
            throw new InvalidOperationException(result.ToString());
    }

    private static bool IsValid(JsonElement payload, string schemaFileName) =>
        Evaluate(payload, schemaFileName).IsValid;

    private static EvaluationResults Evaluate(JsonElement payload, string schemaFileName)
    {
        var schema = s_schemaCache
            .GetOrAdd(
                schemaFileName,
                static fileName => new Lazy<JsonSchema>(
                    () => LoadSchema(fileName),
                    LazyThreadSafetyMode.ExecutionAndPublication
                )
            )
            .Value;
        return schema.Evaluate(payload, new EvaluationOptions { OutputFormat = OutputFormat.Hierarchical });
    }

    private static JsonSchema LoadSchema(string schemaFileName)
    {
        var path = Path.Combine(CliTestRunner.RepositoryRoot.FullName, "docs", "schemas", schemaFileName);
        var text = File.ReadAllText(path);
        return JsonSchema.FromText(text);
    }
}
#pragma warning restore CA1707

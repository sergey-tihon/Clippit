using System.Text.Json.Serialization;
using Clippit.Cli.Commands.Common.Verify;
using Clippit.Cli.Commands.Excel.Create;
using Clippit.Cli.Commands.Install;
using Clippit.Cli.Commands.Pptx.Build;
using Clippit.Cli.Commands.Pptx.Split;
using Clippit.Cli.Commands.Pptx.Verify;
using Clippit.Cli.Commands.Version;
using Clippit.Cli.Commands.Word.Assemble;
using Clippit.Cli.Commands.Word.Build;
using Clippit.Cli.Commands.Word.Compare;
using Clippit.Cli.Commands.Word.Consolidate;
using Clippit.Cli.Infrastructure;

namespace Clippit.Cli;

/// <summary>
/// Source-generation context for all CLI JSON types.
/// Ensures full AOT / NativeAOT compatibility — no reflection at runtime.
///
/// Default context emits compact JSON (one line) for machine consumption.
/// For human-editable artifacts like the scaffolded deck manifest, use
/// <see cref="CliJsonContextIndented"/>.
/// </summary>
[JsonSerializable(typeof(SplitResult))]
[JsonSerializable(typeof(SlideEntry))]
[JsonSerializable(typeof(BuildManifest))]
[JsonSerializable(typeof(DeckEntry))]
[JsonSerializable(typeof(BuildResult))]
[JsonSerializable(typeof(BuildEntryResult))]
[JsonSerializable(typeof(InitResult))]
[JsonSerializable(typeof(VerifyResult))]
[JsonSerializable(typeof(VerifyDiagnostic))]
[JsonSerializable(typeof(VersionResult))]
[JsonSerializable(typeof(ConvertResult))]
[JsonSerializable(typeof(CompareResult))]
[JsonSerializable(typeof(AssembleResult))]
[JsonSerializable(typeof(ConsolidateResult))]
[JsonSerializable(typeof(IReadOnlyList<RevisionInfoResult>))]
[JsonSerializable(typeof(WordBuildManifest))]
[JsonSerializable(typeof(WordEntryItem))]
[JsonSerializable(typeof(WordBuildResult))]
[JsonSerializable(typeof(WordBuildEntryResult))]
[JsonSerializable(typeof(WordBuildInitResult))]
[JsonSerializable(typeof(WorkbookDefinition))]
[JsonSerializable(typeof(CreateResult))]
[JsonSerializable(typeof(InstallResult))]
[JsonSerializable(typeof(InstalledSkillResult))]
[JsonSerializable(typeof(InstallPlanResult))]
[JsonSerializable(typeof(ErrorResult))]
[JsonSourceGenerationOptions(
    WriteIndented = false,
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
)]
internal sealed partial class CliJsonContext : JsonSerializerContext { }

/// <summary>
/// Indented variant used only for human-editable on-disk artifacts (e.g. the
/// manifest file scaffolded by <c>pptx build init</c> or <c>word build init</c>).
/// </summary>
[JsonSerializable(typeof(BuildManifest))]
[JsonSerializable(typeof(WordBuildManifest))]
[JsonSourceGenerationOptions(
    WriteIndented = true,
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
)]
internal sealed partial class CliJsonContextIndented : JsonSerializerContext { }

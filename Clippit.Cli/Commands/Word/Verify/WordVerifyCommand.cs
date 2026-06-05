using System.CommandLine;
using Clippit.Cli.Commands.Common.Verify;
using Clippit.Cli.Infrastructure;
using DocumentFormat.OpenXml;

namespace Clippit.Cli.Commands.Word.Verify;

internal static class WordVerifyCommand
{
    private const FileFormatVersions DefaultOfficeVersion = FileFormatVersions.Microsoft365;
    private static readonly string s_allowedOfficeVersions = string.Join(
        ", ",
        Enum.GetNames<FileFormatVersions>().Where(name => name != nameof(FileFormatVersions.None))
    );

    public static Command Build()
    {
        var inputArg = InputSource.BuildArgument(
            "input",
            "Path to the .docx file to verify. Use '-' to read from stdin."
        );

        var officeVersionOption = new Option<string>("--office-version")
        {
            Description =
                "OpenXml schema version to validate against: "
                + s_allowedOfficeVersions
                + $" (default: {DefaultOfficeVersion}).",
            DefaultValueFactory = _ => DefaultOfficeVersion.ToString(),
        };
        officeVersionOption.Validators.Add(result =>
        {
            var value = result.GetValue(officeVersionOption);
            if (!TryParseOfficeVersion(value, out _))
                result.AddError(
                    $"Invalid value for --office-version: '{value}'. Allowed values are: {s_allowedOfficeVersions}."
                );
        });

        var cmd = new Command(
            "verify",
            "Validate that a .docx file is a correct OpenXml document."
                + "\n\nExamples:"
                + "\n  clippit word verify document.docx"
                + "\n  clippit word verify document.docx --format json"
                + "\n  clippit word verify document.docx --office-version Office2021"
                + "\n  cat document.docx | clippit word verify - --format json"
        );
        cmd.Arguments.Add(inputArg);
        cmd.Options.Add(officeVersionOption);
        var (formatOption, quietOption) = cmd.AddOutputOptions();

        cmd.SetAction(parseResult =>
            CommandRunner.Execute(() =>
                Run(
                    parseResult.GetValue(inputArg)!,
                    ParseOfficeVersion(parseResult.GetValue(officeVersionOption) ?? DefaultOfficeVersion.ToString()),
                    parseResult.GetValue(formatOption),
                    parseResult.GetValue(quietOption)
                )
            )
        );

        return cmd;
    }

    private static int Run(string inputPath, FileFormatVersions officeVersion, OutputFormat format, bool quiet)
    {
        var writer = new OutputWriter(format, quiet);
        var input = InputSource.From(inputPath, "stdin.docx");
        var result = WordVerifyService.Execute(input, officeVersion);
        writer.WriteResult(result, CliJsonContext.Default.VerifyResult, VerifyResult.WriteText);
        return result.Valid ? ExitCodes.Success : ExitCodes.InvalidFormat;
    }

    private static FileFormatVersions ParseOfficeVersion(string value)
    {
        _ = TryParseOfficeVersion(value, out var officeVersion);
        return officeVersion;
    }

    private static bool TryParseOfficeVersion(string? value, out FileFormatVersions officeVersion) =>
        Enum.TryParse(value, ignoreCase: true, out officeVersion) && officeVersion != FileFormatVersions.None;
}

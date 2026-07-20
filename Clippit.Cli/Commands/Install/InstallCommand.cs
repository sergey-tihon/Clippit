using System.CommandLine;
using Clippit.Cli.Infrastructure;

namespace Clippit.Cli.Commands.Install;

internal static class InstallCommand
{
    public static Command Build()
    {
        var skillsOption = new Option<string?>("--skills")
        {
            Description = "Install Clippit agent skills into this workspace. Values: agents (default), claude, all.",
            Arity = ArgumentArity.ZeroOrOne,
        };

        var dryRunOption = new Option<bool>("--dry-run")
        {
            Description = "Print the skill files that would be installed without writing them.",
        };

        var cmd = new Command(
            "install",
            "Install Clippit workspace integrations."
                + "\n\nExamples:"
                + "\n  clippit install --skills"
                + "\n  clippit install --skills=agents"
                + "\n  clippit install --skills=claude"
                + "\n  clippit install --skills=all"
        );
        cmd.Options.Add(skillsOption);
        cmd.Options.Add(dryRunOption);
        var (formatOption, quietOption) = cmd.AddOutputOptions();

        cmd.SetAction(parseResult =>
            CommandRunner.Execute(() =>
            {
                var skillsResult = parseResult.GetResult(skillsOption);
                if (skillsResult is null)
                    throw CliException.InvalidArguments(
                        "Specify what to install, for example: clippit install --skills"
                    );

                var skillsTarget = parseResult.GetValue(skillsOption);
                var writer = new OutputWriter(parseResult.GetValue(formatOption), parseResult.GetValue(quietOption));

                if (parseResult.GetValue(dryRunOption))
                {
                    var plan = InstallSkillsService.Plan(skillsTarget);
                    writer.WriteResult(plan, CliJsonContext.Default.InstallPlanResult, InstallPlanResult.WriteText);
                    return ExitCodes.Success;
                }

                var result = InstallSkillsService.Install(skillsTarget);
                writer.WriteResult(result, CliJsonContext.Default.InstallResult, InstallResult.WriteText);
                return ExitCodes.Success;
            })
        );

        return cmd;
    }
}

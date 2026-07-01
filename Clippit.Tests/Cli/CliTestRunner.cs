using System.Diagnostics;
using System.Text.Json;

namespace Clippit.Tests.Cli;

internal sealed record CliResult(int ExitCode, string StandardOutput, string StandardError)
{
    public JsonDocument ReadStdoutJson() => JsonDocument.Parse(StandardOutput);

    public JsonDocument ReadStderrJson() => JsonDocument.Parse(StandardError);
}

internal sealed record CliBinaryResult(int ExitCode, byte[] StandardOutput, string StandardError);

internal static class CliTestRunner
{
    private static readonly TimeSpan s_defaultTimeout = TimeSpan.FromMinutes(2);

    public static DirectoryInfo RepositoryRoot { get; } = FindRepositoryRoot();

    public static FileInfo TestFile(string relativePath) =>
        new(Path.Combine(RepositoryRoot.FullName, "TestFiles", relativePath));

    public static DirectoryInfo CreateTempDirectory(string name)
    {
        // Use a dedicated subtree under the OS temp dir rather than the repo
        // `temp/`. TestsBase wipes `temp/` lazily on first access from any
        // other test class, which races CLI tests when the full suite runs in
        // parallel.
        var directory = new DirectoryInfo(
            Path.Combine(Path.GetTempPath(), "Clippit.Cli.Tests", $"{name}-{Guid.NewGuid():N}")
        );
        directory.Create();
        return directory;
    }

    public static Task<CliResult> RunManagedAsync(params string[] arguments) =>
        RunAsync("dotnet", [GetManagedCliDll().FullName, .. arguments], RepositoryRoot.FullName, s_defaultTimeout);

    /// <summary>
    /// Runs the managed CLI with content piped into stdin. Returned stdout is
    /// captured as raw bytes so callers can verify binary <c>--output -</c>
    /// pipelines (e.g. <c>pptx build run - --output -</c>).
    /// </summary>
    public static async Task<CliBinaryResult> RunManagedWithStdinAsync(byte[] stdin, params string[] arguments)
    {
        var argList = new List<string> { GetManagedCliDll().FullName };
        argList.AddRange(arguments);

        using var process = new Process();
        process.StartInfo.FileName = "dotnet";
        process.StartInfo.WorkingDirectory = RepositoryRoot.FullName;
        process.StartInfo.RedirectStandardInput = true;
        process.StartInfo.RedirectStandardOutput = true;
        process.StartInfo.RedirectStandardError = true;
        process.StartInfo.UseShellExecute = false;
        foreach (var argument in argList)
            process.StartInfo.ArgumentList.Add(argument);

        process.Start();

        var stdoutTask = Task.Run(async () =>
        {
            using var memory = new MemoryStream();
            await process.StandardOutput.BaseStream.CopyToAsync(memory).ConfigureAwait(false);
            return memory.ToArray();
        });
        var stderrTask = process.StandardError.ReadToEndAsync();

        try
        {
            await process.StandardInput.BaseStream.WriteAsync(stdin).ConfigureAwait(false);
            await process.StandardInput.BaseStream.FlushAsync().ConfigureAwait(false);
        }
        catch (IOException)
        {
            // The process may exit before consuming all stdin (e.g. validation
            // rejects the inputs). Ignore the broken-pipe error and let the
            // caller inspect the exit code and stderr.
        }

        process.StandardInput.Close();

        using var timeoutTokenSource = new CancellationTokenSource(s_defaultTimeout);
        await process.WaitForExitAsync(timeoutTokenSource.Token).ConfigureAwait(false);

        var stdoutBytes = await stdoutTask.ConfigureAwait(false);
        var stderr = await stderrTask.ConfigureAwait(false);
        return new CliBinaryResult(process.ExitCode, stdoutBytes, stderr);
    }

    public static async Task<CliResult> RunAsync(
        string executable,
        IReadOnlyList<string> arguments,
        string workingDirectory,
        TimeSpan timeout
    )
    {
        using var process = new Process();
        process.StartInfo.FileName = executable;
        process.StartInfo.WorkingDirectory = workingDirectory;
        process.StartInfo.RedirectStandardOutput = true;
        process.StartInfo.RedirectStandardError = true;
        process.StartInfo.UseShellExecute = false;

        foreach (var argument in arguments)
            process.StartInfo.ArgumentList.Add(argument);

        process.Start();

        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();

        using var timeoutTokenSource = new CancellationTokenSource(timeout);
        try
        {
            await process.WaitForExitAsync(timeoutTokenSource.Token).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException($"Command timed out: {executable} {string.Join(' ', arguments)}");
        }

        var stdout = await stdoutTask.ConfigureAwait(false);
        var stderr = await stderrTask.ConfigureAwait(false);
        return new CliResult(process.ExitCode, stdout, stderr);
    }

    private static FileInfo GetManagedCliDll()
    {
        var testOutputDirectory = new DirectoryInfo(AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar));
        var targetFramework = testOutputDirectory.Name;
        var configuration = testOutputDirectory.Parent?.Name ?? "Debug";
        var candidates = new[] { configuration, "Release", "Debug" }
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Select(config => new FileInfo(
                Path.Combine(RepositoryRoot.FullName, "Clippit.Cli", "bin", config, targetFramework, "Clippit.Cli.dll")
            ))
            .ToList();

        var cliDll = candidates.FirstOrDefault(file => file.Exists);
        if (cliDll is not null)
            return cliDll;

        throw new FileNotFoundException(
            "The managed CLI assembly was not found. Checked: "
                + string.Join(", ", candidates.Select(file => file.FullName))
        );
    }

    private static DirectoryInfo FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Clippit.slnx")))
                return directory;

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Could not find repository root containing Clippit.slnx.");
    }
}

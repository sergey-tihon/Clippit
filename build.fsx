#r "nuget: Fun.Build, 1.1.17"
#r "nuget: Fake.DotNet.AssemblyInfoFile"

open Fun.Build
open Fake.IO
open Fake.DotNet

type CliRuntime =
    { Rid: string
      NpmDirectory: string
      BinaryName: string }

let cliRuntimes =
    [ { Rid = "win-x64"
        NpmDirectory = "clippit-win32-x64"
        BinaryName = "clippit.exe" }
      { Rid = "osx-x64"
        NpmDirectory = "clippit-darwin-x64"
        BinaryName = "clippit" }
      { Rid = "osx-arm64"
        NpmDirectory = "clippit-darwin-arm64"
        BinaryName = "clippit" }
      { Rid = "linux-x64"
        NpmDirectory = "clippit-linux-x64"
        BinaryName = "clippit" }
      { Rid = "linux-arm64"
        NpmDirectory = "clippit-linux-arm64"
        BinaryName = "clippit" }
      { Rid = "win-arm64"
        NpmDirectory = "clippit-win32-arm64"
        BinaryName = "clippit.exe" } ]

let requestedRids =
    let value = System.Environment.GetEnvironmentVariable("CLIPPIT_PUBLISH_RIDS")

    if System.String.IsNullOrWhiteSpace(value) then
        cliRuntimes
    else
        value.Split(
            ',',
            System.StringSplitOptions.RemoveEmptyEntries
            ||| System.StringSplitOptions.TrimEntries
        )
        |> Array.map (fun rid ->
            cliRuntimes
            |> List.tryFind (fun r -> r.Rid = rid)
            |> Option.defaultWith (fun () -> failwith $"Unsupported CLIPPIT_PUBLISH_RIDS value: {rid}"))
        |> Array.toList

let version =
    Changelog.GetLastVersion __SOURCE_DIRECTORY__
    |> Option.defaultWith (fun () -> failwith "Version is not found")

let cliVersion =
    Changelog.GetLastVersion(System.IO.Path.Combine(__SOURCE_DIRECTORY__, "Clippit.Cli"))
    |> Option.defaultWith (fun () -> failwith "CLI version not found in Clippit.Cli/CHANGELOG.md")

let allCliRuntimesRequested =
    let requested = requestedRids |> List.map _.Rid |> Set.ofList
    let supported = cliRuntimes |> List.map _.Rid |> Set.ofList
    requested = supported

let isGitHubActions =
    System.String.Equals(
        System.Environment.GetEnvironmentVariable("GITHUB_ACTIONS"),
        "true",
        System.StringComparison.OrdinalIgnoreCase
    )

let withCiMsBuildLogger (command: string) =
    if isGitHubActions then
        $"{command} -clp:ErrorsOnly"
    else
        command

let testCommand =
    if isGitHubActions then
        // Microsoft.Testing.Platform treats MSBuild logger options as unknown
        // test-app options. Reuse the Release build stage instead of rebuilding.
        "dotnet test --solution Clippit.slnx -c Release --no-build"
    else
        "dotnet test --solution Clippit.slnx"

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/// Copy a native binary from the publish output into the matching npm platform
/// package directory, making the file executable on non-Windows platforms.
let copyNativeBinary (runtime: CliRuntime) =
    let rid = runtime.Rid
    let npmPkgDir = runtime.NpmDirectory
    let binName = runtime.BinaryName
    let src = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "bin", "cli", rid, binName)
    let destDir = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "npm", npmPkgDir)
    let dest = System.IO.Path.Combine(destDir, binName)

    if not (System.IO.File.Exists(src)) then
        failwith $"Native binary not found for {rid}: {src}"

    Shell.copyFile dest src
    // Make executable on Unix
    if
        binName <> "clippit.exe"
        && not (
            System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(
                System.Runtime.InteropServices.OSPlatform.Windows
            )
        )
    then
        let chmod = System.Diagnostics.Process.Start("chmod", $"+x \"{dest}\"")
        chmod.WaitForExit()

        if chmod.ExitCode <> 0 then
            failwith $"chmod failed ({chmod.ExitCode}) for {dest}"

/// Patch package.json version fields for npm pack. Platform packages only need
/// their top-level version; the wrapper also has clippit-* optional dependencies
/// that must match the same published version.
let patchNpmVersion (pkgJsonPath: string) (v: string) =
    let content = System.IO.File.ReadAllText(pkgJsonPath)
    // Pass 1 — top-level version field
    let pass1 =
        System.Text.RegularExpressions.Regex.Replace(content, "\"version\":\\s*\"[^\"]+\"", $"\"version\": \"{v}\"")
    // Pass 2 — every "pkg-name": "x.y.z" pair anywhere in the file where the
    // value looks like a semver string. Scoped to optionalDependencies by the
    // fact that other string values in the file are not bare semver strings.
    // Pattern: "some-package-name": "digits..."
    let pass2 =
        System.Text.RegularExpressions.Regex.Replace(
            pass1,
            "(\"(?:@[^\"/]+/)?clippit(?:-bin)?-[^\"]+\"\\s*:\\s*)\"[^\"]+\"",
            $"$1\"{v}\""
        )

    System.IO.File.WriteAllText(pkgJsonPath, pass2)

let withOriginalFiles (paths: string list) (action: unit -> unit) =
    let originals =
        paths |> List.map (fun path -> path, System.IO.File.ReadAllText(path))

    try
        action ()
    finally
        for path, content in originals do
            System.IO.File.WriteAllText(path, content)

let runRequiredCommand (workingDir: string) (fileName: string) (arguments: string) =
    let proc =
        System.Diagnostics.Process.Start(
            System.Diagnostics.ProcessStartInfo(
                fileName,
                arguments,
                WorkingDirectory = workingDir,
                UseShellExecute = false
            )
        )

    proc.WaitForExit()

    if proc.ExitCode <> 0 then
        failwith $"Command failed ({proc.ExitCode}): {fileName} {arguments}"

let publishCliBinaries () =
    // NativeAOT is target-toolchain dependent. Set CLIPPIT_PUBLISH_RIDS
    // to a comma-separated subset when building locally on a host that cannot
    // produce every RID. CI builds each RID on its native runner.
    for runtime in requestedRids do
        runRequiredCommand
            __SOURCE_DIRECTORY__
            "dotnet"
            (withCiMsBuildLogger
                $"publish Clippit.Cli/Clippit.Cli.csproj -c Release -r {runtime.Rid} --self-contained -p:NativeAot=true -o bin/cli/{runtime.Rid}")

let packNpmPackages () =
    if not allCliRuntimesRequested then
        failwith
            "NPM wrapper package requires all CLI runtimes. Set CLIPPIT_PUBLISH_RIDS=win-x64,win-arm64,osx-x64,osx-arm64,linux-x64,linux-arm64."

    let npmDir = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "npm")

    let allPackages = [ "clippit" ] @ (cliRuntimes |> List.map _.NpmDirectory)

    let packageJsonFiles =
        allPackages
        |> List.map (fun pkg -> System.IO.Path.Combine(npmDir, pkg, "package.json"))

    withOriginalFiles packageJsonFiles (fun () ->
        try
            // Copy binaries into platform npm packages.
            for runtime in requestedRids do
                copyNativeBinary runtime

            // Patch version in all package.json files only for npm pack.
            for pkgJson in packageJsonFiles do
                patchNpmVersion pkgJson cliVersion.Version

            // Pack platform packages, then the wrapper.
            Shell.mkdir (System.IO.Path.Combine(__SOURCE_DIRECTORY__, "bin", "npm"))

            for pkg in (requestedRids |> List.map _.NpmDirectory) @ [ "clippit" ] do
                let pkgDir = System.IO.Path.Combine(npmDir, pkg)
                let outDir = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "bin", "npm")
                runRequiredCommand pkgDir "npm" $"pack --pack-destination {outDir}"
        finally
            for runtime in requestedRids do
                let generated =
                    System.IO.Path.Combine(npmDir, runtime.NpmDirectory, runtime.BinaryName)

                if System.IO.File.Exists(generated) then
                    System.IO.File.Delete(generated))

// ---------------------------------------------------------------------------
// Build pipeline
// ---------------------------------------------------------------------------

pipeline "build" {
    workingDir __SOURCE_DIRECTORY__

    runBeforeEachStage (fun ctx ->
        if ctx.GetStageLevel() = 0 then
            printfn $"::group::{ctx.Name}")

    runAfterEachStage (fun ctx ->
        if ctx.GetStageLevel() = 0 then
            printfn "::endgroup::")

    stage "Check environment" {
        run "dotnet tool restore"
        run "dotnet restore"
    }

    stage "Check Formatting" { run "dotnet csharpier check ." }

    stage "Clean" {
        run (fun _ ->
            Shell.mkdir "bin"
            Shell.cleanDir "bin")

        run "dotnet clean"
    }

    stage "AssemblyInfo" {
        run (fun _ ->
            let fileName = "Clippit/Properties/AssemblyInfo.g.cs"

            AssemblyInfoFile.createCSharp
                fileName
                [ AssemblyInfo.Title "Clippit"
                  AssemblyInfo.Product "Clippit"
                  AssemblyInfo.Description "Fresh PowerTools for OpenXml"
                  AssemblyInfo.Version version.Version
                  AssemblyInfo.FileVersion version.Version ]

            Shell.copyFile "docs/api/CHANGELOG.md" "CHANGELOG.md")
    }

    stage "Build" { run (withCiMsBuildLogger "dotnet build Clippit.slnx -c Release") }

    stage "RunTests" { run testCommand }

    stage "NuGet" {
        run (fun ctx ->
            let releaseNotes = version.ReleaseNotes.Trim()
            let targetsPath = "Clippit/Directory.Build.targets"

            System.IO.File.WriteAllText(
                targetsPath,
                $"""<Project>
  <PropertyGroup>
    <PackageReleaseNotes><![CDATA[{releaseNotes}]]></PackageReleaseNotes>
  </PropertyGroup>
</Project>"""
            )

            try
                ctx.RunCommand(
                    withCiMsBuildLogger
                        $"dotnet pack Clippit/Clippit.csproj -o bin/ -p:PackageVersion={version.Version}"
                )
            finally
                System.IO.File.Delete(targetsPath))
    }

    stage "PackCliTool" {
        run (fun ctx ->
            ctx.RunCommand(
                withCiMsBuildLogger
                    $"dotnet pack Clippit.Cli/Clippit.Cli.csproj -o bin/ -p:PackageVersion={cliVersion.Version}"
            ))
    }

    runIfOnlySpecified
}

// ---------------------------------------------------------------------------
// Publish pipeline — produces NativeAOT self-contained binaries + npm packages.
// For CI release packaging from downloaded native artifacts, use the pack-npm
// pipeline instead.
// ---------------------------------------------------------------------------

pipeline "publish" {
    workingDir __SOURCE_DIRECTORY__

    runBeforeEachStage (fun ctx ->
        if ctx.GetStageLevel() = 0 then
            printfn $"::group::{ctx.Name}")

    runAfterEachStage (fun ctx ->
        if ctx.GetStageLevel() = 0 then
            printfn "::endgroup::")

    stage "PublishCli" { run (fun _ -> publishCliBinaries ()) }

    stage "PackNpm" { run (fun _ -> packNpmPackages ()) }

    runIfOnlySpecified
}

// ---------------------------------------------------------------------------
// npm packaging pipeline — packs npm packages from binaries already present at
// bin/cli/<rid>/<binary>. Used by GitHub Actions after downloading artifacts
// produced on native OS runners.
// ---------------------------------------------------------------------------

pipeline "pack-npm" {
    workingDir __SOURCE_DIRECTORY__

    runBeforeEachStage (fun ctx ->
        if ctx.GetStageLevel() = 0 then
            printfn $"::group::{ctx.Name}")

    runAfterEachStage (fun ctx ->
        if ctx.GetStageLevel() = 0 then
            printfn "::endgroup::")

    stage "PackNpm" { run (fun _ -> packNpmPackages ()) }

    runIfOnlySpecified
}

tryPrintPipelineCommandHelp ()

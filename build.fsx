#r "nuget: Fun.Build, 1.1.14"
#r "nuget: Fake.DotNet.AssemblyInfoFile"

open Fun.Build
open Fake.IO
open Fake.DotNet

let version =
    Changelog.GetLastVersion(__SOURCE_DIRECTORY__)
    |> Option.defaultWith (fun () -> failwith "Version is not found")

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
                  AssemblyInfo.FileVersion version.Version ])
    }

    stage "Build" { run "dotnet build Clippit.sln -c Release" }

    stage "RunTests" { run "dotnet test Clippit.Tests/" }

    stage "NuGet" {
        run
            $"dotnet pack Clippit/Clippit.csproj -o bin/ -p:PackageVersion={version.Version} -p:PackageReleaseNotes=\"{version.ReleaseNotes}\""
    }

    runIfOnlySpecified
}

tryPrintPipelineCommandHelp ()

#r @"paket:
source https://nuget.org/api/v2
framework netstandard2.0
nuget Fake.Core.Target
nuget Fake.Core.ReleaseNotes 
nuget Fake.DotNet.Paket
nuget Fake.DotNet.AssemblyInfoFile
nuget Fake.DotNet.Cli //"

#if !FAKE
#load "./.fake/build.fsx/intellisense.fsx"
#r "netstandard" // Temp fix for https://github.com/fsharp/FAKE/issues/1985
#endif


// --------------------------------------------------------------------------------------
// FAKE build script
// --------------------------------------------------------------------------------------

open Fake
open Fake.Core
open Fake.Core.TargetOperators
open Fake.DotNet
open Fake.IO
open Fake.IO.Globbing.Operators

let gitName = "Clippit"
let description = "Unofficial fork of OpenXmlPowerTools with new features"
let release = ReleaseNotes.load "docs/RELEASE_NOTES.md"

// Targets
Target.create "Clean" (fun _ ->
    Shell.mkdir "bin"
    Shell.cleanDir "bin"
)

Target.create "AssemblyInfo" (fun _ ->
    let fileName = "OpenXmlPowerTools/Properties/AssemblyInfo.Generated.cs"
    AssemblyInfoFile.createCSharp fileName
      [ AssemblyInfo.Title gitName
        AssemblyInfo.Product gitName
        AssemblyInfo.Description description
        AssemblyInfo.Version release.AssemblyVersion
        AssemblyInfo.FileVersion release.AssemblyVersion ]
)

Target.create "Build" (fun _ ->
    DotNet.exec id "build" "Clippit.sln -c Release" |> ignore
)

Target.create "RunTests" (fun _ ->
    DotNet.test id "OpenXmlPowerTools.Tests/"
)

Target.create "NuGet" (fun _ ->
    Paket.pack(fun p ->
        { p with
            ToolType = ToolType.CreateLocalTool()
            OutputPath = "bin"
            Version = release.NugetVersion
            ReleaseNotes = String.toLines release.Notes})
)

Target.create "All" ignore

// Build order
"Clean"
  ==> "AssemblyInfo"
  ==> "Build"
  ==> "RunTests"
  ==> "NuGet"
  ==> "All"

// start build
Target.runOrDefault "All"
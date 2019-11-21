#r @"paket:
source https://nuget.org/api/v2
framework netstandard2.0
nuget Fake.Core.Target
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

// Targets
Target.create "Build" (fun _ ->
    DotNet.exec id "build" "Clippit.sln -c Release" |> ignore
)

Target.create "RunTests" (fun _ ->
    DotNet.test id "OpenXmlPowerTools.Tests/"
)

Target.create "All" ignore

// Build order
"Build"
  ==> "RunTests"
  ==> "All"

// start build
Target.runOrDefault "All"
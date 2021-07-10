# Clippit [![NuGet Badge](https://buildstats.info/nuget/Clippit)](https://www.nuget.org/packages/Clippit)

[![Build Status](https://github.com/sergey-tihon/Clippit/workflows/Build%20and%20Test/badge.svg?branch=master)](https://github.com/sergey-tihon/Clippit/actions?query=branch%3Amaster)

## Build Instructions

### Prerequisites

- .NET CLI toolchain
- libgdiplus
  - macOS: `brew install mono-libgdiplus`
  - Ubuntu: `sudo apt-get update -y && sudo apt-get install -y libgdiplus`

### Build

Call `.\build.cmd` on Windows or `./build.sh` on other systems.

### Update docs

- Install DocFx
  - Windows : `choco install docfx -y`
  - MacOS: `brew install docfx`

- Run `docfx docs/docfx.json --serve` to start local copy of site/docs.
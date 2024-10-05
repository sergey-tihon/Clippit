# Clippit

![NuGet Version](https://img.shields.io/nuget/v/Clippit) ![NuGet Downloads](https://img.shields.io/nuget/dt/Clippit)

[![Build and Test](https://github.com/sergey-tihon/Clippit/actions/workflows/main.yml/badge.svg)](https://github.com/sergey-tihon/Clippit/actions/workflows/main.yml)

## Build Instructions

### Build

Call `.\build.cmd` on Windows or `./build.sh` on other systems.

### Update docs

- Install DocFx

  - Windows : `choco install docfx -y`
  - MacOS: `brew install docfx`

- Run `docfx docs/docfx.json --serve` to start local copy of site/docs.

# Clippit [![NuGet Badge](https://buildstats.info/nuget/Clippit)](https://www.nuget.org/packages/Clippit)

[![Build and Test](https://github.com/sergey-tihon/Clippit/actions/workflows/main.yml/badge.svg)](https://github.com/sergey-tihon/Clippit/actions/workflows/main.yml)

## Build Instructions

[![Open in Gitpod](https://gitpod.io/button/open-in-gitpod.svg)](https://gitpod.io/#https://github.com/sergey-tihon/Clippit)

### Build

Call `.\build.cmd` on Windows or `./build.sh` on other systems.

### Update docs

- Install DocFx
  - Windows : `choco install docfx -y`
  - MacOS: `brew install docfx`

- Run `docfx docs/docfx.json --serve` to start local copy of site/docs.
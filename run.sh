#!/usr/bin/env bash
# run.sh — convenience wrapper to invoke the Clippit CLI without re-typing
# `dotnet run --project Clippit.Cli -- ...` every time.
#
# Usage:
#   ./run.sh pptx split deck.pptx --slides 1,3
#   ./run.sh pptx build init --output deck.json
#   ./run.sh pptx build run deck.json
#   ./run.sh --version
#
# Environment variables:
#   CLIPPIT_CONFIG   MSBuild configuration to use (default: Release)
#   CLIPPIT_FRAMEWORK Target framework to run (default: net11.0; also net10.0)
#   CLIPPIT_NO_BUILD If set to 1, skip the implicit incremental build
#                    (useful in tight pipe loops once the binary is current).
set -euo pipefail

script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
config="${CLIPPIT_CONFIG:-Release}"
framework="${CLIPPIT_FRAMEWORK:-net11.0}"
project="${script_dir}/Clippit.Cli/Clippit.Cli.csproj"

dotnet_args=(run --project "${project}" --configuration "${config}" --framework "${framework}")
if [[ "${CLIPPIT_NO_BUILD:-0}" == "1" ]]; then
    dotnet_args+=(--no-build)
fi
dotnet_args+=(--)

exec dotnet "${dotnet_args[@]}" "$@"

#!/usr/bin/env bash
# Build the AEDI-S .NET MAUI client. Targets depend on installed workloads.
#
#   ./build.sh                  # build for all installed targets (Debug)
#   ./build.sh release android  # Release build, Android only
#
set -euo pipefail
here="$(cd "$(dirname "$0")" && pwd)"
cd "$here"

cfg="Debug"
tfm=""
if [[ "${1:-}" == "release" ]]; then cfg="Release"; shift; fi
if [[ -n "${1:-}" ]]; then
  case "$1" in
    android) tfm="net10.0-android" ;;
    ios)     tfm="net10.0-ios" ;;
    mac|maccatalyst) tfm="net10.0-maccatalyst" ;;
    windows) tfm="net10.0-windows10.0.19041.0" ;;
    *) echo "unknown target: $1" >&2; exit 2 ;;
  esac
fi

args=(build -c "$cfg")
[[ -n "$tfm" ]] && args+=(-f "$tfm")
echo ">>> dotnet ${args[*]}"
exec dotnet "${args[@]}"

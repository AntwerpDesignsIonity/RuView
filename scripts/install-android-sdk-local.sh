#!/usr/bin/env bash
# One-shot install: JDK 17 aarch64 + Android cmdline-tools + SDK packages.
# Writes only under $HOME. Safe to re-run.
set -euo pipefail

JDK_DIR="$HOME/.jdks/temurin-17"
SDK_DIR="$HOME/Android/Sdk"
CMDLINE_URL="https://dl.google.com/android/repository/commandlinetools-linux-11076708_latest.zip"
JDK_URL="https://api.adoptium.net/v3/binary/latest/17/ga/linux/aarch64/jdk/hotspot/normal/eclipse"

log() { printf '\e[36m>>>\e[0m %s\n' "$*"; }

# ── JDK 17 ─────────────────────────────────────────────────────────────────
if [[ ! -x "$JDK_DIR/bin/java" ]]; then
    log "Download Temurin JDK 17 aarch64"
    mkdir -p "$(dirname "$JDK_DIR")"
    tmp="$(mktemp -d)"
    curl -fL --progress-bar -o "$tmp/jdk.tgz" "$JDK_URL"
    tar -xzf "$tmp/jdk.tgz" -C "$tmp"
    src="$(find "$tmp" -maxdepth 2 -type d -name 'jdk-17*' | head -1)"
    [[ -d "$src" ]] || { echo "JDK extraction failed" >&2; exit 3; }
    mv "$src" "$JDK_DIR"
    rm -rf "$tmp"
fi
log "JDK: $("$JDK_DIR/bin/java" -version 2>&1 | head -1)"

export JAVA_HOME="$JDK_DIR"
export PATH="$JAVA_HOME/bin:$PATH"

# ── Android cmdline-tools ──────────────────────────────────────────────────
if [[ ! -x "$SDK_DIR/cmdline-tools/latest/bin/sdkmanager" ]]; then
    log "Download Android cmdline-tools"
    mkdir -p "$SDK_DIR/cmdline-tools"
    tmp="$(mktemp -d)"
    curl -fL --progress-bar -o "$tmp/ct.zip" "$CMDLINE_URL"
    unzip -q "$tmp/ct.zip" -d "$tmp"
    mkdir -p "$SDK_DIR/cmdline-tools/latest"
    mv "$tmp/cmdline-tools"/* "$SDK_DIR/cmdline-tools/latest/"
    rm -rf "$tmp"
fi

export ANDROID_HOME="$SDK_DIR"
export ANDROID_SDK_ROOT="$SDK_DIR"
export PATH="$SDK_DIR/cmdline-tools/latest/bin:$SDK_DIR/platform-tools:$PATH"

# ── Accept licenses + install SDK packages ─────────────────────────────────
log "Accept licenses"
yes | sdkmanager --licenses >/dev/null 2>&1 || true

log "Install platform-tools, android-34, build-tools 34.0.0"
sdkmanager \
    "platform-tools" \
    "platforms;android-34" \
    "build-tools;34.0.0" 2>&1 | tail -10

log "Done."
echo
echo "JAVA_HOME=$JAVA_HOME"
echo "ANDROID_HOME=$ANDROID_HOME"
df -h "$HOME" | tail -1

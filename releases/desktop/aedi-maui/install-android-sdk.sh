#!/usr/bin/env bash
# install-android-sdk.sh
# Installs the Android SDK + JDK 17 needed to build the AEDI-S MAUI Android
# client on a real dev host (x86_64 Linux, macOS, or Windows via WSL2).
#
#   Usage:
#     ./install-android-sdk.sh                     # auto-detect OS, install to ~/Android/Sdk
#     ANDROID_HOME=/opt/android-sdk ./install-android-sdk.sh
#
# After install, this script prints the env vars to add to ~/.bashrc / ~/.zshrc.
#
# NOT supported: aarch64 Linux (build-tools aapt2 is x86_64-only on Linux).
# On Raspberry Pi/M-series, build via GitHub Actions (.github/workflows/aedi-maui-android.yml).
set -euo pipefail

OS="$(uname -s)"
ARCH="$(uname -m)"
ANDROID_HOME="${ANDROID_HOME:-$HOME/Android/Sdk}"
JAVA_HOME="${JAVA_HOME:-$HOME/.jdks/temurin-17}"
CMDLINE_REV="11076708"   # Android cmdline-tools release (stable)

case "$OS" in
  Linux)
    [[ "$ARCH" != "x86_64" ]] && { echo "error: Android SDK build-tools require x86_64 on Linux (got $ARCH). Use CI workflow instead." >&2; exit 2; }
    ZIP="commandlinetools-linux-${CMDLINE_REV}_latest.zip"
    JDK_ARCH="linux-x64"
    ;;
  Darwin)
    ZIP="commandlinetools-mac-${CMDLINE_REV}_latest.zip"
    case "$ARCH" in
      x86_64)  JDK_ARCH="mac-x64" ;;
      arm64)   JDK_ARCH="mac-aarch64" ;;
      *) echo "unsupported macOS arch: $ARCH" >&2; exit 2 ;;
    esac
    ;;
  *) echo "unsupported OS: $OS (use WSL2 on Windows)" >&2; exit 2 ;;
esac

mkdir -p "$(dirname "$JAVA_HOME")" "$ANDROID_HOME/cmdline-tools"

echo ">>> Installing Temurin JDK 17 → $JAVA_HOME"
if [[ ! -x "$JAVA_HOME/bin/java" ]]; then
  tmp="$(mktemp -d)"
  url="https://api.adoptium.net/v3/binary/latest/17/ga/${OS,,}/${ARCH/x86_64/x64}/jdk/hotspot/normal/eclipse"
  curl -fsSL -o "$tmp/jdk.tar.gz" "$url"
  tar -xzf "$tmp/jdk.tar.gz" -C "$tmp"
  src="$(find "$tmp" -maxdepth 2 -type d -name 'jdk-17*' | head -1)"
  [[ -z "$src" ]] && { echo "JDK extraction failed" >&2; exit 3; }
  mv "$src" "$JAVA_HOME"
  rm -rf "$tmp"
fi
"$JAVA_HOME/bin/java" -version

echo ">>> Installing Android cmdline-tools → $ANDROID_HOME"
if [[ ! -x "$ANDROID_HOME/cmdline-tools/latest/bin/sdkmanager" ]]; then
  tmp="$(mktemp -d)"
  curl -fsSL -o "$tmp/$ZIP" "https://dl.google.com/android/repository/$ZIP"
  unzip -q "$tmp/$ZIP" -d "$tmp"
  mkdir -p "$ANDROID_HOME/cmdline-tools/latest"
  mv "$tmp/cmdline-tools"/* "$ANDROID_HOME/cmdline-tools/latest/"
  rm -rf "$tmp"
fi

export JAVA_HOME ANDROID_HOME
export PATH="$JAVA_HOME/bin:$ANDROID_HOME/cmdline-tools/latest/bin:$ANDROID_HOME/platform-tools:$PATH"

echo ">>> Accepting licenses"
yes | sdkmanager --licenses >/dev/null

echo ">>> Installing platform-tools + Android 34 SDK + build-tools 34.0.0"
sdkmanager \
  "platform-tools" \
  "platforms;android-34" \
  "build-tools;34.0.0"

echo
echo "=== Done. Add these to your shell rc ==="
cat <<EOF
export JAVA_HOME="$JAVA_HOME"
export ANDROID_HOME="$ANDROID_HOME"
export ANDROID_SDK_ROOT="$ANDROID_HOME"
export PATH="\$JAVA_HOME/bin:\$ANDROID_HOME/cmdline-tools/latest/bin:\$ANDROID_HOME/platform-tools:\$PATH"
EOF
echo
echo "Then build the MAUI client:"
echo "  cd releases/desktop/aedi-maui"
echo "  dotnet workload install maui-android"
echo "  ./build.sh release android"

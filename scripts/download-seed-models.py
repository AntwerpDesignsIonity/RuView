#!/usr/bin/env python3
"""
Download pre-trained ONNX models for Cognitum Seed (RPi Zero 2 W).

Uses wget fallback when huggingface-hub is not installed.
Designed for low-RAM edge devices — downloads one file at a time.

Usage:
    python scripts/download-seed-models.py
    python scripts/download-seed-models.py --model-dir /path/to/models
    python scripts/download-seed-models.py --quantized-only   # skip full ONNX encoder
"""

from __future__ import annotations

import argparse
import hashlib
import os
import shutil
import subprocess
import sys
import urllib.request
import urllib.error

HF_REPO = "ionity-global/wifi-densepose-pretrained"
HF_BASE = f"https://huggingface.co/{HF_REPO}/resolve/main"

# Files required for edge inference on RPi Zero 2 W
EDGE_FILES = [
    ("model-q4.bin",           "4-bit quantized model (recommended for 512 MB RAM)"),
    ("presence-head.json",     "Presence detection head weights"),
    ("config.json",            "Model configuration"),
]

FULL_FILES = [
    ("pretrained-encoder.onnx", "ONNX encoder (~2 MB, contrastive TCN backbone)"),
    ("model-q8.bin",            "8-bit quantized model (higher quality, more RAM)"),
    ("model.safetensors",       "SafeTensors full model"),
]

DEFAULT_MODEL_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "models", "pretrained",
)


def download_file(url: str, dest: str) -> bool:
    """Download a single file with progress."""
    tmp = dest + ".tmp"
    try:
        # Try huggingface-hub first
        try:
            from huggingface_hub import hf_hub_download
            filename = os.path.basename(dest)
            local_dir = os.path.dirname(dest)
            hf_hub_download(repo_id=HF_REPO, filename=filename, local_dir=local_dir)
            return True
        except ImportError:
            pass

        # Fall back to wget if available (better progress on terminal)
        if shutil.which("wget"):
            result = subprocess.run(
                ["wget", "-q", "--show-progress", "-O", tmp, url],
                check=False,
            )
            if result.returncode == 0 and os.path.exists(tmp):
                os.rename(tmp, dest)
                return True

        # Last resort: urllib
        print(f"    Downloading via urllib…", end="", flush=True)
        urllib.request.urlretrieve(url, tmp)
        os.rename(tmp, dest)
        print(" done")
        return True

    except (urllib.error.URLError, OSError) as e:
        print(f"\n    Error: {e}")
        if os.path.exists(tmp):
            os.remove(tmp)
        return False


def main():
    parser = argparse.ArgumentParser(description="Download AEDI-S ONNX models for Cognitum Seed")
    parser.add_argument("--model-dir", default=DEFAULT_MODEL_DIR, help="Target directory")
    parser.add_argument("--quantized-only", action="store_true",
                        help="Download only quantized model + config (saves bandwidth)")
    args = parser.parse_args()

    os.makedirs(args.model_dir, exist_ok=True)

    files = list(EDGE_FILES)
    if not args.quantized_only:
        files.extend(FULL_FILES)

    print(f"\n  Cognitum Seed Model Downloader")
    print(f"  Repo: {HF_REPO}")
    print(f"  Dest: {args.model_dir}")
    print(f"  Files: {len(files)}\n")

    downloaded = 0
    skipped = 0
    failed = 0

    for filename, description in files:
        dest = os.path.join(args.model_dir, filename)
        if os.path.exists(dest) and os.path.getsize(dest) > 0:
            size = os.path.getsize(dest)
            print(f"  ✔ {filename} ({size:,} bytes) — already present")
            skipped += 1
            continue

        url = f"{HF_BASE}/{filename}"
        print(f"  ↓ {filename} — {description}")

        if download_file(url, dest):
            size = os.path.getsize(dest)
            print(f"    ✔ Saved ({size:,} bytes)")
            downloaded += 1
        else:
            print(f"    ✖ Failed to download {filename}")
            failed += 1

    print(f"\n  Summary: {downloaded} downloaded, {skipped} already present, {failed} failed\n")
    return 1 if failed > 0 else 0


if __name__ == "__main__":
    sys.exit(main())

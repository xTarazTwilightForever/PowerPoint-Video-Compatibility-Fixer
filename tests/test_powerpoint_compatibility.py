#!/usr/bin/env python3
"""
test_powerpoint_compatibility.py

Verify that converted MP4 files are compatible with Microsoft PowerPoint:
- Video: H.264, High@L4.1, yuv420p
- Audio: AAC LC
"""

import subprocess
from pathlib import Path

OUTPUT_DIR = Path("data/output")

def probe(file: Path, stream: str, entries: str) -> dict:
    """Run ffprobe and return parsed key/value pairs."""
    cmd = [
        "ffprobe", "-v", "error",
        "-select_streams", stream,
        "-show_entries", entries,
        "-of", "default=noprint_wrappers=1",
        str(file)
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    out = {}
    for line in result.stdout.strip().splitlines():
        if "=" in line:
            k, v = line.split("=", 1)
            out[k.strip()] = v.strip()
    return out

def check_file(file: Path) -> bool:
    """Check if a file meets PowerPoint video requirements."""
    try:
        v = probe(file, "v:0", "stream=codec_name,profile,level,pix_fmt")
        a = probe(file, "a:0", "stream=codec_name,profile")

        return (
            v.get("codec_name") == "h264"
            and v.get("profile") == "High"
            and v.get("level") in {"41", "42"}
            and v.get("pix_fmt") == "yuv420p"
            and a.get("codec_name") == "aac"
            and a.get("profile", "").upper() == "LC"
        )
    except Exception as e:
        print(f"  ✗ Error probing {file.name}: {e}")
        return False

def main():
    files = list(OUTPUT_DIR.glob("*.mp4"))
    if not files:
        print("No MP4 files found in data/output")
        return 1

    print(f"Checking {len(files)} files in {OUTPUT_DIR.resolve()}...\n")
    all_ok = True
    for f in files:
        if check_file(f):
            print(f"  ✓ {f.name} — ✅ Compatible with PowerPoint")
        else:
            print(f"  ✗ {f.name} — ❌ NOT compatible with PowerPoint")
            all_ok = False

    print("\nDone.")
    return 0 if all_ok else 1

if __name__ == "__main__":
    raise SystemExit(main())
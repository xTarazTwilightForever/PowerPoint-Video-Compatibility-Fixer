#!/usr/bin/env python3
"""
powerpoint_video_compatibility_fixer.py

Batch converter for making video files PowerPoint-compatible.

This tool:
    - Converts videos into standardized MP4 (H.264 + AAC, yuv420p).
    - Fixes "PowerPoint cannot insert video" errors by enforcing strict codec settings.
    - Supports recursion, overwriting, downscaling, and CRF tuning.

Default behavior:
    Input:  ./data/input
    Output: ./data/output

Requirements:
    Install dependencies before use:
        pip install -r requirements.txt

Usage Example:
    Convert with defaults:
        python src/powerpoint_video_compatibility_fixer.py

    Convert recursively with custom CRF and resolution:
        python src/powerpoint_video_compatibility_fixer.py -i ./data/input -o ./data/output --recursive --crf 20 --max-width 1280 --max-height 720

    Overwrite existing files:
        python src/powerpoint_video_compatibility_fixer.py --overwrite
"""

import argparse
import sys
import subprocess
from pathlib import Path
from typing import Iterable, List, Tuple
from dataclasses import dataclass

try:
    from moviepy.editor import VideoFileClip
    import imageio_ffmpeg
except ImportError:
    print("moviepy is required. Install with: pip install moviepy", file=sys.stderr)
    raise

SUPPORTED_EXTS = {
    ".mp4", ".mov", ".m4v", ".avi", ".mkv", ".wmv", ".webm",
    ".mpg", ".mpeg", ".3gp", ".ts"
}

DEFAULT_INPUT_DIR = "data/input"
DEFAULT_OUTPUT_DIR = "data/output"


@dataclass
class Settings:
    """Container for CLI settings and runtime configuration."""

    input_dir: Path
    output_dir: Path
    recursive: bool
    overwrite: bool
    crf: int
    audio_bitrate: str
    max_width: int
    max_height: int
    quiet: bool


def ensure_ffmpeg() -> Path:
    """Verify that the ffmpeg binary is available.

    Returns:
        Path: Path to ffmpeg executable.

    Raises:
        RuntimeError: If ffmpeg binary is not found.
    """
    exe = Path(imageio_ffmpeg.get_ffmpeg_exe())
    if not exe.exists():
        raise RuntimeError("FFmpeg binary not found.")
    return exe


def discover_videos(root: Path, recursive: bool) -> List[Path]:
    """Collect supported video files from the input directory.

    Args:
        root (Path): Root input directory.
        recursive (bool): Whether to scan subdirectories.

    Returns:
        List[Path]: List of discovered video file paths.

    Raises:
        FileNotFoundError: If input directory does not exist.
    """
    if not root.exists():
        raise FileNotFoundError(f"Input directory not found: {root}")
    pattern = "**/*" if recursive else "*"
    return sorted(
        [p for p in root.glob(pattern) if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS]
    )


def build_output_path(src: Path, out_dir: Path) -> Path:
    """Build output path with same filename and .mp4 extension.

    Args:
        src (Path): Source video file path.
        out_dir (Path): Output directory.

    Returns:
        Path: Output video path with .mp4 extension.
    """
    return (out_dir / src.stem).with_suffix(".mp4")


def resize_if_needed(clip: VideoFileClip, max_w: int, max_h: int) -> VideoFileClip:
    """Downscale video if it exceeds maximum resolution.

    Args:
        clip (VideoFileClip): Input video clip.
        max_w (int): Maximum allowed width.
        max_h (int): Maximum allowed height.

    Returns:
        VideoFileClip: Resized clip or original if within limits.
    """
    if clip.w <= max_w and clip.h <= max_h:
        return clip
    ratio = min(max_w / clip.w, max_h / clip.h)
    return clip.resize(ratio)


def convert_file(
    src: Path,
    dst: Path,
    *,
    crf: int,
    audio_bitrate: str,
    max_w: int,
    max_h: int,
    quiet: bool
) -> Tuple[bool, str]:
    """Convert a single video into standardized MP4.

    Workflow:
        1. Open source video.
        2. Resize if needed.
        3. Re-encode with strict codec settings.
        4. Save as temporary file.
        5. Rename to final output.

    Args:
        src (Path): Source video path.
        dst (Path): Destination video path.
        crf (int): Constant Rate Factor (quality, 18–35).
        audio_bitrate (str): Audio bitrate (e.g., "160k").
        max_w (int): Maximum allowed width.
        max_h (int): Maximum allowed height.
        quiet (bool): Suppress ffmpeg logs.

    Returns:
        Tuple[bool, str]: (Success flag, status message).
    """
    tmp = dst.parent / (dst.stem + "_tmp.mp4")
    try:
        clip = VideoFileClip(str(src))
    except Exception as e:
        return False, f"open-failed: {src.name} ({e})"

    try:
        clip = resize_if_needed(clip, max_w, max_h)
        fps = clip.fps or 30

        ffmpeg_params = [
            "-movflags", "+faststart",
            "-pix_fmt", "yuv420p",
            "-crf", str(crf),
            "-profile:v", "high",
            "-level", "4.1"
        ]
        logger = None if quiet else "bar"

        clip.write_videofile(
            str(tmp),
            codec="libx264",
            audio_codec="aac",
            audio_bitrate=audio_bitrate,
            fps=fps,
            ffmpeg_params=ffmpeg_params,
            logger=logger,
        )
    except Exception as e:
        return False, f"convert-failed: {src.name} ({e})"
    finally:
        try:
            clip.close()
        except Exception:
            pass

    try:
        if dst.exists():
            dst.unlink()
        tmp.rename(dst)
    except Exception as e:
        return False, f"finalize-failed: {src.name} ({e})"

    return True, f"ok: {src.name} -> {dst.name}"


def run_batch(settings: Settings) -> int:
    """Process all videos in batch according to provided settings.

    Args:
        settings (Settings): Runtime configuration.

    Returns:
        int: Exit code (0 = success, 1 = failure).
    """
    ensure_ffmpeg()
    settings.output_dir.mkdir(parents=True, exist_ok=True)

    files = discover_videos(settings.input_dir, settings.recursive)
    if not files:
        print(f"No input files in: {settings.input_dir}")
        return 1

    print(f"Found {len(files)} files")
    print(f"Input: {settings.input_dir.resolve()}")
    print(f"Output: {settings.output_dir.resolve()}")
    print(f"CRF={settings.crf}, Audio={settings.audio_bitrate}, Max={settings.max_width}x{settings.max_height}")
    print("-" * 60)

    success, fail, skip = 0, 0, 0
    for i, src in enumerate(files, 1):
        dst = build_output_path(src, settings.output_dir)

        if dst.exists() and not settings.overwrite:
            print(f"[{i}/{len(files)}] skip (exists): {src.name}")
            skip += 1
            continue

        print(f"[{i}/{len(files)}] convert: {src.name} -> {dst.name}")
        ok, msg = convert_file(
            src, dst,
            crf=settings.crf,
            audio_bitrate=settings.audio_bitrate,
            max_w=settings.max_width,
            max_h=settings.max_height,
            quiet=settings.quiet,
        )
        if ok:
            success += 1
            print(f"  ✓ {msg}")
        else:
            fail += 1
            print(f"  ✗ {msg}")

    print("-" * 60)
    print(f"Done. Success: {success}, Skipped: {skip}, Failed: {fail}")
    return 0 if fail == 0 else 1


def parse_args(argv: Iterable[str]) -> Settings:
    """Parse CLI arguments into Settings.

    Args:
        argv (Iterable[str]): Raw CLI arguments.

    Returns:
        Settings: Parsed runtime configuration.

    Raises:
        SystemExit: If arguments are invalid.
    """
    parser = argparse.ArgumentParser(
        prog="powerpoint_video_compatibility_fixer",
        description="Batch convert videos into MP4 (H.264 + AAC) for PowerPoint."
    )
    parser.add_argument("--input", "-i", type=Path, default=Path(DEFAULT_INPUT_DIR))
    parser.add_argument("--output", "-o", type=Path, default=Path(DEFAULT_OUTPUT_DIR))
    parser.add_argument("--recursive", "-r", action="store_true")
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--crf", type=int, default=22)
    parser.add_argument("--audio-bitrate", default="160k")
    parser.add_argument("--max-width", type=int, default=1920)
    parser.add_argument("--max-height", type=int, default=1080)
    parser.add_argument("--quiet", action="store_true")
    args = parser.parse_args(list(argv))

    if not (18 <= args.crf <= 35):
        parser.error("--crf must be between 18 and 35.")

    return Settings(
        input_dir=args.input.resolve(),
        output_dir=args.output.resolve(),
        recursive=bool(args.recursive),
        overwrite=bool(args.overwrite),
        crf=int(args.crf),
        audio_bitrate=str(args.audio_bitrate),
        max_width=int(args.max_width),
        max_height=int(args.max_height),
        quiet=bool(args.quiet),
    )


def run_post_check() -> None:
    """Run external PowerPoint compatibility test script.

    This executes `tests/test_powerpoint_compatibility.py`
    and prints a summary of results.
    """
    print("\nRunning PowerPoint compatibility check...\n")
    result = subprocess.run(
        ["python", "tests/test_powerpoint_compatibility.py"],
        text=True
    )
    if result.returncode == 0:
        print("\nAll converted files are PowerPoint-compatible ✅")
    else:
        print("\nSome files are NOT PowerPoint-compatible ❌")


def main() -> int:
    """Entry point for script execution.

    Returns:
        int: Exit code (0 = success, non-zero = error).
    """
    try:
        settings = parse_args(sys.argv[1:])
        code = run_batch(settings)
        run_post_check()
        return code
    except KeyboardInterrupt:
        print("Interrupted by user.")
        return 130
    except Exception as e:
        print(f"Critical error: {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
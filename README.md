# PowerPoint-Video-Compatibility-Fixer

![Python](https://img.shields.io/badge/python-3.9%2B-blue)  
![License](https://img.shields.io/badge/license-Apache%202.0-green)  
![Platform](https://img.shields.io/badge/platform-win%20%7C%20mac%20%7C%20linux-lightgrey)  

**PowerPoint-Video-Compatibility-Fixer** is a Python-based tool that batch converts and normalizes video files into fully compliant MP4 format (**H.264 + AAC, yuv420p**).  
It ensures seamless integration into **Microsoft PowerPoint** presentations by eliminating the common error:

> *“PowerPoint cannot insert a video from the selected file.”*

---

## Features
- 🎥 **Batch conversion** of multiple video files at once  
- 🔒 **Guaranteed compatibility** with PowerPoint and other presentation software  
- ⚙️ **Strict codec enforcement** (H.264 video, AAC audio, yuv420p pixel format, +faststart)  
- 📏 **Automatic downscaling** for oversized resolutions  
- 🎚 **Configurable CRF quality** (default 22)  
- 🗂 **Recursive folder scanning** with overwrite control  
- ✅ **Automatic validation** after conversion — ensures every file is truly PowerPoint-compatible  
- 🖥 **Cross-platform** (Windows, macOS, Linux)  

---

## Requirements
- Python 3.9+  
- ffmpeg (installed automatically via `imageio-ffmpeg`)  
- pip for dependency management  

---

## Installation

Clone the repository and install dependencies:

```bash
git clone https://github.com/xTarazTwilightForever/PowerPoint-Video-Compatibility-Fixer.git
cd PowerPoint-Video-Compatibility-Fixer
pip install -r requirements.txt
```

---

## Usage

By default the tool reads from `data/input/` and writes to `data/output/`:

```bash
python src/powerpoint_video_compatibility_fixer.py
```

---

## Options

All CLI options supported by **PowerPoint-Video-Compatibility-Fixer**:

| Flag                | Description                           | Default            |
|---------------------|---------------------------------------|--------------------|
| `--input, -i`       | Input folder                          | `./data/input`     |
| `--output, -o`      | Output folder                         | `./data/output`    |
| `--recursive, -r`   | Scan subdirectories                   | off                |
| `--overwrite`       | Overwrite existing files              | off                |
| `--crf`             | Quality factor (18–28 recommended)    | 22                 |
| `--audio-bitrate`   | AAC bitrate                           | 160k               |
| `--max-width`       | Maximum width before downscaling      | 1920               |
| `--max-height`      | Maximum height before downscaling     | 1080               |
| `--quiet`           | Suppress ffmpeg logs                  | off                |

---

## Examples

Convert with defaults:  

```bash
python src/powerpoint_video_compatibility_fixer.py
```

Recursive scan with custom quality and resolution:  

```bash
python src/powerpoint_video_compatibility_fixer.py -i ./raw -o ./ready --recursive --crf 20 --max-width 1280 --max-height 720
```

Overwrite existing files:  

```bash
python src/powerpoint_video_compatibility_fixer.py --overwrite
```

---

## License

Apache License 2.0 — see [LICENSE](LICENSE).

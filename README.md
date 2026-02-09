# TranscriptFlow Pro

TranscriptFlow Pro is a professional transcription application designed for speed, accuracy, and a seamless user experience. It features a high-performance media engine, a rich-text editor with timecode integration, and a dynamic waveform display.

![TranscriptFlow Splash](splash.png)

## Features

- **High-Performance Media Engine**: Supports MPV (with frame-stepping) and VLC backends.
- **Rich-Text Editor**: Professional editing with bold, italic, underline, and custom font support.
- **Timecode Integration**: Clickable timecodes in text that seek the video instantly.
- **Dynamic Waveform**: Beautiful, interactive waveform display for precise positioning.
- **Pivot-Based Sync**: Advanced logic to keep your transcript perfectly synchronized with media time.
- **Spell Check**: Integrated spell checking with support for custom dictionaries.
- **Customizable Layouts**: Choose from various UI presets (Standard, Stacked, Compact, etc.).
- **Automatic Backups**: Keep your work safe with intelligent background backups.

## Installation

You can download the latest installer from the [Releases](https://github.com/REPLACE_WITH_YOUR_USERNAME/TranscriptFlow/releases) page.

1. Run `TranscriptFlow_Setup_x.x.x.exe`.
2. Follow the on-screen instructions.
3. Launch TranscriptFlow from your desktop or start menu.

## Development Setup

If you want to run the code from source:

### Prerequisites

- Python 3.8+
- [FFmpeg](https://ffmpeg.org/) (for waveform generation)
- [VLC Media Player](https://www.videolan.org/vlc/) (installed on system)

### Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/REPLACE_WITH_YOUR_USERNAME/TranscriptFlow.git
   cd TranscriptFlow
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the application:
   ```bash
   python main.py
   ```

## Building the Installer

We use a custom build script to package the application:

```bash
python build_installer.py --installer
```

This will:
- Bundle the app using PyInstaller.
- Create a Windows installer using Inno Setup.
- Output the setup file to `installer_output/`.

## License

[Add your license information here, e.g., MIT License]

## Acknowledgments

- Built with PyQt6.
- Media powered by libmpv and libvlc.

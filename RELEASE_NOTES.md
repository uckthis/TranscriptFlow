# Release v1.1.7 (Smart Typing, Work Time Tracking & Enhanced Workflow)

This update introduces intelligent auto-capitalization (Smart Typing), live session and file active time tracking, case-preserving bulk spell check corrections, plain text paste options, and smooth auto-scrolling with typewriter mode.

## ✍️ Smart Typing & Auto-Capitalization (New!)
- **Intelligent Capitalization**: Automatically capitalizes the next word following sentence terminators (`.`, `?`, `!`), newlines, or ellipses (`...`).
- **Context-Aware Exception Handling**: Skips capitalization after common honorifics and abbreviations (`Mr.`, `Dr.`, `vs.`, initials), decimal numbers, and video timecodes.
- **Custom Triggers**: Define your own punctuation triggers (such as `--`, `:`, `~`) in Preferences.
- **Ribbon & Preferences Integration**: Toggle Smart Typing instantly via the new split button in Ribbon Zone C or fine-tune trigger rules in **Edit > Options... > General**.

## ⏱️ Live Working Time & Word Statistics
- **Status Bar Productivity Metrics**: Real-time display of active **Session Time**, accumulated **File Time**, and live **Word Count**.
- **Smart Idle Detection**: Timers automatically pause when you are idle or when the application is in the background, only counting active work or media playback. File times persist reliably across editing sessions.

## 🔡 Case-Preserving Spell Check Corrections
- **"Change All (keep caps)"**: Replace misspelled words globally while intelligently preserving the original casing of each individual match (UPPERCASE, Title Case, or lowercase).
- **Single-Step Undo**: Undo bulk spell corrections cleanly with a single `Ctrl+Z`.
- **Context Menu Integration**: Right-click any flagged word to replace all occurrences directly.

## 📄 Paste Plain Text & Smooth Auto-Scrolling
- **Paste Plain Text**: New `Ctrl+Shift+V` shortcut, Edit menu action, and right-click context menu option to strip external formatting and paste pure text.
- **Auto-Scroll & Typewriter Mode**: Toggleable view modes on Ribbon Zone A (*Smart Margin*, *Typewriter Center*, *Advance Bump*, and *Scroll Past End*).
- **1-Line Smooth Wheel Scrolling**: Replaced jerky 3-line jumps with precise, smooth single-line mouse wheel scrolling.

---

# Release v1.1.6 (Stability & UI Overlap Fixes)

This release addresses UI item layout in management dialogs, prevents formatting loss during settings changes, and fixes application freezes during bulk Replace All operations.

## 🎨 UI & Dialog Improvements
- **Snippet & Shortcut Item Layout**: Resolved vertical overlap of list entries in the Snippets and Shortcuts dialogs by dynamically calculating row heights to accommodate multiline snippets and stylesheet padding.

## 🛠️ Formatting & Stability Fixes
- **Preserve Formatting on Settings Update**: Fixed an issue where changing frame rate or transcript settings while text was selected would inadvertently strip font colors, bold, and italic styles from the selection.
- **Find & Replace All Infinite Loop Fix**: Fixed a hang during Replace All operations when replacing text with strings sharing identical prefixes (e.g. `"MATT BAKER:"` with `"MATT BAKER:  "`).
- **Non-Destructive Ribbon Formatting**: Updating font family or font size from the ribbon now merges attributes rather than replacing character format, keeping custom text colors and bold/italic styles intact.

---

# Release v1.1.5 (Hotfix: OCR Dependencies)

This is a critical hotfix to resolve the "ModuleNotFoundError: No module named 'PIL'" error encountered in v1.1.4.

## 🛠️ Fixes
- **Bundled Dependencies**: Fixed a build configuration issue where `PIL` (Pillow) and `pytesseract` were not correctly bundled in the final installer.
- **Improved OCR Stability**: Verified all OCR core dependencies are included in the frozen environment.

---

# Release v1.1.4 (Professional OCR Update)

This update supercharges the OCR engine with multi-engine support, automated installation, and a much cleaner interface for managing languages.

## 🌍 Multi-Engine & Automated OCR
- **Windows + Tesseract**: Now supports both Windows Native OCR (fast) and Tesseract OCR (high precision/multi-language).
- **One-Click Install**: Tesseract engine can now be installed automatically with one click from the Preferences dialog.
- **Friendly Language Names**: Technical codes like `ara` and `spa` have been replaced with human-readable names like `Arabic` and `Spanish`.
- **Searchable Language Catalog**: Search and install from dozens of supported Tesseract languages through a new selection dialog.

## 🎨 UI & UX Refinement
- **High-Contrast Selection**: Overhauled the language selection list with a high-contrast style to ensure text is always visible.
- **Improved Settings**: Renamed ambiguous "osd" model to "Orientation & Script Detection" for better clarity.
- **RTL Stability**: Further refinements to Urdu (RTL) alignment and font persistence.

## 🛠️ Fixes & Under-the-Hood
- **Enhanced Tesseract Pathing**: Improved binary discovery for different Windows configurations.
- **Resource Management**: Optimized loading of OCR drivers to reduce startup latency.

---

# Release v1.1.2 (Sniper OCR Update)

This update brings a powerful new OCR tool to instantly capture text from your screen, along with professional formatting support for Find & Replace and a massive volume boost.

## 🎯 Sniper OCR Tool (New!)
- **Instant Capture**: Use the new **🎯 OCR** button in the ribbon or press `Ctrl+Shift+O` to select any area on your screen and extract text instantly.
- **Auto-Formatting**: Automatically convert captured text to UPPERCASE, lowercase, or Title Case.
- **Smart Prefixes/Suffixes**: Add custom text like `": "` automatically to captured speaker names.
- **Flexible Output**: Choose to copy to clipboard, insert at cursor, or both. Configure these in **Edit > Options... > OCR**.

## 🔍 Format-Aware Find & Replace
- **Find Formatting**: You can now filter search results by formatting. Search for "text" only where it is **Bold** or *Italic*.
- **Replace with Format**: Apply Bold or Italic styles to your replacement text automatically.
- **Improved Logic**: Find Next and Replace All now correctly respect these formatting constraints across the entire document.

## 🔊 Extreme Volume Boost (400%)
- **Triple the Boost**: The Volume Boost control now allows for up to **300% additional gain**, bringing the total volume capability to **400%**. Perfect for quiet field recordings or distant speakers.
- **Smart Initialization**: Your volume, boost, and playback speed settings are now correctly remembered and reapplied automatically when loading new media files.

## 🛠️ Fixes & Refinement
- **Persistent Playback Speed**: Fixed a bug where the playback rate would reset to 1.0x every time a new file was loaded.
- **Engine Optimization**: Adjusted MPV backend initialization to support the expanded volume range without software clipping.

## 📥 Installation
Download `TranscriptFlow_Setup_1.1.2.exe` below and run it. 

---
**Tip**: Combine formatting checkboxes with "Match Case" for ultimate precision when cleaning up transcripts!

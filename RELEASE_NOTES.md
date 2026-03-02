# TranscriptFlow Pro v1.1.3 - Transcription Efficiency Update

This update refines the "Change Case" feature to favor transcription workflows and speaker identification.

## 🔠 Refined Change Case (Shift+F3)
- **UPPERCASE Priority**: Pressing `Shift+F3` on a speaker name (e.g., "Speaker:") now converts it to **UPPERCASE** on the very first press. This eliminates the extra toggles previously required for speaker identification.
- **Refined Cycle**: The logic now intelligently detects Title/Sentence case and prioritizes the most likely desired state for transcribers.

# TranscriptFlow Pro v1.1.2 - Format-Aware Search & Extreme Boost

This update brings professional formatting support to Find & Replace and a massive volume boost for difficult recordings.

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

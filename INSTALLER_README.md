# TranscriptFlow Installer Documentation

## Overview

This document explains how the TranscriptFlow installer works, including architecture detection, permission handling, and directory structure.

## Installation Architecture

### Automatic Architecture Detection

The installer automatically detects your system architecture and installs to the appropriate location:

- **64-bit (x64) Systems**: Installs to `C:\Program Files\TranscriptFlow Pro\`
- **32-bit (x86) Systems**: Installs to `C:\Program Files (x86)\TranscriptFlow Pro\`

The installer uses Inno Setup's `{autopf}` constant which automatically selects the correct Program Files directory based on:
1. The system architecture (32-bit or 64-bit Windows)
2. The installer architecture (configured in the .iss file)

### Current Configuration

The installer is currently configured for **x64 only** with these settings:
```pascal
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
```

**To support both x86 and x64:**
1. Comment out or remove the `ArchitecturesAllowed` and `ArchitecturesInstallIn64BitMode` lines
2. Build separate installers for x86 and x64, OR
3. Create a universal installer (larger file size)

## Permission Handling

### Administrator Privileges

The installer **requires administrator privileges** for the following reasons:

1. **Program Files Installation**: Writing to `C:\Program Files` requires admin rights
2. **File Association**: Registering `.tflow` file extension requires registry access
3. **Start Menu Shortcuts**: Creating system-wide shortcuts requires elevation

The installer is configured with:
```pascal
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog
```

This means:
- The installer will request UAC elevation
- Users can decline and install to a different location (like their user folder)
- If declined, some features (file associations) may not work

### Write Permissions Strategy

TranscriptFlow uses a **dual-directory approach** for security and functionality:

#### 1. Program Files (Read-Only)
**Location**: `C:\Program Files\TranscriptFlow Pro\` or `C:\Program Files (x86)\TranscriptFlow Pro\`

**Contents**:
- Application executable (`TranscriptFlow.exe`)
- Application libraries (DLLs)
- Bundled dictionaries (initial copy)
- Application resources (icons, etc.)

**Permissions**: Read-only for standard users (Windows default)

**Purpose**: Contains the application binaries that should not be modified during normal operation

#### 2. AppData (Read-Write)
**Location**: `C:\Users\[Username]\AppData\Roaming\TranscriptFlow\`

**Contents**:
- `BACKUP\` - Auto-backup files
- `dicts\` - Dictionary files (can be downloaded/updated)
- `custom_dicts\` - User's custom dictionaries
- `waveform_cache\` - Cached waveform data
- `config.json` - User settings and preferences

**Permissions**: Full read-write access for the user (Windows default)

**Purpose**: All user data and files that need to be modified

### Why This Approach?

1. **Security**: Application binaries in Program Files are protected from accidental modification
2. **Multi-User Support**: Each user has their own settings and data
3. **No UAC Prompts**: The application runs without admin rights after installation
4. **Windows Best Practices**: Follows Microsoft's recommended application structure
5. **Easy Backups**: User data is in a standard location that backup software knows about

## Directory Creation Process

### During Installation

The installer performs these steps:

1. **Install to Program Files** (requires admin)
   - Copies all application files
   - Includes initial dictionary files
   - Sets up application structure

2. **Create AppData Structure** (automatic)
   - Creates `%APPDATA%\TranscriptFlow\` directory
   - Creates subdirectories (BACKUP, dicts, custom_dicts, waveform_cache)
   - Copies bundled dictionaries to user's AppData
   - Verifies write permissions

3. **Verify Permissions**
   - Tests write access to AppData directory
   - Shows warning if write access fails (rare)
   - Logs all operations for troubleshooting

### First Run

On first run, the application (`path_manager.py`) will:

1. Check if AppData directories exist
2. Create any missing directories
3. Copy bundled dictionaries if not present
4. Initialize configuration file

This ensures the application works even if the installer's AppData setup fails.

## File Operations

### Dictionary Downloads

When the user downloads a new dictionary:

1. Application downloads dictionary files
2. Saves to `%APPDATA%\TranscriptFlow\dicts\`
3. No admin rights required
4. Works for all users

### Backup Files

When the application creates backups:

1. Saves to `%APPDATA%\TranscriptFlow\BACKUP\`
2. No admin rights required
3. Automatic cleanup of old backups
4. User-specific backups

### Configuration

User settings are saved to:
- `%APPDATA%\TranscriptFlow\config.json`
- Automatically created on first run
- Persists across application updates

## Uninstallation

### Standard Uninstall

When uninstalling via Control Panel:

1. Removes all files from Program Files
2. Removes Start Menu shortcuts
3. Removes desktop icon (if created)
4. **Asks user** about AppData:
   - "Yes" = Remove all user data (backups, settings, dictionaries)
   - "No" = Keep user data for future installations

### Clean Uninstall

To completely remove all traces:

1. Uninstall via Control Panel and select "Yes" to remove data
2. Manually delete (if needed):
   - `%APPDATA%\TranscriptFlow\`
   - Registry keys under `HKEY_CURRENT_USER\Software\TranscriptFlow`

## Troubleshooting

### Installation Fails with "Access Denied"

**Cause**: Insufficient permissions to write to Program Files

**Solution**:
1. Right-click installer and select "Run as administrator"
2. Accept the UAC prompt
3. If still failing, check antivirus software

### Application Can't Save Backups

**Cause**: AppData directory doesn't have write permissions (very rare)

**Solution**:
1. Check if `%APPDATA%\TranscriptFlow\` exists
2. Try creating a file in that directory manually
3. Check folder permissions (should inherit from AppData)
4. Run: `icacls "%APPDATA%\TranscriptFlow" /grant %USERNAME%:(OI)(CI)F`

### Dictionaries Not Downloading

**Cause**: No write access to dicts directory

**Solution**:
1. Verify `%APPDATA%\TranscriptFlow\dicts\` exists
2. Check permissions on the dicts folder
3. Try manually creating a file in that directory
4. Check antivirus/firewall settings

### Application Installed to Wrong Directory

**Cause**: User selected custom directory or architecture mismatch

**Solution**:
1. Uninstall the application
2. Reinstall and use the default directory
3. Ensure you're using the correct installer (x86 vs x64)

## Advanced Configuration

### Installing to Custom Location

If you don't want to install to Program Files:

1. Run the installer
2. When UAC prompt appears, click "No"
3. Choose a custom directory (e.g., `C:\TranscriptFlow\`)
4. Note: File associations may not work without admin rights

### Multiple User Installations

Each Windows user gets their own:
- Settings in their AppData
- Backups in their AppData
- Custom dictionaries

All users share:
- The application executable
- Bundled dictionaries (initial copy)

### Portable Installation

To create a portable version:

1. Copy the contents of `dist\TranscriptFlow\` to a USB drive
2. Copy the `dicts` and `custom_dicts` folders
3. The application will create AppData folders on first run
4. Note: Settings won't roam between computers

## Security Considerations

### Why Admin Rights Are Required

1. **Program Files Protection**: Windows protects this directory to prevent malware
2. **System-Wide Installation**: Benefits all users on the computer
3. **File Associations**: Requires registry modifications
4. **Best Practice**: Separates application code from user data

### What Runs Without Admin Rights

After installation, the application:
- Runs as a standard user (no elevation needed)
- Can read from Program Files
- Can write to AppData
- Cannot modify its own installation
- Cannot access other users' data

### Data Protection

- User data in AppData is protected by Windows user permissions
- Other users cannot access your backups or settings
- Backups are not encrypted (store sensitive data carefully)

## Building the Installer

### Prerequisites

1. Inno Setup 6 (https://jrsoftware.org/isdl.php)
2. PyInstaller (`pip install pyinstaller`)
3. All Python dependencies installed

### Build Commands

```powershell
# Build executable only
python build_installer.py

# Build executable and installer
python build_installer.py --installer

# Clean build directories
python build_installer.py --clean
```

### Manual Build

```powershell
# Step 1: Build executable
pyinstaller TranscriptFlow.spec --clean

# Step 2: Build installer
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" TranscriptFlow.iss
```

### Output

- Executable: `dist\TranscriptFlow\TranscriptFlow.exe`
- Installer: `installer_output\TranscriptFlow_Setup_1.0.0.exe`

## Testing Checklist

Before releasing an installer:

- [ ] Install on clean Windows 10/11 system
- [ ] Verify correct Program Files directory (x86 vs x64)
- [ ] Check AppData directories are created
- [ ] Test dictionary download
- [ ] Test backup creation
- [ ] Test settings persistence
- [ ] Test .tflow file association
- [ ] Test uninstall (keep data)
- [ ] Test uninstall (remove data)
- [ ] Test upgrade from previous version
- [ ] Scan with antivirus software

## Support

For installation issues:

1. Check the installation log: `%TEMP%\Setup Log YYYY-MM-DD #XXX.txt`
2. Verify system requirements
3. Check antivirus/firewall settings
4. Try running installer as administrator
5. Check available disk space

## Version History

### Version 1.0.0
- Initial release
- x64 architecture support
- AppData directory structure
- Automatic dictionary copying
- Permission verification
- Uninstall data preservation option

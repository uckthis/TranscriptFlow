# TranscriptFlow Build Instructions

This document provides step-by-step instructions for building TranscriptFlow into a standalone Windows executable and creating a professional installer.

## Prerequisites

### Required Software

1. **Python 3.8 or higher**
   - Download from: https://www.python.org/downloads/
   - Make sure to check "Add Python to PATH" during installation

2. **PyInstaller**
   ```powershell
   pip install pyinstaller
   ```

3. **Inno Setup 6**
   - Download from: https://jrsoftware.org/isdl.php
   - Install to default location (`C:\Program Files (x86)\Inno Setup 6\`)

4. **All Python Dependencies**
   ```powershell
   pip install -r requirements.txt
   ```

### Verify Installation

Run these commands to verify everything is installed:

```powershell
python --version
pyinstaller --version
pip list | findstr pyenchant
```

## Build Process

### Option 1: Automated Build (Recommended)

Use the provided build script to automate the entire process:

```powershell
# Build executable only
python build_installer.py

# Build executable AND installer
python build_installer.py --installer

# Clean build directories only
python build_installer.py --clean
```

### Option 2: Manual Build

#### Step 1: Clean Previous Builds

```powershell
Remove-Item -Recurse -Force build, dist, installer_output -ErrorAction SilentlyContinue
```

#### Step 2: Build Executable with PyInstaller

```powershell
pyinstaller TranscriptFlow.spec --clean
```

This will create:
- `build/` - Temporary build files (can be deleted)
- `dist/TranscriptFlow/` - The complete application folder

#### Step 3: Test the Executable

Before creating the installer, test the built executable:

```powershell
cd dist\TranscriptFlow
.\TranscriptFlow.exe
```

**Test Checklist:**
- [ ] Application launches without errors
- [ ] Media playback works
- [ ] Spell check is functional (type a misspelled word)
- [ ] Dictionaries are available
- [ ] Backups are created (check `%APPDATA%\TranscriptFlow\BACKUP`)
- [ ] Settings are saved

#### Step 4: Build Installer with Inno Setup

```powershell
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" TranscriptFlow.iss
```

This creates the installer in `installer_output/TranscriptFlow_Setup_1.0.0.exe`

### Installer Features

The installer includes:

- **Automatic Architecture Detection**: Installs to correct Program Files directory (x86 or x64)
- **Admin Permission Handling**: Requests UAC elevation with proper error handling
- **AppData Setup**: Creates user directories with write permissions for:
  - Backups (`%APPDATA%\TranscriptFlow\BACKUP`)
  - Dictionaries (`%APPDATA%\TranscriptFlow\dicts`)
  - Custom dictionaries (`%APPDATA%\TranscriptFlow\custom_dicts`)
  - Waveform cache (`%APPDATA%\TranscriptFlow\waveform_cache`)
  - Configuration (`%APPDATA%\TranscriptFlow\config.json`)
- **Permission Verification**: Tests write access and warns if issues detected
- **Smart File Copying**: Preserves existing user files during upgrades
- **File Association**: Registers `.tflow` file extension
- **Clean Uninstall**: Option to keep or remove user data

For detailed information about the installer, see [INSTALLER_README.md](INSTALLER_README.md).

## Testing the Installer

### Pre-Installation Testing

1. **Check installer size**: Should be approximately 100-200MB
2. **Scan with antivirus**: Ensure no false positives

### Installation Testing

1. **Run the installer** (requires admin privileges)
2. **Verify installation directory**:
   - 64-bit: `C:\Program Files\TranscriptFlow Pro\`
   - 32-bit: `C:\Program Files (x86)\TranscriptFlow Pro\`

3. **Check Start Menu shortcuts**:
   - Start Menu → TranscriptFlow Pro

4. **Verify AppData structure**:
   ```powershell
   explorer %APPDATA%\TranscriptFlow
   ```
   Should contain:
   - `BACKUP/` - Backup files
   - `dicts/` - Dictionary files
   - `custom_dicts/` - Custom dictionaries
   - `waveform_cache/` - Waveform cache
   - `config.json` - Configuration file

### Functional Testing

After installation, test:

1. **Launch application** from Start Menu
2. **Load a media file**
3. **Test spell checking**:
   - Type a misspelled word
   - Right-click to see suggestions
   - Verify red underline appears

4. **Download a new dictionary**:
   - Go to spell check settings
   - Download a language dictionary
   - Verify it appears in `%APPDATA%\TranscriptFlow\dicts`

5. **Create backups**:
   - Make changes to a transcript
   - Wait for auto-backup
   - Verify backup file in `%APPDATA%\TranscriptFlow\BACKUP`

6. **Test .tflow file association**:
   - Save a transcript as `.tflow`
   - Double-click the file
   - Verify it opens in TranscriptFlow

### Uninstallation Testing

1. **Uninstall via Control Panel**
2. **When prompted**, choose to keep or remove user data
3. **Verify**:
   - Program Files directory is removed
   - If "remove data" was selected, `%APPDATA%\TranscriptFlow` is deleted
   - If "keep data" was selected, `%APPDATA%\TranscriptFlow` remains

### Upgrade Testing

1. **Install version 1.0.0**
2. **Create some user data** (transcripts, settings, backups)
3. **Install version 1.0.1** (or newer)
4. **Verify**:
   - User data is preserved
   - Settings are maintained
   - Backups are still accessible

## Troubleshooting

### PyInstaller Issues

**Problem**: `ModuleNotFoundError: No module named 'enchant'`

**Solution**: Ensure pyenchant is installed:
```powershell
pip install pyenchant
```

**Problem**: Missing DLL errors when running the executable

**Solution**: Check that all binaries are collected:
```powershell
# View what's being bundled
pyinstaller TranscriptFlow.spec --clean --log-level DEBUG
```

### Inno Setup Issues

**Problem**: `ISCC.exe` not found

**Solution**: Update the path in `build_installer.py` or install Inno Setup to the default location.

**Problem**: Installer fails to create AppData directories

**Solution**: Run the installer as administrator.

### Spell Check Issues

**Problem**: Spell check not working in built executable

**Solution**: 
1. Verify dictionaries are in `dist/TranscriptFlow/dicts/`
2. Check that `ENCHANT_DATA_PATH` is set correctly
3. Test with:
   ```powershell
   cd dist\TranscriptFlow
   .\TranscriptFlow.exe
   # Check console output for enchant errors
   ```

**Problem**: Dictionaries not copying to AppData

**Solution**: Check `path_manager.py` - the `get_dicts_dir()` function should copy bundled dictionaries on first run.

## Release Checklist

Before releasing a new version:

- [ ] Update version number in `TranscriptFlow.iss` (`#define MyAppVersion`)
- [ ] Test on clean Windows 10/11 system
- [ ] Test on both 32-bit and 64-bit systems (if applicable)
- [ ] Verify all features work in installed version
- [ ] Test upgrade from previous version
- [ ] Scan installer with antivirus
- [ ] Create release notes
- [ ] Tag release in version control
- [ ] Upload installer to distribution platform

## File Structure

After building, your directory should look like this:

```
TranscriptFlow/
├── build/                          # Temporary build files (can delete)
├── dist/
│   └── TranscriptFlow/             # Standalone application
│       ├── TranscriptFlow.exe      # Main executable
│       ├── dicts/                  # Bundled dictionaries
│       ├── custom_dicts/           # Custom dictionaries
│       ├── app_icon.ico            # Application icon
│       ├── mpv-1.dll               # Media player library
│       └── [many other DLLs]       # Python and dependency libraries
├── installer_output/
│   └── TranscriptFlow_Setup_1.0.0.exe  # Windows installer
├── TranscriptFlow.spec             # PyInstaller configuration
├── TranscriptFlow.iss              # Inno Setup configuration
├── build_installer.py              # Build automation script
└── requirements.txt                # Python dependencies
```

## Support

For build issues or questions:
1. Check this documentation
2. Review error logs in `build/` directory
3. Test in development mode first: `python main.py`
4. Verify all dependencies are installed: `pip list`

## Advanced Configuration

### Customizing the Installer

Edit `TranscriptFlow.iss` to customize:
- Installation directory
- Start Menu folder name
- Desktop icon creation
- File associations
- Uninstaller behavior

### Optimizing Build Size

To reduce the executable size:

1. **Exclude unused modules** in `TranscriptFlow.spec`:
   ```python
   excludes=[
       'matplotlib',
       'scipy',
       'pandas',
       'PIL',
       'tkinter',
       # Add more unused modules
   ],
   ```

2. **Use UPX compression** (already enabled in spec file)

3. **Remove debug symbols**:
   ```python
   debug=False,
   strip=True,
   ```

### Code Signing (Optional)

To sign the executable and installer:

1. Obtain a code signing certificate
2. Use `signtool.exe` from Windows SDK:
   ```powershell
   signtool sign /f certificate.pfx /p password /t http://timestamp.digicert.com dist\TranscriptFlow\TranscriptFlow.exe
   signtool sign /f certificate.pfx /p password /t http://timestamp.digicert.com installer_output\TranscriptFlow_Setup_1.0.0.exe
   ```

This prevents Windows SmartScreen warnings and builds user trust.

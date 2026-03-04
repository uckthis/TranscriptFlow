"""
Path Manager for TranscriptFlow

Handles path resolution for both development and frozen (PyInstaller) environments.
Ensures proper directory structure for user data in AppData.
"""

import os
import sys
import shutil


def is_frozen():
    """Check if running as a PyInstaller bundle"""
    return getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')


def get_base_path():
    """Get the base application directory"""
    if is_frozen():
        # Running as compiled executable
        return os.path.dirname(sys.executable)
    else:
        # Running as script
        return os.path.dirname(os.path.abspath(__file__))


def get_appdata_dir():
    """Get the AppData directory for user-writable data"""
    if sys.platform == 'win32':
        appdata = os.getenv('APPDATA')
        if not appdata:
            appdata = os.path.expanduser('~\\AppData\\Roaming')
    else:
        # macOS/Linux
        appdata = os.path.expanduser('~/.config')
    
    app_dir = os.path.join(appdata, 'TranscriptFlow')
    
    # Create if it doesn't exist
    if not os.path.exists(app_dir):
        os.makedirs(app_dir, exist_ok=True)
    
    return app_dir


def get_dicts_dir():
    """Get the dictionaries directory (writable location in AppData)"""
    dicts_dir = os.path.join(get_appdata_dir(), 'dicts')
    
    # Create if it doesn't exist
    if not os.path.exists(dicts_dir):
        os.makedirs(dicts_dir, exist_ok=True)
        
        # Copy bundled dictionaries on first run
        bundled_dicts = os.path.join(get_base_path(), 'dicts')
        if os.path.exists(bundled_dicts):
            try:
                for item in os.listdir(bundled_dicts):
                    src = os.path.join(bundled_dicts, item)
                    dst = os.path.join(dicts_dir, item)
                    if os.path.isfile(src) and not os.path.exists(dst):
                        shutil.copy2(src, dst)
            except Exception as e:
                print(f"Warning: Could not copy bundled dictionaries: {e}")
    
    return dicts_dir


def get_custom_dicts_dir():
    """Get the custom dictionaries directory"""
    custom_dicts_dir = os.path.join(get_appdata_dir(), 'custom_dicts')
    
    if not os.path.exists(custom_dicts_dir):
        os.makedirs(custom_dicts_dir, exist_ok=True)
        
        # Copy bundled custom dictionaries on first run
        bundled_custom = os.path.join(get_base_path(), 'custom_dicts')
        if os.path.exists(bundled_custom):
            try:
                for item in os.listdir(bundled_custom):
                    src = os.path.join(bundled_custom, item)
                    dst = os.path.join(custom_dicts_dir, item)
                    if os.path.isfile(src) and not os.path.exists(dst):
                        shutil.copy2(src, dst)
            except Exception as e:
                print(f"Warning: Could not copy bundled custom dictionaries: {e}")
    
    return custom_dicts_dir


def get_backup_dir():
    """Get the backup directory"""
    backup_dir = os.path.join(get_appdata_dir(), 'BACKUP')
    
    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir, exist_ok=True)
    
    return backup_dir


def get_waveform_cache_dir():
    """Get the waveform cache directory"""
    cache_dir = os.path.join(get_appdata_dir(), 'waveform_cache')
    
    if not os.path.exists(cache_dir):
        os.makedirs(cache_dir, exist_ok=True)
    
    return cache_dir


def get_logs_dir():
    """Get the directory where logs are stored"""
    logs_dir = os.path.join(get_appdata_dir(), 'logs')
    
    if not os.path.exists(logs_dir):
        os.makedirs(logs_dir, exist_ok=True)
    
    return logs_dir


def get_config_path():
    """Get the configuration file path"""
    return os.path.join(get_appdata_dir(), 'config.json')


def get_bin_dir():
    """Get the directory for external binaries (e.g. Tesseract)"""
    bin_dir = os.path.join(get_appdata_dir(), 'bin')
    if not os.path.exists(bin_dir):
        os.makedirs(bin_dir, exist_ok=True)
    return bin_dir


def get_tesseract_dir():
    """Get the directory where Tesseract is installed"""
    tess_dir = os.path.join(get_bin_dir(), 'tesseract')
    if not os.path.exists(tess_dir):
        os.makedirs(tess_dir, exist_ok=True)
    return tess_dir


def get_tessdata_dir():
    """Get the directory for Tesseract language data"""
    tessdata_dir = os.path.join(get_tesseract_dir(), 'tessdata')
    if not os.path.exists(tessdata_dir):
        os.makedirs(tessdata_dir, exist_ok=True)
    return tessdata_dir


def get_tesseract_exe():
    """Get the path to the Tesseract executable, checking multiple locations"""
    # 1. Check app-specific bin directory (AppData)
    app_tess = os.path.join(get_tesseract_dir(), 'tesseract.exe')
    if os.path.exists(app_tess):
        return app_tess
        
    # 2. Check standard system paths
    system_paths = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
    ]
    for p in system_paths:
        if os.path.exists(p):
            return p
            
    # 3. Fallback to just 'tesseract' and hope it's on PATH
    import shutil
    on_path = shutil.which("tesseract")
    if on_path:
        return on_path
        
    return app_tess # Default to expected app path if not found


def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    if is_frozen():
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    else:
        base_path = get_base_path()
    
    return os.path.join(base_path, relative_path)


def initialize_app_directories():
    """Initialize all application directories on startup"""
    # Create all necessary directories
    get_appdata_dir()
    get_dicts_dir()
    get_custom_dicts_dir()
    get_backup_dir()
    get_waveform_cache_dir()
    get_logs_dir()
    get_bin_dir()
    get_tesseract_dir()
    get_tessdata_dir()
    
    # Copy bundled config.json if it exists and AppData config doesn't
    appdata_config = get_config_path()
    if not os.path.exists(appdata_config):
        bundled_config = os.path.join(get_base_path(), 'config.json')
        if os.path.exists(bundled_config):
            try:
                shutil.copy2(bundled_config, appdata_config)
            except:
                pass
    
    # Do not print directly here as it can crash windowed apps if stdout is None
    # main.py will log these paths instead

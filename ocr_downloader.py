import os
import requests
import zipfile
import logging
import subprocess
from PyQt6.QtCore import QThread, pyqtSignal
from path_manager import get_tesseract_dir, get_tessdata_dir, get_bin_dir

logger = logging.getLogger('TranscriptFlow.OCRDownloader')

# Official UB Mannheim installers are the standard for Windows.
TESSERACT_INSTALLER_URL = "https://github.com/tesseract-ocr/tesseract/releases/download/5.5.0/tesseract-ocr-w64-setup-5.5.0.20241111.exe"
WINGET_CMD = ["winget", "install", "Tesseract-OCR.Tesseract-OCR", "--silent", "--accept-package-agreements", "--accept-source-agreements"]
LANG_DATA_BASE_URL = "https://github.com/tesseract-ocr/tessdata_fast/raw/main/"

class DownloadThread(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str)

    def __init__(self, url, dest_path, is_zip=False):
        super().__init__()
        self.url = url
        self.dest_path = dest_path
        self.is_zip = is_zip

    def run(self):
        try:
            logger.info(f"Downloading from {self.url} to {self.dest_path}")
            response = requests.get(self.url, stream=True, timeout=30)
            response.raise_for_status()
            
            total_size = int(response.headers.get('content-length', 0))
            downloaded_size = 0
            
            with open(self.dest_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded_size += len(chunk)
                        if total_size > 0:
                            percent = int((downloaded_size / total_size) * 100)
                            self.progress.emit(percent, f"Downloading... {percent}%")
            
            if self.is_zip:
                self.progress.emit(99, "Extracting files...")
                extract_to = os.path.dirname(self.dest_path)
                with zipfile.ZipFile(self.dest_path, 'r') as zip_ref:
                    zip_ref.extractall(extract_to)
                os.remove(self.dest_path) # Clean up zip
            
            self.finished.emit(True, "Download successful.")
        except Exception as e:
            logger.error(f"Download failed: {e}")
            self.finished.emit(False, str(e))

class TesseractInstallThread(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str)

    def run(self):
        try:
            # Step 1: Try Winget (Fastest)
            self.progress.emit(10, "Attempting install via Winget (Silent)...")
            logger.info("Running winget install...")
            try:
                # Use shell=True for winget as it's often a shell alias/wrapper on some systems
                result = subprocess.run(WINGET_CMD, capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
                if result.returncode == 0:
                    self.finished.emit(True, "Tesseract installed successfully via Winget.")
                    return
                logger.warning(f"Winget returned non-zero code: {result.returncode}")
            except FileNotFoundError:
                logger.warning("Winget command not found. Falling back to manual download.")
            except Exception as e:
                logger.warning(f"Winget error: {e}. Falling back to manual download.")

            # Step 2: Fallback to downloading installer and running /S
            self.progress.emit(20, "Downloading official installer (approx. 40MB)...")
            dest_path = os.path.join(get_bin_dir(), "tesseract_installer.exe")
            
            response = requests.get(TESSERACT_INSTALLER_URL, stream=True, timeout=30)
            response.raise_for_status()
            total_size = int(response.headers.get('content-length', 0))
            download_size = 0
            
            with open(dest_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        download_size += len(chunk)
                        if total_size > 0:
                            p = int((download_size / total_size) * 60) + 20
                            self.progress.emit(p, f"Downloading installer... {int((download_size/total_size)*100)}%")
            
            self.progress.emit(85, "Launching installer... please follow UAC prompts.")
            
            # WinError 740 Fix: Explicitly request elevation via PowerShell
            # This triggers the standard Windows UAC prompt.
            install_cmd = f'Start-Process "{dest_path}" -ArgumentList "/S" -Verb RunAs -Wait'
            logger.info(f"Executing: {install_cmd}")
            
            # We use subprocess.run with powershell to wait for completion
            proc = subprocess.run(["powershell", "-Command", install_cmd], 
                                capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
            
            if proc.returncode == 0:
                self.finished.emit(True, "Tesseract installed successfully via Installer.")
            else:
                self.finished.emit(False, f"Installer failed or cancelled (Code {proc.returncode})")
            
        except Exception as e:
            logger.error(f"Installation failed: {e}", exc_info=True)
            self.finished.emit(False, str(e))

def download_tesseract(on_finished, on_progress):
    """
    Automated one-click installation.
    """
    thread = TesseractInstallThread()
    thread.progress.connect(on_progress)
    thread.finished.connect(on_finished)
    thread.start()
    return thread

def download_language(lang_code, on_finished, on_progress):
    """
    Language data is still reliably hosted on GitHub.
    """
    url = f"{LANG_DATA_BASE_URL}{lang_code}.traineddata"
    dest_path = os.path.join(get_tessdata_dir(), f"{lang_code}.traineddata")
    
    thread = DownloadThread(url, dest_path)
    thread.progress.connect(on_progress)
    thread.finished.connect(on_finished)
    thread.start()
    return thread

import logging
import asyncio
import os
import subprocess
from PIL import Image
import io

# Optional WinRT imports
try:
    from winrt.windows.media.ocr import OcrEngine
    from winrt.windows.graphics.imaging import BitmapDecoder
    from winrt.windows.storage.streams import InMemoryRandomAccessStream, DataWriter
    WINRT_AVAILABLE = True
except ImportError:
    WINRT_AVAILABLE = False

# Optional Tesseract import
try:
    import pytesseract
    TESSERACT_AVAILABLE = True
except ImportError:
    TESSERACT_AVAILABLE = False

from path_manager import get_tesseract_exe, get_tessdata_dir

logger = logging.getLogger('TranscriptFlow.OCREngine')

# Comprehensive mapping of Tesseract codes to human-readable names
TESS_LANG_MAP = {
    "afr": "Afrikaans", "amh": "Amharic", "ara": "Arabic", "asm": "Assamese",
    "aze": "Azerbaijani", "aze_cyrl": "Azerbaijani - Cyrillic", "bel": "Belarusian",
    "ben": "Bengali", "bod": "Tibetan", "bos": "Bosnian", "bre": "Breton",
    "bul": "Bulgarian", "cat": "Catalan", "ceb": "Cebuano", "ces": "Czech",
    "chi_sim": "Chinese - Simplified", "chi_tra": "Chinese - Traditional",
    "chr": "Cherokee", "cos": "Corsican", "cym": "Welsh", "dan": "Danish",
    "deu": "German", "dzo": "Dzongkha", "ell": "Greek", "eng": "English",
    "enm": "English - Middle", "epo": "Esperanto", "est": "Estonian",
    "eus": "Basque", "fas": "Persian", "fin": "Finnish", "fra": "French",
    "frm": "French - Middle", "fry": "Western Frisian", "gla": "Scottish Gaelic",
    "gle": "Irish", "glg": "Galician", "guj": "Gujarati", "hat": "Haitian",
    "heb": "Hebrew", "hin": "Hindi", "hrv": "Croatian", "hun": "Hungarian",
    "iku": "Inuktitut", "ind": "Indonesian", "isl": "Icelandic", "ita": "Italian",
    "jav": "Javanese", "jpn": "Japanese", "kan": "Kannada", "kat": "Georgian",
    "kaz": "Kazakh", "khm": "Central Khmer", "kir": "Kyrgyz", "kor": "Korean",
    "kur": "Kurdish", "lao": "Lao", "lat": "Latin", "lav": "Latvian",
    "lit": "Lithuanian", "mal": "Malayalam", "mar": "Marathi", "mkd": "Macedonian",
    "mlt": "Maltese", "mon": "Mongolian", "mya": "Burmese", "nep": "Nepali",
    "nld": "Dutch", "nor": "Norwegian", "ori": "Oriya", "pan": "Punjabi",
    "pol": "Polish", "por": "Portuguese", "pus": "Pashto", "que": "Quechua",
    "ron": "Romanian", "rus": "Russian", "san": "Sanskrit", "sin": "Sinhala",
    "slk": "Slovak", "slv": "Slovenian", "spa": "Spanish", "srp": "Serbian",
    "srp_latn": "Serbian - Latin", "sun": "Sundanese", "swa": "Swahili",
    "swe": "Swedish", "syr": "Syriac", "tam": "Tamil", "tat": "Tatar",
    "tel": "Telugu", "tgk": "Tajik", "tha": "Thai", "tir": "Tigrinya",
    "ton": "Tongan", "tur": "Turkish", "uig": "Uyghur", "ukr": "Ukrainian",
    "urd": "Urdu", "uzb": "Uzbek", "uzb_cyrl": "Uzbek - Cyrillic",
    "vie": "Vietnamese", "yid": "Yiddish", "yor": "Yoruba",
    "osd": "Orientation & Script Detection"
}

def get_lang_name(code):
    return TESS_LANG_MAP.get(code, code)

def get_lang_code(name):
    # Reverse lookup
    for code, n in TESS_LANG_MAP.items():
        if n == name:
            return code
    return name

class OCRAbstractDriver:
    async def perform_ocr(self, image_bytes, settings):
        raise NotImplementedError

class WindowsNativeDriver(OCRAbstractDriver):
    async def perform_ocr(self, image_bytes, settings):
        if not WINRT_AVAILABLE:
            logger.error("WinRT modules not available for Windows Native OCR.")
            return None
            
        try:
            logger.debug("Starting Windows Native OCR...")
            stream = InMemoryRandomAccessStream()
            writer = DataWriter(stream.get_output_stream_at(0))
            writer.write_bytes(image_bytes) 
            await writer.store_async()
            await writer.flush_async()
            
            decoder = await BitmapDecoder.create_async(stream)
            software_bitmap = await decoder.get_software_bitmap_async()
            
            engine = OcrEngine.try_create_from_user_profile_languages()
            if not engine:
                logger.error("Could not create Windows OCR engine.")
                return None
                
            result = await engine.recognize_async(software_bitmap)
            if not result or not result.lines:
                return None
                
            return result.text.strip()
        except Exception as e:
            logger.error(f"Windows Native OCR Error: {e}")
            return None

class TesseractDriver(OCRAbstractDriver):
    def is_installed(self):
        exe_path = get_tesseract_exe()
        return os.path.exists(exe_path)

    async def perform_ocr(self, image_bytes, settings):
        if not TESSERACT_AVAILABLE:
            logger.error("pytesseract not installed.")
            return "Error: pytesseract not installed."
            
        exe_path = get_tesseract_exe()
        if not os.path.exists(exe_path):
            logger.error(f"Tesseract executable not found at {exe_path}")
            return "Error: Tesseract not installed. Please download it in Options."

        try:
            logger.debug("Starting Tesseract OCR...")
            pytesseract.pytesseract.tesseract_cmd = exe_path
            
            # Load image from bytes
            image = Image.open(io.BytesIO(image_bytes))
            
            lang = settings.get('tesseract_lang', 'eng')
            
            # Check if we should use the app-specific tessdata path
            custom_tessdata = get_tessdata_dir()
            config = ""
            if os.path.exists(os.path.join(custom_tessdata, f"{lang}.traineddata")):
                config = f'--tessdata-dir "{custom_tessdata}"'

            # Run Tesseract in a thread
            loop = asyncio.get_event_loop()
            text = await loop.run_in_executor(None, 
                lambda: pytesseract.image_to_string(image, lang=lang, config=config))
            
            return text.strip()
        except Exception as e:
            logger.error(f"Tesseract OCR Error: {e}")
            return f"Error: {str(e)}"

class OCREngineManager:
    def __init__(self):
        self.drivers = {
            'windows': WindowsNativeDriver(),
            'tesseract': TesseractDriver()
        }

    async def perform_ocr(self, image_bytes, settings):
        engine_type = settings.get('engine', 'windows')
        driver = self.drivers.get(engine_type)
        
        if not driver:
            logger.error(f"Unknown OCR engine type: {engine_type}")
            return None
            
        text = await driver.perform_ocr(image_bytes, settings)
        if not text:
            return None
            
        # Common Post-Processing
        # 1. Handle Case Conversion
        case = settings.get('case_conversion', 'none')
        if case == 'upper':
            text = text.upper()
        elif case == 'lower':
            text = text.lower()
        elif case == 'title':
            text = text.title()
            
        # 2. Add Prefix/Suffix
        prefix = settings.get('prefix', '')
        suffix = settings.get('suffix', '')
        
        final_text = f"{prefix}{text}{suffix}"
        return final_text

# Global manager instance
manager = OCREngineManager()

def perform_ocr(image_bytes, settings):
    """
    Main entry point for OCR, dispatches to selected engine.
    """
    try:
        logger.info(f"Running perform_ocr with engine: {settings.get('engine', 'windows')}")
        return asyncio.run(manager.perform_ocr(image_bytes, settings))
    except Exception as e:
        logger.error(f"OCR Manager Error: {e}", exc_info=True)
        return None

# Helper for UI to check Tesseract status
def is_tesseract_installed():
    return TesseractDriver().is_installed()

def get_tesseract_version():
    """
    Returns the version string of the installed Tesseract engine.
    """
    if not TESSERACT_AVAILABLE:
        return "Not Available (pytesseract missing)"
        
    exe_path = get_tesseract_exe()
    if not os.path.exists(exe_path):
        return "Not Found"
        
    try:
        pytesseract.pytesseract.tesseract_cmd = exe_path
        version = pytesseract.get_tesseract_version()
        return str(version)
    except Exception as e:
        logger.error(f"Error getting Tesseract version: {e}")
        return "Unknown"

def get_installed_tesseract_langs():
    langs = set()
    
    # 1. Check App-specific tessdata
    tessdata_app = get_tessdata_dir()
    if os.path.exists(tessdata_app):
        langs.update([f.replace('.traineddata', '') for f in os.listdir(tessdata_app) if f.endswith('.traineddata')])
    
    # 2. Check System-wide tessdata (if found)
    exe_path = get_tesseract_exe()
    if exe_path and os.path.exists(exe_path):
        tessdata_sys = os.path.join(os.path.dirname(exe_path), 'tessdata')
        if os.path.exists(tessdata_sys):
            langs.update([f.replace('.traineddata', '') for f in os.listdir(tessdata_sys) if f.endswith('.traineddata')])
            
    return sorted(list(langs))

import logging
import asyncio
from winrt.windows.media.ocr import OcrEngine
from winrt.windows.graphics.imaging import BitmapDecoder
from winrt.windows.storage.streams import InMemoryRandomAccessStream, DataWriter

logger = logging.getLogger('TranscriptFlow.OCREngine')

async def _perform_ocr_async(image_bytes, settings):
    """Internal async OCR implementation using native winrt modules."""
    try:
        logger.debug("Starting OCR async process...")
        # Create a stream from bytes
        stream = InMemoryRandomAccessStream()
        writer = DataWriter(stream.get_output_stream_at(0))
        writer.write_bytes(image_bytes) 
        logger.debug("Writing bytes to stream...")
        await writer.store_async()
        await writer.flush_async()
        logger.debug("Bytes stored and flushed.")
        
        # Decode the image
        logger.debug("Decoding image...")
        decoder = await BitmapDecoder.create_async(stream)
        software_bitmap = await decoder.get_software_bitmap_async()
        logger.debug("Image decoded to SoftwareBitmap.")
        
        # Use OCR Engine
        logger.debug("Initializing OCR Engine...")
        engine = OcrEngine.try_create_from_user_profile_languages()
        if not engine:
            logger.error("Could not create OCR engine.")
            return None
            
        logger.debug("Performing recognition...")
        result = await engine.recognize_async(software_bitmap)
        
        if not result or not result.lines:
            logger.debug("No text detected.")
            return None
            
        text = result.text.strip()
        logger.debug(f"OCR result: {text[:50]}...")
        
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
        
    except Exception as e:
        logger.error(f"OCR Error: {e}", exc_info=True)
        return None

def perform_ocr(image_bytes, settings):
    """
    Performs OCR on image bytes using Windows Native WinRT OCR.
    """
    try:
        logger.info("Running perform_ocr...")
        return asyncio.run(_perform_ocr_async(image_bytes, settings))
    except Exception as e:
        logger.error(f"Async Run Error: {e}", exc_info=True)
        return None

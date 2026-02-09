import vlc
import sys
import time
from PyQt6.QtCore import QObject, pyqtSignal, QTimer, QDateTime, Qt
import urllib.request
import os
import ctypes
import ctypes.util
import struct
import zipfile
import io
import logging

logger = logging.getLogger('TranscriptFlow.MediaEngine')

class MediaEngineBackend(QObject):
    positionChanged = pyqtSignal(int)
    durationChanged = pyqtSignal(int)
    
    def __init__(self):
        super().__init__()

    def load(self, source_type, path): pass
    def play_pause(self): pass
    def play(self): pass
    def pause(self): pass
    def stop(self): pass
    def seek(self, ms): pass
    def set_rate(self, rate): pass
    def set_volume(self, vol): pass
    def set_display_handle(self, win_id): pass
    def get_audio_tracks(self): return []
    def get_subtitle_tracks(self): return []
    def set_audio_track(self, tid): pass
    def set_subtitle_track(self, tid): pass
    def get_time(self): return 0
    def is_playing(self): return False
    def get_duration(self): return 0
    def frame_step(self, forward=True): pass
    def release(self): pass
    def get_backend_type(self): return "generic"

class VLCBackend(MediaEngineBackend):
    def __init__(self, instance):
        super().__init__()
        self.instance = instance
        self.player = self.instance.media_player_new()
        self.events = self.player.event_manager()
        self.events.event_attach(vlc.EventType.MediaPlayerTimeChanged, self._on_vlc_time)
        self.events.event_attach(vlc.EventType.MediaPlayerLengthChanged, self._on_vlc_length)
        self.released = False

    def load(self, source_type, path):
        media = self.instance.media_new(path)
        self.player.set_media(media)
        self.player.play()
        # Safe timer: only pause if not released and player exists
        QTimer.singleShot(100, lambda: self.player.pause() if (not self.released and self.player) else None)

    def play_pause(self):
        if self.player.is_playing(): self.player.pause()
        else: self.player.play()

    def play(self): self.player.play()
    def pause(self): self.player.set_pause(1)
    def stop(self): self.player.stop()
    def release(self):
        self.released = True
        # Explicitly stop first
        if self.player:
            try: self.player.stop()
            except: pass
            
        if self.events:
            try:
                self.events.event_detach(vlc.EventType.MediaPlayerTimeChanged)
                self.events.event_detach(vlc.EventType.MediaPlayerLengthChanged)
            except: pass
        
        if self.player:
            # Reset window handle to avoid collision with new engine
            # 0 or None might be needed here to fully unbind from the winId
            if sys.platform == "win32": 
                try: self.player.set_hwnd(0)
                except: pass
            
            try: self.player.release()
            except: pass
            
        self.player = None
        self.events = None
        # VLC instance is shared, we don't release it here
    def seek(self, ms): self.player.set_time(int(ms))
    def set_rate(self, rate): self.player.set_rate(rate)
    def set_volume(self, vol): self.player.audio_set_volume(int(vol))
    def get_audio_tracks(self):
        # Returns list of (id, name)
        tracks = self.player.audio_get_track_description()
        return [(t[0], t[1].decode('utf-8', 'ignore')) for t in tracks] if tracks else []
    def get_subtitle_tracks(self):
        tracks = self.player.video_get_spu_description()
        return [(t[0], t[1].decode('utf-8', 'ignore')) for t in tracks] if tracks else []
    def set_audio_track(self, track_id): self.player.audio_set_track(int(track_id))
    def set_subtitle_track(self, track_id): self.player.video_set_spu(int(track_id))
    def set_display_handle(self, win_id):
        if sys.platform == "win32": self.player.set_hwnd(win_id)
        elif sys.platform == "darwin": self.player.set_nsobject(win_id)
        else: self.player.set_xwindow(win_id)

    def set_aspect_ratio(self, ratio):
        if self.player:
            self.player.video_set_aspect_ratio(ratio.encode() if ratio else None)
    def set_pitch_lock(self, enabled):
        if self.player:
            # VLC uses play_resampling_mode for this. 0 is normal, 1 is no resampling (shifts pitch if speed changes)
            # Actually, the standard '--audio-time-stretch' option is what handles pitch lock in many versions.
            # We try to use the most common method available via ctypes/libvlc.
            try:
                # 0 = No pitch correction (pitch shifts with speed)
                # 1 = Pitch correction enabled (pitch stays same)
                self.player.audio_set_play_resampling_mode(0 if enabled else 1)
            except: pass
    def get_time(self): return self.player.get_time()
    def is_playing(self): return self.player.is_playing()
    def get_duration(self): return self.player.get_length()

    def frame_step(self, forward=True):
        if forward:
            self.player.next_frame()
        else:
            # VLC doesn't have a direct "previous frame"
            # We calculate step from FPS (defaulting to 25 if unknown)
            fps = self.player.get_fps()
            if fps <= 0: fps = 25.0
            step_ms = int(1000 / fps)
            curr = self.player.get_time()
            self.player.set_time(max(0, curr - step_ms))

    def _on_vlc_time(self, event):
        if not self.released:
            self.positionChanged.emit(self.player.get_time())
    def _on_vlc_length(self, event):
        if not self.released:
            self.durationChanged.emit(self.player.get_length())
    def get_backend_type(self): return "vlc"

class MPVBackend(MediaEngineBackend):
    def __init__(self, lib_path=None, win_id=None):
        super().__init__()
        self.mpv_available = False
        self.lib_path = lib_path
        self._pending_wid = win_id
        self.handle = None
        self._lib = None
        
        # 1. Locate DLL
        if not self.lib_path:
            detected = self.discover_dlls()
            if detected: self.lib_path = detected[0]['path']
            
        if self.lib_path:
            logger.info(f"MPV: Attempting to load library from {self.lib_path}")
            self._try_load(self.lib_path)
        
        if self.mpv_available:
            logger.info("MPV: Library loaded successfully")
            self._init_mpv()
        else:
            logger.error("MPV: Library could not be loaded")
            
    def _try_load(self, path):
        try:
            # Safe PATH handling for DLL loading
            app_dir = os.path.dirname(os.path.abspath(__file__))
            if app_dir not in os.environ.get("PATH", ""):
                os.environ["PATH"] = app_dir + os.pathsep + os.environ.get("PATH", "")

            self._lib = ctypes.CDLL(path)
            self.mpv_available = True
            
            # Signatures
            self._lib.mpv_create.restype = ctypes.c_void_p
            self._lib.mpv_initialize.argtypes = [ctypes.c_void_p]
            self._lib.mpv_initialize.restype = ctypes.c_int
            self._lib.mpv_terminate_destroy.argtypes = [ctypes.c_void_p]
            self._lib.mpv_set_option_string.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_char_p]
            self._lib.mpv_set_option_string.restype = ctypes.c_int
            self._lib.mpv_command_string.argtypes = [ctypes.c_void_p, ctypes.c_char_p]
            self._lib.mpv_command_string.restype = ctypes.c_int
            self._lib.mpv_get_property_string.argtypes = [ctypes.c_void_p, ctypes.c_char_p]
            self._lib.mpv_get_property_string.restype = ctypes.c_void_p
            self._lib.mpv_set_property_string.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_char_p]
            self._lib.mpv_set_property_string.restype = ctypes.c_int
            self._lib.mpv_free.argtypes = [ctypes.c_void_p]
        except:
            self.mpv_available = False

    def _init_mpv(self):
        self.handle = self._lib.mpv_create()
        if self.handle:
            # Options BEFORE initialize
            self._lib.mpv_set_option_string(self.handle, b"keep-open", b"yes")
            self._lib.mpv_set_option_string(self.handle, b"input-default-bindings", b"no")
            self._lib.mpv_set_option_string(self.handle, b"hwdec", b"auto")
            self._lib.mpv_set_option_string(self.handle, b"osc", b"no")
            self._lib.mpv_set_option_string(self.handle, b"osd-bar", b"no")
            self._lib.mpv_set_option_string(self.handle, b"osd-level", b"0")
            
            # Critical: Set WID as OPTION before initialization for rendering stability
            if self._pending_wid:
                self._lib.mpv_set_option_string(self.handle, b"wid", str(int(self._pending_wid)).encode())

            res = self._lib.mpv_initialize(self.handle)
            
            # Start status poller
            self.timer = QTimer()
            self.timer.timeout.connect(self._poll_status)
            self.timer.start(100)
            self.last_pos = -1

    def _poll_status(self):
        if not self.handle: return
        try:
            # Position
            ptr = self._lib.mpv_get_property_string(self.handle, b"time-pos")
            if ptr:
                try:
                    res = ctypes.string_at(ptr).decode()
                    ms = int(float(res) * 1000)
                    if ms != self.last_pos:
                        self.positionChanged.emit(ms)
                        self.last_pos = ms
                finally: self._lib.mpv_free(ptr)
                
            # Duration
            ptr = self._lib.mpv_get_property_string(self.handle, b"duration")
            if ptr:
                try:
                    res = ctypes.string_at(ptr).decode()
                    ms = int(float(res) * 1000)
                    self.durationChanged.emit(ms)
                finally: self._lib.mpv_free(ptr)
        except Exception as e:
            if hasattr(self, 'handle') and self.handle:
                logging.getLogger('TranscriptFlow').debug(f"MPV Poll Error: {e}")

    def load(self, source_type, path):
        if not self.handle: return
        path_clean = path.replace('\\', '/')
        cmd = f"loadfile \"{path_clean}\"".encode('utf-8')
        self._lib.mpv_command_string(self.handle, cmd)

    def play_pause(self):
        if not self.handle: return
        ptr = self._lib.mpv_get_property_string(self.handle, b"pause")
        if ptr:
            try:
                is_paused = ctypes.string_at(ptr).decode() == "yes"
                self.pause() if not is_paused else self.play()
            finally: self._lib.mpv_free(ptr)

    def play(self):
        if self.handle: 
            logger.info("MPV: Play requested")
            self._lib.mpv_set_property_string(self.handle, b"pause", b"no")

    def pause(self):
        if self.handle: self._lib.mpv_set_property_string(self.handle, b"pause", b"yes")

    def stop(self):
        if self.handle: self._lib.mpv_command_string(self.handle, b"stop")

    def seek(self, ms):
        if self.handle:
            cmd = f"seek {ms/1000.0} absolute".encode()
            self._lib.mpv_command_string(self.handle, cmd)

    def set_rate(self, rate):
        if self.handle: self._lib.mpv_set_property_string(self.handle, b"speed", str(rate).encode())

    def set_volume(self, vol):
        if self.handle: self._lib.mpv_set_property_string(self.handle, b"volume", str(vol).encode())

    def set_pitch_lock(self, enabled):
        if self.handle:
            val = b"yes" if enabled else b"no"
            self._lib.mpv_set_property_string(self.handle, b"audio-pitch-correction", val)

    def set_aspect_ratio(self, ratio):
        if self.handle:
            # MPV video-aspect-override needs to be a string or -1 for reset
            # Value can be "16:9", "4:3", etc.
            val = ratio.encode() if ratio and ratio != "default" else b"-1"
            self._lib.mpv_set_property_string(self.handle, b"video-aspect-override", val)

    def get_audio_tracks(self):
        # Full track list extraction via ctypes is complex; returning empty for now
        return []

    def get_subtitle_tracks(self):
        return []

    def set_audio_track(self, tid):
        if self.handle:
            self._lib.mpv_set_property_string(self.handle, b"aid", str(tid).encode())

    def set_subtitle_track(self, tid):
        if self.handle:
            self._lib.mpv_set_property_string(self.handle, b"sid", str(tid).encode())

    def set_display_handle(self, win_id):
        self._pending_wid = win_id
        if self.handle:
            # Update property if already initialized
            try: self._lib.mpv_set_property_string(self.handle, b"wid", str(int(win_id)).encode())
            except: pass

    def get_time(self):
        if not self.handle: return 0
        ptr = self._lib.mpv_get_property_string(self.handle, b"time-pos")
        if ptr:
            try: 
                res = ctypes.string_at(ptr).decode()
                return int(float(res) * 1000)
            except: return 0
            finally: self._lib.mpv_free(ptr)
        return 0

    def is_playing(self):
        if not self.handle: return False
        
        # Check if we actually have a file loaded; if not pos, we're effectively idle
        pos_ptr = self._lib.mpv_get_property_string(self.handle, b"time-pos")
        if not pos_ptr:
            return False
        self._lib.mpv_free(pos_ptr)

        ptr = self._lib.mpv_get_property_string(self.handle, b"pause")
        if ptr:
            try: 
                return ctypes.string_at(ptr).decode() == "no"
            finally: self._lib.mpv_free(ptr)
        return False

    def get_duration(self):
        if not self.handle: return 0
        ptr = self._lib.mpv_get_property_string(self.handle, b"duration")
        if ptr:
            try: 
                res = ctypes.string_at(ptr).decode()
                return int(float(res) * 1000)
            except: return 0
            finally: self._lib.mpv_free(ptr)
        return 0

    def frame_step(self, forward=True):
        if self.handle:
            cmd = b"frame-step" if forward else b"frame-back-step"
            self._lib.mpv_command_string(self.handle, cmd)

    def release(self):
        if hasattr(self, 'timer'):
            try: self.timer.stop()
            except: pass
        if self.handle:
            try: self._lib.mpv_terminate_destroy(self.handle)
            except: pass
            self.handle = None
        self._lib = None

    def get_backend_type(self): return "mpv"

    @staticmethod
    def discover_dlls():
        """Returns a list of dicts with name and full path for all detected MPV DLLs"""
        # Search locations: 1. Current dir, 2. App Root (frozen), 3. _internal (frozen), 4. System PATH
        search_dirs = [os.path.dirname(os.path.abspath(__file__))]
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
            if base_dir not in search_dirs:
                search_dirs.append(base_dir)
            internal_dir = os.path.join(base_dir, '_internal')
            if internal_dir not in search_dirs:
                search_dirs.append(internal_dir)
            
        dll_names = ["libmpv-2.dll", "mpv-2.dll", "mpv-1.dll", "mpv.dll"]
        found = []
        
        for d in search_dirs:
            for name in dll_names:
                p = os.path.join(d, name)
                if os.path.exists(p):
                    try:
                        ctypes.CDLL(p)
                        found.append({'name': name, 'path': p})
                        logger.info(f"MPV: Found DLL at {p}")
                    except Exception as e:
                        logger.warning(f"MPV: DLL found at {p} but could not be loaded: {e}")
                        continue
        
        if not found:
            logger.info("MPV: No local DLLs found, searching system PATH")
            for name in dll_names:
                p = ctypes.util.find_library(name)
                if p:
                    try:
                        ctypes.CDLL(p)
                        found.append({'name': f"System {name}", 'path': p})
                        logger.info(f"MPV: Found system DLL at {p}")
                    except: continue
        return found

class MediaEngine(QObject):
    positionChanged = pyqtSignal(int)
    durationChanged = pyqtSignal(int)
    
    def __init__(self):
        super().__init__()
        self.vlc_instance = vlc.Instance('--avcodec-hw=none', '--no-video-title-show', '--quiet')
        self.backend = None
        self.last_win_id = None
        self._check_and_set_best_backend()
        
        self.mode = "file"
        self.skip_on_pause = 0
        
    def set_backend(self, backend_type):
        """Swaps the current playback engine with maximum stability guards"""
        # Block signals during transition to prevent UI callbacks on semi-released objects
        self.blockSignals(True)
        try:
            if self.backend:
                # 1. STOP & DISCONNECT
                try: self.backend.stop()
                except: pass
                
                try: self.backend.positionChanged.disconnect()
                except: pass
                try: self.backend.durationChanged.disconnect()
                except: pass
                
                # 2. RELEASE
                try: self.backend.release()
                except: pass
                
                self.backend = None
                
            # Brief OS/Library settle time
            time.sleep(0.05)
            
            # 3. CREATE NEW
            if backend_type == "vlc":
                self.backend = VLCBackend(self.vlc_instance)
            elif backend_type == "mpv":
                saved_path = None
                if hasattr(self, 'config') and 'playback' in self.config:
                    saved_path = self.config['playback'].get('mpv_path')
                self.backend = MPVBackend(saved_path, win_id=self.last_win_id)
            elif backend_type in ["ffmpeg", "quicktime"]:
                # These currently fall back to VLC
                self.backend = VLCBackend(self.vlc_instance)
            else:
                # Default fallback
                self.backend = VLCBackend(self.vlc_instance)
                
            # 4. ATTACH & CONNECT
            if self.backend:
                if self.last_win_id:
                    try: self.backend.set_display_handle(self.last_win_id)
                    except: pass
                    
                try: self.backend.positionChanged.connect(self.positionChanged.emit)
                except: pass
                try: self.backend.durationChanged.connect(self.durationChanged.emit)
                except: pass
            else:
                logger.error(f"MediaEngine: Failed to initialize any backend for type {backend_type}")
                # Fallback to a dummy backend or similar if needed
            
        finally:
            self.blockSignals(False)

    def load_source(self, source_type, path):
        logger.info(f"MediaEngine: Loading {source_type} from {path}")
        self.mode = source_type
        self.current_path = path
        try:
            self.backend.load(source_type, path)
            logger.info("MediaEngine: Backend load call completed")
        except Exception as e:
            logger.exception(f"MediaEngine: Error in backend.load: {e}")
            raise

    def play_pause(self):
        if self.backend.is_playing():
            # Pausing: Apply skip immediately
            self.backend.pause()
            if self.skip_on_pause > 0:
                curr = self.backend.get_time()
                self.backend.seek(max(0, curr - self.skip_on_pause))
        else:
            # Playing: Just start
            self.backend.play()

    def stop(self): self.backend.stop()
    def play(self): self.backend.play()
    def pause(self): self.backend.pause()
    def seek(self, ms): self.backend.seek(ms)
    def seek_relative(self, ms):
        curr = self.backend.get_time()
        self.backend.seek(max(0, curr + ms))
        
    def frame_step(self, forward=True):
        self.backend.frame_step(forward)

    def set_rate(self, rate): self.backend.set_rate(rate)
    def set_volume(self, vol): self.backend.set_volume(vol)
    def get_audio_tracks(self): return self.backend.get_audio_tracks()
    def get_subtitle_tracks(self): return self.backend.get_subtitle_tracks()
    def set_audio_track(self, tid): self.backend.set_audio_track(tid)
    def set_subtitle_track(self, tid): self.backend.set_subtitle_track(tid)
    def set_display_handle(self, win_id):
        self.last_win_id = win_id
        if self.backend:
            self.backend.set_display_handle(win_id)
    def set_aspect_ratio(self, ratio): self.backend.set_aspect_ratio(ratio)
    def set_pitch_lock(self, enabled): self.backend.set_pitch_lock(enabled)
    def get_time(self): return self.backend.get_time()
    def is_playing(self): return self.backend.is_playing()
    def get_duration(self): return self.backend.get_duration()

    def download_mpv(self, progress_callback=None):
        """Downloads MPV DLL directly and renames it to mpv-1.dll"""
        # Primary direct link from user
        primary_link = "https://download0526.sfile.co/downloadfile/2203611/2/9a78cdce0a66f9fa2c178ace77c4f984/mpv-1.dll?k=dde201b091536b55dd44748b311db28f"
        # Backup link
        backup_link = "https://sfile.co/SjmPjUXcgvO"
        
        # Determine the correct target directory
        if getattr(sys, 'frozen', False):
            # Running as a PyInstaller bundle
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            
        target_path = os.path.join(base_dir, "mpv-1.dll")
        
        # CHECK FOR PERMISSIONS: In Program Files, we might not have write access
        if not os.access(base_dir, os.W_OK):
            print(f"CRITICAL: No write permission in {base_dir}. The application likely needs to be run as Administrator.")
            return "PERMISSION_DENIED"

        for url in [primary_link, "https://download0426.sfile.co/downloadfile/2203611/2/9d8307fba55270bf6b6efc4335b95e11/mpv-1.dll?k=dde201b091536b55dd44748b311db28f"]:
            try:
                print(f"Attempting download from: {url}")
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=15) as response:
                    total_size = int(response.info().get('Content-Length', -1))
                    block_size = 16384 # 16KB blocks
                    downloaded = 0
                    
                    with open(target_path, "wb") as f:
                        while True:
                            block = response.read(block_size)
                            if not block:
                                break
                            f.write(block)
                            downloaded += len(block)
                            if total_size > 0 and progress_callback:
                                progress_callback(min(100, int(downloaded * 100 / total_size)))

                print(f"Downloaded and saved MPV DLL: {target_path}")
                # Re-check backend after download
                self._check_and_set_best_backend()
                return True
            except Exception as e:
                print(f"Download from {url} failed: {e}")
                continue # Try next link
                
        return False

    def _check_and_set_best_backend(self):
        preferred = "vlc"
        if hasattr(self, 'config') and self.config:
            preferred = self.config.get('engine', 'vlc')
            
        # Try preferred first
        if preferred == "mpv":
            # Check for a saved MPV library path in current config
            saved_path = None
            if hasattr(self, 'config') and 'playback' in self.config:
                saved_path = self.config['playback'].get('mpv_path')
            
            mpv = MPVBackend(saved_path)
            if mpv.mpv_available:
                self.backend = mpv
            else:
                self.backend = VLCBackend(self.vlc_instance)
        else:
            self.backend = VLCBackend(self.vlc_instance)
        
        # Initial Signal Connections
        try: self.backend.positionChanged.connect(self.positionChanged.emit)
        except: pass
        try: self.backend.durationChanged.connect(self.durationChanged.emit)
        except: pass

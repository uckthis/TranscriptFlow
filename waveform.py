from PyQt6.QtWidgets import QWidget
from PyQt6.QtCore import pyqtSignal, Qt, QThread, QPoint, QSize, QLine
from PyQt6.QtGui import QPainter, QColor, QPen, QLinearGradient, QPixmap
import subprocess
import numpy as np
import shutil
import os
import hashlib
import time
from path_manager import get_waveform_cache_dir

class WaveformWorker(QThread):
    finished = pyqtSignal(object) # Returns numpy array

    def __init__(self, path, retention_months=3):
        super().__init__()
        self.path = path
        self.retention_months = retention_months
        self.cache_dir = self._get_cache_dir()
        self._cleanup_old_cache()

    def _get_cache_dir(self):
        """Get the waveform cache directory from path manager"""
        return get_waveform_cache_dir()

    def _cleanup_old_cache(self):
        """Delete cache files older than retention_months"""
        if self.retention_months <= 0:
            return
            
        if not os.path.exists(self.cache_dir):
            return
        
        # Approximate 30 days per month
        cutoff = time.time() - (self.retention_months * 30 * 24 * 60 * 60)
        
        for filename in os.listdir(self.cache_dir):
            filepath = os.path.join(self.cache_dir, filename)
            if os.path.isfile(filepath) and filename.endswith('.npy'):
                try:
                    file_mtime = os.path.getmtime(filepath)
                    if file_mtime < cutoff:
                        os.remove(filepath)
                        print(f"Cleaned up old waveform cache: {filename}")
                except Exception as e:
                    print(f"Error cleaning cache file {filename}: {e}")

    @staticmethod
    def clear_cache():
        """Deletes all cached waveform files"""
        cache_dir = get_waveform_cache_dir()
        if not os.path.exists(cache_dir):
            return True
            
        success = True
        for filename in os.listdir(cache_dir):
            if filename.endswith('.npy'):
                try:
                    os.remove(os.path.join(cache_dir, filename))
                except Exception as e:
                    print(f"Error clearing waveform cache file {filename}: {e}")
                    success = False
        return success

    def _get_cache_path(self):
        """Generate cache file path based on media file path hash"""
        # Create a hash of the file path for cache key
        path_hash = hashlib.md5(self.path.encode('utf-8')).hexdigest()
        return os.path.join(self.cache_dir, f"{path_hash}.npy")

    def run(self):
        # Check cache first
        cache_path = self._get_cache_path()
        if os.path.exists(cache_path):
            try:
                data = np.load(cache_path)
                print(f"Loaded waveform from cache: {os.path.basename(cache_path)}")
                self.finished.emit(data)
                return
            except Exception as e:
                print(f"Error loading cache, regenerating: {e}")
        
        # Generate waveform if not in cache
        if not shutil.which("ffmpeg"):
            print("FFmpeg not found")
            self.finished.emit(None)
            return
            
        try:
            cmd = [
                'ffmpeg', '-i', self.path, 
                '-f', 's16le', '-ac', '1', '-acodec', 'pcm_s16le', '-ar', '400', 
                'pipe:1'
            ]
            process = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
            raw = process.stdout
            if not raw:
                self.finished.emit(None)
                return
            
            # Process raw bytes into amplitudes
            data = np.frombuffer(raw, dtype=np.int16)
            
            # Save to cache
            try:
                np.save(cache_path, data)
                print(f"Saved waveform to cache: {os.path.basename(cache_path)}")
            except Exception as e:
                print(f"Error saving to cache: {e}")
            
            self.finished.emit(data)
        except:
            self.finished.emit(None)

class WaveformWidget(QWidget):
    seekRequested = pyqtSignal(int)
    generateRequested = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setMinimumHeight(10)
        self.data = None
        self.visual_data = None # Precomputed downsampled peaks
        self.duration_ms = 1
        self.current_ms = 0
        self.window_ms = 90000  # Show 1.5 minutes at a time
        self.amplitude_zoom = 1.0
        self.cache_pixmap = None
        self.cache_page_index = -1
        self.cache_size = QSize(0, 0)
        self.last_sync_ms = 0
        self.last_sync_time = 0
        self.is_playing = False
        self.playback_rate = 1.0
        
        from PyQt6.QtCore import QTimer
        self.smooth_timer = QTimer(self)
        self.smooth_timer.timeout.connect(self.update)
        self.smooth_timer.setInterval(16) # ~60fps
        
        self.setStyleSheet("background-color: #1a1a1a;")

    def load_data(self, data):
        self.data = data
        self.cache_pixmap = None # Invalidate cache
        if self.data is not None:
            # Fallback duration calculation: Worker uses 400Hz (1 sample = 2.5ms)
            proposed_ms = len(self.data) * 2.5
            # Only use if current duration is unset or significantly different
            if self.duration_ms <= 1 or abs(self.duration_ms - proposed_ms) > 1000:
                self.duration_ms = proposed_ms
                
        self.precompute_visual_data()
        self.update()

    def precompute_visual_data(self):
        """Precompute downsampled peaks for fast rendering"""
        if self.data is None or len(self.data) == 0:
            self.visual_data = None
            self.cache_pixmap = None
            self.cache_page_index = -1
            self.cache_size = QSize(0, 0)
            return
            
        # Target ~10,000 points for the whole file is enough for smooth overview
        target_resolution = 10000 
        num_samples = len(self.data)
        step = max(1, num_samples // target_resolution)
        
        # Reshape and find max peaks
        visual_points = []
        for i in range(0, num_samples, step):
            chunk = self.data[i : i + step]
            if len(chunk) > 0:
                visual_points.append(np.max(np.abs(chunk)))
            else:
                visual_points.append(0)
        
        self.visual_data = np.array(visual_points, dtype=np.float32)

    def set_position(self, ms):
        self.current_ms = ms
        self.last_sync_ms = ms
        self.last_sync_time = time.time()
        self.update()

    def set_playing(self, playing):
        self.is_playing = playing
        if playing:
            self.smooth_timer.start()
        else:
            self.smooth_timer.stop()
        self.update()

    def set_playback_rate(self, rate):
        self.playback_rate = rate

    def set_duration(self, ms):
        self.duration_ms = max(1, ms)

    def set_amplitude_zoom(self, zoom):
        self.amplitude_zoom = max(0.1, zoom)
        self.cache_pixmap = None # Invalidate cache to force redraw of bars
        self.update()

    def set_timeline_zoom(self, zoom_pct):
        """zoom_pct: 0.0 (wide) to 1.0 (narrow/zoomed in)"""
        # Range: 180s (3min) at 0% zoom down to 5s at 100% zoom
        min_window = 5000 
        max_window = 180000 
        self.window_ms = max_window - (zoom_pct * (max_window - min_window))
        self.window_ms = max(min_window, self.window_ms)
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        w, h = self.width(), self.height()
        if self.visual_data is None or len(self.visual_data) == 0:
            painter.fillRect(self.rect(), QColor("#1a1a1a"))
            painter.setPen(QPen(QColor("#00d2ff"), 1))
            font = painter.font()
            font.setPointSize(10)
            font.setBold(True)
            painter.setFont(font)
            painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, "CLICK TO GENERATE WAVEFORM")
            return

        # Simple Paging
        page_index = int(self.current_ms // self.window_ms)
        
        # --- Cache Logic for Smooth Playback ---
        if (self.cache_pixmap is None or 
            self.cache_page_index != page_index or 
            self.cache_size != self.size()):
            
            self.cache_pixmap = QPixmap(self.size())
            self.cache_pixmap.fill(QColor("#1a1a1a"))
            cache_painter = QPainter(self.cache_pixmap)
            
            start_ms = page_index * self.window_ms
            end_ms = start_ms + self.window_ms
            num_visual_points = len(self.visual_data)
            start_v_idx = int((start_ms / self.duration_ms) * num_visual_points)
            end_v_idx = int((end_ms / self.duration_ms) * num_visual_points)
            window_points = self.visual_data[start_v_idx:end_v_idx]
            
            if len(window_points) > 0:
                half_h = h // 2
                # Rudimentary Baseline
                cache_painter.setPen(QPen(QColor(100, 100, 100, 40), 1))
                cache_painter.drawLine(0, half_h, w, half_h)

                # Waveform (Thicker 3px bars, drawn as batch for extreme speed)
                cache_painter.setPen(QPen(QColor("#00d2ff"), 2)) 
                v_ptr_step = len(window_points) / max(1, w)
                
                lines = []
                for x in range(0, w, 3): # 3px steps: good balance of detail/speed
                    v_idx = int(x * v_ptr_step)
                    if v_idx < len(window_points):
                        val = abs(window_points[v_idx]) / 32768.0
                        val = min(1.0, val * 4.0 * self.amplitude_zoom)
                        if val > 0.01:
                            bar_h = int(val * h * 0.45)
                            lines.append(QLine(int(x), half_h - bar_h, int(x), half_h + bar_h))
                
                if lines:
                    cache_painter.drawLines(lines)
            
            cache_painter.end()
            self.cache_page_index = page_index
            self.cache_size = self.size()

        # Fast Background Draw
        painter.drawPixmap(0, 0, self.cache_pixmap)

        # INTERPOLATED Playhead for "Tik Tok" Free Motion
        render_ms = self.current_ms
        if self.is_playing:
            now = time.time()
            # Safety: If engine hasn't pulsed for > 1500ms, stop interpolation to prevent "ghost" movement
            if now - self.last_sync_time > 1.5:
                # Don't set self.is_playing = False here as it might cause flickering, 
                # but use the last known sync position instead.
                render_ms = self.last_sync_ms
            else:
                elapsed = (now - self.last_sync_time) * 1000
                render_ms = self.last_sync_ms + (elapsed * self.playback_rate)
            
            # Clamp to next page or duration to avoid overshooting
            max_ms = min(self.duration_ms, (page_index + 1) * self.window_ms)
            render_ms = min(render_ms, max_ms)

        # Simple Vertical Thread Playhead
        start_ms = page_index * self.window_ms
        end_ms = start_ms + self.window_ms
        window_duration = max(1, end_ms - start_ms)
        relative_ms = render_ms - start_ms
        pos_x = (relative_ms / window_duration) * w
        
        painter.setPen(QPen(QColor("#ffde00"), 1))
        painter.drawLine(int(pos_x), 0, int(pos_x), h)
        
        # Small rect at top
        painter.fillRect(int(pos_x)-3, 0, 6, 6, QColor("#ffde00"))

    def mousePressEvent(self, event):
        if self.data is None:
            self.generateRequested.emit()
            return
        
        # Paging Logic Seek
        page_index = int(self.current_ms // self.window_ms)
        start_ms = page_index * self.window_ms
        end_ms = start_ms + self.window_ms
        window_duration = max(1, end_ms - start_ms)
        
        pos_x = event.position().x()
        pct = pos_x / max(1, self.width())
        ms = int(start_ms + pct * window_duration)
        ms = max(0, min(self.duration_ms, ms))
        
        self.current_ms = ms
        self.update()
        self.seekRequested.emit(ms)

    def wheelEvent(self, event):
        if self.duration_ms <= 0: return
        delta = event.angleDelta().y()
        if delta == 0: return
        
        seek_step = 2000 # 2 seconds
        if delta > 0:
            new_ms = min(self.duration_ms, self.current_ms + seek_step)
        else:
            new_ms = max(0, self.current_ms - seek_step)
            
        if new_ms != self.current_ms:
            self.current_ms = new_ms
            self.update()
            self.seekRequested.emit(new_ms)
        event.accept()

import sys
import os
import re
import datetime
import time
import logging
from logging.handlers import RotatingFileHandler
import constants

# Configure Spell Check Dictionaries BEFORE importing enchant
# Import path manager to handle directory resolution
from path_manager import get_dicts_dir, initialize_app_directories, get_logs_dir

# Initialize all app directories (creates AppData structure)
initialize_app_directories()

def setup_logging():
    """Configure logging to write to a file in the AppData logs directory"""
    logs_dir = get_logs_dir()
    log_file = os.path.join(logs_dir, 'transcriptflow.log')
    
    handlers = [
        RotatingFileHandler(log_file, maxBytes=10*1024*1024, backupCount=5)
    ]
    
    # Only add StreamHandler if stdout exists (not in windowed mode)
    if sys.stdout is not None:
        handlers.append(logging.StreamHandler(sys.stdout))

    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=handlers
    )
    
    logger = logging.getLogger('TranscriptFlow')
    logger.info("Application starting up...")
    logger.info(f"Log file: {log_file}")
    
    def exception_hook(exctype, value, traceback):
        """Global exception hook to capture unhandled exceptions"""
        import traceback as tb
        err_msg = "".join(tb.format_exception(exctype, value, traceback))
        logger.critical("Uncaught Exception:\n" + err_msg)
        # Call the original excepthook if you want the default behavior too
        sys.__excepthook__(exctype, value, traceback)

    # Install the exception hook
    sys.excepthook = exception_hook

    # Log critical application paths
    from path_manager import get_appdata_dir, get_dicts_dir, get_backup_dir
    logger.info(f"App Data: {get_appdata_dir()}")
    logger.info(f"Dictionaries: {get_dicts_dir()}")
    logger.info(f"Backup Dir: {get_backup_dir()}")

    # Add _internal directory to DLL search path for Windows
    if getattr(sys, 'frozen', False) and sys.platform == 'win32':
        internal_dir = os.path.join(os.path.dirname(sys.executable), '_internal')
        if os.path.exists(internal_dir):
            try:
                # Use os.add_dll_directory if available (Python 3.8+)
                if hasattr(os, 'add_dll_directory'):
                    os.add_dll_directory(internal_dir)
                    logger.info(f"Added DLL search directory: {internal_dir}")
            except Exception as e:
                logger.warning(f"Failed to add DLL directory {internal_dir}: {e}")

    return logger

# Setup logging
logger = setup_logging()

def get_resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)

# Set enchant to use the AppData dicts directory
dicts_path = get_dicts_dir()
os.environ["ENCHANT_DATA_PATH"] = dicts_path

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QSlider, QFileDialog, 
                             QMessageBox, QSplitter, QLabel, QFrame, QDockWidget,
                             QInputDialog, QTextBrowser, QDialog, QProgressDialog,
                             QToolBar, QColorDialog, QFontComboBox, QComboBox, QGridLayout, QSizePolicy, QMenu,
                             QSplashScreen)
from PyQt6.QtCore import Qt, QUrl, QTimer, QCoreApplication, QByteArray, QRegularExpression, QLocale
from PyQt6.QtGui import (QAction, QIcon, QKeySequence, QActionGroup, QFont, QColor, 
                         QTextCharFormat, QFontDatabase, QTextCursor, QPixmap, QPainter, QLinearGradient, QBrush,
                         QDesktopServices, QPen, QGuiApplication, QTextOption)
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog, QPageSetupDialog

# Set Application User Model ID for Windows Taskbar Icon stability
if sys.platform == "win32":
    import ctypes
    myappid = 'mycompany.mytranscriber.transcriptflow.1.0'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

# Import Modules
from media_engine import MediaEngine
from editor import TranscriptEditor
from waveform import WaveformWidget, WaveformWorker
from dialogs import (MediaSourceDialog, SnippetsManagerDialog, 
                      ShortcutsManagerDialog, TranscriptSettingsDialog, 
                      ExportSettingsDialog, AdjustTimecodesDialog, MediaOffsetDialog,
                      SpellCheckDialog, SyncTranscriptDialog, FindReplaceDialog)
from backup_dialog import BackupSettingsDialog
import hardware
from utils import Exporter, TimecodeHelper, SettingsManager, FileManager, BackupManager, TranscriptParser

class ClickableSlider(QSlider):
    def __init__(self, orientation, parent=None, overlay_text=""):
        super().__init__(orientation, parent)
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.overlay_text = overlay_text

    def paintEvent(self, event):
        super().paintEvent(event)
        if self.overlay_text:
            painter = QPainter(self)
            # Use semi-transparent white/gray for a subtle look
            painter.setPen(QPen(QColor(255, 255, 255, 100))) 
            font = painter.font()
            font.setPointSize(8)
            font.setBold(True)
            font.setLetterSpacing(QFont.SpacingType.AbsoluteSpacing, 1.2)
            painter.setFont(font)
            
            # Draw centered text
            painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, self.overlay_text)
            painter.end()

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            if self.orientation() == Qt.Orientation.Horizontal:
                val = self.minimum() + ((self.maximum() - self.minimum()) * event.position().x()) / self.width()
            else:
                # Vertical sliders: click position is top-down, so we invert
                val = self.maximum() - ((self.maximum() - self.minimum()) * event.position().y()) / self.height()
            self.setValue(int(val))
            event.accept()
        super().mousePressEvent(event)

class MainWindow(QMainWindow):
    def __init__(self, splash=None):
        self.is_initializing = True
        self.splash = splash
        if self.splash:
            self.splash.showMessage("Loading settings...", Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignCenter, QColor("white"))
        
        super().__init__()
        self.setWindowTitle("TranscriptFlow Pro")
        self.resize(1400, 900)
        self.current_file_path = None
        self.current_media_path = None # Initialize media path
        self.printer = QPrinter()
        
        # Set App Icon
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_icon.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        
        # --- Data & Settings ---
        self.sm = SettingsManager()
        
        default_config = {
            'settings': {
                'font': 'Tahoma', 
                'size': 14, 
                'fps': 23.976,
                'timecode_format': '[00:01:23:29]',
                'omit_frames': False,
                'media_offset': 0,
                'timecode_color': '#00aa00',
                'font_color': 'black',
                'print_margins': {
                    'units': 'Inches',
                    'top': 1.0, 'bottom': 1.0, 'left': 1.0, 'right': 1.0
                },
                'recognize_unbracketed': True,
                'timecode_new_line': True,
                'spell_check': True,
                'spell_check_lang': 'en_GB',
                'autosave_enabled': False,
                'autosave_interval': 5,
            },
            'backup_interval': 1,
            'default_window_size': [1024, 800],
            'playback': {
                'volume': 36,
                'rate': 1.0,
                'pitch_lock': True,
                'preferred_player': 'mpv'
            },
            'shortcuts': [
                {'name': 'Toggle Pause and Play', 'trigger': 'Tab', 'command': 'Toggle Pause and Play', 'skip': 0.5},
                {'name': 'Skipback', 'trigger': 'Ctrl+Left', 'command': 'Skipback', 'skip': 2.0},
                {'name': 'Fast Forward', 'trigger': 'Ctrl+Right', 'command': 'Fast Forward', 'skip': 5.0},
                {'name': 'Rewind', 'trigger': 'Alt+Left', 'command': 'Rewind', 'skip': 0},
                {'name': 'Insert Time', 'trigger': 'Ctrl+;', 'command': 'Insert Current Time'},
                {'name': 'Frame Fwd', 'trigger': 'Alt+Right', 'command': 'Advance One Frame'},
                {'name': 'Speed Up', 'trigger': 'Ctrl+Up', 'command': 'Increase Play Rate', 'value': 0.1},
                {'name': 'Slow Down', 'trigger': 'Ctrl+Down', 'command': 'Decrease Play Rate', 'value': 0.1},
                {'name': 'Increase Volume', 'trigger': 'Ctrl+Shift+Up', 'command': 'Increase Volume', 'value': 10},
                {'name': 'Decrease Volume', 'trigger': 'Ctrl+Shift+Down', 'command': 'Decrease Volume', 'value': 10},
            ],
            'snippets': [
                {'name': 'Interviewer', 'trigger': 'Ctrl+3', 'text': '{$time}\nInterviewer: ', 'color': '#0000ff'},
                {'name': 'Interviewee', 'trigger': 'Ctrl+4', 'text': '{$time}\nInterviewee: ', 'color': '#800080'},
                {'name': 'Laughter', 'trigger': 'Ctrl+L', 'text': '[laughter] ', 'color': '#ff00ff'},
                {'name': 'BOY', 'trigger': 'Ctrl+1', 'text': '{$time}\nBOY: ', 'color': '#ff007f'},
                {'name': 'WOMAN', 'trigger': 'Ctrl+7', 'text': '{$time}\nWOMAN: ', 'color': '#ffaa00'},
            ],
            'engine': 'mpv',
            'ui': {
                'show_timeline': True,
                'show_remote': True,
                'show_playrate': True,
                'show_volume': True,
                'show_waveform': False,
                'dark_mode': False,
                'theme': 'Sepia',
                'video_player_size': None,
                'left_splitter_state': None,
                'player_scaling_behavior': 'proportional'
            },
            'playback': {
                'volume': 36,
                'speed': 100,
                'waveform_amplitude_zoom': 38,
                'waveform_timeline_zoom': 0.1
            },
            'backup_interval': 1,
            'waveform_retention_months': 3,
            'auto_generate_waveform': False,
            'last_export_dir': os.path.expanduser("~"),
            'autoplay_on_load': False,
            'theme_builder': {
                'custom_themes': {},
                'hidden_themes': []
            },
            'recent_files': []
        }
        
        self.find_dialog = None # Initialize Find & Replace dialog state
        self.config = self.sm.load(default_config)
        self.settings = self.config['settings']
        self.shortcuts = self.config['shortcuts']
        self.snippets = self.config['snippets']
        self.recent_files = self.config.get('recent_files', [])

        # UI visibility flags
        self.show_timeline = self.config['ui']['show_timeline']
        self.show_remote = self.config['ui']['show_remote']
        self.show_playrate = self.config['ui']['show_playrate']
        self.show_volume = self.config['ui']['show_volume']
        self.show_waveform = self.config['ui']['show_waveform']
        self.dark_mode = self.config['ui'].get('dark_mode', False)
        self.current_theme = self.config['ui'].get('theme', 'Light')
        self.rtl_mode = self.config['ui'].get('rtl_mode', False)
        
        # Playback settings
        self.playback_config = self.config.get('playback', {
            'volume': 70, 'speed': 100, 
            'waveform_amplitude_zoom': 10, 'waveform_timeline_zoom': 0.2
        })
        
        # --- Engine ---
        self.engine = MediaEngine()
        self.engine.config = self.config
        self.engine.positionChanged.connect(self.on_position_changed)
        self.engine.durationChanged.connect(self.on_duration_changed)
        
        if self.splash:
            self.splash.showMessage("Initializing UI...", Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignCenter, QColor("white"))
        
        # --- UI Setup ---
        self.init_ui()
        self.init_menus()
        
        # Delayed/Safe Initialization
        # This mitigates the memory access crash on startup by letting the OS
        # fully register the main window environment before loading media DLLs or complex themes.
        QTimer.singleShot(1500, self.late_initialization)
        
        # --- CLI File Opening ---
        QTimer.singleShot(2000, self._check_cli_args)
        
        # Apply saved playback settings to UI controls
        self.apply_saved_playback_settings()
        
        # ENSURE INITIAL STATE IS PAUSED (Amber/Play Icon)
        self.update_playback_visuals(force_playing=False)
        
        # Apply Logic
        self.update_playback_skip()
        self.update_tc_helper()
        self.last_seek_time = 0
        self.setAcceptDrops(True)
        
        # --- Register Urdu Font (Jameel Noori Nastaleeq) ---
        font_path = get_resource_path("Jameel Noori Nastaleeq.ttf")
        self.urdu_font_family = "Jameel Noori Nastaleeq" # Fallback
        if os.path.exists(font_path):
            f_id = QFontDatabase.addApplicationFont(font_path)
            if f_id != -1:
                families = QFontDatabase.applicationFontFamilies(f_id)
                if families:
                    self.urdu_font_family = families[0]
                    logger.info(f"Custom font '{self.urdu_font_family}' registered successfully.")

        # --- Monitor Input Language ---
        QApplication.inputMethod().localeChanged.connect(self.on_locale_changed)
        
        self.is_dirty = False
        self.editor.textChanged.connect(self._on_text_changed_internal)

        # --- Backup System ---
        self.backup_manager = BackupManager(self.sm)
        # Initialize to a past date to allow immediate backup on FIRST change
        self.last_backup_time = datetime.datetime.now() - datetime.timedelta(days=1)
        self.last_autosave_time = datetime.datetime.now()
        self.last_typed_time = datetime.datetime.now()
        self.media_just_loaded = False
        self.is_backup_dirty = False
        
        # hardware
        self.pedal_manager = hardware.FootPedalManager(self.config)
        self.pedal_manager.pedalPressed.connect(self.on_pedal_pressed)
        
        # Protective Delay for Hardware & Start-up
        # This addresses the memory access crash reported on first start
        QTimer.singleShot(2000, self.pedal_manager.start)
        
        self.backup_timer = QTimer(self)
        self.backup_timer.timeout.connect(self.perform_auto_backup)
        self.backup_timer.start(1000) # Check every second for precision

        # Initial RTL sync
        self.is_initializing = False
        
        # Initial locale sync (triggers correct font/alignment for current keyboard)
        self.on_locale_changed()
        
        # Application-wide shortcuts (Primary: Playback toggle)
        self.setup_global_shortcuts()

    def setup_global_shortcuts(self):
        self.tab_act = QAction("Global Toggle Play", self)
        self.tab_act.setShortcut(QKeySequence(Qt.Key.Key_Tab))
        # Ensure it works globally regardless of where focus is
        self.tab_act.setShortcutContext(Qt.ShortcutContext.ApplicationShortcut)
        self.tab_act.triggered.connect(self.toggle_play)
        self.addAction(self.tab_act)

    def _on_text_changed_internal(self):
        """Central handler for all text changes to track dirty state and backup timing"""
        self.is_backup_dirty = True
        self.last_typed_time = datetime.datetime.now()
        self._mark_as_dirty()

    def _mark_as_dirty(self):
        self.is_dirty = True
        self.editor.document().setModified(True)

    def init_ui(self):
        # Main Layout
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0,0,0,0)
        main_layout.setSpacing(0)
        
        # Hide standard status bar to save space; use transient toast for messages
        self.statusBar().hide()
        
        # Floating Toast for status messages
        self.toast_label = QLabel(self)
        self.toast_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.toast_label.setStyleSheet("""
            background: rgba(0,0,0,0.8); color: white; border-radius: 10px; 
            padding: 5px 15px; font-weight: bold; border: 1px solid #444;
        """)
        self.toast_label.hide()
        self.toast_timer = QTimer(self)
        self.toast_timer.timeout.connect(self.toast_label.hide)

        # Core Component Construction (Splitters, Panels, & Editor)
        self.main_splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # LEFT SIDE: Video & Controls
        self.left_panel = QWidget()
        self.left_layout = QVBoxLayout(self.left_panel)
        self.left_layout.setContentsMargins(0,0,0,0) # Tightly align with Splitter
        self.left_layout.setSpacing(0)
        
        self.left_splitter = QSplitter(Qt.Orientation.Vertical)
        
        # Video Frame Container
        self.video_container = QWidget()
        video_c_layout = QVBoxLayout(self.video_container)
        # --- CHANGE THE GHOST GAP SIZE HERE ---
        # The 3rd value (10) below is the right margin (player side of the gap)
        video_c_layout.setContentsMargins(10,10,10,10) 
        self.video_frame = QFrame()
        self.video_frame.setMinimumSize(160, 90)
        video_c_layout.addWidget(self.video_frame)
        self.left_splitter.addWidget(self.video_container)

        # Controls Container
        self.controls_panel = QWidget()
        self.controls_layout = QVBoxLayout(self.controls_panel)
        self.controls_layout.setContentsMargins(10,0,10,10) # Restored right margin
        
        # Time Labels & Slider
        self.time_container = QWidget()
        time_layout = QHBoxLayout(self.time_container)
        time_layout.setContentsMargins(0,5,0,5)
        self.lbl_curr = QLabel("00:00:00.00")
        self.lbl_curr.setStyleSheet("font-weight: bold; font-size: 12px;")
        self.slider = ClickableSlider(Qt.Orientation.Horizontal)
        self.slider.sliderMoved.connect(self.engine.seek)
        self.slider.valueChanged.connect(self._on_slider_clicked)
        self.lbl_total = QLabel("00:00:00.00")
        self.lbl_total.setStyleSheet("font-size: 12px;")
        time_layout.addWidget(self.lbl_curr); time_layout.addWidget(self.slider); time_layout.addWidget(self.lbl_total)
        self.controls_layout.addWidget(self.time_container)
        
        # Play Buttons
        self.buttons_container = QWidget()
        btn_layout = QGridLayout(self.buttons_container)
        btn_layout.setContentsMargins(0, 5, 0, 5); btn_layout.setSpacing(8)
        btn_styles = "QPushButton { border-radius: 18px; min-width: 40px; min-height: 40px; font-size: 14px; font-weight: bold; }"
        
        self.btn_start = QPushButton("⏮"); self.btn_start.setStyleSheet(btn_styles); self.btn_start.clicked.connect(self.go_to_start)
        self.btn_start.setToolTip("Go to Start")
        
        self.btn_back5 = QPushButton("⏪ 5s"); self.btn_back5.setStyleSheet(btn_styles); self.btn_back5.clicked.connect(self.skip_back)
        self.btn_back5.setToolTip("Rewind 5 Seconds")
        
        self.btn_play = QPushButton("▶"); self.btn_play.setStyleSheet(btn_styles); self.btn_play.clicked.connect(self.toggle_play)
        self.btn_play.setObjectName("PlayPauseButton")
        
        self.btn_fwd5 = QPushButton("5s ⏩"); self.btn_fwd5.setStyleSheet(btn_styles); self.btn_fwd5.clicked.connect(self.skip_forward)
        self.btn_fwd5.setToolTip("Fast Forward 5 Seconds")
        
        self.btn_end = QPushButton("⏭"); self.btn_end.setStyleSheet(btn_styles); self.btn_end.clicked.connect(self.go_to_end)
        self.btn_end.setToolTip("Go to End")
        
        self.btn_ins_tc = QPushButton("🕒"); self.btn_ins_tc.setStyleSheet(btn_styles); self.btn_ins_tc.clicked.connect(self.insert_current_time)
        self.btn_ins_tc.setToolTip("Insert Current Timecode (Ctrl+;)")
        
        # Prevent focus loss when clicking playback buttons
        for btn in [self.btn_start, self.btn_back5, self.btn_play, self.btn_fwd5, self.btn_end, self.btn_ins_tc]:
            btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        
        btn_layout.addWidget(self.btn_start, 0, 0); btn_layout.addWidget(self.btn_back5, 0, 1); btn_layout.addWidget(self.btn_play, 0, 2)
        btn_layout.addWidget(self.btn_fwd5, 1, 0); btn_layout.addWidget(self.btn_end, 1, 1); btn_layout.addWidget(self.btn_ins_tc, 1, 2)
        self.controls_layout.addWidget(self.buttons_container)
        
        # Speed/Vol
        self.sliders_container = QWidget(); sliders_layout = QVBoxLayout(self.sliders_container)
        self.sl_rate = ClickableSlider(Qt.Orientation.Horizontal, overlay_text="SPEED"); self.sl_rate.setRange(25, 300); self.sl_rate.setValue(100); self.sl_rate.valueChanged.connect(self.on_speed_changed)
        self.speed_label = QLabel("1.0x"); sliders_layout.addWidget(self.sl_rate); sliders_layout.addWidget(self.speed_label)
        self.sl_vol = ClickableSlider(Qt.Orientation.Horizontal, overlay_text="VOLUME"); self.sl_vol.setRange(0, 100); self.sl_vol.setValue(100); self.sl_vol.valueChanged.connect(self.on_volume_changed)
        self.volume_label = QLabel("100%"); sliders_layout.addWidget(self.sl_vol); sliders_layout.addWidget(self.volume_label)
        
        self.sl_boost = ClickableSlider(Qt.Orientation.Horizontal, overlay_text="BOOST"); self.sl_boost.setRange(0, 300); self.sl_boost.setValue(0); self.sl_boost.valueChanged.connect(self.on_volume_changed)
        self.boost_label = QLabel("Boost: 0%"); sliders_layout.addWidget(self.sl_boost); sliders_layout.addWidget(self.boost_label)
        self.controls_layout.addWidget(self.sliders_container)
        
        self.left_splitter.addWidget(self.controls_panel); self.left_layout.addWidget(self.left_splitter)
        
        # RIGHT SIDE: Editor
        self.right_panel = QWidget(); self.right_layout = QVBoxLayout(self.right_panel)
        # --- CHANGE THE GHOST GAP SIZE HERE ---
        # The 1st value (10) below is the left margin (editor side of the gap)
        self.right_layout.setContentsMargins(10,0,10,10) 
        self.editor = TranscriptEditor(self.shortcuts, self.snippets, self.settings)
        # Note: Connection is handled in __init__ after init_ui
        self.editor.seekRequested.connect(self.engine.seek)
        self.editor.commandTriggered.connect(self.handle_command)
        self.editor.snippetTriggered.connect(self.handle_snippet)
        self.editor.settingsChanged.connect(self.handle_editor_settings_change)
        self.right_layout.addWidget(self.editor)
        
        self.main_splitter.addWidget(self.left_panel); self.main_splitter.addWidget(self.right_panel)
        self.main_splitter.setHandleWidth(8)
        self.main_splitter.setStyleSheet("QSplitter::handle { background: transparent; }") # INVISIBLE HANDLE
        self.main_splitter.setStretchFactor(0, 1); self.main_splitter.setStretchFactor(1, 2)

        # Ribbon Construction - VERSION 2: INDEPENDENT GROUP NAVIGATION
        self.format_bar = QToolBar("Formatting")
        self.format_bar.setObjectName("FormattingRibbon")
        self.format_bar.setMovable(False)
        self.format_bar.setFloatable(False)
        main_layout.addWidget(self.format_bar)

        # The Ribbon Splitter (5 Segments for Independent Movement)
        self.ribbon_splitter = QSplitter(Qt.Orientation.Horizontal)
        self.ribbon_splitter.setHandleWidth(8)
        self.ribbon_splitter.setStyleSheet("QSplitter::handle { background: transparent; }") # INVISIBLE HANDLE

        # Segment 1: Zone A (App Tools)
        self.zone_a = QFrame()
        zone_a_layout = QHBoxLayout(self.zone_a)
        zone_a_layout.setContentsMargins(5, 0, 5, 0)
        zone_a_layout.setSpacing(6)
        
        self.btn_ribbon_wf = QPushButton("🌊 Waveform")
        self.btn_ribbon_wf.setCheckable(True)
        self.btn_ribbon_wf.setChecked(self.show_waveform)
        self.btn_ribbon_wf.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.btn_ribbon_wf.clicked.connect(self.toggle_waveform)
        zone_a_layout.addWidget(self.btn_ribbon_wf)

        self.btn_ribbon_layouts = QPushButton("🗄️ Layouts")
        self.btn_ribbon_layouts.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        layout_menu = QMenu(self)
        self.ribbon_layout_group = QActionGroup(self)
        self.ribbon_layout_group.setExclusive(True)
        presets = [
            ("Standard", "standard"), ("Standard (No Waveform)", "standard_no_wf"),
            ("Wide", "wide"), ("Reversed", "reversed"), ("Reversed (No Waveform)", "reversed_no_wf"),
            ("Audio Only", "audio_only"), ("Stacked", "stacked"), ("Focus Mode", "focus"), ("Compact Player", "compact")
        ]
        for name, preset in presets:
            act = QAction(name, self); act.setCheckable(True); act.setData(preset)
            act.triggered.connect(lambda chk, p=preset: self.apply_layout_preset(p))
            layout_menu.addAction(act); self.ribbon_layout_group.addAction(act)
            if preset == self.config['ui'].get('layout', 'standard'): act.setChecked(True)
        self.btn_ribbon_layouts.setMenu(layout_menu)
        zone_a_layout.addWidget(self.btn_ribbon_layouts)

        self.btn_ribbon_themes = QPushButton("🎨 Themes")
        self.btn_ribbon_themes.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        theme_menu = QMenu(self)
        self.ribbon_theme_group = QActionGroup(self)
        self.ribbon_theme_group.setExclusive(True)
        themes = ['Light', 'Dark', 'Midnight', 'Emerald', 'Ocean', 'Sunset', 'Forest', 'Sepia', 'Velvet', 'Cyber', 'Coffee', 'Rose Gold', 'Nord', 'Crimson', 'Onyx', 'Lavender', 'Sandstone', 'Ember']
        for t in themes:
            act = QAction(t, self); act.setCheckable(True)
            act.triggered.connect(lambda chk, tm=t: self.set_theme(tm))
            theme_menu.addAction(act); self.ribbon_theme_group.addAction(act)
            if t == self.current_theme: act.setChecked(True)
        self.btn_ribbon_themes.setMenu(theme_menu)
        zone_a_layout.addWidget(self.btn_ribbon_themes)
        
        self.ribbon_splitter.addWidget(self.zone_a)

        # Segment 2: Transition Gap 1 (Draggable spacer)
        self.gap_1 = QWidget()
        self.ribbon_splitter.addWidget(self.gap_1)

        # Segment 3: Zone B (Format Tools)
        self.zone_b = QFrame()
        zone_b_layout = QHBoxLayout(self.zone_b)
        zone_b_layout.setContentsMargins(10, 0, 10, 0)
        zone_b_layout.setSpacing(6)
        
        def create_format_btn(text, slot, shortcut=None, checkable=False):
            btn = QPushButton(text)
            btn.setFixedSize(38, 38)
            btn.setCheckable(checkable)
            btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
            
            # Use standard font & re-assert text to ensure it's not overridden
            btn.setFont(QFont("Arial", 12, QFont.Weight.Bold))
            
            if shortcut: btn.setShortcut(shortcut)

            # SMART FOCUS WRAPPER: Always return focus to editor after click
            def on_btn_clicked():
                slot()
                self.editor.setFocus()

            btn.clicked.connect(on_btn_clicked)
            
            # ABSOLUTE VISIBILITY: Use solid colors that ignore theme overrides
            btn.setStyleSheet("""
                QPushButton { 
                    border-radius: 6px; 
                    background: rgba(255, 255, 255, 0.05);
                    padding: 0;
                }
                QPushButton:hover { 
                    background: rgba(255, 255, 255, 0.2); 
                }
                QPushButton:checked { 
                    background: #222; /* High contrast black */
                    color: #0f0;      /* Bright neon green label - UNMISSABLE */
                    border: 2px solid #0f0;
                }
            """)
            return btn

        self.btn_bold = create_format_btn("B", self.editor.toggle_bold, "Ctrl+B", True)
        self.btn_italic = create_format_btn("I", self.editor.toggle_italic, "Ctrl+I", True)
        self.btn_underline = create_format_btn("U", self.editor.toggle_underline, "Ctrl+U", True)
        
        zone_b_layout.addWidget(self.btn_bold)
        zone_b_layout.addWidget(self.btn_italic)
        zone_b_layout.addWidget(self.btn_underline)
        
        self.font_combo = QFontComboBox()
        self.font_combo.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.font_combo.setFontFilters(QFontComboBox.FontFilter.ScalableFonts) # Prevent terminal errors
        self.font_combo.setWritingSystem(QFontDatabase.WritingSystem.Latin)   # Exclude complex scripts causing warnings
        self.font_combo.setFixedWidth(160)
        self.font_combo.setCurrentFont(QFont(self.settings.get('font', 'Tahoma')))
        self.font_combo.currentFontChanged.connect(lambda f: self.editor.set_font_family(f.family()))
        zone_b_layout.addWidget(self.font_combo)
        
        self.size_combo = QComboBox()
        self.size_combo.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.size_combo.setFixedWidth(60)
        self.size_combo.addItems([str(s) for s in [8, 10, 12, 14, 16, 18, 20, 24, 28, 36, 48]])
        self.size_combo.setCurrentText(str(self.settings.get('size', 14)))
        self.size_combo.currentTextChanged.connect(self.editor.set_font_size)
        zone_b_layout.addWidget(self.size_combo)
        
        self.btn_color = QPushButton()
        self.btn_color.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.btn_color.setFixedSize(38, 38)
        self.btn_color.clicked.connect(self.choose_color)
        self.update_color_button_ui(self.settings.get('font_color', 'black'))
        zone_b_layout.addWidget(self.btn_color)
        
        self.ribbon_splitter.addWidget(self.zone_b)

        # Segment 4: Transition Gap 2 (Draggable spacer)
        self.gap_2 = QWidget()
        self.ribbon_splitter.addWidget(self.gap_2)

        # Segment 5: Zone C (Utility Shortcuts)
        self.zone_c = QFrame()
        zone_c_layout = QHBoxLayout(self.zone_c)
        zone_c_layout.setContentsMargins(5, 0, 15, 0)
        zone_c_layout.setSpacing(6)
        
        def create_util_btn(text, slot):
            btn = QPushButton(text)
            btn.setMinimumHeight(38)
            btn.setFocusPolicy(Qt.FocusPolicy.NoFocus) # Fix editor focus loss
            btn.clicked.connect(slot)
            return btn

        self.btn_spell = create_util_btn("🔡 Spell Check", self.toggle_spell_check)
        self.btn_spell.setCheckable(True)
        self.btn_spell.setChecked(self.settings.get('spell_check', True))
        
        self.btn_adj_tc = create_util_btn("⏱ Adjust TC", self.adjust_timecodes)
        self.btn_ribbon_snippets = create_util_btn("📋 Snippets", self.open_snippets_dialog)
        self.btn_ribbon_shortcuts = create_util_btn("⌨️ Shortcuts", self.open_shortcuts_dialog)
        self.btn_ribbon_offset = create_util_btn("⏳ Offset", self.set_media_offset)
        self.btn_settings = create_util_btn("⚙️ Settings", self.open_transcript_settings)
        
        zone_c_layout.addWidget(self.btn_spell)
        zone_c_layout.addWidget(self.btn_adj_tc); zone_c_layout.addWidget(self.btn_ribbon_snippets)
        zone_c_layout.addWidget(self.btn_ribbon_shortcuts); zone_c_layout.addWidget(self.btn_ribbon_offset)
        zone_c_layout.addWidget(self.btn_settings)
        
        self.ribbon_splitter.addWidget(self.zone_c)
        
        # Add the whole splitter to the toolbar
        self.format_bar.addWidget(self.ribbon_splitter)
        
        # Restore or Set INITIAL SIZES
        self.ribbon_splitter.setStretchFactor(1, 1) # Make gaps the expanding parts
        self.ribbon_splitter.setStretchFactor(3, 1)
        
        ribbon_state = self.config['ui'].get('ribbon_splitter_state')
        if ribbon_state:
            try:
                # restoreState should be called after stretch factors are set
                self.ribbon_splitter.restoreState(QByteArray.fromHex(ribbon_state.encode()))
            except Exception as e:
                logger.debug(f"Failed to restore ribbon state: {e}")
                self.ribbon_splitter.setSizes([260, 300, 440, 0, 400])
        else:
            self.ribbon_splitter.setSizes([260, 300, 440, 0, 400])

        self.editor.cursorPositionChanged.connect(self.update_ribbon_format)
        self.editor.currentCharFormatChanged.connect(lambda f: self.update_ribbon_format())
        
        main_layout.addWidget(self.main_splitter)

        # Waveform Dock (Detachable & Full Width)
        self.waveform_dock = QDockWidget("Timeline / Waveform", self)
        self.waveform_dock.setObjectName("WaveformDock")
        self.waveform_dock.setAllowedAreas(Qt.DockWidgetArea.BottomDockWidgetArea | Qt.DockWidgetArea.TopDockWidgetArea)
        self.waveform_dock.setTitleBarWidget(QWidget()) # HIDE TITLE BAR
        
        # Waveform + Zoom Control Container
        wf_container = QWidget(); wf_layout = QHBoxLayout(wf_container); wf_layout.setContentsMargins(5, 2, 5, 2)
        self.waveform = WaveformWidget()
        self.waveform.seekRequested.connect(self._on_waveform_seek)
        self.waveform.generateRequested.connect(self.trigger_waveform_generation)
        
        # Amplitude Zoom Slider
        azoom_layout = QVBoxLayout(); azoom_layout.setSpacing(2)
        self.zoom_slider = ClickableSlider(Qt.Orientation.Vertical); self.zoom_slider.setRange(10, 200); self.zoom_slider.setValue(10); self.zoom_slider.valueChanged.connect(self.on_waveform_zoom_changed)
        azoom_layout.addWidget(self.zoom_slider); wf_layout.addWidget(self.waveform, stretch=1); wf_layout.addLayout(azoom_layout)
        
        # Ultra-Minimal Waveform Navigation Row
        self.mini_slider = ClickableSlider(Qt.Orientation.Horizontal); self.mini_slider.setFixedHeight(6); self.mini_slider.sliderMoved.connect(self.engine.seek); self.mini_slider.valueChanged.connect(self._on_slider_clicked)
        self.lbl_mini_curr = QLabel("00:00:00.00"); self.lbl_mini_curr.setStyleSheet("font-size: 8px; font-family: 'Consolas';")
        self.lbl_mini_total = QLabel("00:00:00.00"); self.lbl_mini_total.setStyleSheet("font-size: 8px; font-family: 'Consolas';")
        mini_nav_layout = QHBoxLayout(); mini_nav_layout.setContentsMargins(5, 0, 5, 0); mini_nav_layout.setSpacing(8); mini_nav_layout.addWidget(self.lbl_mini_curr); mini_nav_layout.addWidget(self.mini_slider); mini_nav_layout.addWidget(self.lbl_mini_total)
        
        wf_master_container = QWidget(); wf_master_layout = QVBoxLayout(wf_master_container); wf_master_layout.setContentsMargins(0, 0, 0, 0); wf_master_layout.setSpacing(0); wf_master_layout.addWidget(wf_container); wf_master_layout.addLayout(mini_nav_layout)
        self.waveform.window_ms = 90000 
        
        self.waveform_dock.setWidget(wf_master_container)
        self.addDockWidget(Qt.DockWidgetArea.BottomDockWidgetArea, self.waveform_dock)
        self.waveform_dock.visibilityChanged.connect(self.on_waveform_visibility_changed)
        self.waveform_dock.setVisible(self.show_waveform)

        # Restore workspace state
        try:
            if 'geometry' in self.config['ui']: self.restoreGeometry(QByteArray.fromHex(self.config['ui']['geometry'].encode()))
            if 'window_state' in self.config['ui']: self.restoreState(QByteArray.fromHex(self.config['ui']['window_state'].encode()))
            if 'splitter_state' in self.config['ui']: self.main_splitter.restoreState(QByteArray.fromHex(self.config['ui']['splitter_state'].encode()))
            if 'left_splitter_state' in self.config['ui'] and self.config['ui']['left_splitter_state']: self.left_splitter.restoreState(QByteArray.fromHex(self.config['ui']['left_splitter_state'].encode()))
        except: pass
        
        # Prevent widgets from being collapsed to zero size (Fixes disappearing player bug)
        self.main_splitter.setCollapsible(0, False)
        self.main_splitter.setCollapsible(1, False)
        self.left_splitter.setCollapsible(0, False)
        self.left_splitter.setCollapsible(1, False)
        
        self.video_frame.show()
        
        self.apply_player_scaling()


    def _on_waveform_seek(self, ms):
        self.engine.seek(ms)

    def on_waveform_zoom_changed(self, value):
        self.waveform.amp_zoom = value / 10.0
        self.waveform.update()
        self.save_config()

    def get_builtin_themes(self):
        return {
            'Light': {
                'main_bg': '#f8fafc', 'text': '#1e293b', 
                'btn_grad': ['#ffffff', '#f1f5f9', '#e2e8f0'],
                'edit_bg': 'white', 'top_bar': '#facc15', 'accent': '#3b82f6', 
                'border': '#cbd5e1', 'glass': 'rgba(255, 255, 255, 0.7)',
                'shadow': 'rgba(0, 0, 0, 0.1)', 'glow': '#3b82f6'
            },
            'Dark': {
                'main_bg': '#0f172a', 'text': '#f1f5f9', 
                'btn_grad': ['#334155', '#1e293b', '#0f172a'],
                'edit_bg': '#e2e8f0', 'top_bar': '#10b981', 'accent': '#38bdf8', 
                'border': '#334155', 'glass': 'rgba(30, 41, 59, 0.7)',
                'shadow': 'rgba(0, 0, 0, 0.5)', 'glow': '#38bdf8'
            },
            'Midnight': {
                'main_bg': '#020617', 'text': '#f8fafc', 
                'btn_grad': ['#1e3a8a', '#1e40af', '#172554'],
                'edit_bg': '#dbeafe', 'top_bar': '#3b82f6', 'accent': '#60a5fa', 
                'border': '#1e3a8a', 'glass': 'rgba(30, 58, 138, 0.5)',
                'shadow': 'rgba(0, 0, 0, 0.7)', 'glow': '#60a5fa'
            },
            'Sepia': {
                'main_bg': '#fdf6e3', 'text': '#5b4636', 
                'btn_grad': ['#fdf6e3', '#eee8d5', '#decba4'],
                'edit_bg': '#fefcf0', 'top_bar': '#b58900', 'accent': '#8b4513', 
                'border': '#d3c6aa', 'glass': 'rgba(253, 246, 227, 0.7)',
                'shadow': 'rgba(91, 70, 54, 0.15)', 'glow': '#b58900'
            },
            'Emerald': {
                'main_bg': '#022c22', 'text': '#ecfdf5', 
                'btn_grad': ['#065f46', '#064e3b', '#022c22'],
                'edit_bg': '#d1fae5', 'top_bar': '#10b981', 'accent': '#34d399', 
                'border': '#064e3b', 'glass': 'rgba(6, 95, 70, 0.5)',
                'shadow': 'rgba(0, 0, 0, 0.6)', 'glow': '#34d399'
            },
            'Ocean': {
                'main_bg': '#0c4a6e', 'text': '#e0f2fe', 
                'btn_grad': ['#0e7490', '#0891b2', '#06b6d4'],
                'edit_bg': '#cffafe', 'top_bar': '#06b6d4', 'accent': '#22d3ee', 
                'border': '#155e75', 'glass': 'rgba(14, 116, 144, 0.6)',
                'shadow': 'rgba(0, 0, 0, 0.5)', 'glow': '#22d3ee'
            },
            'Sunset': {
                'main_bg': '#7c2d12', 'text': '#fed7aa', 
                'btn_grad': ['#ea580c', '#f97316', '#fb923c'],
                'edit_bg': '#fed7aa', 'top_bar': '#f97316', 'accent': '#fb923c', 
                'border': '#9a3412', 'glass': 'rgba(234, 88, 12, 0.5)',
                'shadow': 'rgba(0, 0, 0, 0.4)', 'glow': '#fb923c'
            },
            'Forest': {
                'main_bg': '#14532d', 'text': '#d1fae5', 
                'btn_grad': ['#15803d', '#16a34a', '#22c55e'],
                'edit_bg': '#d1fae5', 'top_bar': '#16a34a', 'accent': '#4ade80', 
                'border': '#166534', 'glass': 'rgba(21, 128, 61, 0.6)',
                'shadow': 'rgba(0, 0, 0, 0.5)', 'glow': '#4ade80'
            },
            'Velvet': {
                'main_bg': '#3b0764', 'text': '#f3e8ff', 
                'btn_grad': ['#6b21a8', '#7c3aed', '#8b5cf6'],
                'edit_bg': '#e9d5ff', 'top_bar': '#7c3aed', 'accent': '#a78bfa', 
                'border': '#5b21b6', 'glass': 'rgba(107, 33, 168, 0.5)',
                'shadow': 'rgba(0, 0, 0, 0.6)', 'glow': '#a78bfa'
            },
            'Cyber': {
                'main_bg': '#0a0a0a', 'text': '#00ffff', 
                'btn_grad': ['#1a1a2e', '#16213e', '#0f3460'],
                'edit_bg': '#ccffff', 'top_bar': '#00ffff', 'accent': '#ff00ff', 
                'border': '#00ffff', 'glass': 'rgba(0, 255, 255, 0.1)',
                'shadow': 'rgba(0, 255, 255, 0.3)', 'glow': '#ff00ff'
            },
            'Coffee': {
                'main_bg': '#271711', 'text': '#ede0d4', 
                'btn_grad': ['#3e2723', '#4e342e', '#5d4037'],
                'edit_bg': '#f5f5f5', 'top_bar': '#795548', 'accent': '#a1887f', 
                'border': '#4e342e', 'glass': 'rgba(121, 85, 72, 0.5)',
                'shadow': 'rgba(0, 0, 0, 0.6)', 'glow': '#d7ccc8'
            },
            'Rose Gold': {
                'main_bg': '#fce7f3', 'text': '#831843', 
                'btn_grad': ['#fce7f3', '#fbcfe8', '#f9a8d4'],
                'edit_bg': '#fdf2f8', 'top_bar': '#ec4899', 'accent': '#f472b6', 
                'border': '#f9a8d4', 'glass': 'rgba(252, 231, 243, 0.8)',
                'shadow': 'rgba(131, 24, 67, 0.15)', 'glow': '#ec4899'
            },
            'Nord': {
                'main_bg': '#2e3440', 'text': '#eceff4', 
                'btn_grad': ['#3b4252', '#434c5e', '#4c566a'],
                'edit_bg': '#d8dee9', 'top_bar': '#88c0d0', 'accent': '#81a1c1', 
                'border': '#4c566a', 'glass': 'rgba(59, 66, 82, 0.7)',
                'shadow': 'rgba(0, 0, 0, 0.5)', 'glow': '#88c0d0'
            },
            'Crimson': {
                'main_bg': '#1a1a1a', 'text': '#fecaca', 
                'btn_grad': ['#1f1212', '#2d1a1a', '#3f2323'],
                'edit_bg': '#e5cfcf', 'top_bar': '#dc2626', 'accent': '#ef4444', 
                'border': '#7f1d1d', 'glass': 'rgba(31, 18, 18, 0.6)',
                'shadow': 'rgba(127, 29, 29, 0.3)', 'glow': '#f87171'
            },
            'Onyx': {
                'main_bg': '#000000', 'text': '#fcd34d', 
                'btn_grad': ['#0a0a0a', '#111111', '#171717'],
                'edit_bg': '#e5e0d0', 'top_bar': '#fbbf24', 'accent': '#f59e0b', 
                'border': '#451a03', 'glass': 'rgba(23, 23, 23, 0.6)',
                'shadow': 'rgba(251, 191, 36, 0.1)', 'glow': '#fef3c7'
            },
            'Lavender': {
                'main_bg': '#1e1b4b', 'text': '#e0e7ff', 
                'btn_grad': ['#312e81', '#3730a3', '#4338ca'],
                'edit_bg': '#e0e7ff', 'top_bar': '#818cf8', 'accent': '#a5b4fc', 
                'border': '#3730a3', 'glass': 'rgba(49, 46, 129, 0.6)',
                'shadow': 'rgba(0, 0, 0, 0.5)', 'glow': '#c7d2fe'
            },
            'Sandstone': {
                'main_bg': '#451a03', 'text': '#ffedd5', 
                'btn_grad': ['#78350f', '#92400e', '#b45309'],
                'edit_bg': '#ffedd5', 'top_bar': '#f59e0b', 'accent': '#fbbf24', 
                'border': '#92400e', 'glass': 'rgba(120, 53, 15, 0.6)',
                'shadow': 'rgba(0, 0, 0, 0.5)', 'glow': '#fdba74'
            },
            'Ember': {
                'main_bg': '#0c0a09', 'text': '#f97316', 
                'btn_grad': ['#1c1917', '#292524', '#44403c'],
                'edit_bg': '#f5e9e2', 'top_bar': '#ea580c', 'accent': '#f97316', 
                'border': '#44403c', 'glass': 'rgba(28, 25, 23, 0.8)',
                'shadow': 'rgba(0, 0, 0, 0.7)', 'glow': '#fb923c'
            }
        }

    def apply_theme(self, theme_name=None):
        if theme_name is None:
            theme_name = self.current_theme
            
        # Merge builtin and custom themes
        themes = self.get_builtin_themes()
        builder_cfg = self.config.get('theme_builder', {})
        custom = builder_cfg.get('custom_themes', {})
        themes.update(custom)
        
        theme = themes.get(theme_name, themes.get('Light'))
        self.active_theme_colors = theme
        
        # Use a solid version of the background for menus/popups to avoid 'black hole' transparency issues on some systems
        menu_bg = theme['main_bg']
        
        style = f"""
            QMainWindow {{ 
                background-color: {theme['main_bg']}; 
                color: {theme['text']}; 
                font-family: 'Segoe UI', 'Inter', system-ui, sans-serif;
            }}
            
            QWidget {{ 
                background-color: {theme['main_bg']}; 
                color: {theme['text']}; 
            }}
            
            QLabel {{ 
                color: {theme['text']}; 
                font-size: 13px; 
            }}
            
            /* Premium Buttons with Shadows & Hover Effects */
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                    stop:0 {theme['btn_grad'][0]}, 
                    stop:0.5 {theme['btn_grad'][1]}, 
                    stop:1 {theme['btn_grad'][2]});
                border: 2px solid {theme['border']}; 
                border-radius: 8px; 
                padding: 8px 20px; 
                color: {theme['text']}; 
                font-weight: 600;
                font-size: 13px;
            }}
            
            QPushButton:hover {{ 
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {theme['accent']},
                    stop:1 {theme['glow']});
                color: white; 
                border: 2px solid white;
                padding: 7px 19px 9px 21px;
            }}
            
            QPushButton:pressed {{
                background-color: {theme['main_bg']};
                padding: 9px 19px 7px 21px;
            }}
            
            QPushButton:checked, QToolButton:checked {{
                background: {theme['accent']};
                color: white;
                border: 2px solid {theme['glow']};
            }}

            /* Text Editor with Premium Styling */
            QTextEdit {{ 
                background-color: {theme['edit_bg']}; 
                color: black; 
                border: 2px solid {theme['border']}; 
                border-radius: 10px;
                padding: 16px; 
                line-height: 1.6;
                selection-background-color: {theme['accent']};
                selection-color: white;
            }}

            /* Modern Sliders with Enhanced Handles */
            QSlider::groove:horizontal {{ 
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {theme['border']},
                    stop:1 {theme['accent']});
                height: 8px; 
                border-radius: 4px; 
            }}
            
            QSlider::handle:horizontal {{ 
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 white,
                    stop:1 {theme['accent']}); 
                width: 22px; 
                height: 22px;
                margin: -7px 0; 
                border-radius: 11px; 
                border: 3px solid {theme['glow']};
            }}
            
            QSlider::handle:horizontal:hover {{
                width: 26px;
                height: 26px;
                margin: -9px 0;
                border-radius: 13px;
            }}
            
            QSlider::groove:vertical {{ 
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {theme['accent']},
                    stop:1 {theme['border']});
                width: 8px; 
                border-radius: 4px; 
            }}
            
            QSlider::handle:vertical {{ 
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 white,
                    stop:1 {theme['accent']}); 
                width: 22px; 
                height: 22px;
                margin: 0 -7px; 
                border-radius: 11px; 
                border: 3px solid {theme['glow']};
            }}
            
            /* Menu Styling with Glassmorphism (Solid fallback for stability) */
            QMenu {{ 
                background-color: {menu_bg}; 
                color: {theme['text']}; 
                border: 2px solid {theme['border']}; 
                border-radius: 8px;
                padding: 8px;
            }}
            
            QMenu::item {{ 
                padding: 10px 30px; 
                border-radius: 6px; 
                margin: 2px;
            }}
            
            QMenu::item:selected {{ 
                background-color: {theme['accent']}; 
                color: white; 
            }}
            
            /* MenuBar Styling */
            QMenuBar {{
                background-color: {theme['main_bg']};
                color: {theme['text']};
                border-bottom: 1px solid {theme['border']};
                padding: 4px;
            }}
            
            QMenuBar::item {{
                padding: 6px 12px;
                border-radius: 4px;
            }}
            
            QMenuBar::item:selected {{
                background-color: {theme['accent']};
                color: white;
            }}
            
            /* Tooltip Styling */
            QToolTip {{
                background-color: {menu_bg};
                color: {theme['text']};
                border: 2px solid {theme['border']};
                padding: 8px 12px;
                border-radius: 6px;
                font-size: 12px;
            }}
            
            /* ComboBox Styling */
            QComboBox {{
                background-color: {theme['edit_bg']};
                color: black;
                border: 2px solid {theme['border']};
                border-radius: 6px;
                padding: 5px 10px;
            }}
            
            QComboBox:hover {{
                border: 2px solid {theme['accent']};
            }}
            
            QComboBox::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border: none;
            }}
            
            
            QComboBox::down-arrow {{
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid {theme['text']};
                margin-top: 2px;
            }}
            QComboBox QAbstractItemView {{
                background-color: {theme['edit_bg']};
                color: black;
                selection-background-color: {theme['accent']};
                selection-color: white;
            }}
            
            /* QListWidget Styling (for dialogs) */
            QListWidget {{
                background-color: {theme['edit_bg']};
                color: black;
                border: 2px solid {theme['border']};
                border-radius: 6px;
            }}
            
            QListWidget::item {{
                padding: 5px;
                color: black;
            }}
            
            QListWidget::item:selected {{
                background-color: {theme['accent']};
                color: white;
            }}
            
            QListWidget::item:hover {{
                background-color: {theme['border']};
            }}
            
            /* QDialog Styling - Use main background, not editor background */
            QDialog {{
                background-color: {theme['main_bg']};
                color: {theme['text']};
            }}
            
            /* QLineEdit Styling */
            QLineEdit {{
                background-color: {theme['edit_bg']};
                color: black;
                border: 2px solid {theme['border']};
                border-radius: 4px;
                padding: 5px;
            }}
            
            QLineEdit:focus {{
                border: 2px solid {theme['accent']};
            }}
            
            QDialog QTextEdit {{
                background-color: {theme['edit_bg']};
                color: black;
                border: 2px solid {theme['border']};
                border-radius: 6px;
                padding: 8px;
            }}
            
            /* QSpinBox and QDoubleSpinBox */
            QSpinBox, QDoubleSpinBox {{
                background-color: {theme['edit_bg']};
                color: black;
                border: 2px solid {theme['border']};
                border-radius: 4px;
                padding: 3px;
            }}
            
            /* Premium ToolBar Styling - Floating Glass Card */
            QToolBar {{
                background: {theme['glass']};
                border: 1px solid {theme['border']};
                border-radius: 12px;
                spacing: 12px;
                padding: 8px;
                margin: 10px;
            }}
            
            QToolButton {{
                background: transparent;
                border-radius: 8px;
                padding: 6px;
                min-width: 30px;
            }}
            
            QToolButton:hover {{
                background: {theme['glass']};
                border: 1px solid {theme['accent']};
            }}

            /* Modern ScrollBars */
            QScrollBar:vertical {{
                border: none;
                background: transparent;
                width: 10px;
                margin: 0px 0px 0px 0px;
            }}
            QScrollBar::handle:vertical {{
                background: {theme['accent']};
                min-height: 30px;
                border-radius: 5px;
                margin: 2px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: {theme['glow']};
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0px;
            }}
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
                background: transparent;
            }}
            
            QScrollBar:horizontal {{
                border: none;
                background: transparent;
                height: 10px;
                margin: 0px 0px 0px 0px;
            }}
            QScrollBar::handle:horizontal {{
                background: {theme['accent']};
                min-width: 30px;
                border-radius: 5px;
                margin: 2px;
            }}
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
                width: 0px;
            }}
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{
                background: transparent;
            }}
            
            /* DockWidget Styling */
            QDockWidget {{
                color: {theme['text']};
                titlebar-close-icon: url(close.png);
                titlebar-normal-icon: url(float.png);
            }}
            
            QDockWidget::title {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {theme['accent']},
                    stop:1 {theme['glow']});
                color: white;
                padding: 8px;
                border-radius: 6px 6px 0 0;
                font-weight: bold;
            }}

            /* QCheckBox Styling - High Visibility */
            QCheckBox {{
                spacing: 12px;
                color: {theme['text']};
            }}
            QCheckBox::indicator {{
                width: 18px;
                height: 18px;
                border: 2px solid {theme['border']};
                border-radius: 4px;
                background-color: {theme['edit_bg']};
            }}
            QCheckBox::indicator:hover {{
                border: 2px solid {theme['accent']};
            }}
            QCheckBox::indicator:checked {{
                background-color: {theme['accent']};
                border: 2px solid {theme['glow']};
            }}
        """
        # Apply to global application instance so dialogs and popups follow theme even if parented differently
        QApplication.instance().setStyleSheet(style)
        
        # Premium Global Ribbon Gradient
        # We reuse the theme's top_bar color for the ribbon now it's global
        if hasattr(self, 'format_bar'):
            self.format_bar.setStyleSheet(f"""
                QToolBar {{
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, 
                        stop:0 {theme['top_bar']}, 
                        stop:0.5 {theme['accent']},
                        stop:1 {theme['glow']}); 
                    border-bottom: 2px solid {theme['border']};
                    padding: 2px;
                }}
                QToolButton {{ color: white; border-radius: 4px; padding: 4px; margin: 2px; }}
                QToolButton:hover {{ background: rgba(255, 255, 255, 0.2); }}
                QComboBox, QFontComboBox {{ border-radius: 4px; padding: 2px; }}
                QPushButton {{ 
                    background: rgba(255, 255, 255, 0.1); 
                    color: white; 
                    border: 1px solid rgba(255,255,255,0.2); 
                    padding: 4px 8px; 
                    border-radius: 4px;
                    font-size: 11px;
                }}
                QPushButton:hover {{ background: rgba(255, 255, 255, 0.25); }}
                QPushButton::menu-indicator {{ image: none; }}
            """)
        
        from utils import get_contrast_color
        
        # Premium Video Frame with Shadow
        self.video_frame.setStyleSheet(f"""
            background-color: #000; 
            border: 3px solid {theme['accent']}; 
            border-radius: 12px;
        """)
        
        # Enhanced Time Labels
        for label in [self.lbl_curr, getattr(self, 'lbl_mini_curr', None)]:
            if label:
                label.setStyleSheet(f"""
                    color: {theme['glow']}; 
                    font-weight: 800; 
                    font-size: 12px; 
                    font-family: 'Consolas', 'Courier New', monospace;
                """)
                
        for label in [self.lbl_total, getattr(self, 'lbl_mini_total', None)]:
            if label:
                label.setStyleSheet(f"""
                    color: {theme['text']}; 
                    font-size: 11px;
                    font-family: 'Consolas', 'Courier New', monospace;
                """)
        
        # High-visibility Waveform Zoom Buttons
        wf_btn_style = f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {theme['btn_grad'][0]},
                    stop:1 {theme['border']});
                color: {theme['text']};
                border: 2px solid {theme['accent']};
                font-weight: 900;
                font-size: 20px;
                border-radius: 4px;
            }}
            QPushButton:hover {{
                background: {theme['accent']};
                color: white;
                border: 2px solid {theme['glow']};
            }}
        """
        if hasattr(self, 'btn_zi'): 
            self.btn_zi.setStyleSheet(wf_btn_style)
            self.btn_zo.setStyleSheet(wf_btn_style)
            
        # Theme mini slider
        if hasattr(self, 'mini_slider'):
            self.mini_slider.setStyleSheet(f"""
                QSlider::groove:horizontal {{ border: 1px solid {theme['border']}; height: 4px; background: {theme['main_bg']}; margin: 2px 0; border-radius: 2px; }}
                QSlider::handle:horizontal {{ background: {theme['accent']}; border: 1px solid {theme['border']}; width: 14px; margin: -5px 0; border-radius: 7px; }}
                QSlider::handle:horizontal:hover {{ background: {theme['glow']}; }}
            """)
            
        # Player Buttons Styling (Premium Pill Shape)
        player_btn_style = f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                    stop:0 {theme['btn_grad'][0]}, 
                    stop:1 {theme['btn_grad'][2]});
                color: {theme['text']};
                border: 2px solid {theme['border']};
                border-radius: 22px;
                font-weight: 700;
                font-size: 15px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {theme['accent']},
                    stop:1 {theme['glow']});
                color: white;
                border: 2px solid {theme['glow']};
            }}
        """
        player_btns = ['btn_start', 'btn_back5', 'btn_play', 'btn_fwd5', 'btn_end', 'btn_ins_tc']
        for bname in player_btns:
            btn = getattr(self, bname, None)
            if btn: btn.setStyleSheet(player_btn_style)
        
        # Special style for play button (larger & more prominent)
        if hasattr(self, 'btn_play'):
            self.btn_play.setMinimumSize(55, 55)
            # We let update_playback_visuals handle the colors for maximum state feedback
        
        # Sync playback state visuals with new theme colors
        self.update_playback_visuals()

        # Formatting Toolbar with Glassmorphism
        if hasattr(self, 'format_bar'):
            self.format_bar.setStyleSheet(f"""
                QToolBar {{
                    background: {theme['glass']}; 
                    border-bottom: 2px solid {theme['border']};
                    spacing: 6px;
                }}
            """)
        
        # Force redraw
        self.update()
        
        # Force editor to update
        self.editor.setStyleSheet(self.editor.styleSheet())

    def apply_saved_playback_settings(self):
        """Apply saved playback settings to UI controls"""
        # Volume
        saved_volume = self.playback_config.get('volume', 70)
        saved_boost = self.playback_config.get('boost', 0)
        self.sl_vol.setValue(saved_volume)
        self.sl_boost.setValue(saved_boost)
        self.volume_label.setText(f"{saved_volume}%")
        self.boost_label.setText(f"Boost: {saved_boost}%")
        self.engine.set_volume(saved_volume + saved_boost)
        
        # Speed
        saved_speed = self.playback_config.get('speed', 100)
        self.sl_rate.setValue(saved_speed)
        rate = saved_speed / 100.0
        self.speed_label.setText(f"{rate:.2f}x")
        self.engine.set_rate(rate)
        
        # Waveform amplitude zoom
        saved_amp_zoom = self.playback_config.get('waveform_amplitude_zoom', 10)
        self.zoom_slider.setValue(saved_amp_zoom)
        self.waveform.set_amplitude_zoom(saved_amp_zoom / 10.0)

    def set_theme(self, theme_name):
        self.current_theme = theme_name
        self.config['ui']['theme'] = theme_name
        self.apply_theme()
        self.save_config()
        
        # Sync Ribbon Ticks
        if hasattr(self, 'ribbon_theme_group'):
            for act in self.ribbon_theme_group.actions():
                if act.text() == theme_name:
                    act.setChecked(True)
                    break
        
        # Sync Main Menu Ticks
        if hasattr(self, 'theme_group'):
            for act in self.theme_group.actions():
                if act.text() == theme_name:
                    act.setChecked(True)
                    break

    def go_to_start(self):
        self.engine.seek(0)
        self.editor.setFocus()
        
    def go_to_end(self):
        self.engine.seek(self.engine.get_duration())
        self.editor.setFocus()
        
    def skip_back(self):
        self.engine.seek_relative(-5000)
        self.editor.setFocus()
        
    def skip_forward(self):
        self.engine.seek_relative(5000)
        self.editor.setFocus()

    def on_speed_changed(self, value):
        """Handle speed slider change"""
        rate = value / 100.0
        self.engine.set_rate(rate)
        self.waveform.set_playback_rate(rate)
        self.speed_label.setText(f"{rate:.2f}x")
        # Save to config
        self.playback_config['speed'] = value
        self.config['playback'] = self.playback_config
        self.save_config()

    def on_volume_changed(self, value=None):
        """Handle volume or boost slider change"""
        vol = self.sl_vol.value()
        boost = self.sl_boost.value()
        total = vol + boost
        self.engine.set_volume(total)
        self.volume_label.setText(f"{vol}%")
        self.boost_label.setText(f"Boost: {boost}%")
        
        # Save to config
        self.playback_config['volume'] = vol
        self.playback_config['boost'] = boost
        self.config['playback'] = self.playback_config
        self.save_config()

    def on_waveform_zoom_changed(self, value):
        zoom = value / 10.0
        self.waveform.set_amplitude_zoom(zoom)
        # Save to config
        self.playback_config['waveform_amplitude_zoom'] = value
        self.config['playback'] = self.playback_config
        self.save_config()
        
    def _on_slider_clicked(self, value):
        # Only seek if a slider was clicked/changed not by timer
        sender = self.sender()
        is_main = (sender == self.slider or self.slider.isSliderDown())
        is_mini = (hasattr(self, 'mini_slider') and (sender == self.mini_slider or self.mini_slider.isSliderDown()))
        
        if is_main or is_mini:
             self.engine.seek(value)
             self.last_seek_time = time.time()
             # INSTANT FEEDBACK: Force UI update while paused
             self.on_position_changed(value, force=True)
             self.editor.setFocus()

    def _on_waveform_seek(self, ms):
        self.engine.seek(ms)
        self.last_seek_time = time.time()
        # INSTANT FEEDBACK
        self.on_position_changed(ms, force=True)
        self.editor.setFocus()

    def on_engine_selected(self, action):
        engine_type = action.data()
        
        # Check MPV availability if selected
        if engine_type == "mpv":
            from media_engine import MPVBackend
            def check_mpv():
                found = MPVBackend.discover_dlls()
                return len(found) > 0

            if not check_mpv():
                ans = QMessageBox.question(self, "MPV Not Found", 
                    "MPV engine (mpv-1.dll) was not found in the application directory.\n\n"
                    "Would you like to download and install it automatically?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                
                if ans == QMessageBox.StandardButton.Yes:
                    self.start_mpv_download()
                else:
                    self.vlc_act.setChecked(True)
                return

        self.statusBar().showMessage(f"Switching engine to: {engine_type.upper()}...", 2000)
        
        # Save current state for continuation
        pos = self.engine.get_time()
        was_playing = self.engine.is_playing()
        
        # NUCLEAR OPTION: Hide the video container to force OS to detach whatever is rendering there
        self.video_container.hide()
        
        # Stop current playback and process events to ensure clean release
        self.engine.stop()
        self.stop_waveform_worker() # Ensure worker is dead
        
        # Disconnect any lingering position updates from UI to avoid processing on semi-dead backend
        try: self.engine.positionChanged.disconnect(self.on_position_changed)
        except: pass
        
        from PyQt6.QtCore import QCoreApplication
        QCoreApplication.processEvents()
        time.sleep(0.15) # Brief pause to allow OS and threads to settle
        
        # Perform the actual switch
        self.engine.set_backend(engine_type)
        self.config['engine'] = engine_type
        self.save_config()
        
        # RE-CONNECT UI SIGNALS to new backend
        self.engine.positionChanged.connect(self.on_position_changed)
        
        # Another processEvents to ensure the new backend is fully initialized
        QCoreApplication.processEvents()
        
        # Update visuals IMMEDIATELY to show the engine is stopped (PLAY icon ▶)
        self.update_playback_visuals(force_playing=False)
        
        # If media is loaded, reload it in the new engine
        if hasattr(self, 'current_media_path') and self.current_media_path:
            # Re-show container BEFORE setting display handle
            self.video_container.show()
            QCoreApplication.processEvents()
            
            # Re-set display handle explicitly before loading (AS INT)
            win_id = int(self.video_frame.winId())
            self.engine.set_display_handle(win_id)
            
            # Brief delay before loading to ensure OS has mapped the frame
            QTimer.singleShot(200, lambda: self.load_media_file(self.current_media_path))
            
            # Restore position after reload
            def final_restore():
                self.engine.seek(pos)
                # ALWAYS PAUSE after switch as requested
                self.engine.pause()
                self.update_playback_visuals(force_playing=False)
                
                # Clear focus from buttons to remove highlighting
                self.editor.setFocus()
            
            # Reduced delay for snappier feel but still enough for stability
            QTimer.singleShot(1200, final_restore)
        else:
            self.video_container.show()
            self.update_playback_visuals(force_playing=False)
            self.editor.setFocus()

    def start_mpv_download(self):
        self.progress = QProgressDialog("Downloading MPV Engine...", "Cancel", 0, 100, self)
        self.progress.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress.show()
        
        def on_progress(val):
            self.progress.setValue(val)
            if self.progress.wasCanceled():
                pass # Logic to stop download if needed
                
        def do_download():
            result = self.engine.download_mpv(on_progress)
            if result is True:
                QMessageBox.information(self, "Success", "MPV Engine installed! You can now select it from the menu.")
                self.mpv_act.setEnabled(True)
                self.mpv_act.setChecked(True)
                self.on_engine_selected(self.mpv_act)
            elif result == "PERMISSION_DENIED":
                QMessageBox.critical(self, "Permission Denied", 
                    "TranscriptFlow does not have permission to write to the application folder.\n\n"
                    "Please run the application as ADMINISTRATOR (Right-click -> Run as Administrator) to download MPV automatically.")
                self.vlc_act.setChecked(True)
            else:
                manual_link = "https://sfile.co/SjmPjUXcgvO"
                QMessageBox.critical(self, "Error", 
                    "Failed to download MPV automatically.\n"
                    "Please run the application as ADMINISTRATOR (Right-click -> Run as Administrator) to download MPV automatically.\n"
                    f"Or... Please download manually from: {manual_link}\n"
                    "Then rename the file to 'mpv-1.dll' and place it in the application folder.")
                self.vlc_act.setChecked(True)

        QTimer.singleShot(100, do_download)

    def add_action_safe(self, menu, text, slot, shortcut=None):
        """Helper to safely add actions with shortcuts"""
        action = QAction(text, self)
        if slot:
            action.triggered.connect(slot)
        if shortcut:
            if isinstance(shortcut, QKeySequence.StandardKey):
                action.setShortcut(shortcut)
            else:
                action.setShortcut(QKeySequence(shortcut))
        menu.addAction(action)
        return action

    def init_menus(self):
        menu = self.menuBar()
        
        # ==================== FILE MENU ====================
        file_menu = menu.addMenu("&File")
        
        self.add_action_safe(file_menu, "New Document", self.new_doc, QKeySequence.StandardKey.New)
        self.add_action_safe(file_menu, "New Document with Media File...", self.new_with_media, "Ctrl+Shift+N")
        self.add_action_safe(file_menu, "Open Document...", self.open_doc, QKeySequence.StandardKey.Open)
        
        self.recent_menu = file_menu.addMenu("Open Recent")
        self.rebuild_recent_menu()
        
        
        self.add_action_safe(file_menu, "Save", self.save_doc, QKeySequence.StandardKey.Save)
        self.add_action_safe(file_menu, "Save As...", self.save_as_doc, QKeySequence.StandardKey.SaveAs)

        
        file_menu.addSeparator()
        
        # Import Submenu
        imp_menu = file_menu.addMenu("Import")
        self.add_action_safe(imp_menu, "Plain Text...", lambda: self.import_file('txt'))
        self.add_action_safe(imp_menu, "Excel / CSV (.csv)...", lambda: self.import_file('csv'))
        self.add_action_safe(imp_menu, "Scenarist Closed Caption (SCC)...", lambda: self.import_file('scc'))
        self.add_action_safe(imp_menu, "Spruce STL Format...", lambda: self.import_file('stl'))
        self.add_action_safe(imp_menu, "Tab-delimited Text...", lambda: self.import_file('tab'))
        self.add_action_safe(imp_menu, "XML...", lambda: self.import_file('xml'))
        self.add_action_safe(imp_menu, "SubRip (.srt)...", lambda: self.import_file('srt'))
        
        # Export Submenu
        exp_main_menu = file_menu.addMenu("Export")
        
        # Subtitles / Closed Captions Group
        sub_menu = exp_main_menu.addMenu("Subtitles / Closed Captions")
        self.add_action_safe(sub_menu, "SubRip Format (.srt)...", lambda: self.export_file('srt'))
        self.add_action_safe(sub_menu, "Scenarist Closed Caption (.scc)...", lambda: self.export_file('scc_export'))
        self.add_action_safe(sub_menu, "Spruce STL Format...", lambda: self.export_file('stl'))
        
        # Transcript / Data Group
        trans_menu = exp_main_menu.addMenu("Transcript / Data")
        self.add_action_safe(trans_menu, "Open Document Format (.odt)...", lambda: self.export_file('odf'))
        self.add_action_safe(trans_menu, "HTML...", lambda: self.export_file('html'))
        self.add_action_safe(trans_menu, "Plain Text...", lambda: self.export_file('txt'))
        self.add_action_safe(trans_menu, "Excel / CSV (.csv)...", lambda: self.export_file('csv'))
        self.add_action_safe(trans_menu, "Tab-delimited Text...", lambda: self.export_file('tab'))
        self.add_action_safe(trans_menu, "XML...", lambda: self.export_file('xml'))
        trans_menu.addSeparator()
        self.add_action_safe(trans_menu, "Final Cut Pro XML...", lambda: self.export_file('fcpxml'))
        self.add_action_safe(trans_menu, "Final Cut Pro Markers...", lambda: self.export_file('fcpmarkers'))
        
        file_menu.addSeparator()
        
        self.add_action_safe(file_menu, "Page Setup...", self.page_setup)
        self.add_action_safe(file_menu, "Print...", self.print_doc, QKeySequence.StandardKey.Print)
        
        file_menu.addSeparator()
        self.add_action_safe(file_menu, "Exit", self.close)

        # ==================== EDIT MENU ====================
        edit_menu = menu.addMenu("&Edit")
        
        self.add_action_safe(edit_menu, "Undo", self.editor.undo, QKeySequence.StandardKey.Undo)
        self.add_action_safe(edit_menu, "Redo", self.editor.redo, QKeySequence.StandardKey.Redo)
        edit_menu.addSeparator()
        self.add_action_safe(edit_menu, "Cut", self.editor.cut, QKeySequence.StandardKey.Cut)
        self.add_action_safe(edit_menu, "Copy", self.editor.copy, QKeySequence.StandardKey.Copy)
        self.add_action_safe(edit_menu, "Paste", self.editor.paste, QKeySequence.StandardKey.Paste)
        self.add_action_safe(edit_menu, "Delete", lambda: self.editor.textCursor().removeSelectedText(), "Delete")
        self.add_action_safe(edit_menu, "Select All", self.editor.selectAll, QKeySequence.StandardKey.SelectAll)
        edit_menu.addSeparator()
        self.add_action_safe(edit_menu, "Copy Time to Clipboard", self.copy_time, "Ctrl+Shift+C")
        self.add_action_safe(edit_menu, "Insert Time", self.insert_current_time, "Ctrl+;")
        edit_menu.addSeparator()
        
        self.spell_act = QAction("Check Spelling...", self)
        self.spell_act.setShortcut("F7") # Industry standard for spell check
        self.spell_act.triggered.connect(self.open_spell_check_dialog)
        edit_menu.addAction(self.spell_act)
        edit_menu.addSeparator()
        
        self.add_action_safe(edit_menu, "Options...", self.open_options)
        self.add_action_safe(edit_menu, "Edit Shortcuts...", self.open_shortcuts_dialog)
        self.add_action_safe(edit_menu, "Edit Snippets...", self.open_snippets_dialog)
        edit_menu.addSeparator()
        self.add_action_safe(edit_menu, "Find...", self.find_text, QKeySequence.StandardKey.Find)
        self.add_action_safe(edit_menu, "Find Next", self.find_next_silent, "F3")
        self.add_action_safe(edit_menu, "Find Previous", self.find_prev_silent, "Alt+F3")
        self.add_action_safe(edit_menu, "Change Case", self.editor.cycle_case, "Shift+F3")
        self.add_action_safe(edit_menu, "Replace...", self.replace_text, QKeySequence.StandardKey.Replace)
        edit_menu.addSeparator()
        self.add_action_safe(edit_menu, "Set Up Foot Pedal...", self.setup_foot_pedal)
        self.add_action_safe(edit_menu, "Manage USB Devices...", self.manage_usb_devices)

        # ==================== VIEW MENU ====================
        view_menu = menu.addMenu("&View")
        
        vid_size_menu = view_menu.addMenu("Video Size")
        self.add_action_safe(vid_size_menu, "50%", lambda: self.set_video_size(0.5))
        self.add_action_safe(vid_size_menu, "100%", lambda: self.set_video_size(1.0))
        self.add_action_safe(vid_size_menu, "200%", lambda: self.set_video_size(2.0))
        
        self.waveform_act = QAction("Show Waveform", self)
        self.waveform_act.setCheckable(True)
        self.waveform_act.setChecked(self.show_waveform)
        self.waveform_act.triggered.connect(self.toggle_waveform)
        view_menu.addAction(self.waveform_act)
        
        view_menu.addSeparator()
        
        # --- Aspect Ratio Submenu ---
        aspect_menu = view_menu.addMenu("Aspect Ratio")
        self.aspect_group = QActionGroup(self)
        self.aspect_group.setExclusive(True)
        
        ratios = [
            ("Default", "default"),
            ("16:9 (Widescreen)", "16:9"),
            ("4:3 (Standard)", "4:3"),
            ("1:1 (Square)", "1:1"),
            ("2.35:1 (CinemaScope)", "2.35:1"),
            ("2.39:1 (Anamorphic)", "2.39:1"),
            ("1.85:1 (Flat)", "1.85:1"),
            ("21:9 (Ultrawide)", "21:9")
        ]
        
        for name, data in ratios:
            act = QAction(name, self)
            act.setCheckable(True)
            if data == "default":
                act.setChecked(True)
            act.triggered.connect(lambda checked, r=data: self.set_aspect_ratio(r))
            aspect_menu.addAction(act)
            self.aspect_group.addAction(act)
        
        view_menu.addSeparator()
        
        # --- Layouts Submenu ---
        layout_menu = view_menu.addMenu("Layout")
        self.layout_group = QActionGroup(self)
        self.layout_group.setExclusive(True)
        
        layouts = [
            ("Standard", "standard"),
            ("Standard (No Waveform)", "standard_no_wf"),
            ("Wide Player", "wide"),
            ("Reversed (Editor Left)", "reversed"),
            ("Reversed (No Waveform)", "reversed_no_wf"),
            ("Audio Only", "audio_only"),
            ("Stacked (Top-Bottom)", "stacked"),
            ("Focus Mode (Editor Only)", "focus"),
            ("Compact Player", "compact")
        ]
        
        for name, data in layouts:
            act = QAction(name, self)
            act.setCheckable(True)
            act.setData(data)
            if data == self.config['ui'].get('layout', 'standard'):
                act.setChecked(True)
            act.triggered.connect(lambda checked, d=data: self.apply_layout_preset(d))
            layout_menu.addAction(act)
            self.layout_group.addAction(act)
        
        view_menu.addSeparator()
        
        # --- Multi-Theme Submenu ---
        self.themes_menu = view_menu.addMenu("Themes")
        self.update_theme_menu()
        
        view_menu.addSeparator()
        
        self.rtl_act = QAction("Right-to-Left (RTL) Mode", self)
        self.rtl_act.setCheckable(True)
        self.rtl_act.setChecked(self.rtl_mode)
        self.rtl_act.triggered.connect(self.toggle_rtl)
        view_menu.addAction(self.rtl_act)
        
        view_menu.addSeparator()

        # ==================== MEDIA MENU ====================
        media_menu = menu.addMenu("&Media")
        self.add_action_safe(media_menu, "Select Media...", self.open_media, "Ctrl+D")
        media_menu.addSeparator()
        self.add_action_safe(media_menu, "Reload Media", self.reload_media, "Ctrl+R")
        self.add_action_safe(media_menu, "Go to Time...", self.go_to_time, "Ctrl+T")
        self.add_action_safe(media_menu, "Set Media Offset...", self.set_media_offset)
        media_menu.addSeparator()
        
        # Engine selection
        engine_group = QActionGroup(self)
        
        # MPV Engine (Default)
        import ctypes.util
        def check_mpv():
            dll_names = ["libmpv-2.dll", "mpv-2.dll", "mpv-1.dll", "mpv.dll"]
            base_dir = os.path.dirname(os.path.abspath(__file__))
            for name in dll_names:
                if ctypes.util.find_library(name): return True
                local_path = os.path.join(base_dir, name)
                if os.path.exists(local_path):
                    try:
                        ctypes.CDLL(local_path)
                        return True
                    except: continue
            return False
            
        mpv_exists = check_mpv()
        
        self.mpv_act = QAction("Engine: MPV (Recommended)", self)
        self.mpv_act.setCheckable(True)
        self.mpv_act.setChecked(mpv_exists)
        self.mpv_act.setData("mpv")
        engine_group.addAction(self.mpv_act)
        media_menu.addAction(self.mpv_act)
        
        # VLC Engine
        self.vlc_act = QAction("Engine: VLC", self)
        self.vlc_act.setCheckable(True)
        self.vlc_act.setChecked(not mpv_exists)
        self.vlc_act.setData("vlc")
        engine_group.addAction(self.vlc_act)
        media_menu.addAction(self.vlc_act)
        
        # FFmpeg Engine
        self.ffmpeg_act = QAction("Engine: FFmpeg (Internal)", self)
        self.ffmpeg_act.setCheckable(True)
        self.ffmpeg_act.setData("ffmpeg")
        engine_group.addAction(self.ffmpeg_act)
        media_menu.addAction(self.ffmpeg_act)
        
        # QuickTime Engine
        self.qt_act = QAction("Engine: QuickTime", self)
        self.qt_act.setCheckable(True)
        self.qt_act.setData("quicktime")
        engine_group.addAction(self.qt_act)
        media_menu.addAction(self.qt_act)
        
        engine_group.triggered.connect(self.on_engine_selected)

        # ==================== TRANSCRIPT MENU ====================
        trans_menu = menu.addMenu("&Transcript")
        self.add_action_safe(trans_menu, "Transcript Settings...", self.open_transcript_settings)
        self.add_action_safe(trans_menu, "Adjust Time Codes...", self.adjust_timecodes)
        self.add_action_safe(trans_menu, "Word Count...", self.show_word_count)
        self.add_action_safe(trans_menu, "Sync Transcript With Current Time", self.sync_transcript, "Ctrl+K")

        # ==================== WINDOW MENU ====================
        win_menu = menu.addMenu("&Window")
        self.add_action_safe(win_menu, "Show Snippets", self.open_snippets_dialog)
        self.add_action_safe(win_menu, "Show Shortcuts", self.open_shortcuts_dialog)
        self.add_action_safe(win_menu, "Show Backups", self.show_backups)

        # ==================== HELP MENU ====================
        help_menu = menu.addMenu("&Help")
        self.add_action_safe(help_menu, "Documentation", self.show_documentation)
        self.add_action_safe(help_menu, "Snippet Variables Reference", self.show_variables_reference)
        self.add_action_safe(help_menu, "Check for Updates", self.check_updates)
        self.add_action_safe(help_menu, "About", self.show_about)

    # ==================== CORE FUNCTIONALITY ====================
    
    def toggle_play(self):
        """Toggle play/pause with visual button update"""
        was_playing = self.engine.is_playing()
        self.engine.play_pause()
        # Immediately update UI based on EXPECTED state change for snappy feel
        self.update_playback_visuals(force_playing=not was_playing)
        self.editor.setFocus()

    def update_playback_visuals(self, force_playing=None):
        """Syncs UI buttons and waveform state with the engine's actual state with DRAMATIC visual feedback"""
        playing = force_playing if force_playing is not None else self.engine.is_playing()
        
        # FAIL-SAFE: If no video is loaded yet, never show the pause button
        if not self.current_media_path:
            playing = False
        
        if hasattr(self, 'btn_play'):
            self.btn_play.setText(" || " if playing else " ▶ ")
            
            # --- DRAMATIC DYNAMIC VISUALS ---
            # Using absolute, high-contrast colors so the user cannot miss the change
            if playing:
                # VIBRANT SUCCESS GREEN (Active)
                c1, c2, border = "#10b981", "#34d399", "#059669"
                shadow = "rgba(16, 185, 129, 0.4)"
            else:
                # VIBRANT WARNING AMBER/ORANGE (Paused)
                c1, c2, border = "#f59e0b", "#fbbf24", "#d97706"
                shadow = "rgba(245, 158, 11, 0.3)"
            
            # Apply style DIRECTLY to ensure it wins over global theme selectors
            self.btn_play.setStyleSheet(f"""
                QPushButton#PlayPauseButton {{
                    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {c1}, stop:1 {c2});
                    color: white;
                    border: 3px solid {border};
                    border-radius: 27px;
                    font-weight: 900;
                    font-size: 26px;
                }}
                QPushButton#PlayPauseButton:hover {{
                    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {c2}, stop:1 {c1});
                    border: 3px solid white;
                }}
            """)
            
            # FORCE RE-POLISH: This tells Qt to re-calculate styles immediately
            self.btn_play.style().unpolish(self.btn_play)
            self.btn_play.style().polish(self.btn_play)
            self.btn_play.update()
            
        if hasattr(self, 'waveform'):
            self.waveform.set_playing(playing)
            self.waveform.set_playback_rate(self.sl_rate.value() / 100.0)
    
    def show_toast(self, message, duration=3000):
        self.toast_label.setText(message)
        self.toast_label.adjustSize()
        # Position at bottom center of the window
        x = (self.width() - self.toast_label.width()) // 2
        y = self.height() - self.toast_label.height() - 40
        self.toast_label.move(x, y)
        self.toast_label.show()
        self.toast_label.raise_()
        self.toast_timer.start(duration)

    def engine_command(self, cmd):
        """Helper to route button clicks to handle_command logic"""
        self.handle_command({'command': cmd, 'skip': 5.0, 'value': 0.1})


    def handle_editor_settings_change(self, updated_settings):
        """Persists on-the-fly style changes (font, size, color) to the config"""
        if self.is_initializing: return
        self.config['settings'].update(updated_settings)
        self.save_config()

    def handle_command(self, cmd_dict):
        """Dispatches commands received from Editor shortcuts"""
        cmd = cmd_dict['command']
        skip = cmd_dict.get('skip', 0)
        val = cmd_dict.get('value', 0)
        
        if cmd == 'Toggle Pause and Play':
            self.engine.skip_on_pause = int(skip * 1000)
            self.toggle_play()
        elif cmd == 'Skipback':
            self.engine.seek_relative(-int(skip * 1000))
        elif cmd == 'Fast Forward':
            self.engine.seek_relative(int(skip * 1000))
        elif cmd == 'Rewind':
            self.engine.seek_relative(-5000)
        elif cmd == 'Stop':
            self.engine.stop()
        elif cmd == 'Pause':
            self.engine.pause()
        elif cmd == 'Play':
            self.engine.play()
        elif cmd == 'Advance One Frame':
            self.engine.frame_step(True)
        elif cmd == 'Rewind One Frame':
            self.engine.frame_step(False)
        elif cmd == 'Jump To End':
            self.engine.seek(self.engine.get_duration())
        elif cmd == 'Jump To Beginning':
            self.engine.seek(0)
        elif cmd == 'Insert Current Time':
            self.insert_current_time()
        elif cmd == 'Increase Play Rate':
            new_val = min(300, self.sl_rate.value() + int(val * 100))
            self.sl_rate.setValue(new_val)
        elif cmd == 'Decrease Play Rate':
            new_val = max(25, self.sl_rate.value() - int(val * 100))
            self.sl_rate.setValue(new_val)
        elif cmd == 'Increase Volume':
            self.sl_vol.setValue(min(100, self.sl_vol.value() + int(val)))
        elif cmd == 'Decrease Volume':
            self.sl_vol.setValue(max(0, self.sl_vol.value() - int(val)))
        elif cmd == 'Go To Next Time Code':
            self.editor.go_to_timecode(forward=True)
        elif cmd == 'Go To Previous Time Code':
            self.editor.go_to_timecode(forward=False)
            
        # Ensure UI reflects any state changes (Pause/Play)
        self.update_playback_visuals()

    def handle_snippet(self, snip_dict):
        """Parses variables in snippet and inserts into editor"""
        text = snip_dict['text']
        color = snip_dict.get('color', 'black')
        
        # Get current values
        ms = self.engine.get_time()
        helper = TimecodeHelper(self.settings['fps'], self.settings.get('media_offset', 0))
        
        # Replace all variables
        text = self.replace_snippet_variables(text, ms, helper)
        
        # Insert
        self.editor.insert_processed_content(
            text, 
            color, 
            carry_format=snip_dict.get('carry_format', False),
            bold=snip_dict.get('bold', True),
            italic=snip_dict.get('italic', False),
            underline=snip_dict.get('underline', False)
        )

    def replace_snippet_variables(self, text, ms, helper):
        """Replace all snippet variables with their values"""
        import re
        from datetime import datetime
        
        # Time variables
        if '{$time}' in text or '${time}' in text:
            tc = helper.ms_to_timestamp(ms)
            timecode_format = self.settings.get('timecode_format', '[00:01:23.29]')
            if '[' in timecode_format and not tc.startswith('['):
                tc = f'[{tc}]'
            elif '[' not in timecode_format and tc.startswith('['):
                tc = tc[1:-1]
            text = text.replace('{$time}', tc).replace('${time}', tc)
        
        if '{$time_raw}' in text or '${time_raw}' in text:
            tc_raw = helper.ms_to_timestamp(ms, bracket="")
            text = text.replace('{$time_raw}', tc_raw).replace('${time_raw}', tc_raw)
        
        if '{$time_hours}' in text or '${time_hours}' in text:
            hours = str(ms // 3600000).zfill(2)
            text = text.replace('{$time_hours}', hours).replace('${time_hours}', hours)
        
        if '{$time_minutes}' in text or '${time_minutes}' in text:
            minutes = str((ms % 3600000) // 60000).zfill(2)
            text = text.replace('{$time_minutes}', minutes).replace('${time_minutes}', minutes)
        
        if '{$time_seconds}' in text or '${time_seconds}' in text:
            seconds = str((ms % 60000) // 1000).zfill(2)
            text = text.replace('{$time_seconds}', seconds).replace('${time_seconds}', seconds)
        
        if '{$time_frames}' in text or '${time_frames}' in text:
            frame = int((ms % 1000) / (1000 / self.settings['fps']))
            text = text.replace('{$time_frames}', str(frame).zfill(2)).replace('${time_frames}', str(frame).zfill(2))
        
        # Date/Clock variables
        now = datetime.now()
        
        if '{$date_short}' in text or '${date_short}' in text:
            date_str = now.strftime("%m/%d/%Y")
            text = text.replace('{$date_short}', date_str).replace('${date_short}', date_str)
        
        if '{$date_long}' in text or '${date_long}' in text:
            date_str = now.strftime("%A, %B %d, %Y")
            text = text.replace('{$date_long}', date_str).replace('${date_long}', date_str)
        
        if '{$date_abbrev}' in text or '${date_abbrev}' in text:
            date_str = now.strftime("%a, %b %d, %Y")
            text = text.replace('{$date_abbrev}', date_str).replace('${date_abbrev}', date_str)
        
        if '{$clock_short}' in text or '${clock_short}' in text:
            clock_str = now.strftime("%I:%M %p")
            text = text.replace('{$clock_short}', clock_str).replace('${clock_short}', clock_str)
        
        if '{$clock_long}' in text or '${clock_long}' in text:
            clock_str = now.strftime("%I:%M:%S %p")
            text = text.replace('{$clock_long}', clock_str).replace('${clock_long}', clock_str)
        
        # Media variables
        if '{$media_name}' in text or '${media_name}' in text:
            media_name = os.path.basename(self.current_media_path) if hasattr(self, 'current_media_path') and self.current_media_path else 'Unknown Media'
            text = text.replace('{$media_name}', media_name).replace('${media_name}', media_name)
        
        if '{$media_path}' in text or '${media_path}' in text:
            media_path = self.current_media_path if hasattr(self, 'current_media_path') and self.current_media_path else 'Unknown Path'
            text = text.replace('{$media_path}', media_path).replace('${media_path}', media_path)
        
        # Document variables
        if '{$doc_name}' in text or '${doc_name}' in text:
            doc_name = os.path.basename(self.current_file_path) if self.current_file_path else "Untitled"
            text = text.replace('{$doc_name}', doc_name).replace('${doc_name}', doc_name)
        
        if '{$doc_path}' in text or '${doc_path}' in text:
            doc_path = self.current_file_path if self.current_file_path else "Unsaved"
            text = text.replace('{$doc_path}', doc_path).replace('${doc_path}', doc_path)
        
        # Selection variable
        if '{$selection}' in text or '${selection}' in text:
            selection = self.editor.textCursor().selectedText()
            text = text.replace('{$selection}', selection).replace('${selection}', selection)
        
        # Version
        if '{$version}' in text or '${version}' in text:
            text = text.replace('{$version}', '1.0').replace('${version}', '1.0')
        
        # Handle time_offset
        offset_pattern = r'\{\$time_offset\(([^)]+)\)\}|\$\{time_offset\(([^)]+)\)\}'
        for match in re.finditer(offset_pattern, text):
            offset_str = match.group(1) or match.group(2)
            try:
                offset_ms = helper.timestamp_to_ms(offset_str)
                new_ms = ms + offset_ms
                new_tc = helper.ms_to_timestamp(new_ms)
                text = text.replace(match.group(0), new_tc)
            except:
                pass
        
        # Handle time_raw_offset
        offset_pattern = r'\{\$time_raw_offset\(([^)]+)\)\}|\$\{time_raw_offset\(([^)]+)\)\}'
        for match in re.finditer(offset_pattern, text):
            offset_str = match.group(1) or match.group(2)
            try:
                offset_ms = helper.timestamp_to_ms(offset_str)
                new_ms = ms + offset_ms
                new_tc = helper.ms_to_timestamp(new_ms, bracket="")
                text = text.replace(match.group(0), new_tc)
            except:
                pass
        
        return text

    def open_media(self):
        dlg = MediaSourceDialog(self)
        if dlg.exec():
            stype, path = dlg.get_data()
            if stype == 'file':
                self.load_media_file(path)
            elif stype:
                self.engine.set_display_handle(int(self.video_frame.winId()))
                self.engine.load_source(stype, path)
                self.show_toast(f"Mode: {stype}")

    def load_media_file(self, path):
        """Unified method to load a media file from dialog or drag-drop"""
        logger.info(f"Loading media file: {path}")
        if not path or not os.path.exists(path):
            logger.warning(f"File does not exist: {path}")
            return
            
        try:
            self.current_media_path = path
            win_id = int(self.video_frame.winId())
            logger.info(f"Video Frame WinID: {win_id}")
            
            self.engine.set_display_handle(win_id)
            self.engine.load_source('file', path)
            self.show_toast(f"Loaded: {os.path.basename(path)}")
            self.media_just_loaded = True
            logger.info(f"Successfully called engine.load_source for {path}")
        except Exception as e:
            logger.exception(f"Error during media engine loading: {e}")
            self.show_toast("Error loading media")
            return
        
        # RESET OFFSET: User wants media to start at 00:00 by default
        self.settings['media_offset'] = 0
        self.update_tc_helper()
        
        # FORCE UI RESET: Ensure labels show 00:00 immediately
        self.lbl_curr.setText(self.tc_helper.ms_to_timestamp(0, bracket="", omit_frames=False))
        self.slider.setValue(0)
        self.waveform.set_position(0)
        
        # EXPLICIT ENGINE SEEK: Ensure engine actually starts at 0
        QTimer.singleShot(300, lambda: self.engine.seek(0))
        
        # Apply saved volume and speed after a short delay to ensure media is loaded
        def apply_playback():
            vol = self.sl_vol.value()
            boost = self.sl_boost.value()
            self.engine.set_volume(vol + boost)
            self.engine.set_rate(self.sl_rate.value() / 100.0)
            
        QTimer.singleShot(300, apply_playback)
        
        # CONDITIONAL AUTO-PLAY
        def check_autoplay():
            should_play = self.config.get('autoplay_on_load', True)
            if should_play:
                self.engine.play()
                self.update_playback_visuals(force_playing=True)
            else:
                self.engine.pause()
                self.update_playback_visuals(force_playing=False)
                
        QTimer.singleShot(500, check_autoplay)
        
        # Reset Waveform UI
        self.waveform.load_data(None)
        self.stop_waveform_worker()
        
        # CONDITIONAL AUTO-GENERATE
        auto_gen = self.config.get('auto_generate_waveform', False)
        if auto_gen and self.show_waveform:
            self.trigger_waveform_generation()
        

    def trigger_waveform_generation(self):
        """Starts the waveform generation process if media is loaded"""
        media_path = getattr(self.engine, 'current_path', None)
        if not media_path or not os.path.exists(media_path):
            return
            
        # Already running?
        if hasattr(self, 'worker') and self.worker and self.worker.isRunning():
            return
            
        # Already loaded?
        if self.waveform.data is not None:
            return
            
        self.stop_waveform_worker()
        retention = self.config.get('waveform_retention_months', 3)
        self.worker = WaveformWorker(media_path, retention_months=retention)
        self.worker.finished.connect(self.waveform.load_data)
        self.worker.start()
        logger.info(f"Waveform generation triggered for {media_path}")

    def stop_waveform_worker(self):
        """Safely shuts down the waveform worker if it's running"""
        if hasattr(self, 'worker') and self.worker:
            try: self.worker.finished.disconnect()
            except: pass
            if self.worker.isRunning():
                self.worker.terminate()
                self.worker.wait()
            self.worker = None

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        logger.info(f"Drop event: {len(files)} files")
        if files:
            logger.info(f"Dropped file: {files[0]}")
            # Load the first file dropped
            self.load_media_file(files[0])

    def update_playback_skip(self):
        for sc in self.shortcuts:
            if sc['command'] == 'Toggle Pause and Play':
                self.engine.skip_on_pause = int(sc.get('skip', 0) * 1000)

    def on_position_changed(self, ms, force=False):
        # Cooldown: Ignore engine updates for 500ms after a manual seek
        # to prevent jumping back to old position while engine is seeking
        # 'force' bypasses this for instant manual selection feedback
        if not force and (time.time() - self.last_seek_time < 0.5):
            return
            
        if not self.slider.isSliderDown():
            self.slider.blockSignals(True)
            self.slider.setValue(ms)
            self.slider.blockSignals(False)
            
        if hasattr(self, 'mini_slider') and not self.mini_slider.isSliderDown():
            self.mini_slider.blockSignals(True)
            self.mini_slider.setValue(ms)
            self.mini_slider.blockSignals(False)
        
        # Sync playing state and position
        self.update_playback_visuals()
        if hasattr(self, 'waveform'):
            self.waveform.set_position(ms)
            
        # Use cached helper for performance during high-frequency pulses
        tc = self.tc_helper.ms_to_timestamp(ms, bracket="", omit_frames=self.settings.get('omit_frames', False))
        self.lbl_curr.setText(tc)
        if hasattr(self, 'lbl_mini_curr'):
            self.lbl_mini_curr.setText(tc)

    def on_duration_changed(self, ms):
        self.slider.setMaximum(ms)
        if hasattr(self, 'mini_slider'):
            self.mini_slider.setMaximum(ms)
        self.waveform.set_duration(ms)
        tc_dur = self.tc_helper.ms_to_timestamp(ms, bracket="", omit_frames=self.settings.get('omit_frames', False))
        self.lbl_total.setText(tc_dur)
        if hasattr(self, 'lbl_mini_total'):
            self.lbl_mini_total.setText(tc_dur)

    def update_tc_helper(self):
        """Recreates the cached timecode helper when settings change"""
        self.tc_helper = TimecodeHelper(self.settings.get('fps', 23.976), self.settings.get('media_offset', 0))

    def insert_current_time(self):
        ms = self.engine.get_time()
        
        # Determine format parameters
        # Default now includes frames with : separator
        tc_format_str = self.settings.get('timecode_format', '[00:01:23:29]')
        omit_frames = self.settings.get('omit_frames', False)
        
        # Detect separator from format string (usually : or .)
        # If user has a custom format string like [00:00:00.00], we respect the dot.
        sep = ":"
        if "." in tc_format_str: sep = "."
        
        # Get brackets from format string
        brackets = ""
        if tc_format_str.startswith('[') and tc_format_str.endswith(']'): brackets = "[]"
        elif tc_format_str.startswith('(') and tc_format_str.endswith(')'): brackets = "()"
        elif tc_format_str.startswith('{') and tc_format_str.endswith('}'): brackets = "{}"
        elif tc_format_str.startswith('<') and tc_format_str.endswith('>'): brackets = "<>"
        
        helper = TimecodeHelper(self.settings['fps'], self.settings.get('media_offset', 0))
        tc = helper.ms_to_timestamp(ms, bracket=brackets, omit_frames=omit_frames, use_frames_sep=sep)
        
        tc_color = self.settings.get('timecode_color', 'green')
        style = self.settings.get('timecode_style', 'Bold')
        
        self.editor.insert_processed_content(tc + " ", tc_color, 
                                             bold=(style == "Bold"),
                                             italic=(style == "Italic"),
                                             underline=(style == "Underline"))

    def copy_time(self):
        ms = self.engine.get_time()
        tc = TimecodeHelper(self.settings['fps'], self.settings.get('media_offset', 0)).ms_to_timestamp(ms)
        QApplication.clipboard().setText(tc)
        self.statusBar().showMessage(f"Copied: {tc}", 2000)

    # ==================== FILE OPERATIONS ====================
    
    def new_doc(self):
        if self.is_dirty or self.editor.document().isModified():
            # ... UI prompt ...
            reply = QMessageBox.question(
                self, "Unsaved Changes", 
                "You have unsaved changes. Clear transcript and start new?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return
        self.editor.clear()
        self.current_file_path = None
        self.setWindowTitle("TranscriptFlow Pro - Untitled")
        self.is_dirty = False
        self.editor.document().setModified(False)
        
        # New Requirement: Show settings on new doc
        self.open_transcript_settings()
    
    def new_with_media(self):
        self.new_doc()
        self.open_media()
    
    def open_doc(self):
        last_dir = self.config.get('last_export_dir', os.path.expanduser("~"))
        f, _ = QFileDialog.getOpenFileName(
            self, "Open Document", last_dir, 
            "TranscriptFlow Files (*.tflow);;HTML Files (*.html);;Text Files (*.txt)"
        )
        if f:
            self.config['last_export_dir'] = os.path.dirname(f)
            self.save_config()
            if f.endswith('.tflow'):
                data = FileManager.load_tflow(f)
                if data:
                    self.restore_state(data)
                    self.current_file_path = f
                    self.setWindowTitle(f"TranscriptFlow Pro - {os.path.basename(f)}")
            else:
                with open(f, 'r', encoding='utf-8') as file:
                    content = file.read()
                    if f.endswith('.html') or '<html>' in content:
                        self.editor.setHtml(content)
                    else:
                        self.editor.setPlainText(content)
                self.current_file_path = f
                self.setWindowTitle(f"TranscriptFlow Pro - {os.path.basename(f)}")
            
            # Reset modified flag after load
            self.is_dirty = False
            self.editor.document().setModified(False)
            
            # Update recent files
            self.update_recent_files(f)

    def _check_cli_args(self):
        """Check if a file was passed via command line and load it"""
        if len(sys.argv) > 1:
            f = sys.argv[1]
            if os.path.exists(f) and f.endswith('.tflow'):
                logger.info(f"CLI: Opening file {f}")
                data = FileManager.load_tflow(f)
                if data:
                    self.restore_state(data)
                    self.current_file_path = f
                    self.setWindowTitle(f"TranscriptFlow Pro - {os.path.basename(f)}")
                    self.update_recent_files(f)
                    self.is_dirty = False
                    self.editor.document().setModified(False)

    def update_recent_files(self, file_path):
        """Add file to recent list and rebuild menu"""
        if not file_path:
            return
            
        file_path = os.path.abspath(file_path)
        
        # Remove if already in list to move to top
        if file_path in self.recent_files:
            self.recent_files.remove(file_path)
            
        # Add to top
        self.recent_files.insert(0, file_path)
        
        # Keep only top 10
        self.recent_files = self.recent_files[:10]
        
        # Save to config
        self.config['recent_files'] = self.recent_files
        self.save_config()
        
        # UI Update
        self.rebuild_recent_menu()

    def rebuild_recent_menu(self):
        """Clear and repopulate the recent files menu"""
        if not hasattr(self, 'recent_menu'):
            return
            
        self.recent_menu.clear()
        
        if not self.recent_files:
            self.recent_menu.addAction("(No recent files)").setEnabled(False)
            return
            
        for f in self.recent_files:
            if os.path.exists(f):
                action = self.recent_menu.addAction(os.path.basename(f))
                action.setToolTip(f)
                action.triggered.connect(lambda checked, path=f: self._load_recent_file(path))
            else:
                # Optional: Remove non-existent files?
                pass
                
        if self.recent_files:
            self.recent_menu.addSeparator()
            self.recent_menu.addAction("Clear List", self._clear_recent_files)

    def _load_recent_file(self, path):
        """Load a file from the recent files menu"""
        if not os.path.exists(path):
            QMessageBox.warning(self, "File Not Found", f"The file '{path}' no longer exists.")
            self.recent_files.remove(path)
            self.update_recent_files(None)
            return
            
        data = FileManager.load_tflow(path)
        if data:
            self.restore_state(data)
            self.current_file_path = path
            self.setWindowTitle(f"TranscriptFlow Pro - {os.path.basename(path)}")
            self.update_recent_files(path)
            self.is_dirty = False
            self.editor.document().setModified(False)

    def _clear_recent_files(self):
        """Clear the entire recent files list"""
        self.recent_files = []
        self.config['recent_files'] = []
        self.save_config()
        self.rebuild_recent_menu()

    def save_doc(self):
        if self.current_file_path:
            if self.current_file_path.endswith('.tflow'):
                data = self.capture_state()
                if FileManager.save_tflow(self.current_file_path, data):
                    self.statusBar().showMessage(f"Saved: {self.current_file_path}", 2000)
                    self.is_dirty = False
                    self.is_backup_dirty = False
                    self.editor.document().setModified(False)
                    self.update_recent_files(self.current_file_path)
            else:
                with open(self.current_file_path, 'w', encoding='utf-8') as f:
                    f.write(self.editor.toHtml())
                self.statusBar().showMessage(f"Saved: {self.current_file_path}", 2000)
                self.is_dirty = False
                self.is_backup_dirty = False
                self.editor.document().setModified(False)
        else:
            self.save_as_doc()

    def save_as_doc(self):
        import datetime
        last_dir = self.config.get('last_export_dir', os.path.expanduser("~"))
        
        # Priority: If media just loaded and we haven't saved since, use media dir
        if self.media_just_loaded and hasattr(self, 'current_media_path') and self.current_media_path:
            media_dir = os.path.dirname(self.current_media_path)
            if os.path.exists(media_dir):
                last_dir = media_dir
        
        # Generate default filename
        if self.current_file_path:
            default_name = os.path.basename(self.current_file_path)
        elif hasattr(self, 'current_media_path') and self.current_media_path:
            # Suggest name based on media
            media_base = os.path.splitext(os.path.basename(self.current_media_path))[0]
            default_name = f"{media_base}.tflow"
        else:
            default_name = f"transcript_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.tflow"
        
        f, _ = QFileDialog.getSaveFileName(
            self, "Save As", os.path.join(last_dir, default_name), 
            "TranscriptFlow Files (*.tflow);;HTML Files (*.html);;Text Files (*.txt)"
        )
        if f:
            self.config['last_export_dir'] = os.path.dirname(f)
            self.media_just_loaded = False # Reset once we have a manual save
            self.save_config()
            self.current_file_path = f
            self.save_doc()
            self.setWindowTitle(f"TranscriptFlow Pro - {os.path.basename(f)}")

    def reload_media(self):
        if hasattr(self, 'current_media_path') and self.current_media_path:
            pos = self.engine.get_time()
            self.load_media_file(self.current_media_path)
            # Restore position after reload
            QTimer.singleShot(1000, lambda: self.engine.seek(pos))
            self.statusBar().showMessage("Media reloaded", 2000)
        else:
            self.statusBar().showMessage("No media loaded to reload", 2000)

    def save_current_frame(self):
        QMessageBox.information(self, "Save Frame", "Frame capture functionality coming soon.")

    def save_time_series(self):
        QMessageBox.information(self, "Save Time Series", "Time series export coming soon.")

    def save_subtitled_movie(self):
        QMessageBox.information(self, "Save Subtitled Movie", "Subtitled movie export coming soon.")

    def page_setup(self):
        dialog = QPageSetupDialog(self.printer, self)
        dialog.exec()

    def import_file(self, fmt):
        last_dir = self.config.get('last_export_dir', os.path.expanduser("~"))
        f, _ = QFileDialog.getOpenFileName(
            self, f"Import {fmt.upper()}", last_dir, 
            f"{fmt.upper()} Files (*.{fmt});;All Files (*.*)"
        )
        if not f: return
        
        self.config['last_export_dir'] = os.path.dirname(f)
        self.save_config()
        
        try:
            with open(f, 'r', encoding='utf-8', errors='ignore') as file:
                content = file.read()
            
            from utils import Importer
            fps = self.settings.get('fps', 30.0)
            
            if fmt == 'srt':
                transcript = Importer.from_srt(content, fps)
            elif fmt == 'csv':
                transcript = Importer.from_csv(content, fps)
            elif fmt == 'scc':
                transcript = Importer.from_scc(content, fps)
            elif fmt == 'stl':
                transcript = Importer.from_stl(content, fps)
            elif fmt == 'tab':
                transcript = Importer.from_tab(content, fps)
            elif fmt == 'txt':
                transcript = content
            else:
                transcript = content
                
            self.editor.setHtml("")
            self.editor.insertPlainText(transcript)
            QMessageBox.information(self, "Import", f"Successfully imported {f}")
        except Exception as e:
            QMessageBox.critical(self, "Import Error", f"Failed to import {f}: {e}")

    def export_file(self, fmt):
        show_dialog = fmt in ['html', 'odf', 'csv', 'tab']
        settings = None
        dest_path = ""
        
        last_dir = self.config.get('last_export_dir', os.path.expanduser("~"))
        default_base = "export"
        if self.current_file_path:
            default_base = os.path.splitext(os.path.basename(self.current_file_path))[0]

        if show_dialog:
            # Load persistent settings from config
            prev_settings = self.config.get('export_settings', {})
            
            dlg = ExportSettingsDialog(self, initial_format=fmt, 
                                       current_file_path=self.current_file_path,
                                       initial_settings=prev_settings)
            
            # Ensure dialog knows about last dir if not in prev_settings
            if 'target_path' not in prev_settings:
                init_target = os.path.join(last_dir, f"{default_base}.{dlg._get_ext(fmt)}")
                dlg.target_edit.setText(init_target)
            
            if not dlg.exec():
                return
            settings = dlg.get_settings()
            
            # Save settings back to config
            self.config['export_settings'] = settings
            self.save_config()
            
            dest_path = settings['target_path']
            fmt = settings['format']
        else:
            # Direct Save Dialog (SRT, STL, XML, TXT, etc.)
            ext_map = {'srt': 'srt', 'stl': 'stl', 'xml': 'xml', 'fcpxml': 'fcpxml', 'fcpmarkers': 'txt', 'txt': 'txt', 'scc_export': 'scc'}
            ext = ext_map.get(fmt, 'txt')
            
            f, _ = QFileDialog.getSaveFileName(self, "Export", os.path.join(last_dir, f"{default_base}.{ext}"), f"{fmt.upper()} Files (*.{ext})")
            if not f: return
            dest_path = f
            
            # Save the directory for next time
            self.config['last_export_dir'] = os.path.dirname(dest_path)
            self.save_config()
            
            # Default settings for background export
            settings = {
                'format': fmt,
                'export_out_points': True,
                'export_durations': False,
                'export_speakers': True,
                'export_sfx': True,
                'speaker_delimiter': ':',
                'replace_existing': False, # Default to ask for direct saves
                'encoding': 'UTF-8'
            }

        # Check overwrite logic (if not auto-replace)
        if os.path.exists(dest_path) and not settings.get('replace_existing', False):
            reply = QMessageBox.question(self, "Replace File", 
                f"File '{os.path.basename(dest_path)}' already exists. Replace it?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.No:
                return

        raw_text = self.editor.toPlainText()
        fps = self.settings.get('fps', 30.0)
        from utils import Exporter
        
        data = None
        
        try:
            if fmt == 'odf':
                data = Exporter.to_odf(self.editor, settings)
                if data is None:
                    QMessageBox.critical(self, "Export Error", "ODF export failed.")
                    return
            elif fmt == 'srt':
                data = Exporter.to_srt(raw_text, fps, rich=False, settings=settings)
            elif fmt == 'rich_srt':
                data = Exporter.to_srt(self.editor.toHtml(), fps, rich=True, settings=settings)
            elif fmt == 'csv':
                data = Exporter.to_csv(raw_text, settings)
            elif fmt == 'scc_export':
                data = Exporter.to_scc(raw_text, fps, settings)
            elif fmt == 'html':
                data = Exporter.to_html(self.editor.toHtml(), fps, settings=settings)
            elif fmt == 'fcpxml':
                data = Exporter.to_fcpxml(raw_text, settings)
            elif fmt == 'fcpmarkers':
                data = Exporter.to_fcp_markers(raw_text, settings)
            elif fmt == 'stl':
                data = Exporter.to_stl(raw_text, fps, settings)
            elif fmt == 'tab':
                data = Exporter.to_tab(raw_text, settings=settings)
            elif fmt == 'xml':
                data = f'<?xml version="1.0" encoding="UTF-8"?>\n<transcript>\n{raw_text}\n</transcript>'
            elif fmt == 'txt':
                data = raw_text
            if data is not None:
                # Map encoding name to python standard
                enc_name = settings.get('encoding', 'UTF-8')
                if "with BOM" in enc_name:
                    encoding = "utf-8-sig"
                elif "Latin-1" in enc_name:
                    encoding = "iso-8859-1"
                elif "UTF-16" in enc_name:
                    encoding = "utf-16"
                elif "ASCII" in enc_name:
                    encoding = "ascii"
                else:
                    encoding = "utf-8"

                mode = 'wb' if isinstance(data, bytes) else 'w'
                encoding = None if isinstance(data, bytes) else encoding
                
                with open(dest_path, mode, encoding=encoding) as file:
                    file.write(data)
                    
                self.statusBar().showMessage(f"Successfully exported to {os.path.basename(dest_path)}", 3000)
                
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to export: {e}")

    def print_doc(self):
        # Configure margins based on settings
        margins = self.settings.get('print_margins', {})
        units = margins.get('units', 'Inches')
        
        # Convert to points (72 per inch, ~28.35 per cm)
        factor = 72.0 if units == 'Inches' else 28.3465
        
        from PyQt6.QtGui import QPageLayout, QPageSize
        from PyQt6.QtCore import QMarginsF
        
        m_left = margins.get('left', 1.0) * factor
        m_right = margins.get('right', 1.0) * factor
        m_top = margins.get('top', 1.0) * factor
        m_bottom = margins.get('bottom', 1.0) * factor
        
        page_layout = self.printer.pageLayout()
        page_layout.setMargins(QMarginsF(m_left, m_top, m_right, m_bottom))
        self.printer.setPageLayout(page_layout)

        dialog = QPrintDialog(self.printer, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.editor.print(self.printer)

    # ==================== SETTINGS DIALOGS ====================

    def open_shortcuts_dialog(self):
        dlg = ShortcutsManagerDialog(self, self.shortcuts)
        if dlg.exec():
            self.shortcuts = dlg.shortcuts
            self.editor.shortcuts = self.shortcuts
            self.update_playback_skip()
            self.save_config()
            self.statusBar().showMessage("Shortcuts updated and saved", 2000)

    def open_snippets_dialog(self):
        dlg = SnippetsManagerDialog(self, self.snippets)
        if dlg.exec():
            self.snippets = dlg.snippets
            self.editor.snippets = self.snippets
            self.save_config()
            self.statusBar().showMessage("Snippets updated and saved", 2000)

    def open_transcript_settings(self):
        # Default settings for "Use Defaults" button
        app_defaults = {
            'font': 'Tahoma', 'size': 12, 'fps': 23.976,
            'timecode_format': '[00:01:23.29]', 'omit_frames': False,
            'media_offset': 0, 'timecode_color': 'green',
            'font_color': 'black',
            'print_margins': {'units': 'Inches', 'top': 1.0, 'bottom': 1.0, 'left': 1.0, 'right': 1.0},
            'recognize_unbracketed': True
        }
        
        dlg = TranscriptSettingsDialog(self, self.settings, app_defaults)
        if dlg.exec():
            self.settings = dlg.get_settings()
            self.editor.settings = self.settings
            self.editor.update_font()
            self.update_tc_helper()
            # Sync ribbon controls
            self.font_combo.setCurrentFont(QFont(self.settings.get('font', 'Tahoma')))
            self.size_combo.setCurrentText(str(self.settings.get('size', 12)))
            self.save_config()
            self.statusBar().showMessage("Transcript settings updated and saved", 2000)

    def open_options(self):
        # Default settings for "Use Defaults" button
        app_defaults = {
            'font': 'Tahoma', 'size': 12, 'fps': 23.976,
            'timecode_format': '[00:01:23.29]', 'omit_frames': False,
            'media_offset': 0, 'timecode_color': 'green',
            'font_color': 'black',
            'print_margins': {'units': 'Inches', 'top': 1.0, 'bottom': 1.0, 'left': 1.0, 'right': 1.0},
            'recognize_unbracketed': True
        }
        
        from dialogs import PreferencesDialog
        dlg = PreferencesDialog(self, self.config, app_defaults)
        if dlg.exec():
            # Get updated config
            self.config = dlg.get_config()
            
            # Apply immediate settings
            self.settings = self.config['settings']
            self.editor.settings = self.settings
            self.editor.update_font()
            self.update_tc_helper()
            self.apply_player_scaling()
            
            # Pitch Lock
            self.engine.set_pitch_lock(self.config.get('playback', {}).get('pitch_lock', True))
            
            # Re-initialize last_autosave_time to respect new interval
            self.last_autosave_time = datetime.datetime.now()
            
            self.save_config()
            self.statusBar().showMessage("Preferences updated and saved", 2000)

    def toggle_spell_check(self):
        # Sync the Ribbon button toggle only (Menu item is now a separate launcher)
        enabled = self.btn_spell.isChecked()
        self.editor.set_spell_check_enabled(enabled)
        # Ensure deep update to settings dict which is part of config
        self.config['settings']['spell_check'] = enabled
        self.save_config()
        msg = "Live Spell Check: ON" if enabled else "Live Spell Check: OFF"
        self.show_toast(msg)

    def open_spell_check_dialog(self):
        """Opens the full Subtitle-Edit-Style spell check dialog"""
        dlg = SpellCheckDialog(self, self.editor)
        if dlg.exec():
            # If language changed in dialog, save it
            self.config['settings']['spell_check_lang'] = self.editor.highlighter.lang
            self.save_config()

    def update_ribbon_format(self):
        fmt = self.editor.currentCharFormat()
        
        # Block signals to avoid feedback loops if we add listeners to buttons later
        self.btn_bold.blockSignals(True)
        self.btn_italic.blockSignals(True)
        self.btn_underline.blockSignals(True)
        self.font_combo.blockSignals(True)
        self.size_combo.blockSignals(True)
        
        self.btn_bold.setChecked(fmt.fontWeight() == QFont.Weight.Bold)
        self.btn_italic.setChecked(fmt.fontItalic())
        self.btn_underline.setChecked(fmt.fontUnderline())
        
        # Sync color button
        self.update_color_button_ui(fmt.foreground().color())
        
        self.font_combo.setCurrentFont(fmt.font())
        curr_size = str(int(fmt.fontPointSize())) if fmt.fontPointSize() > 0 else self.size_combo.currentText()
        self.size_combo.setCurrentText(curr_size)
        
        self.btn_bold.blockSignals(False)
        self.btn_italic.blockSignals(False)
        self.btn_underline.blockSignals(False)
        self.font_combo.blockSignals(False)
        self.size_combo.blockSignals(False)

    def setup_foot_pedal(self):
        from dialogs import FootPedalSetupDialog
        dlg = FootPedalSetupDialog(self, self.config)
        if dlg.exec():
            # Restart manager with new config if needed
            self.pedal_manager.start()
            self.save_config()

    def manage_usb_devices(self):
        from dialogs import ManageUSBDevicesDialog
        dlg = ManageUSBDevicesDialog(self)
        dlg.exec()
        
    def on_pedal_pressed(self, button_id):
        """Handle hardware pedal events"""
        # Mapping logic here
        pass

    # ==================== VIEW MENU FUNCTIONS ====================

    def set_video_size(self, scale):
        """Resizes the video player using the left panel splitter"""
        try:
            # We use a base height of 400 for the video frame at 100%
            target_video_height = int(400 * scale)
            
            # Total height available in the left panel
            total_height = self.left_panel.height()
            
            # Ensure we don't exceed reasonable limits or hide controls
            if target_video_height > total_height - 100:
                target_video_height = total_height - 100
                
            # Set splitter sizes. Second item is the controls panel.
            self.left_splitter.setSizes([target_video_height, total_height - target_video_height])
            
            self.statusBar().showMessage(f"Video size set to {int(scale*100)}%", 2000)
            self.save_config()
        except Exception as e:
            logger.error(f"Error resizing video: {e}")

    def set_aspect_ratio(self, ratio):
        """Sets the media engine aspect ratio"""
        if ratio == "default":
            self.engine.set_aspect_ratio(None)
        else:
            self.engine.set_aspect_ratio(ratio)
        self.statusBar().showMessage(f"Aspect ratio set to {ratio}", 2000)

    def toggle_timeline(self):
        self.show_timeline = self.timeline_act.isChecked()
        self.time_container.setVisible(self.show_timeline)
        self.save_config()

    def toggle_remote(self):
        self.show_remote = self.remote_act.isChecked()
        self.buttons_container.setVisible(self.show_remote)
        self.save_config()

    def toggle_playrate(self):
        self.show_playrate = self.playrate_act.isChecked()
        self.update_sliders_visibility()
        self.save_config()

    def toggle_volume_control(self):
        self.show_volume = self.volume_act.isChecked()
        self.update_sliders_visibility()
        self.save_config()

    def update_sliders_visibility(self):
        self.sliders_container.setVisible(self.show_playrate or self.show_volume)

    def toggle_waveform(self, checked=None):
        if checked is None:
            # Called from button or action without parameter
            self.show_waveform = not self.show_waveform
        else:
            self.show_waveform = checked
            
        if hasattr(self, 'waveform_act'):
            self.waveform_act.blockSignals(True)
            self.waveform_act.setChecked(self.show_waveform)
            self.waveform_act.blockSignals(False)
            
        if hasattr(self, 'btn_ribbon_wf'):
            self.btn_ribbon_wf.blockSignals(True)
            self.btn_ribbon_wf.setChecked(self.show_waveform)
            self.btn_ribbon_wf.blockSignals(False)
            
        self.waveform_dock.setVisible(self.show_waveform)
        
        # If toggled ON, and no data is loaded, try to generate
        if self.show_waveform and self.waveform.data is None:
            self.trigger_waveform_generation()
            
        self.save_config()

    def on_waveform_visibility_changed(self, visible):
        """Syncs the menu action when dock is closed via 'x'"""
        if self.is_initializing:
            return
        self.show_waveform = visible
        if hasattr(self, 'waveform_act'):
            self.waveform_act.blockSignals(True)
            self.waveform_act.setChecked(visible)
            self.waveform_act.blockSignals(False)
            
        # Trigger generation if shown
        if visible and self.waveform.data is None:
            self.trigger_waveform_generation()
            
        self.save_config()

    def on_locale_changed(self):
        """Automatically switch fonts and alignment when switching to Urdu keyboard"""
        if self.is_initializing: return
        
        locale = QApplication.inputMethod().locale()
        
        if locale.language() == QLocale.Language.Urdu:
            # 1. Switch to Urdu-specific font visually (Temporary override)
            urdu_font = self.urdu_font_family
            self.editor.set_temporary_font_override(urdu_font)
            
            # Update Ribbon UI to reflect current visual font
            self.font_combo.blockSignals(True)
            self.font_combo.setCurrentFont(QFont(urdu_font))
            self.font_combo.blockSignals(False)
            
            # 2. Force RTL for Editor specifically (BiDi cursor support)
            self.editor.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
            self.editor.setAlignment(Qt.AlignmentFlag.AlignRight)
            
            # Force current block to right alignment (prevents BiDi cursor jumping)
            cursor = self.editor.textCursor()
            block_format = cursor.blockFormat()
            block_format.setAlignment(Qt.AlignmentFlag.AlignRight)
            cursor.setBlockFormat(block_format)
            
            # 3. Ensure future blocks default to right
            text_option = self.editor.document().defaultTextOption()
            text_option.setAlignment(Qt.AlignmentFlag.AlignRight)
            self.editor.document().setDefaultTextOption(text_option)
            
            self.statusBar().showMessage(f"Urdu Keyboard Detected: Using {urdu_font}.", 4000)
        else:
            # 1. Clear temporary override and revert to standard font
            self.editor.set_temporary_font_override(None)
            standard_font = self.config['settings'].get('font', 'Tahoma')
            
            self.font_combo.blockSignals(True)
            self.font_combo.setCurrentFont(QFont(standard_font))
            self.font_combo.blockSignals(False)
            
            # 2. Synchronize with manual RTL toggle or default to LTR
            direction = Qt.LayoutDirection.RightToLeft if self.rtl_mode else Qt.LayoutDirection.LeftToRight
            alignment = Qt.AlignmentFlag.AlignRight if self.rtl_mode else Qt.AlignmentFlag.AlignLeft
            
            # Only flip global UI if it matches the manual toggle (don't flip for Urdu keyboard only)
            QApplication.setLayoutDirection(direction)
            self.editor.setLayoutDirection(direction)
            self.editor.setAlignment(alignment)
            
            text_option = self.editor.document().defaultTextOption()
            text_option.setAlignment(alignment)
            self.editor.document().setDefaultTextOption(text_option)

    def toggle_rtl(self):
        """Toggles the global application and editor layout direction for RTL languages"""
        self.rtl_mode = self.rtl_act.isChecked()
        
        direction = Qt.LayoutDirection.RightToLeft if self.rtl_mode else Qt.LayoutDirection.LeftToRight
        alignment = Qt.AlignmentFlag.AlignRight if self.rtl_mode else Qt.AlignmentFlag.AlignLeft
        
        # 1. Flip Global UI
        QApplication.setLayoutDirection(direction)
        
        # 2. Flip Editor specifically and force immediate alignment
        self.editor.setLayoutDirection(direction)
        self.editor.setAlignment(alignment)
        
        # 3. Force document-level alignment for all new blocks
        text_option = self.editor.document().defaultTextOption()
        text_option.setAlignment(alignment)
        self.editor.document().setDefaultTextOption(text_option)
        
        # 4. Handle Urdu specifically if enabled
        if self.rtl_mode:
            # If switching TO RTL, and we have Urdu keyboard, ensure font is ready
            locale = QApplication.inputMethod().locale()
            if locale.language() == QLocale.Language.Urdu:
                self.on_locale_changed()
        
        # 5. Update Status Bar
        msg = "RTL Mode Enabled (Right-Aligned)" if self.rtl_mode else "RTL Mode Disabled (Left-Aligned)"
        self.statusBar().showMessage(msg, 3000)
        
        # 6. Save to config
        self.config['ui']['rtl_mode'] = self.rtl_mode
        self.save_config()


    def display_tracks(self):
        from PyQt6.QtWidgets import QMenu
        from PyQt6.QtGui import QCursor
        
        audio = self.engine.get_audio_tracks()
        subs = self.engine.get_subtitle_tracks()
        
        # Log for the user/debugging
        logger.debug(f"Found {len(audio)} audio tracks, {len(subs)} subtitle tracks.")
        
        if not audio and not subs:
            QMessageBox.information(self, "Tracks", "No additional audio or subtitle tracks found in this media.")
            return
            
        menu = QMenu(self)
        menu.setTitle("Select Track") # For logging/debug visibility
        
        if audio:
            menu.addSection("Audio Tracks")
            for tid, name in audio:
                # Use data to ensure persistence
                act = menu.addAction(name)
                act.triggered.connect(lambda checked, i=tid: self.engine.set_audio_track(i))
        
        if subs:
            menu.addSection("Subtitle Tracks")
            for tid, name in subs:
                act = menu.addAction(name)
                act.triggered.connect(lambda checked, i=tid: self.engine.set_subtitle_track(i))
                
        # Ensuring the menu is correctly parented and displayed
        menu.popup(QCursor.pos())

    def go_to_time(self):
        t_str, ok = QInputDialog.getText(self, "Go to Time", "Enter time (HH:MM:SS or HH:MM:SS.FF):")
        if ok and t_str:
            try:
                helper = TimecodeHelper(self.settings['fps'], self.settings.get('media_offset', 0))
                ms = helper.timestamp_to_ms(t_str)
                self.engine.seek(ms)
            except:
                QMessageBox.warning(self, "Invalid Time", "Please enter time in format HH:MM:SS or HH:MM:SS.FF")

    def set_media_offset(self):
        """Allows setting a start time offset for the media (e.g. 01:00:00:00)"""
        helper = TimecodeHelper(self.settings['fps'])
        current_offset = self.settings.get('media_offset', 0)
        current_str = helper.ms_to_timestamp(current_offset, bracket="", use_offset=False)
        
        # Get actual media path from engine
        media_path = getattr(self.engine, 'current_path', None)

        dlg = MediaOffsetDialog(self, current_str, media_path)
        if dlg.exec():
            offset_str = dlg.get_offset()
            try:
                new_offset = helper.timestamp_to_ms(offset_str, use_offset=False)
                self.settings['media_offset'] = new_offset
                self.update_tc_helper()
                self.save_config()
                
                # Update UI immediately
                self.on_position_changed(self.engine.get_time())
                
                # Check if user wants to shift existing timecodes
                diff = new_offset - current_offset
                if diff != 0:
                    ans = QMessageBox.question(self, "Adjust Existing Timecodes?",
                        f"You changed the media offset by {helper.ms_to_timestamp(abs(diff), bracket='', use_offset=False)}.\n\n"
                        "Would you like to automatically adjust all existing timecodes in the transcript to reflect this change?",
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                    
                    if ans == QMessageBox.StandardButton.Yes:
                        self.editor.adjust_timecodes(diff, selection_only=False)
                        self.statusBar().showMessage(f"Offset set to {offset_str} and transcript adjusted", 3000)
                    else:
                        self.statusBar().showMessage(f"Offset set to {offset_str}", 2000)
                else:
                    self.statusBar().showMessage(f"Offset set to {offset_str}", 2000)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Invalid time format: {e}")

    def adjust_timecodes(self):
        dlg = AdjustTimecodesDialog(self)
        if dlg.exec():
            data = dlg.get_data()
            amount_str = data['amount']
            operation = data['operation']
            selection_only = data['selection_only']
            
            # Parse HH:MM:SS.FF
            try:
                fps = self.settings.get('fps', 30.0)
                helper = TimecodeHelper(fps)
                ms = helper.timestamp_to_ms(amount_str)
                
                if operation == "Subtract":
                    ms = -ms
                
                if ms == 0:
                    return
                    
                self.editor.adjust_timecodes(ms, selection_only)
                self.statusBar().showMessage(f"Adjusted timecodes by {amount_str}", 3000)
            except Exception as e:
                QMessageBox.critical(self, "Adjustment Error", f"Invalid time format or error: {e}")


    def show_word_count(self):
        txt = self.editor.toPlainText()
        words = len(txt.split())
        chars = len(txt)
        lines = len(txt.split('\n'))
        QMessageBox.information(
            self, "Word Count", 
            f"Words: {words}\nCharacters: {chars}\nLines: {lines}"
        )

    def sync_transcript(self):
        """Aligns transcript timecodes with current playback position using a pivot."""
        # Current time as DISPLAYED to the user
        current_video_ms = self.engine.get_time()
        current_display_ms = current_video_ms + self.tc_helper.offset_ms
        tc_display_str = self.tc_helper.ms_to_timestamp(current_video_ms, bracket="", omit_frames=False)
        
        dlg = SyncTranscriptDialog(tc_display_str, self)
        if not dlg.exec():
            return
            
        opts = dlg.get_options()
        scope = opts['scope']
        
        doc = self.editor.document()
        # Find the first timecode in scope to act as PIVOT
        regex = QRegularExpression(self.tc_helper.get_regex())
        
        search_start = 0 if scope == 'all' else self.editor.textCursor().position()
        
        # 1. Find the PIVOT Match
        pivot_cursor = doc.find(regex, search_start)
        if pivot_cursor.isNull():
            self.show_toast("No timecodes found in the selected range.")
            return
            
        pivot_str = pivot_cursor.selectedText()
        try:
            # Parse the pivot's FACE VALUE (which is a display time)
            pivot_ms = self.tc_helper.timestamp_to_ms(pivot_str, use_offset=False)
        except:
            self.show_toast("Error parsing pivot timecode.")
            return
            
        # Calculate Delta: (Where video IS) - (Where pivot was)
        delta_ms = current_display_ms - pivot_ms
        
        if delta_ms == 0:
            self.show_toast("Transcript is already in sync.")
            return

        # 2. Collect all matches in scope to avoid replacement drift
        # (We still iterate backward to be safe with lengths, even with find)
        cursors_to_update = []
        curr_cursor = doc.find(regex, search_start)
        while not curr_cursor.isNull():
            # Clone cursor so we don't lose position
            cursors_to_update.append(QTextCursor(curr_cursor))
            curr_cursor = doc.find(regex, curr_cursor.position())
            
        if not cursors_to_update:
            return

        # 3. Apply the Ripple (Backward)
        main_cursor = self.editor.textCursor()
        main_cursor.beginEditBlock()
        try:
            for c in reversed(cursors_to_update):
                orig_tc_str = c.selectedText()
                try:
                    orig_ms = self.tc_helper.timestamp_to_ms(orig_tc_str, use_offset=False)
                    new_ms = max(0, orig_ms + delta_ms)
                    
                    # Detect original stylistic details
                    bracket = ""
                    if orig_tc_str.startswith('['): bracket = "[]"
                    elif orig_tc_str.startswith('('): bracket = "()"
                    elif orig_tc_str.startswith('{'): bracket = "{}"
                    elif orig_tc_str.startswith('<'): bracket = "<>"
                    sep = "." if "." in orig_tc_str else ":"
                    
                    new_tc_str = self.tc_helper.ms_to_timestamp(new_ms, bracket=bracket, use_frames_sep=sep, use_offset=False)
                    
                    # PRESERVE FORMATTING: Only replace text content within the selection
                    # The cursor 'c' already has the range selected.
                    fmt = c.charFormat()
                    c.insertText(new_tc_str)
                    
                    # Ensure the newly inserted text retains the old format
                    # insertText usually inherits, but we'll re-apply just in case of complex boundaries
                    c.setPosition(c.position() - len(new_tc_str), QTextCursor.MoveMode.KeepAnchor)
                    c.setCharFormat(fmt)
                    
                except Exception as e:
                    logger.debug(f"Sync error on specific TC: {e}")
                    continue
                    
            shift_display = f"{delta_ms/1000:+.2f}s"
            self.show_toast(f"Transcript Synced (Shift: {shift_display})")
        finally:
            main_cursor.endEditBlock()
            self.is_dirty = True

    def display_timecode_format(self):
        current_format = self.settings.get('timecode_format', '[00:01:23.29]')
        QMessageBox.information(
            self, "Time Code Format", 
            f"Current format: {current_format}\n\nChange this in Transcript Settings."
        )

    def set_transcript_width(self):
        width, ok = QInputDialog.getInt(
            self, "Set Transcript Width", 
            "Width (characters):", 80, 40, 200
        )
        if ok:
            self.statusBar().showMessage(f"Transcript width set to {width} characters", 2000)

    def show_backups(self):
        dlg = BackupSettingsDialog(self, self.backup_manager)
        if dlg.exec():
            # If restore was clicked (accepted)
            path = dlg.get_selected_backup_path()
            if path:
                data = FileManager.load_tflow(path)
                if data:
                    reply = QMessageBox.question(
                        self, "Restore Backup", 
                        "This will overwrite current changes. Continue?",
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                    )
                    if reply == QMessageBox.StandardButton.Yes:
                        self.restore_state(data)
                        self.current_file_path = None 
                        self.setWindowTitle("TranscriptFlow Pro - Restored Backup")
                        self.statusBar().showMessage("Backup restored successfully", 3000)

    def capture_state(self):
        """Captures entire application state for saving"""
        return {
            'content': self.editor.toHtml(),
            'media_path': self.current_media_path if hasattr(self, 'current_media_path') else None,
            'cursor_position': self.editor.textCursor().position(),
            'playback_position': self.engine.get_time(),
            'timestamp': datetime.datetime.now().isoformat(),
            'version': '1.0'
        }

    def restore_state(self, data):
        """Restores application state from dictionary"""
        if 'content' in data:
            self.editor.setHtml(data['content'])
            
        if 'media_path' in data and data['media_path']:
             self.load_media_file(data['media_path'])
             
        if 'playback_position' in data:
            QTimer.singleShot(500, lambda: self.engine.seek(data['playback_position']))
            
        if 'cursor_position' in data:
            cursor = self.editor.textCursor()
            cursor.setPosition(data['cursor_position'])
            self.editor.setTextCursor(cursor)
            self.editor.ensureCursorVisible()

    def perform_auto_backup(self, force=False):
        if not self.editor.toPlainText().strip():
            return
            
        now = datetime.datetime.now()
        # Responsive Backup Logic:
        # Trigger if:
        # 1. 1 minute has passed since last backup (Safety Fallback)
        # 2. OR User is idle for 3 seconds after a change (Responsive Save)
        time_since_last_backup = (now - self.last_backup_time).total_seconds()
        time_since_typing = (now - self.last_typed_time).total_seconds()
        
        interval_min = self.config.get('backup_interval', 1) # Default to 1 min security
        interval_reached = time_since_last_backup >= (interval_min * 60)
        is_idle = time_since_typing >= 3
        
        if force or is_idle or interval_reached:
            data = self.capture_state()
            # Use current filename as prefix, or 'unsaved'
            prefix = "unsaved"
            if self.current_file_path:
                prefix = os.path.splitext(os.path.basename(self.current_file_path))[0]
            
            if self.backup_manager.save_backup(data, prefix=prefix):
                self.last_backup_time = now
                self.is_backup_dirty = False # Content successfully backed up
                self.statusBar().showMessage("Background backup created.", 2000)
                # Note: We intentionally DO NOT reset self.is_dirty or 
                # self.editor.document().isModified() here so that the app
                # still prompts to save on close.
        
        # Also check for Autosave
        self.perform_autosave()

    def perform_autosave(self):
        """Automatically saves the document to its original path if enabled and interval reached."""
        if not self.config['settings'].get('autosave_enabled', False):
            return
            
        if not self.current_file_path or not self.is_dirty:
            return
            
        now = datetime.datetime.now()
        interval_min = self.config['settings'].get('autosave_interval', 5)
        
        if (now - self.last_autosave_time).total_seconds() >= (interval_min * 60):
            # perform actual save
            self.save_doc()
            self.last_autosave_time = now
            self.statusBar().showMessage(f"Autosaved: {os.path.basename(self.current_file_path)}", 3000)

    def late_initialization(self):
        """Heavy lifting that happens after the main window is stable."""
        if self.splash:
            self.splash.showMessage("Loading Media Engine...", Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignCenter, QColor("white"))
        self.engine._check_and_set_best_backend()
        
        if self.splash:
            self.splash.showMessage("Applying Themes...", Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignCenter, QColor("white"))
        self.apply_theme()
        
        self.is_initializing = False
        if self.splash:
            self.splash.finish(self)
        
        # FINAL REVEAL: Window is now themed and ready
        self.show()
        logger.info("Late initialization complete and window revealed.")

    def show_documentation(self):
        """Opens the professional HELP.html documentation file."""
        # 1. Try root directory first (ideal for installed app)
        if getattr(sys, 'frozen', False):
            install_dir = os.path.dirname(sys.executable)
            root_path = os.path.join(install_dir, "HELP.html")
            if os.path.exists(root_path):
                import webbrowser
                webbrowser.open(f"file:///{root_path}")
                return

        # 2. Try the resource path helper
        help_path = get_resource_path("HELP.html")
        if os.path.exists(help_path):
            import webbrowser
            webbrowser.open(f"file:///{help_path}")
            return
            
        # 3. Last ditch fallback
        QMessageBox.information(self, "Documentation", f"Found no HELP.html in installation directory.")

    def show_variables_reference(self):
        """Show comprehensive snippet variables reference"""
        ref_text = """
        <h2>Snippet Variable Reference</h2>
        <p>Use these variables in snippets - they'll be replaced with actual values when inserted.</p>
        
        <h3>Time Variables</h3>
        <table border="1" cellpadding="5" style="border-collapse: collapse;">
        <tr><th>Variable</th><th>Example Output</th><th>Description</th></tr>
        <tr><td><b>{$time}</b> or <b>${time}</b></td><td>[00:01:23.05]</td><td>Current timecode with formatting</td></tr>
        <tr><td><b>{$time_raw}</b></td><td>00:01:23.05</td><td>Timecode without brackets</td></tr>
        <tr><td><b>{$time_hours}</b></td><td>00</td><td>Hours component only</td></tr>
        <tr><td><b>{$time_minutes}</b></td><td>01</td><td>Minutes component only</td></tr>
        <tr><td><b>{$time_seconds}</b></td><td>23</td><td>Seconds component only</td></tr>
        <tr><td><b>{$time_frames}</b></td><td>05</td><td>Frames component only</td></tr>
        <tr><td><b>${time_offset(00:00:10.00)}</b></td><td>[00:01:33.05]</td><td>Timecode offset by amount</td></tr>
        <tr><td><b>${time_raw_offset(-00:00:05.00)}</b></td><td>00:01:18.05</td><td>Raw timecode with offset</td></tr>
        </table>
        
        <h3>Date & Clock Variables</h3>
        <table border="1" cellpadding="5" style="border-collapse: collapse;">
        <tr><th>Variable</th><th>Example Output</th><th>Description</th></tr>
        <tr><td><b>{$date_short}</b></td><td>12/31/2024</td><td>Short date format</td></tr>
        <tr><td><b>{$date_long}</b></td><td>Wednesday, December 31, 2024</td><td>Long date format</td></tr>
        <tr><td><b>{$date_abbrev}</b></td><td>Wed, Dec 31, 2024</td><td>Abbreviated date</td></tr>
        <tr><td><b>{$clock_short}</b></td><td>2:23 PM</td><td>Short clock time</td></tr>
        <tr><td><b>{$clock_long}</b></td><td>2:23:50 PM</td><td>Long clock time</td></tr>
        </table>
        
        <h3>Media & Document Variables</h3>
        <table border="1" cellpadding="5" style="border-collapse: collapse;">
        <tr><th>Variable</th><th>Example Output</th><th>Description</th></tr>
        <tr><td><b>{$media_name}</b></td><td>example.mov</td><td>Media file name</td></tr>
        <tr><td><b>{$media_path}</b></td><td>/path/to/example.mov</td><td>Full media path</td></tr>
        <tr><td><b>{$doc_name}</b></td><td>transcript.html</td><td>Document file name</td></tr>
        <tr><td><b>{$doc_path}</b></td><td>/path/to/transcript.html</td><td>Full document path</td></tr>
        </table>
        
        <h3>Special Variables</h3>
        <table border="1" cellpadding="5" style="border-collapse: collapse;">
        <tr><th>Variable</th><th>Example Output</th><th>Description</th></tr>
        <tr><td><b>{$selection}</b></td><td>selected text</td><td>Currently selected text</td></tr>
        <tr><td><b>{$version}</b></td><td>1.0</td><td>Application version</td></tr>
        </table>
        
        <h3>Usage Examples</h3>
        <p><b>Speaker with timestamp:</b> <code>SPEAKER: {$time} </code></p>
        <p><b>Note with date:</b> <code>[NOTE - {$date_short}] </code></p>
        <p><b>Time offset:</b> <code>Started at ${time_offset(-00:00:05.00)}</code></p>
        <p><b>Wrap selection:</b> <code>&lt;emphasis&gt;{$selection}&lt;/emphasis&gt;</code></p>
        """
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Snippet Variables Reference")
        dialog.resize(800, 700)
        layout = QVBoxLayout(dialog)
        
        browser = QTextBrowser()
        browser.setHtml(ref_text)
        layout.addWidget(browser)
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn)
        
        dialog.exec()

    def check_updates(self):
        """Opens the GitHub releases page to check for the latest version"""
        url = "https://github.com/uckthis/TranscriptFlow/releases/latest"
        QDesktopServices.openUrl(QUrl(url))


    def show_about(self):
        QMessageBox.about(
            self, "About TranscriptFlow Pro",
            "<h2>TranscriptFlow Pro</h2>"
            "<p>Version 1.1.2</p>"
            "<p>A professional transcription application.</p>"
            "<p><b>Features:</b></p>"
            "<ul>"
            "<li>Professional timecode support</li>"
            "<li>Customizable snippets with variables</li>"
            "<li>Waveform visualization</li>"
            "<li>Multiple export formats</li>"
            "<li>VLC-powered media playback</li>"
            "<li>Dark and light themes</li>"
            "</ul>"
            "<p>Visit our GitHub for updates and community support:</p>"
            "<p><a href='https://github.com/uckthis/TranscriptFlow'>https://github.com/uckthis/TranscriptFlow</a></p>"
            "<p><i>Built with PyQt6, libmpv, and libvlc</i></p>"
        )

    def find_text(self):
        """Open Find dialog (Modeless)"""
        if not self.find_dialog:
            self.find_dialog = FindReplaceDialog(self, initial_mode="find")
        else:
            # Switch to find mode if it's already open
            self.find_dialog.replace_input.hide()
            self.find_dialog.btn_replace.hide()
            self.find_dialog.btn_replace_all.hide()
            self.find_dialog.form_layout.labelForField(self.find_dialog.replace_input).hide()
            self.find_dialog.setWindowTitle("Find")
            
        self.find_dialog.show()
        self.find_dialog.raise_()
        self.find_dialog.activateWindow()
        self.find_dialog.find_input.setFocus()
        self.find_dialog.find_input.selectAll()

    def replace_text(self):
        """Open Replace dialog (Modeless)"""
        if not self.find_dialog:
            self.find_dialog = FindReplaceDialog(self, initial_mode="replace")
        else:
            # Switch to replace mode
            self.find_dialog.replace_input.show()
            self.find_dialog.btn_replace.show()
            self.find_dialog.btn_replace_all.show()
            self.find_dialog.form_layout.labelForField(self.find_dialog.replace_input).show()
            self.find_dialog.setWindowTitle("Find & Replace")
            
        self.find_dialog.show()
        self.find_dialog.raise_()
        self.find_dialog.activateWindow()
        self.find_dialog.find_input.setFocus()
        self.find_dialog.find_input.selectAll()

    def find_next_silent(self):
        """Perform 'Find Next' using last search criteria"""
        if self.find_dialog and self.find_dialog.find_input.text():
            self.find_dialog.find_next()
        else:
            self.find_text()

    def find_prev_silent(self):
        """Perform 'Find Previous' using last search criteria"""
        if self.find_dialog and self.find_dialog.find_input.text():
            self.find_dialog.find_previous()
        else:
            self.find_text()

    def choose_color(self):
        # Use existing color as default
        current_fmt = self.editor.currentCharFormat()
        initial_color = current_fmt.foreground().color() if current_fmt.foreground().style() != Qt.BrushStyle.NoBrush else Qt.GlobalColor.black
        
        color = QColorDialog.getColor(initial_color, self)
        if color.isValid():
            self.editor.set_text_color(color.name())
            self.update_color_button_ui(color)

    def update_color_button_ui(self, color):
        """Updates the ribbon color button's appearance"""
        if isinstance(color, str):
            color = QColor(color)
        
        # Determine border color based on theme or luminance for visibility
        border_col = "#666" if not self.dark_mode else "#ccc"
        
        self.btn_color.setStyleSheet(f"""
            QPushButton {{ 
                border-radius: 6px; 
                background-color: {color.name()};
                border: 2px solid {border_col};
            }}
            QPushButton:hover {{ border: 2px solid white; }}
        """)

    def apply_player_scaling(self):
        """Applies one of 3 refined player control scaling/resize behaviors"""
        if not hasattr(self, 'controls_panel'): return
        
        behavior = self.config['ui'].get('player_scaling_behavior', 'proportional')
        
        # 1. Reset Global Panel constraints (Let it stretch for Video)
        self.controls_panel.setMaximumWidth(16777215)
        self.controls_panel.setMinimumWidth(0)
        self.controls_panel.setFixedHeight(16777215)
        self.controls_panel.setMinimumHeight(0)
        
        # Identify inner containers to constrain
        containers = [self.time_container, self.buttons_container, self.sliders_container]
        
        # 2. Reset Inner Container Constraints
        for c in containers:
            if c:
                c.setMaximumWidth(16777215)
                c.setMinimumWidth(0)
                # Clear fixed heights that cause clipping
                c.setFixedHeight(16777215)
                c.setMinimumHeight(0)
        
        if behavior == 'cap':
            # 2. Limit Control Width (Wide Prevention)
            for c in containers:
                if c: c.setMaximumWidth(900)
            
        elif behavior == 'island':
            # 3. Center Controls (Modern Island)
            for c in containers:
                if c:
                    c.setFixedWidth(650)
            
        # Ensure centered alignment for non-proportional modes
        if behavior in ['cap', 'island']:
            self.controls_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        else:
            self.controls_layout.setAlignment(Qt.AlignmentFlag.AlignTop) # Standard
            
        self.update()

    def apply_layout_preset(self, preset):
        """Applies a named layout preset"""
        self.config['ui']['layout'] = preset
        self.save_config()
        
        self.main_splitter.show()
        self.left_panel.show()
        self.editor.show() # Should always be shown
        self.waveform_dock.show()
        
        # Ensure correct order for horizontal layouts
        if preset in ["standard", "wide", "audio_only"]:
            self.main_splitter.insertWidget(0, self.left_panel)
            self.main_splitter.insertWidget(1, self.right_panel)
            self.main_splitter.setOrientation(Qt.Orientation.Horizontal)
        elif "reversed" in preset:
            self.main_splitter.insertWidget(0, self.right_panel)
            self.main_splitter.insertWidget(1, self.left_panel)
            self.main_splitter.setOrientation(Qt.Orientation.Horizontal)
        elif preset in ["stacked", "compact"]:
            # Editor on TOP, Player/Controls at BOTTOM
            self.main_splitter.insertWidget(0, self.right_panel)
            self.main_splitter.insertWidget(1, self.left_panel)
            self.main_splitter.setOrientation(Qt.Orientation.Vertical)

        if preset == "standard":
            self.main_splitter.setSizes([400, 800])
        elif preset == "wide":
            self.main_splitter.setSizes([800, 400])
        elif preset == "reversed":
            self.main_splitter.setSizes([800, 400])
        elif preset == "audio_only":
            self.left_panel.hide()
            self.waveform_dock.show()
        elif "no_wf" in preset:
            self.waveform_dock.hide()
        elif preset == "stacked":
            # Editor at top gets 600, Player at bottom gets 400
            self.main_splitter.setSizes([600, 400])
        elif preset == "focus":
            self.left_panel.hide()
            self.waveform_dock.hide()
        elif preset == "compact":
            self.waveform_dock.hide()
            # Editor at top gets most space, Player at bottom stays minimal
            self.main_splitter.setSizes([850, 150])
            
        
    def update_theme_menu(self):
        """Populates the theme menu dynamically based on visibility and custom themes"""
        if not hasattr(self, 'themes_menu'): return
        self.themes_menu.clear()
        
        self.theme_group = QActionGroup(self)
        self.theme_group.setExclusive(True)
        
        themes = self.get_builtin_themes()
        builder_cfg = self.config.get('theme_builder', {})
        custom = builder_cfg.get('custom_themes', {})
        hidden = builder_cfg.get('hidden_themes', [])
        
        all_themes = list(themes.keys()) + list(custom.keys())
        
        for tname in all_themes:
            if tname in hidden: continue
            
            act = QAction(tname, self)
            act.setCheckable(True)
            act.setChecked(tname == self.current_theme)
            act.setData(tname)
            act.triggered.connect(lambda checked, n=tname: self.set_theme(n))
            self.themes_menu.addAction(act)
            self.theme_group.addAction(act)
            
        self.themes_menu.addSeparator()
        self.add_action_safe(self.themes_menu, "Theme Builder...", self.open_theme_builder)

    def open_theme_builder(self):
        from dialogs import ThemeBuilderDialog
        builtin = self.get_builtin_themes()
        builder_cfg = self.config.setdefault('theme_builder', {'custom_themes': {}, 'hidden_themes': []})
        
        dlg = ThemeBuilderDialog(self, builtin, builder_cfg['custom_themes'], builder_cfg['hidden_themes'])
        if dlg.exec():
            # Update config
            builder_cfg['custom_themes'] = dlg.custom_themes
            builder_cfg['hidden_themes'] = dlg.hidden_themes
            
            # Refresh UI
            self.update_theme_menu()
            self.apply_theme() # Refresh current theme in case it was edited
            self.save_config()
            self.statusBar().showMessage("Themes updated successfully", 3000)


    def resizeEvent(self, event):
        """Ensure ribbon stays aligned on window resize"""
        super().resizeEvent(event)

    def eventFilter(self, obj, event):
        """Track video frame resize events and save size to config"""
        if obj == self.video_frame and event.type() == event.Type.Resize:
            # Save the new size to config
            new_size = [self.video_frame.width(), self.video_frame.height()]
            self.config['ui']['video_player_size'] = new_size
        return super().eventFilter(obj, event)

    def closeEvent(self, event):
        """Warn about unsaved changes on close"""
        if self.is_dirty or self.editor.document().isModified():
            reply = QMessageBox.question(self, "Unsaved Changes",
                "You have unsaved changes in your document. Do you want to save them before exiting?",
                QMessageBox.StandardButton.Save | QMessageBox.StandardButton.Discard | QMessageBox.StandardButton.Cancel)
            
            if reply == QMessageBox.StandardButton.Save:
                self.save_doc()
                self.save_config()
                event.accept()
            elif reply == QMessageBox.StandardButton.Discard:
                self.save_config()
                event.accept()
            else:
                event.ignore()
        else:
            self.save_config()
            event.accept()

    def save_config(self):
        """Updates and persists the unified configuration"""
        if self.is_initializing:
            return
        self.config['shortcuts'] = self.shortcuts
        self.config['snippets'] = self.snippets
        self.config['settings'] = self.settings
        self.config['ui']['show_timeline'] = self.show_timeline
        self.config['ui']['show_remote'] = self.show_remote
        self.config['ui']['show_playrate'] = self.show_playrate
        self.config['ui']['show_volume'] = self.show_volume
        self.config['ui']['show_waveform'] = self.show_waveform
        self.config['ui']['dark_mode'] = self.dark_mode

        # Save window and splitter states
        try:
            self.config['ui']['geometry'] = self.saveGeometry().toHex().data().decode()
            self.config['ui']['window_state'] = self.saveState().toHex().data().decode()
            self.config['ui']['splitter_state'] = self.main_splitter.saveState().toHex().data().decode()
            self.config['ui']['left_splitter_state'] = self.left_splitter.saveState().toHex().data().decode()
            self.config['ui']['ribbon_splitter_state'] = self.ribbon_splitter.saveState().toHex().data().decode()
        except:
            pass
        
        # Video player size is already saved in eventFilter, no need to save again here

        self.sm.save(self.config)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    # --- Splash Screen ---
    splash_path = get_resource_path("splash.png")
    if os.path.exists(splash_path):
        original_pixmap = QPixmap(splash_path)
        # Scale to 50% for a professional size (Photoshop style)
        pixmap = original_pixmap.scaled(
            original_pixmap.width() // 2, 
            original_pixmap.height() // 2, 
            Qt.AspectRatioMode.KeepAspectRatio, 
            Qt.TransformationMode.SmoothTransformation
        )
    else:
        # Create a beautiful programmatic fallback splash
        pixmap = QPixmap(600, 400)
        pixmap.fill(QColor("#1e1e2e")) # Midnight Blue
        painter = QPainter(pixmap)
        
        # Gradient
        gradient = QLinearGradient(0, 0, 600, 400)
        gradient.setColorAt(0, QColor("#1e1e2e"))
        gradient.setColorAt(1, QColor("#313244"))
        painter.fillRect(pixmap.rect(), QBrush(gradient))
        
        # Title
        painter.setPen(QColor("white"))
        font = QFont("Inter", 28, QFont.Weight.Bold)
        painter.setFont(font)
        painter.drawText(pixmap.rect(), Qt.AlignmentFlag.AlignCenter, "TranscriptFlow Pro")
        
        # Subtitle
        font.setPointSize(12)
        font.setWeight(QFont.Weight.Normal)
        painter.setFont(font)
        painter.drawText(pixmap.rect().adjusted(0, 80, 0, 0), Qt.AlignmentFlag.AlignCenter, "Initializing Professional Workspace...")
        painter.end()
        
    splash = QSplashScreen(pixmap, Qt.WindowType.WindowStaysOnTopHint)
    splash.show()
    app.processEvents()
    
    # window.show() is DEFERRED to late_initialization for a "Clean Reveal"
    window = MainWindow(splash)
    sys.exit(app.exec())
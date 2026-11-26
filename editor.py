from PyQt6.QtWidgets import QTextEdit, QMenu, QApplication, QLabel
from PyQt6.QtCore import (pyqtSignal, Qt, QRegularExpression, QPropertyAnimation,
                          QEasingCurve, QRectF, QPointF, QTimer)
from PyQt6.QtGui import (QKeyEvent, QKeySequence, QTextCursor, QTextDocument,
                         QFont, QFontMetrics, QColor, QTextCharFormat,
                         QSyntaxHighlighter, QAction, QPainter, QPen, QPolygonF)
from utils import TimecodeHelper
import smart_typing
import re
import os
import time

# Configure enchant to use AppData dicts folder BEFORE importing enchant
from path_manager import get_dicts_dir
_dicts_path = get_dicts_dir()
os.environ["DICPATH"] = _dicts_path

import enchant

class SpellHighlighter(QSyntaxHighlighter):
    def __init__(self, editor):
        super().__init__(editor.document())
        self.editor = editor
        self.skip_list = set() # Words to ignore in current session
        self.lang = "en_US"
        self._init_dict()
            
        self.format = QTextCharFormat()
        self.format.setUnderlineColor(QColor("red"))
        self.format.setUnderlineStyle(QTextCharFormat.UnderlineStyle.SpellCheckUnderline)
        self.enabled = True

    def _init_dict(self):
        try:
            self.dict = enchant.Dict(self.lang)
            self.valid = True
        except:
            self.valid = False

    def highlightBlock(self, text):
        if not self.enabled or not text or not self.valid:
            return

        # Track cursor to avoid highlighting word being typed
        cursor = self.editor.textCursor()
        cursor_block_num = cursor.block().blockNumber()
        cursor_pos_in_block = cursor.positionInBlock()
        is_current_block = (self.currentBlock().blockNumber() == cursor_block_num)

        # Regex to find words while ignoring timecodes and brackets
        for match in re.finditer(r"\b[A-Za-z']+\b", text):
            word = match.group(0)
            start = match.start()
            end = match.end()
            
            # Skip if word in skip list
            if word.lower() in self.skip_list:
                continue

            # 1. Skip if cursor is currently within or at the end of this word (typing in progress)
            if is_current_block:
                if start <= cursor_pos_in_block <= end:
                    continue

            # 2. Heuristic: skip if inside brackets
            if start > 0 and text[start-1] == '[': continue
            
            # 3. Spelling check
            if not self.dict.check(word):
                self.setFormat(start, len(word), self.format)

    def set_enabled(self, enabled):
        self.enabled = enabled
        self.rehighlight()

    def set_language(self, lang):
        self.lang = lang
        self._init_dict()
        self.rehighlight()

class TranscriptEditor(QTextEdit):
    # Signals to Main Window
    seekRequested = pyqtSignal(int)
    commandTriggered = pyqtSignal(dict) # Sends the shortcut dict to Main
    snippetTriggered = pyqtSignal(dict) # Sends the snippet dict to Main
    settingsChanged = pyqtSignal(dict) # Notifies Main of style changes (font, size, etc.)

    def __init__(self, shortcuts_list, snippets_list, settings):
        super().__init__()
        self.setAcceptRichText(True)
        self.shortcuts = shortcuts_list
        self.snippets = snippets_list
        self.settings = settings
        self.setMouseTracking(True)
        self.temporary_font_override = None # For Urdu auto-switching
        self._smart_retro_marker = None  # Last retroactive cap, for '...' undo
        self.update_font()
        
        # Spell Checker Integration
        self.highlighter = SpellHighlighter(self)
        self.highlighter.set_language(settings.get('spell_check_lang', 'en_US'))
        self.highlighter.set_enabled(settings.get('spell_check', True))
        
        # Extra selections are rebuilt so spell/timecode highlights coexist
        # with the current-line indicator.
        self._transient_selections = []
        self._cursor_line_selection = None
        self.cursorPositionChanged.connect(self._refresh_extra_selections)
        
        # Redundant Modification Safeguard:
        # 1. Standard text change tracking
        self.document().contentsChanged.connect(lambda: self.document().setModified(True))

        # Auto-scroll: one follow mode with configurable lines ahead
        self.auto_scroll_enabled = True
        self.auto_scroll_lines = 6
        self.auto_scroll_pause_manual = False
        self.auto_scroll_type_only = False
        self.auto_scroll_pause_selecting = True
        self._mouse_selection_down = False
        self._modifier_selection_held = False
        self._selection_pause_until = 0.0
        self._selection_resume_delay = 0.35
        # Watchdog timestamps: a mouse release or key release can be swallowed
        # by a focus change. Without a self-expiry path, one missed release
        # leaves the guard latched on forever and auto scroll never resumes.
        self._mouse_down_at = 0.0
        self._mouse_last_activity_at = 0.0
        self._mouse_press_pos = None
        self._mouse_dragging = False
        self._modifier_held_at = 0.0
        # Diagnostics for the auto-scroll follow decision, surfaced by the
        # F12 debug readout in main.py.
        self._last_follow_probe = None
        self._last_follow_skip = ''
        # Scroll offset at the moment a follow started. If a relayout collapses
        # the scrollbar range during an edit, Qt can clamp the value down which
        # reads as the view jumping to the top of the page; this lets the editor
        # put it back. Disarmed by any deliberate user scroll or click.
        self._follow_floor = -1
        self._follow_guard_active = False
        self._mouse_watchdog_seconds = 1.5
        self._modifier_watchdog_seconds = 2.0
        self.bottom_padding_percent = 25
        self._last_manual_scroll_at = 0.0
        self._last_manual_nav_at = 0.0
        self._manual_pause_seconds = 3.0
        # cursorPositionChanged fires from inside super().keyPressEvent(),
        # before the new block's layout has settled. Following there starts an
        # animation from stale geometry which the authoritative post-keypress
        # call then retargets - visible as a jump to the top and back on Enter.
        self._in_key_press = False
        self.cursorPositionChanged.connect(self._on_cursor_position_changed)
        self.verticalScrollBar().valueChanged.connect(
            self._on_scroll_value_changed)

        # Wheel scroll animation
        self._scroll_target = None
        self._scroll_anim = QPropertyAnimation(self.verticalScrollBar(), b"value")
        self._scroll_anim.setDuration(100)
        self._scroll_anim.setEasingCurve(QEasingCurve.Type.OutCubic)

        # Current-line indicator
        self.current_line_enabled = True
        self.current_line_style = 'background'
        self.current_line_scope = 'line'
        self.current_line_opacity = 12
        self.current_line_color = ''
        self.current_line_rail_color = ''
        self.current_line_theme_color = '#378ADD'
        self.current_line_focus_only = True
        self.current_line_hide_selection = True
        self.current_line_rail_shape = 'solid'
        self.current_line_rail_width = 3

    def set_auto_scroll(self, enabled=True, lines=6,
                        pause_manual=False, type_only=False,
                        bottom_padding=25, pause_selecting=True):
        self.auto_scroll_enabled = bool(enabled)
        self.auto_scroll_lines = max(3, int(lines))
        self.auto_scroll_pause_manual = bool(pause_manual)
        self.auto_scroll_type_only = bool(type_only)
        self.auto_scroll_pause_selecting = bool(pause_selecting)
        self._mouse_selection_down = False
        self._modifier_selection_held = False
        self._selection_pause_until = 0.0
        self._mouse_down_at = 0.0
        self._mouse_last_activity_at = 0.0
        self._mouse_press_pos = None
        self._mouse_dragging = False
        self._modifier_held_at = 0.0
        self.bottom_padding_percent = max(0, min(40, int(bottom_padding)))
        self.update_bottom_margin()
        if self.auto_scroll_enabled:
            self._handle_auto_scroll()

    def set_current_line_indicator(self, enabled=True, style='background',
                                   opacity=12, color='', focus_only=True,
                                   hide_selection=True, rail_shape='solid',
                                   rail_width=3, rail_color='',
                                   scope='line'):
        self.current_line_enabled = bool(enabled)
        self.current_line_style = style if style in (
            'background', 'rail', 'both') else 'background'
        self.current_line_scope = scope if scope in ('line', 'paragraph') else 'line'
        self.current_line_opacity = max(5, min(25, int(opacity)))
        self.current_line_color = color or ''
        self.current_line_rail_color = rail_color or ''
        self.current_line_focus_only = bool(focus_only)
        self.current_line_hide_selection = bool(hide_selection)
        self.current_line_rail_shape = rail_shape if rail_shape in (
            'solid', 'rounded', 'tick', 'arrow') else 'solid'
        self.current_line_rail_width = max(2, min(8, int(rail_width)))
        self.update_bottom_margin()
        self._refresh_extra_selections()

    def set_current_line_theme_color(self, color):
        self.current_line_theme_color = color or '#378ADD'
        self._refresh_extra_selections()

    def _current_line_color(self):
        color = QColor(self.current_line_color) if self.current_line_color \
            else QColor()
        if not color.isValid():
            color = QColor(self.current_line_theme_color)
        if not color.isValid():
            color = QColor('#378ADD')
        return color

    def _rail_color(self):
        color = QColor(self.current_line_rail_color) \
            if self.current_line_rail_color else QColor()
        if not color.isValid():
            color = self._current_line_color()
        return color

    def _should_show_current_line(self):
        if not self.current_line_enabled:
            return False
        if self.current_line_style not in ('background', 'rail', 'both'):
            return False
        if self.current_line_focus_only and not self.hasFocus():
            return False
        if (self.current_line_hide_selection and
                self.textCursor().hasSelection() and
                not self._transient_selections):
            return False
        return True

    def _refresh_extra_selections(self):
        selection = None
        if (self._should_show_current_line() and
                self.current_line_style in ('background', 'both')):
            cursor = QTextCursor(self.textCursor())
            cursor.clearSelection()
            if self.current_line_scope == 'paragraph':
                cursor.movePosition(QTextCursor.MoveOperation.StartOfBlock)
                cursor.movePosition(QTextCursor.MoveOperation.EndOfBlock,
                                    QTextCursor.MoveMode.KeepAnchor)
            else:
                cursor.movePosition(QTextCursor.MoveOperation.StartOfLine)
                cursor.movePosition(QTextCursor.MoveOperation.EndOfLine,
                                    QTextCursor.MoveMode.KeepAnchor)
            selection = QTextEdit.ExtraSelection()
            selection.cursor = cursor
            fmt = QTextCharFormat()
            color = self._current_line_color()
            color.setAlpha(self.current_line_opacity * 255 // 100)
            fmt.setBackground(color)
            selection.format = fmt
        self._cursor_line_selection = selection
        selections = ([selection] if selection else []) + \
            list(self._transient_selections)
        self.setExtraSelections(selections)
        self.viewport().update()

    def paintEvent(self, event):
        super().paintEvent(event)
        if (not self._should_show_current_line() or
                self.current_line_style not in ('rail', 'both')):
            return
        rect = self.cursorRect()
        if rect.isNull():
            return
        color = self._rail_color()
        color.setAlpha(255)
        width = self.current_line_rail_width
        height = max(rect.height(), self.fontMetrics().lineSpacing())
        y = rect.top()
        x = 3
        painter = QPainter(self.viewport())
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        shape = self.current_line_rail_shape
        if shape == 'solid':
            painter.fillRect(QRectF(x, y, width, height), color)
        elif shape == 'rounded':
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(color)
            painter.drawRoundedRect(QRectF(x, y + 1, width, height - 2),
                                    width / 2.0, width / 2.0)
        elif shape == 'tick':
            painter.setPen(QPen(color, 2))
            mid = int(y + height / 2)
            painter.drawLine(int(x), mid, int(x + width + 6), mid)
        elif shape == 'arrow':
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(color)
            mid = y + height / 2.0
            arrow_height = min(height - 4, max(10, height * 0.55))
            tip = x + width + 6
            painter.drawPolygon(QPolygonF([
                QPointF(x, mid - arrow_height / 2),
                QPointF(tip, mid),
                QPointF(x, mid + arrow_height / 2),
            ]))
        painter.end()

    def focusInEvent(self, event):
        super().focusInEvent(event)
        self._refresh_extra_selections()

    def focusOutEvent(self, event):
        super().focusOutEvent(event)
        self._mouse_selection_down = False
        self._modifier_selection_held = False
        self._mouse_down_at = 0.0
        self._mouse_last_activity_at = 0.0
        self._modifier_held_at = 0.0
        self._refresh_extra_selections()

    def _effective_line_height(self):
        """Line height at the cursor, accounting for the current font/RTL size."""
        return max(1, QFontMetrics(self.currentFont()).lineSpacing())

    def max_lines_ahead(self):
        """Largest useful lines-ahead value: 85% of the visible lines."""
        viewport_h = self.viewport().height()
        if viewport_h <= 0:
            return 12
        visible = viewport_h / self._effective_line_height()
        return max(3, int(visible * 0.85))

    def _rail_gutter_width(self):
        if self.current_line_enabled and self.current_line_style in (
                'rail', 'both'):
            return max(18, self.current_line_rail_width + 10)
        return 4

    def update_bottom_margin(self):
        root = self.document().rootFrame()
        if not root:
            return
        fmt = root.frameFormat()
        viewport_h = self.viewport().height()
        bottom = int(viewport_h * self.bottom_padding_percent / 100.0) \
            if viewport_h > 0 else 0
        was_modified = self.document().isModified()
        fmt.setBottomMargin(float(bottom))
        fmt.setLeftMargin(float(self._rail_gutter_width()))
        root.setFrameFormat(fmt)
        self.document().setModified(was_modified)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update_bottom_margin()

    def _selection_guard_active(self):
        """True only while a selection is actively being made.

        Deliberately trusts editor-local event state only. Two earlier versions
        of this guard queried QApplication.mouseButtons()/keyboardModifiers()
        and both made auto-scroll die outright, because a global query can
        report a held button or modifier at times the widget never saw a
        matching release. A selection that merely *exists* must not block
        either: commands such as 'Go To Next Time Code' select the timecode on
        purpose, and gating on hasSelection() froze auto-scroll for the whole
        document.
        """
        if not self.auto_scroll_pause_selecting:
            return False
        now = time.monotonic()

        # Self-healing watchdogs. An event-driven flag alone is not enough in
        # the real app: a release can be swallowed by a focus change (clicking
        # a timecode seeks the media and can steal focus), leaving the flag on
        # forever. A genuine drag keeps renewing its activity timestamp.
        if self._mouse_selection_down:
            # After a short grace period, a globally released button proves
            # the click/drag is over even if this widget missed the release
            # event. During a real drag the button stays down, so the guard
            # remains active exactly when it should.
            if (now - self._mouse_down_at > 0.08 and
                    QApplication.mouseButtons() == Qt.MouseButton.NoButton):
                self._mouse_selection_down = False
            elif (now - self._mouse_last_activity_at >
                  self._mouse_watchdog_seconds):
                self._mouse_selection_down = False
        if self._modifier_selection_held:
            if now - self._modifier_held_at > self._modifier_watchdog_seconds:
                self._modifier_selection_held = False

        if self._modifier_selection_held:
            return True
        if self._mouse_selection_down:
            return True
        return now < self._selection_pause_until

    def follow_diagnostics(self):
        """One-line human-readable report on the last follow decision."""
        probe = self._last_follow_probe or {}
        if self._last_follow_skip:
            reason = self._last_follow_skip
        elif probe.get('needs'):
            reason = 'scrolled'
        elif probe:
            reason = 'cursor_in_zone'
        else:
            reason = 'not_run'
        if not probe:
            return f"follow: {reason}"
        return (f"follow: {reason} | caret_bottom={probe['cursor_bottom']} "
                f"threshold={probe['threshold']} viewport={probe['viewport_h']} "
                f"line_h={probe['line_height']} ahead={probe['lines_ahead']} "
                f"scroll={probe['scroll']} | flags "
                f"mouse={'1' if self._mouse_selection_down else '0'} "
                f"mod={'1' if self._modifier_selection_held else '0'} "
                f"drag={'1' if self._mouse_dragging else '0'} "
                f"pause={max(0.0, self._selection_pause_until - time.monotonic()):.2f}s")

    def toggle_follow_diagnostics(self, enabled):
        """Show the live follow readout as an overlay (temporary debug aid)."""
        if not hasattr(self, '_diag_label'):
            self._diag_label = QLabel(self.viewport())
            self._diag_label.setStyleSheet(
                "background: rgba(0,0,0,190); color: #7CFF7C;"
                " padding: 4px 6px; font-family: Consolas, monospace;"
                " font-size: 11px; border-radius: 3px;")
            self._diag_label.setWordWrap(True)
            self._diag_label.setMouseTracking(True)
            self._diag_label.setAttribute(
                Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        if enabled:
            self._diag_label.setGeometry(8, 8, self.viewport().width() - 24, 54)
            self._diag_label.show()
            self._diag_label.raise_()
            self._diag_timer = getattr(self, '_diag_timer', None)
            if self._diag_timer is None:
                from PyQt6.QtCore import QTimer
                self._diag_timer = QTimer(self)
                self._diag_timer.timeout.connect(self._refresh_diag_label)
            self._diag_timer.start(200)
            self._refresh_diag_label()
        else:
            if getattr(self, '_diag_timer', None) is not None:
                self._diag_timer.stop()
            self._diag_label.hide()

    def _refresh_diag_label(self):
        if not hasattr(self, '_diag_label'):
            return
        self._diag_label.setGeometry(
            8, 8, self.viewport().width() - 24, 54)
        self._diag_label.setText(self.follow_diagnostics())

    def _pause_auto_scroll_for_selection(self):
        """Hold the follow while a selection gesture is in progress.

        The resume delay is armed only for real gestures (mouse drag or a held
        Shift/Ctrl). A plain click just records the gesture so the follow does
        not fight the pointer, without delaying the next follow.
        """
        if self._mouse_selection_down or self._modifier_selection_held:
            self._selection_pause_until = (
                time.monotonic() + self._selection_resume_delay)

    def _manual_pause_active(self):
        if not self.auto_scroll_pause_manual:
            return False
        now = time.monotonic()
        return (now - self._last_manual_scroll_at < self._manual_pause_seconds or
                now - self._last_manual_nav_at < self._manual_pause_seconds)

    def _on_cursor_position_changed(self):
        """Follow cursor movement from external sources, but never mid-keypress.

        During keyPressEvent the document layout is not yet final when Qt emits
        this signal. keyPressEvent calls _handle_auto_scroll itself once the
        edit is complete, so following here as well produced a two-animation
        fight and the whole-page jump on Enter.
        """
        if self._in_key_press:
            return
        self._handle_auto_scroll()

    def _on_scroll_value_changed(self, value):
        """Put the view back if a relayout clamped it toward the top.

        Adding a paragraph can transiently shrink the document's layout, which
        makes Qt clamp the scrollbar value down. On a long file the clamp can be
        enormous, and it looked exactly like the page jumping to the top when
        Enter was pressed. If the value dropped far below where the follow
        asked to be, while the range is still settling, restore it.
        """
        if not self._follow_guard_active or self._follow_floor < 0:
            return
        # Only correct a large backward clamp, not normal small adjustments.
        if value >= self._follow_floor - max(4, self._effective_line_height() * 2):
            return
        self.verticalScrollBar().setValue(self._follow_floor)

    def _arm_follow_guard(self):
        """Track a follow so a range collapse during relayout can be undone.

        The floor is deliberately the offset shown *before* the follow, not the
        requested target: it must never drag the view forward, only refuse to
        be clamped backwards. Recording the target instead would also undo a
        legitimate scroll made inside the guard window.
        """
        self._follow_floor = self.verticalScrollBar().value()
        self._follow_guard_active = True

        def disarm():
            self._follow_guard_active = False

        QTimer.singleShot(600, disarm)

    def _follow_scroll_to(self, target):
        """Move the view so the caret clears the bottom threshold.

        Applied immediately, with no easing. An earlier build animated this,
        but a follow advances one line per keystroke (~22 px): far too little
        for any easing curve to be perceptible, so it only added lag. Worse, a
        QPropertyAnimation that is still running while the document reflows
        can be yanked backwards, which is what made Enter flash the view.
        """
        vbar = self.verticalScrollBar()
        target = max(vbar.minimum(), min(vbar.maximum(), int(target)))
        if target == vbar.value():
            return
        # A wheel animation still coasting owns the bar; its internal value can
        # drag things backwards when cancelled mid flight.
        if self._scroll_anim.state() == QPropertyAnimation.State.Running:
            self._scroll_anim.stop()
            self._scroll_target = None
        vbar.setValue(target)

    def _handle_auto_scroll(self, from_typing=False, content_changed=False):
        """Follow the caret, keeping `lines ahead` of text below it.

        content_changed=True marks a genuine insertion (typing, paste,
        snippet). That cancels a stranded Ctrl/Shift latch - pressing Ctrl+V is
        not a selection gesture, but it does latch Ctrl - while a live mouse
        drag still wins, so dragging never gets yanked by the follow.
        """
        self._last_follow_skip = ''
        if not self.auto_scroll_enabled:
            self._last_follow_skip = 'disabled'
            return
        if self._manual_pause_active():
            self._last_follow_skip = 'manual_pause'
            return
        if self.auto_scroll_type_only and not from_typing:
            self._last_follow_skip = 'type_only'
            return
        if content_changed and not self._mouse_selection_down:
            # The text moved; a modifier left over from the chord is not a
            # selection in progress. Clear it so the follow can resume.
            self._modifier_selection_held = False
            self._selection_pause_until = 0.0
        if self._selection_guard_active():
            if self._modifier_selection_held:
                self._last_follow_skip = 'guard:modifier'
            elif self._mouse_selection_down:
                self._last_follow_skip = 'guard:mouse'
            else:
                self._last_follow_skip = 'guard:pause_window'
            return

        viewport_h = self.viewport().height()
        if viewport_h <= 0:
            self._last_follow_skip = 'no_viewport'
            return

        line_height = self._effective_line_height()
        lines_ahead = min(self.auto_scroll_lines, self.max_lines_ahead())
        bottom_margin = min(lines_ahead * line_height,
                            int(viewport_h * 0.85))
        cursor_rect = self.cursorRect()

        self._last_follow_probe = {
            'cursor_bottom': cursor_rect.bottom(),
            'threshold': viewport_h - bottom_margin,
            'viewport_h': viewport_h,
            'needs': cursor_rect.bottom() > viewport_h - bottom_margin,
            'scroll': self.verticalScrollBar().value(),
            'line_height': line_height,
            'lines_ahead': lines_ahead,
        }

        if cursor_rect.bottom() > viewport_h - bottom_margin:
            delta = cursor_rect.bottom() - (viewport_h - bottom_margin)
            # Move a full line clear of the boundary so the caret does not sit
            # exactly on the threshold (where each keypress nudges the bar by
            # 2-6 px and looks stuck).
            delta = max(delta, line_height)
            current = self.verticalScrollBar().value()
            target = current + delta
            self._last_follow_probe['target'] = target
            # Arm the anti-clamp guard around the edit that is about to reflow
            # the document, so a range collapse cannot dump us at the top.
            self._arm_follow_guard()
            self._follow_scroll_to(target)

    def wheelEvent(self, event):
        # Allow standard zoom when Ctrl is held
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            super().wheelEvent(event)
            return

        delta_y = event.angleDelta().y()
        if delta_y != 0:
            self._last_manual_scroll_at = time.monotonic()
            # Deliberate scrolling wins over the anti-clamp restore.
            self._follow_guard_active = False
            # 120 units per standard wheel notch -> exactly 1 line per notch!
            steps = delta_y / 120.0
            line_height = self.fontMetrics().lineSpacing()
            delta_pixels = -int(round(steps * line_height))

            vbar = self.verticalScrollBar()
            current_val = vbar.value()
            base_val = self._scroll_target if (self._scroll_anim.state() == QPropertyAnimation.State.Running and self._scroll_target is not None) else current_val
            target_val = max(vbar.minimum(), min(vbar.maximum(), base_val + delta_pixels))

            self._scroll_target = target_val
            self._scroll_anim.stop()
            self._scroll_anim.setStartValue(current_val)
            self._scroll_anim.setEndValue(target_val)
            self._scroll_anim.start()
            event.accept()
        else:
            super().wheelEvent(event)

    def keyReleaseEvent(self, event):
        if event.key() in (Qt.Key.Key_Shift, Qt.Key.Key_Control):
            # Arm the delay while the flag is still set, then release it.
            self._pause_auto_scroll_for_selection()
            self._modifier_selection_held = False
            self._modifier_held_at = 0.0
        super().keyReleaseEvent(event)

    def paste_plain_text(self):
        clipboard = QApplication.clipboard()
        text = clipboard.text()
        if text:
            cursor = self.textCursor()
            cursor.beginEditBlock()
            cursor.insertText(text)
            cursor.endEditBlock()
            self.ensureCursorVisible()
            if self.auto_scroll_enabled:
                self._handle_auto_scroll(from_typing=True,
                                         content_changed=True)

    def update_font(self):
        font_name = self.settings.get('font', 'Tahoma')
        font_size = int(self.settings.get('size', 14))
        
        font = QFont(font_name, font_size)
        self.setFont(font)
        self.document().setDefaultFont(font)
        
        # CRITICAL: Never apply default format to an active selection!
        # In Qt, setCurrentCharFormat directly overrides the format of any selected text,
        # which would strip colors, bold, and italics from the selection.
        cursor = self.textCursor()
        if not cursor.hasSelection():
            self.setCurrentCharFormat(self.get_default_format())

    def keyPressEvent(self, event: QKeyEvent):
        # Suppress the cursorPositionChanged follow for the whole keypress:
        # Qt emits it before the new block is laid out, so following there
        # scrolls against stale geometry and then gets retargeted.
        self._in_key_press = True
        try:
            self._key_press_event(event)
        finally:
            self._in_key_press = False

    def _key_press_event(self, event: QKeyEvent):
        key = event.key()
        modifiers = event.modifiers()

        if key in (Qt.Key.Key_Up, Qt.Key.Key_Down, Qt.Key.Key_Left,
                   Qt.Key.Key_Right, Qt.Key.Key_PageUp, Qt.Key.Key_PageDown,
                   Qt.Key.Key_Home, Qt.Key.Key_End):
            self._last_manual_nav_at = time.monotonic()
        _sel_mods = (Qt.KeyboardModifier.ShiftModifier |
                     Qt.KeyboardModifier.ControlModifier)
        if modifiers & _sel_mods:
            # Only latch while a modifier is actually held. Cleared as soon as
            # any later key arrives without it, so a dropped keyRelease (focus
            # change, OS quirk) can never strand the flag on and kill auto-scroll.
            self._modifier_selection_held = True
            self._modifier_held_at = time.monotonic()
            self._pause_auto_scroll_for_selection()
        elif self._modifier_selection_held:
            self._modifier_selection_held = False
            self._modifier_held_at = 0.0
        
        # 1. Handle TAB specifically (Standard TranscriptFlow behavior)
        if key == Qt.Key.Key_Tab and not modifiers:
            # Find the "Toggle Pause and Play" shortcut
            for sc in self.shortcuts:
                if sc['command'] == 'Toggle Pause and Play':
                    self.commandTriggered.emit(sc)
                    return
            # Fallback if shortcut deleted
            self.commandTriggered.emit({'command': 'Toggle Pause and Play', 'skip': 0})
            return

        # Handle Paste as Plain Text (Ctrl+Shift+V)
        if key == Qt.Key.Key_V and (modifiers & Qt.KeyboardModifier.ControlModifier) and (modifiers & Qt.KeyboardModifier.ShiftModifier):
            self.paste_plain_text()
            return

        # 2. Reset format on ENTER (Reset on Enter feature)
        if key in [Qt.Key.Key_Return, Qt.Key.Key_Enter]:
            # Remember alignment before enter
            is_right_aligned = self.alignment() == Qt.AlignmentFlag.AlignRight or self.layoutDirection() == Qt.LayoutDirection.RightToLeft
            group_edit = self._smart_typing_should_group(event, is_newline=True)
            edit_cursor = self.textCursor() if group_edit else None
            if edit_cursor:
                edit_cursor.beginEditBlock()
            super().keyPressEvent(event)
            self.setCurrentCharFormat(self.get_default_format())
            # Maintain alignment for the new block
            if is_right_aligned:
                self.setAlignment(Qt.AlignmentFlag.AlignRight)
            self._smart_typing_after_key(event, is_newline=True)
            if edit_cursor:
                edit_cursor.endEditBlock()
            self._handle_auto_scroll(from_typing=True, content_changed=True)
            return

        # 3. Build Shortcut String (e.g. "Ctrl+Shift+P")
        parts = []
        if modifiers & Qt.KeyboardModifier.ControlModifier: parts.append("Ctrl")
        if modifiers & Qt.KeyboardModifier.AltModifier: parts.append("Alt")
        if modifiers & Qt.KeyboardModifier.ShiftModifier: parts.append("Shift")
        
        key_text = QKeySequence(key).toString()
        # Filter out modifier-only presses
        if key_text and key not in [Qt.Key.Key_Control, Qt.Key.Key_Alt, Qt.Key.Key_Shift]:
            parts.append(key_text)
        
        shortcut_str = "+".join(parts)
        
        # 4. Check Snippets (Priority 1)
        if shortcut_str:
            for snip in self.snippets:
                if self.matches_trigger(shortcut_str, snip.get('trigger', '')):
                    self.snippetTriggered.emit(snip)
                    return

        # 5. Check Shortcuts (Priority 2)
            for sc in self.shortcuts:
                if self.matches_trigger(shortcut_str, sc.get('trigger', '')):
                    self.commandTriggered.emit(sc)
                    return

        # 6. Default Typing
        group_edit = self._smart_typing_should_group(event)
        if not group_edit:
            self._smart_retro_marker = None
        edit_cursor = self.textCursor() if group_edit else None
        if edit_cursor:
            edit_cursor.beginEditBlock()
        super().keyPressEvent(event)

        # 6b. Smart Typing & Formatting (after the character is inserted)
        self._smart_typing_after_key(event)
        if edit_cursor:
            edit_cursor.endEditBlock()
        self._handle_auto_scroll(from_typing=bool(event.text()),
                                 content_changed=bool(event.text()))

        # 7. Force block alignment for RTL to prevent BiDi cursor jumping
        # Qt's BiDi algorithm tries to be "smart" and moves RTL text to the left
        # We need to explicitly force the alignment after each keystroke
        if self.layoutDirection() == Qt.LayoutDirection.RightToLeft:
            cursor = self.textCursor()
            block_format = cursor.blockFormat()
            if block_format.alignment() != Qt.AlignmentFlag.AlignRight:
                block_format.setAlignment(Qt.AlignmentFlag.AlignRight)
                cursor.setBlockFormat(block_format)

    def _smart_typing_should_group(self, event, is_newline=False):
        """True when a key may need its text and smart edit in one undo step."""
        rules = smart_typing.rules_from_settings(self.settings)
        if not rules['enabled']:
            return False
        if event.isAutoRepeat():
            return False
        blockers = (Qt.KeyboardModifier.ControlModifier |
                    Qt.KeyboardModifier.AltModifier |
                    Qt.KeyboardModifier.MetaModifier)
        if event.modifiers() & blockers:
            return False
        if is_newline:
            return True
        text = event.text()
        custom = smart_typing.parse_custom_triggers(rules['custom_triggers'])
        custom_char = any(text and token.endswith(text) for token in custom)
        return (text in ('.', '?', '!') or custom_char or
                (len(text) == 1 and 'a' <= text <= 'z'))

    def _smart_typing_after_key(self, event, is_newline=False):
        """Apply Smart Typing rules after the key's text is in the document."""
        if not self._smart_typing_should_group(event, is_newline):
            return
        rules = smart_typing.rules_from_settings(self.settings)
        text = event.text()
        cursor = self.textCursor()
        text_before = cursor.block().text()[:cursor.positionInBlock()]
        custom_match = smart_typing.ending_custom_trigger(text_before, rules)

        if is_newline:
            self._smart_retro_marker = None
            self._smart_capitalize_ahead(rules, True)
            return

        if text == '.':
            run = self._trailing_dot_run()
            if run == 3:
                if rules.get('capitalize_after_ellipsis'):
                    self._smart_capitalize_ahead(rules, False)
                else:
                    self._revert_retro_marker()
                    self._smart_retro_marker = None
                return
            if run >= 4:
                self._revert_retro_marker()
                self._smart_capitalize_ahead(rules, False)
                return
            self._smart_capitalize_ahead(rules, False)
            return

        if text in ('?', '!'):
            self._smart_retro_marker = None
            self._smart_capitalize_ahead(rules, False)
            return

        if custom_match:
            self._smart_retro_marker = None
            self._smart_capitalize_ahead(rules, False)
            return

        if len(text) == 1 and 'a' <= text <= 'z':
            self._smart_retro_marker = None
            self._smart_capitalize_typed(text, rules)

    def _trailing_dot_run(self):
        """Count consecutive dots immediately before the cursor."""
        cursor = self.textCursor()
        text = cursor.block().text()[:cursor.positionInBlock()]
        count = 0
        for char in reversed(text):
            if char != '.':
                break
            count += 1
        return count

    def _revert_retro_marker(self):
        """Undo the capitalization made when the first dot was typed."""
        marker = self._smart_retro_marker
        if not marker:
            return
        cursor = self.textCursor()
        position = cursor.position() + marker['offset']
        block = cursor.block()
        offset = position - block.position()
        text = block.text()
        if offset < 0 or offset >= len(text):
            return
        char = text[offset]
        if char.isupper() and char.lower() == marker['original'].lower():
            self._replace_char_at(position, marker['original'],
                                  restore_position=cursor.position())

    def _replace_char_at(self, position, char, restore_position=None):
        """Replace one character in one undo step, keeping its char format."""
        if position < 0 or position + 1 >= self.document().characterCount():
            return False
        cursor = QTextCursor(self.document())
        cursor.setPosition(position)
        char_format = cursor.charFormat()
        cursor.beginEditBlock()
        cursor.setPosition(position + 1, QTextCursor.MoveMode.KeepAnchor)
        cursor.insertText(char)
        cursor.setPosition(position)
        cursor.setPosition(position + 1, QTextCursor.MoveMode.KeepAnchor)
        cursor.setCharFormat(char_format)
        cursor.clearSelection()
        cursor.endEditBlock()
        if restore_position is not None:
            cursor.setPosition(restore_position)
            self.setTextCursor(cursor)
        return True

    def _smart_capitalize_typed(self, letter, rules):
        """Uppercase the letter just typed when the preceding text arms it."""
        cursor = self.textCursor()
        if cursor.hasSelection():
            return
        block = cursor.block()
        offset = cursor.position() - block.position()
        text = block.text()
        if offset < 1 or offset > len(text) or text[offset - 1] != letter:
            return
        if not smart_typing.should_capitalize(text[:offset - 1], rules):
            return
        self._replace_char_at(cursor.position() - 1, letter.upper())

    def _smart_capitalize_ahead(self, rules, is_newline):
        """Uppercase an existing lowercase word right after the trigger."""
        cursor = self.textCursor()
        if cursor.hasSelection():
            return
        block = cursor.block()
        offset = cursor.position() - block.position()
        text_before = '\n' if is_newline else block.text()[:offset]
        if not smart_typing.should_capitalize(text_before, rules):
            return
        text_after = block.text()[offset:]
        index = smart_typing.find_next_lowercase(text_after)
        if index < 0:
            return
        if self._replace_char_at(cursor.position() + index,
                                 text_after[index].upper(),
                                 restore_position=cursor.position()):
            self._smart_retro_marker = {'offset': index,
                                        'original': text_after[index]}

    def matches_trigger(self, pressed, trigger):
        """Normalize and compare keys"""
        if not pressed or not trigger: return False
        return pressed.upper().replace(" ", "") == trigger.upper().replace(" ", "")

    def mousePressEvent(self, event):
        """Click to seek logic"""
        self._last_manual_nav_at = time.monotonic()
        # A click is deliberate; never undo the user's chosen position.
        self._follow_guard_active = False
        self._mouse_selection_down = True
        self._mouse_down_at = time.monotonic()
        self._mouse_last_activity_at = self._mouse_down_at
        self._mouse_press_pos = event.pos()
        self._mouse_dragging = False
        # Deliberately DO NOT arm the pause here. A click that just places the
        # caret is not a selection gesture; arming on press was what made the
        # caret stop following after every mouse placement. The pause is armed
        # on the first real drag movement instead.
        cursor = self.cursorForPosition(event.pos())
        cursor.select(cursor.SelectionType.WordUnderCursor)
        text = cursor.selectedText()
        
        # Regex for standard timecodes [00:00:00.00]
        # Clean the text of brackets for regex checking
        clean_text = text.replace("[", "").replace("]", "").replace("(", "").replace(")", "")
        if re.match(r'\d{1,2}:\d{2}:\d{2}', clean_text):
            tc_helper = TimecodeHelper(fps=self.settings.get('fps', 30.0))
            ms = tc_helper.timestamp_to_ms(text)
            self.seekRequested.emit(ms)
        
        super().mousePressEvent(event)

    def insert_processed_content(self, text, color_name="black", is_html=False, carry_format=False, 
                                 bold=True, italic=False, underline=False):
        """Inserts text/html at cursor"""
        cursor = self.textCursor()
        
        # Auto-newline logic (Configurable)
        if self.settings.get('timecode_new_line', True) and cursor.position() > 0:
            doc = self.toPlainText()
            char_before = doc[cursor.position() - 1]
            if char_before != '\n':
                cursor.insertText('\n')

        if is_html:
            cursor.insertHtml(text)
        else:
            # Convert newlines for HTML and preserve spaces
            text_html = text.replace('\n', '<br>')
            # User restoration: Header should follow snippet style
            # CONTRAST GUARD: Ensure text is visible against editor background
            bg_color = self.palette().color(self.backgroundRole())
            text_color = QColor(color_name)
            
            # Simple luminance-based distance
            def get_lum(c): return (0.299 * c.red() + 0.587 * c.green() + 0.114 * c.blue()) / 255.0
            
            bg_lum = get_lum(bg_color)
            txt_lum = get_lum(text_color)
            
            # Base snippet style
            style = f'color:{color_name}; white-space: pre-wrap;'
            if bold: style += ' font-weight:bold;'
            if italic: style += ' font-style:italic;'
            if underline: style += ' text-decoration:underline;'
            
            # If colors are too close (delta < 0.25), add a subtle "halo" background
            halo_style = ""
            if abs(bg_lum - txt_lum) < 0.25:
                # If background is dark, use a light halo, and vice versa
                halo_bg = "rgba(255, 255, 255, 0.15)" if bg_lum < 0.5 else "rgba(0, 0, 0, 0.08)"
                halo_style = f" background-color: {halo_bg}; border-radius: 3px; padding: 0 2px;"

            html = f'<span style="{style}{halo_style}">{text_html}</span>'
            cursor.insertHtml(html)
            
        # Move main text cursor to where the insertion cursor ended
        self.setTextCursor(cursor)
        
        # CRITICAL: Set the format for subsequent typing
        if not carry_format:
            # Reset to default
            self.setCurrentCharFormat(self.get_default_format())
        else:
            # Carry the snippet's style
            fmt = self.get_default_format()
            fmt.setForeground(QColor(color_name))
            if bold: fmt.setFontWeight(QFont.Weight.Bold)
            fmt.setFontItalic(italic)
            fmt.setFontUnderline(underline)
            self.setCurrentCharFormat(fmt)
            self.setCurrentCharFormat(fmt)
            
        self.ensureCursorVisible()
        self.document().setModified(True)
        # Snippets and timecode stamps insert text without a key event, so the
        # follow has to be driven here or the caret can walk off screen.
        self._handle_auto_scroll(from_typing=True, content_changed=True)

    def mouseReleaseEvent(self, event):
        """Detect click on timecode and seek"""
        # Only a real drag earns a resume delay. A plain caret-placing click
        # leaves auto scroll free to keep following immediately.
        if self._mouse_dragging:
            self._mouse_selection_down = True
            self._pause_auto_scroll_for_selection()
        self._mouse_selection_down = False
        self._mouse_dragging = False
        self._mouse_press_pos = None
        self._mouse_last_activity_at = time.monotonic()
        if event.button() == Qt.MouseButton.LeftButton:
            pos = event.pos()
            cursor = self.cursorForPosition(pos)
            
            # PHYSICAL HIT TEST: Ensure we aren't clicking in whitespace
            # If the cursor rect doesn't contain the x-coordinate, we are in the margin/whitespace
            crect = self.cursorRect(cursor)
            # Add a small buffer for easier clicking
            if pos.x() < crect.left() - 5 or pos.x() > crect.right() + 5:
                super().mouseReleaseEvent(event)
                return

            block = cursor.block()
            block_text = block.text()
            block_pos = cursor.positionInBlock()
            
            recognize_unbracketed = self.settings.get('recognize_unbracketed', True)
            regex = TimecodeHelper.get_regex(bracketed_only=not recognize_unbracketed)

            for match in re.finditer(regex, block_text):
                if match.start() <= block_pos <= match.end():
                    # Double Check: Ensure the actual character we clicked is part of the match
                    tc_string = match.group(0)
                    try:
                        helper = TimecodeHelper(self.settings.get('fps', 30.0), self.settings.get('media_offset', 0))
                        ms = helper.timestamp_to_ms(tc_string)
                        self.seekRequested.emit(ms)
                        
                        # 1. VISUAL FEEDBACK (Professional way)
                        # We use ExtraSelections so we don't actually change the cursor selection.
                        # This avoids triggering ribbon rich-text updates.
                        sel = QTextEdit.ExtraSelection()
                        sel.cursor = QTextCursor(block)
                        sel.cursor.setPosition(block.position() + match.start())
                        sel.cursor.setPosition(block.position() + match.end(), QTextCursor.MoveMode.KeepAnchor)
                        
                        # Determine highlight color (Subtle theme-aware or light blue)
                        sel.format.setBackground(QColor(173, 216, 230, 100)) # Lightblue semi-transparent
                        self._transient_selections = [sel]
                        self._refresh_extra_selections()
                        
                        # 2. REAL CURSOR PLACEMENT
                        # Move real cursor to the end of timecode (without selection)
                        new_cursor = self.textCursor()
                        new_cursor.setPosition(block.position() + match.end())
                        self.setTextCursor(new_cursor)
                        
                        # 3. FORMAT GUARD
                        # Force "Normal" format for any typing that follows
                        self.setCurrentCharFormat(self.get_default_format())
                        return
                    except: pass
                    
        super().mouseReleaseEvent(event)

    def mouseMoveEvent(self, event):
        """Change cursor when hovering over a timecode"""
        if event.buttons() and self._mouse_selection_down:
            # First real movement with the button held turns a click into a
            # drag; that is the moment selection (and its pause) begins.
            if (not self._mouse_dragging and self._mouse_press_pos is not None
                    and (event.pos() - self._mouse_press_pos).manhattanLength()
                    > 4):
                self._mouse_dragging = True
                self._pause_auto_scroll_for_selection()
            self._mouse_last_activity_at = time.monotonic()
        elif not event.buttons() and self._mouse_selection_down:
            # Button released outside the widget.
            if self._mouse_dragging:
                self._pause_auto_scroll_for_selection()
            self._mouse_selection_down = False
            self._mouse_dragging = False
            self._mouse_press_pos = None
            self._mouse_last_activity_at = time.monotonic()
        pos = event.pos()
        cursor = self.cursorForPosition(pos)
        
        # PHYSICAL HIT TEST
        crect = self.cursorRect(cursor)
        if pos.x() < crect.left() - 5 or pos.x() > crect.right() + 5:
            self.viewport().setCursor(Qt.CursorShape.IBeamCursor)
            super().mouseMoveEvent(event)
            return

        block = cursor.block()
        block_text = block.text()
        block_pos = cursor.positionInBlock()
        
        recognize_unbracketed = self.settings.get('recognize_unbracketed', True)
        regex = TimecodeHelper.get_regex(bracketed_only=not recognize_unbracketed)
            
        found = False
        for match in re.finditer(regex, block_text):
            if match.start() <= block_pos <= match.end():
                found = True
                break
        
        if found:
            self.viewport().setCursor(Qt.CursorShape.PointingHandCursor)
        else:
            self.viewport().setCursor(Qt.CursorShape.IBeamCursor)
            
        super().mouseMoveEvent(event)

    def go_to_timecode(self, forward=True):
        """Find next/previous timecode in text and seek to it"""
        cursor = self.textCursor()
        pos = cursor.position()
        text = self.toPlainText()
        
        recognize_unbracketed = self.settings.get('recognize_unbracketed', True)
        regex = TimecodeHelper.get_regex(bracketed_only=not recognize_unbracketed)

        timecodes = []
        for match in re.finditer(regex, text):
            timecodes.append((match.start(), match.group(0)))
            
        if not timecodes: return
        
        target_idx = -1
        if forward:
            # Find first timecode AFTER current position
            for i, (start, tc) in enumerate(timecodes):
                if start > pos:
                    target_idx = i
                    break
            # Wrap to beginning if none found after
            if target_idx == -1: target_idx = 0
        else:
            # Find first timecode BEFORE current position
            for i in range(len(timecodes) - 1, -1, -1):
                start, tc = timecodes[i]
                if start < pos:
                    target_idx = i
                    break
            # Wrap to end if none found before
            if target_idx == -1: target_idx = len(timecodes) - 1
                
        if target_idx != -1:
            start_pos, target_tc = timecodes[target_idx]
            # Select the timecode to show we found it
            cursor.setPosition(start_pos)
            cursor.setPosition(start_pos + len(target_tc), QTextCursor.MoveMode.KeepAnchor)
            self.setTextCursor(cursor)
            self.ensureCursorVisible()
            
            # Trigger seek
            tc_helper = TimecodeHelper(fps=self.settings.get('fps', 30.0))
            ms = tc_helper.timestamp_to_ms(target_tc)
            self.seekRequested.emit(ms)

    def toggle_bold(self):
        fmt = QTextCharFormat()
        is_bold = self.currentCharFormat().fontWeight() == QFont.Weight.Bold
        fmt.setFontWeight(QFont.Weight.Normal if is_bold else QFont.Weight.Bold)
        self.mergeCurrentCharFormat(fmt)

    def toggle_italic(self):
        fmt = QTextCharFormat()
        fmt.setFontItalic(not self.currentCharFormat().fontItalic())
        self.mergeCurrentCharFormat(fmt)

    def toggle_underline(self):
        fmt = QTextCharFormat()
        fmt.setFontUnderline(not self.currentCharFormat().fontUnderline())
        self.mergeCurrentCharFormat(fmt)

    def cycle_case(self):
        """Cycles selected text through lowercase -> UPPERCASE -> Title Case (Word-style)"""
        cursor = self.textCursor()
        if not cursor.hasSelection():
            # If no selection, select the word under cursor
            cursor.select(QTextCursor.SelectionType.WordUnderCursor)
            
        text = cursor.selectedText()
        if not text:
            return

        # Handle unique logic for empty or whitespace-only
        if not text.strip():
            return

        # Cycle: lower -> UPPER -> Title -> lower
        if text.isupper():
            # From UPPERCASE to Title Case
            new_text = text.title()
        elif text.istitle() or (len(text) > 0 and text[0].isupper() and text[1:].islower()):
            # From Title Case to lowercase
            new_text = text.lower()
        else:
            # From lowercase (or mixed) to UPPERCASE
            new_text = text.upper()
            
        # Perform replacement
        cursor.beginEditBlock()
        cursor.insertText(new_text)
        
        # Re-select the replaced text to allow continuous cycling
        cursor.setPosition(cursor.position() - len(new_text), QTextCursor.MoveMode.KeepAnchor)
        self.setTextCursor(cursor)
        cursor.endEditBlock()

    def set_text_color(self, color):
        self.settings['font_color'] = color
        fmt = QTextCharFormat()
        fmt.setForeground(QColor(color))
        self.mergeCurrentCharFormat(fmt)
        self.settingsChanged.emit(self.settings)

    def set_font_family(self, font_name):
        self.settings['font'] = font_name
        fmt = QTextCharFormat()
        fmt.setFontFamilies([font_name])
        self.mergeCurrentCharFormat(fmt)
        self.settingsChanged.emit(self.settings)

    def set_font_size(self, size):
        self.settings['size'] = int(size)
        fmt = QTextCharFormat()
        fmt.setFontPointSize(float(size))
        self.mergeCurrentCharFormat(fmt)
        self.settingsChanged.emit(self.settings)

    def set_temporary_font_override(self, font_name):
        """Sets a font override without updating permanent settings (for auto-Urdu)"""
        self.temporary_font_override = font_name
        # Apply immediately to current typing
        self.setCurrentCharFormat(self.get_default_format())

    def get_default_format(self):
        """Returns a character format based on current settings or temporary override"""
        from PyQt6.QtGui import QTextCharFormat
        fmt = QTextCharFormat()
        # Use override if present (for Urdu), otherwise config setting
        font_name = self.temporary_font_override if self.temporary_font_override else self.settings.get('font', 'Tahoma')
        font_size = self.settings.get('size', 14)
        fmt.setFont(QFont(font_name, font_size))
        fmt.setForeground(QColor(self.settings.get('font_color', 'black')))
        fmt.setFontWeight(QFont.Weight.Normal)
        fmt.setFontItalic(False)
        fmt.setFontUnderline(False)
        return fmt

    def adjust_timecodes(self, offset_ms, selection_only=False):
        """Adjusts all timecodes in the text or selection by offset_ms."""
        helper = TimecodeHelper(self.settings.get('fps', 30.0))
        recognize_unbracketed = self.settings.get('recognize_unbracketed', True)
        regex = TimecodeHelper.get_regex(bracketed_only=not recognize_unbracketed)
        
        cursor = self.textCursor()
        cursor.beginEditBlock()
        
        if selection_only:
            selection_start = cursor.selectionStart()
            selection_end = cursor.selectionEnd()
            cursor.setPosition(selection_start)
            
            while True:
                found_cursor = self.document().find(QRegularExpression(regex), cursor)
                if found_cursor.isNull() or found_cursor.selectionStart() >= selection_end:
                    break
                
                tc_text = found_cursor.selectedText()
                try:
                    current_ms = helper.timestamp_to_ms(tc_text)
                    new_ms = max(0, current_ms + offset_ms)
                    
                    has_brackets = tc_text.startswith('[')
                    new_tc = helper.ms_to_timestamp(new_ms, 
                                                  bracket="[]" if has_brackets else "",
                                                  use_frames_sep=":" if ":" in tc_text else ".")
                    
                    found_cursor.insertText(new_tc)
                except:
                    pass
                cursor = found_cursor
        else:
            cursor.movePosition(QTextCursor.MoveOperation.Start)
            while True:
                found_cursor = self.document().find(QRegularExpression(regex), cursor)
                if found_cursor.isNull():
                    break
                
                tc_text = found_cursor.selectedText()
                try:
                    current_ms = helper.timestamp_to_ms(tc_text)
                    new_ms = max(0, current_ms + offset_ms)
                    
                    has_brackets = tc_text.startswith('[')
                    new_tc = helper.ms_to_timestamp(new_ms, 
                                                  bracket="[]" if has_brackets else "",
                                                  use_frames_sep=":" if ":" in tc_text else ".")
                    
                    found_cursor.insertText(new_tc)
                except:
                    pass
                cursor = found_cursor

        cursor.endEditBlock()

    def contextMenuEvent(self, event):
        if self._mouse_dragging:
            self._pause_auto_scroll_for_selection()
        self._mouse_selection_down = False
        self._mouse_dragging = False
        self._mouse_press_pos = None
        self._mouse_last_activity_at = time.monotonic()
        menu = self.createStandardContextMenu()
        
        # Insert "Paste as Plain Text" right after standard "Paste"
        paste_act = None
        for act in menu.actions():
            act_text = act.text().replace("&", "").split('\t')[0].strip()
            if act_text == "Paste" or act.shortcut() == QKeySequence.StandardKey.Paste:
                paste_act = act
                break
                
        plain_paste_act = QAction("Paste as Plain Text\tCtrl+Shift+V", self)
        plain_paste_act.setShortcut(QKeySequence("Ctrl+Shift+V"))
        plain_paste_act.triggered.connect(self.paste_plain_text)
        clipboard = QApplication.clipboard()
        plain_paste_act.setEnabled(bool(clipboard.text()))

        if paste_act:
            actions = menu.actions()
            idx = actions.index(paste_act)
            if idx + 1 < len(actions):
                menu.insertAction(actions[idx + 1], plain_paste_act)
            else:
                menu.addAction(plain_paste_act)
        else:
            menu.addAction(plain_paste_act)

        # If spell check is enabled, add suggestions
        if self.highlighter.enabled and self.highlighter.valid:
            cursor = self.cursorForPosition(event.pos())
            cursor.select(QTextCursor.SelectionType.WordUnderCursor)
            word = cursor.selectedText().strip()
            
            if word and not self.highlighter.dict.check(word):
                # We insert in reverse order so the last one inserted ends up at the top
                menu.insertSeparator(menu.actions()[0])
                
                # 3. Add to Dictionary
                add_act = QAction(f"Add '{word}' to Dictionary", self)
                add_act.triggered.connect(lambda checked, w=word: self.add_to_dictionary(w))
                menu.insertAction(menu.actions()[0], add_act)

                # 2. Ignore All (current session)
                ignore_act = QAction(f"Ignore All ('{word}')", self)
                ignore_act.triggered.connect(lambda checked, w=word: self.add_to_skip_list(w))
                menu.insertAction(menu.actions()[0], ignore_act)
                
                menu.insertSeparator(menu.actions()[0])

                # 1. Suggestions
                suggestions = self.highlighter.dict.suggest(word)
                if suggestions:
                    suggestion_list = suggestions[:5] # Top 5
                    for s in reversed(suggestion_list):
                        act = QAction(s, self)
                        act.triggered.connect(lambda checked, replacement=s, c=cursor: self.replace_word(c, replacement))
                        menu.insertAction(menu.actions()[0], act)

                        all_act = QAction(f"Change All to '{s}'", self)
                        all_act.setToolTip(
                            "Replace every whole-word instance, keeping each one's capitalization")
                        all_act.triggered.connect(
                            lambda checked, w=word, replacement=s:
                            self.replace_all_words(w, replacement))
                        menu.insertAction(menu.actions()[0], all_act)
                
                header_act = QAction(f"Spelling: '{word}'", self)
                header_act.setEnabled(False)
                menu.insertAction(menu.actions()[0], header_act)
        
        menu.exec(event.globalPos())

    def replace_word(self, cursor, replacement):
        cursor.beginEditBlock()
        cursor.insertText(replacement)
        cursor.endEditBlock()

    def set_spell_check_enabled(self, enabled):
        self.highlighter.set_enabled(enabled)
        self.settings['spell_check'] = enabled
        # Note: main.py should save config

    def find_next_error(self, start_pos=0):
        """Finds the next misspelled word starting from start_pos"""
        if not self.highlighter.valid: return None
        
        text = self.toPlainText()
        # Regex to find words
        for match in re.finditer(r"\b[A-Za-z']+\b", text):
            # Check if this match contains or starts at/after start_pos
            if match.start() < start_pos: continue
            
            word = match.group(0)
            # Skip if in skip list
            if word.lower() in self.highlighter.skip_list: continue
            # Skip if bracketed
            if match.start() > 0 and text[match.start()-1] == '[': continue
            
            # Check dictionary
            if not self.highlighter.dict.check(word):
                return {
                    'word': word,
                    'start': match.start(),
                    'end': match.end()
                }
        return None

    def highlight_word(self, start, end):
        """High-visibility focus on a specific word in the editor"""
        cursor = self.textCursor()
        cursor.setPosition(start)
        cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
        self.setTextCursor(cursor)
        
        # Professional "Glow" Highlight via ExtraSelections
        sel = QTextEdit.ExtraSelection()
        sel.cursor = cursor
        # Use a vibrant but transparent yellow/gold highlight
        sel.format.setBackground(QColor(255, 215, 0, 100)) 
        sel.format.setProperty(QTextCharFormat.Property.OutlinePen, QColor("orange"))
        self._transient_selections = [sel]
        self._refresh_extra_selections()
        
        self.ensureCursorVisible()

    def replace_selection(self, start, end, new_text):
        """Replaces text at specified range and moves cursor"""
        cursor = self.textCursor()
        cursor.setPosition(start)
        cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
        cursor.insertText(new_text)
        self.setTextCursor(cursor)

    @staticmethod
    def _match_case(original, replacement):
        """Apply the capitalization pattern of original to replacement."""
        if not original or not replacement:
            return replacement
        if original.isupper():
            return replacement.upper()
        if original[0].isupper() and (len(original) == 1 or original[1:].islower()):
            return replacement[0].upper() + replacement[1:]
        return replacement

    def replace_all_words(self, old_word, new_word):
        """Replace every whole-word occurrence, preserving case and formatting.

        Bracketed annotations such as [teh] are skipped, matching the spell
        highlighter's own exclusions. All replacements form one undo step.
        """
        if not old_word or not new_word or old_word == new_word:
            return 0
        doc = self.document()
        plain = self.toPlainText()
        group = QTextCursor(doc)
        group.beginEditBlock()
        search = QTextCursor(doc)
        count = 0
        while True:
            found = doc.find(old_word, search,
                             QTextDocument.FindFlag.FindWholeWords)
            if found.isNull():
                break
            start = found.selectionStart()
            search = found
            if start > 0 and plain[start - 1] == '[':
                continue
            matched = found.selectedText()
            replacement = self._match_case(matched, new_word)
            char_format = found.charFormat()
            found.insertText(replacement)
            found.setPosition(start)
            found.setPosition(start + len(replacement),
                              QTextCursor.MoveMode.KeepAnchor)
            found.setCharFormat(char_format)
            found.clearSelection()
            search = found
            count += 1
        group.endEditBlock()
        return count

    def add_to_skip_list(self, word):
        """Add word to session skip list and re-highlight"""
        self.highlighter.skip_list.add(word.lower())
        self.highlighter.rehighlight()

    def remove_from_skip_list(self, word):
        """Remove word from session skip list"""
        if word.lower() in self.highlighter.skip_list:
            self.highlighter.skip_list.remove(word.lower())
            self.highlighter.rehighlight()

    def add_to_dictionary(self, word):
        """Permanently add word to custom dictionary"""
        try:
            self.highlighter.dict.add(word)
            self.highlighter.rehighlight()
        except Exception as e:
            print(f"Error adding to dictionary: {e}")

    def remove_from_dictionary(self, word):
        """Remove word from user dictionary (if possible)"""
        try:
            self.highlighter.dict.remove(word)
            self.highlighter.rehighlight()
        except: pass

    def set_spell_language(self, lang_code):
        """Switch dictionary language"""
        self.highlighter.set_language(lang_code)


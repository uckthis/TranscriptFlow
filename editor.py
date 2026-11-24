from PyQt6.QtWidgets import QTextEdit, QMenu, QApplication
from PyQt6.QtCore import pyqtSignal, Qt, QRegularExpression, QPropertyAnimation, QEasingCurve
from PyQt6.QtGui import QKeyEvent, QKeySequence, QTextCursor, QTextDocument, QFont, QColor, QTextCharFormat, QSyntaxHighlighter, QAction
from utils import TimecodeHelper
import smart_typing
import re
import os

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
        
        # Clear highlights when cursor moves
        self.cursorPositionChanged.connect(lambda: self.setExtraSelections([]))
        
        # Redundant Modification Safeguard:
        # 1. Standard text change tracking
        self.document().contentsChanged.connect(lambda: self.document().setModified(True))

        # Auto-scroll and overscroll padding
        self.auto_scroll_enabled = True
        self.auto_scroll_mode = 'smart_margin'  # 'smart_margin', 'typewriter', 'advance_bump'
        self.scroll_past_end = True
        self.cursorPositionChanged.connect(self._handle_auto_scroll)

        # Smooth 1-line mouse wheel scrolling
        self._scroll_target = None
        self._scroll_anim = QPropertyAnimation(self.verticalScrollBar(), b"value")
        self._scroll_anim.setDuration(100)
        self._scroll_anim.setEasingCurve(QEasingCurve.Type.OutCubic)

    def set_auto_scroll(self, enabled=True, mode='smart_margin', scroll_past_end=True):
        self.auto_scroll_enabled = enabled
        self.auto_scroll_mode = mode
        self.scroll_past_end = scroll_past_end
        self.update_bottom_margin()
        if self.auto_scroll_enabled:
            self._handle_auto_scroll()

    def update_bottom_margin(self):
        root = self.document().rootFrame()
        if not root:
            return
        fmt = root.frameFormat()
        if self.scroll_past_end:
            viewport_h = self.viewport().height()
            margin = max(150, int(viewport_h * 0.5)) if viewport_h > 0 else 200
            fmt.setBottomMargin(float(margin))
        else:
            fmt.setBottomMargin(0.0)
        root.setFrameFormat(fmt)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update_bottom_margin()

    def _handle_auto_scroll(self):
        if not self.auto_scroll_enabled:
            return

        viewport_h = self.viewport().height()
        if viewport_h <= 0:
            return

        cursor_rect = self.cursorRect()
        line_height = self.fontMetrics().lineSpacing()
        vbar = self.verticalScrollBar()

        if self.auto_scroll_mode == 'smart_margin':
            # Guarantee 5-8 lines of upcoming text below cursor
            bottom_margin = max(5 * line_height, int(viewport_h * 0.25))
            top_margin = max(3 * line_height, int(viewport_h * 0.15))

            if cursor_rect.bottom() > viewport_h - bottom_margin:
                delta = cursor_rect.bottom() - (viewport_h - bottom_margin)
                vbar.setValue(vbar.value() + delta)
            elif cursor_rect.top() < top_margin:
                delta = top_margin - cursor_rect.top()
                vbar.setValue(vbar.value() - delta)

        elif self.auto_scroll_mode == 'typewriter':
            target_y = viewport_h // 2
            diff = cursor_rect.center().y() - target_y
            if abs(diff) > line_height // 2:
                vbar.setValue(vbar.value() + diff)

        elif self.auto_scroll_mode == 'advance_bump':
            bottom_threshold = int(viewport_h * 0.75)
            if cursor_rect.bottom() > bottom_threshold:
                bump = max(4 * line_height, int(viewport_h * 0.25))
                vbar.setValue(vbar.value() + bump)

    def wheelEvent(self, event):
        # Allow standard zoom when Ctrl is held
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            super().wheelEvent(event)
            return

        delta_y = event.angleDelta().y()
        if delta_y != 0:
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
                self._handle_auto_scroll()

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
        key = event.key()
        modifiers = event.modifiers()
        
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

    def mouseReleaseEvent(self, event):
        """Detect click on timecode and seek"""
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
                        self.setExtraSelections([sel])
                        
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
        self.setExtraSelections([sel])
        
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


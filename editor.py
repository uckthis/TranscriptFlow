from PyQt6.QtWidgets import QTextEdit, QMenu
from PyQt6.QtCore import pyqtSignal, Qt, QRegularExpression
from PyQt6.QtGui import QKeyEvent, QKeySequence, QTextCursor, QFont, QColor, QTextCharFormat, QSyntaxHighlighter, QAction
from utils import TimecodeHelper
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

    def __init__(self, shortcuts_list, snippets_list, settings):
        super().__init__()
        self.setAcceptRichText(True)
        self.shortcuts = shortcuts_list
        self.snippets = snippets_list
        self.settings = settings
        self.setMouseTracking(True)
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

    def update_font(self):
        font_name = self.settings.get('font', 'Tahoma')
        font_size = int(self.settings.get('size', 14))
        color = self.settings.get('font_color', 'black')
        
        self.setFont(QFont(font_name, font_size))
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

        # 2. Reset format on ENTER (Reset on Enter feature)
        if key in [Qt.Key.Key_Return, Qt.Key.Key_Enter]:
            super().keyPressEvent(event)
            self.setCurrentCharFormat(self.get_default_format())
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
        super().keyPressEvent(event)

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

    def set_text_color(self, color):
        fmt = self.currentCharFormat()
        fmt.setForeground(QColor(color))
        self.setCurrentCharFormat(fmt)

    def set_font_family(self, font_name):
        fmt = self.currentCharFormat()
        fmt.setFontFamilies([font_name])
        self.setCurrentCharFormat(fmt)

    def set_font_size(self, size):
        fmt = self.currentCharFormat()
        fmt.setFontPointSize(float(size))
        self.setCurrentCharFormat(fmt)

    def get_default_format(self):
        """Returns a character format based on current settings"""
        from PyQt6.QtGui import QTextCharFormat
        fmt = QTextCharFormat()
        font_name = self.settings.get('font', 'Tahoma')
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
        
        # If spell check is enabled, add suggestions
        if self.highlighter.enabled and self.highlighter.valid:
            cursor = self.cursorForPosition(event.pos())
            cursor.select(QTextCursor.SelectionType.WordUnderCursor)
            word = cursor.selectedText().strip()
            
            if word and not self.highlighter.dict.check(word):
                # Add suggestions to top of menu
                suggestions = self.highlighter.dict.suggest(word)
                if suggestions:
                    suggestion_list = suggestions[:5] # Top 5
                    
                    menu.insertSeparator(menu.actions()[0])
                    for s in reversed(suggestion_list):
                        act = QAction(s, self)
                        act.triggered.connect(lambda checked, replacement=s, c=cursor: self.replace_word(c, replacement))
                        menu.insertAction(menu.actions()[0], act)
                    
                    header_act = QAction(f"Enchant: '{word}'", self)
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

    def add_to_skip_list(self, word):
        """Add word to session skip list and re-highlight"""
        self.highlighter.skip_list.add(word.lower())
        self.highlighter.rehighlight()

    def remove_from_skip_list(self, word):
        """Remove word from session skip list"""
        if word.lower() in self.highlighter.skip_list:
            self.highlighter.skip_list.remove(word.lower())
            self.highlighter.rehighlight()

    def remove_from_dictionary(self, word):
        """Remove word from user dictionary (if possible)"""
        try:
            self.highlighter.dict.remove(word)
            self.highlighter.rehighlight()
        except: pass

    def set_spell_language(self, lang_code):
        """Switch dictionary language"""
        self.highlighter.set_language(lang_code)


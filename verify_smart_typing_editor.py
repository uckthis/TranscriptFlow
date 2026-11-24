"""Offscreen keyboard-level checks for the editor integration."""

import os
import unittest

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

from PyQt6.QtCore import Qt
from PyQt6.QtTest import QTest
from PyQt6.QtWidgets import QApplication

from editor import TranscriptEditor


app = QApplication.instance() or QApplication([])


def make_editor(text, position, smart_settings=None):
    smart = {'enabled': True}
    if smart_settings:
        smart.update(smart_settings)
    settings = {'font': 'Tahoma', 'size': 14, 'fps': 23.976,
                'spell_check': False, 'smart_typing': smart}
    editor = TranscriptEditor([], [], settings)
    editor.setPlainText(text)
    cursor = editor.textCursor()
    cursor.setPosition(position)
    editor.setTextCursor(cursor)
    return editor


def cursor_at(editor, position):
    cursor = editor.textCursor()
    cursor.setPosition(position)
    editor.setTextCursor(cursor)


class EditorSmartTypingTests(unittest.TestCase):
    def test_predictive_capitalization(self):
        editor = make_editor('hello. ', 7)
        QTest.keyClick(editor, Qt.Key.Key_W)
        self.assertEqual(editor.toPlainText(), 'hello. W')

    def test_retroactive_capitalization(self):
        editor = make_editor('hello world', 5)
        QTest.keyClick(editor, Qt.Key.Key_Period)
        self.assertEqual(editor.toPlainText(), 'hello. World')
        self.assertEqual(editor.textCursor().position(), 6)

    def test_predictive_change_is_one_undo_step(self):
        editor = make_editor('hello. ', 7)
        QTest.keyClick(editor, Qt.Key.Key_W)
        QTest.keyClick(editor, Qt.Key.Key_Z, Qt.KeyboardModifier.ControlModifier)
        self.assertEqual(editor.toPlainText(), 'hello. ')

    def test_retroactive_change_is_one_undo_step(self):
        editor = make_editor('hello world', 5)
        QTest.keyClick(editor, Qt.Key.Key_Period)
        QTest.keyClick(editor, Qt.Key.Key_Z, Qt.KeyboardModifier.ControlModifier)
        self.assertEqual(editor.toPlainText(), 'hello world')

    def test_exactly_three_dots_blocks_capitalization(self):
        editor = make_editor('hello... ', 9)
        QTest.keyClick(editor, Qt.Key.Key_W)
        self.assertEqual(editor.toPlainText(), 'hello... w')

    def test_typing_three_dots_uncaps_already_capitalized_word(self):
        editor = make_editor('hello world', 5)
        for _ in range(3):
            QTest.keyClick(editor, Qt.Key.Key_Period)
        self.assertEqual(editor.toPlainText(), 'hello... world')

    def test_typing_four_dots_caps_again(self):
        editor = make_editor('hello world', 5)
        for _ in range(4):
            QTest.keyClick(editor, Qt.Key.Key_Period)
        self.assertEqual(editor.toPlainText(), 'hello.... World')

    def test_ellipsis_cap_mode_keeps_capitalization(self):
        editor = make_editor('hello world', 5,
                             {'capitalize_after_ellipsis': True})
        for _ in range(3):
            QTest.keyClick(editor, Qt.Key.Key_Period)
        self.assertEqual(editor.toPlainText(), 'hello... World')

    def test_custom_single_char_trigger(self):
        editor = make_editor('hello world', 5, {'custom_triggers': '--, :'})
        QTest.keyClick(editor, Qt.Key.Key_Colon)
        self.assertEqual(editor.toPlainText(), 'hello: World')

    def test_custom_two_char_sequence_waits_for_second_char(self):
        editor = make_editor('hello world', 5, {'custom_triggers': '--, :'})
        QTest.keyClick(editor, Qt.Key.Key_Minus)
        self.assertEqual(editor.toPlainText(), 'hello- world')
        QTest.keyClick(editor, Qt.Key.Key_Minus)
        self.assertEqual(editor.toPlainText(), 'hello-- World')

    def test_custom_two_char_sequence_arms_next_letter(self):
        editor = make_editor('hello-- ', 8, {'custom_triggers': '--, :'})
        QTest.keyClick(editor, Qt.Key.Key_W)
        self.assertEqual(editor.toPlainText(), 'hello-- W')

    def test_two_dots_still_capitalizes(self):
        editor = make_editor('hello.. ', 8)
        QTest.keyClick(editor, Qt.Key.Key_W)
        self.assertEqual(editor.toPlainText(), 'hello.. W')

    def test_enter_capitalizes_existing_next_line(self):
        editor = make_editor('first linesecond', 10)
        QTest.keyClick(editor, Qt.Key.Key_Return)
        self.assertEqual(editor.toPlainText(), 'first line\nSecond')

    def test_disabled_master_leaves_text_alone(self):
        editor = make_editor('hello. ', 7, {'enabled': False})
        QTest.keyClick(editor, Qt.Key.Key_W)
        self.assertEqual(editor.toPlainText(), 'hello. w')

    def test_question_and_exclamation(self):
        for text, position, key in (('Really? ', 8, Qt.Key.Key_Y),
                                    ('Wow! ', 5, Qt.Key.Key_O)):
            with self.subTest(text=text):
                editor = make_editor(text, position)
                QTest.keyClick(editor, key)
                self.assertEqual(editor.toPlainText()[position], 'Y' if key == Qt.Key.Key_Y else 'O')

    def test_timecode_is_not_touched(self):
        editor = make_editor('[00:01:23.29] ', 14)
        QTest.keyClick(editor, Qt.Key.Key_W)
        self.assertEqual(editor.toPlainText(), '[00:01:23.29] w')


if __name__ == '__main__':
    unittest.main()

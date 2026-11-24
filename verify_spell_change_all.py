"""Checks for case-aware whole-document spell replacement."""

import os
import unittest

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

from unittest.mock import patch

from PyQt6.QtCore import QPoint
from PyQt6.QtGui import (QColor, QContextMenuEvent, QTextCharFormat,
                         QTextCursor)
from PyQt6.QtWidgets import QApplication, QMenu

from editor import TranscriptEditor


app = QApplication.instance() or QApplication([])


def make_editor(text):
    settings = {'font': 'Tahoma', 'size': 14, 'fps': 23.976,
                'spell_check': False, 'smart_typing': {'enabled': False}}
    editor = TranscriptEditor([], [], settings)
    editor.setPlainText(text)
    return editor


class FakeDict:
    def check(self, word):
        return word.lower() != 'teh'

    def suggest(self, word):
        return ['the', 'tech']


class ReplaceAllTests(unittest.TestCase):
    def test_case_preserving_whole_word_replacement(self):
        editor = make_editor('teh Teh TEH teh. [teh] theteh\nNext teh')
        count = editor.replace_all_words('teh', 'the')
        self.assertEqual(count, 5)
        self.assertEqual(
            editor.toPlainText(),
            'the The THE the. [teh] theteh\nNext the')

    def test_replacement_is_one_undo_step(self):
        original = 'teh Teh TEH'
        editor = make_editor(original)
        editor.replace_all_words('teh', 'the')
        editor.undo()
        self.assertEqual(editor.toPlainText(), original)

    def test_character_format_is_preserved(self):
        editor = make_editor('teh teh')
        cursor = editor.textCursor()
        cursor.setPosition(0)
        cursor.setPosition(3, QTextCursor.MoveMode.KeepAnchor)
        fmt = QTextCharFormat()
        fmt.setForeground(QColor('#ff0000'))
        cursor.setCharFormat(fmt)

        editor.replace_all_words('teh', 'the')

        check = editor.textCursor()
        check.setPosition(0)
        check.setPosition(3, QTextCursor.MoveMode.KeepAnchor)
        self.assertEqual(check.charFormat().foreground().color().name(), '#ff0000')

    def test_context_menu_offers_change_all(self):
        editor = make_editor('teh')
        editor.highlighter.dict = FakeDict()
        editor.highlighter.valid = True
        editor.highlighter.enabled = True
        editor.show()
        app.processEvents()

        captured = {}

        def fake_exec(self, *args, **kwargs):
            captured['actions'] = [action.text() for action in self.actions()]
            return None

        event = QContextMenuEvent(QContextMenuEvent.Reason.Mouse,
                                  QPoint(2, 2), QPoint(2, 2))
        with patch.object(QMenu, 'exec', fake_exec):
            editor.contextMenuEvent(event)

        actions = captured.get('actions', [])
        self.assertIn("Change All to 'the'", actions)
        self.assertIn("Change All to 'tech'", actions)
        self.assertIn("Ignore All ('teh')", actions)

    def test_no_op_for_identical_words(self):
        editor = make_editor('the the')
        self.assertEqual(editor.replace_all_words('the', 'the'), 0)
        self.assertEqual(editor.toPlainText(), 'the the')


if __name__ == '__main__':
    unittest.main()

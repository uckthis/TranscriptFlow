"""Checks for the current-line indicator and its coexistence with highlights."""

import os
import unittest

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

from PyQt6.QtWidgets import QApplication

from dialogs import PreferencesDialog
from editor import TranscriptEditor


app = QApplication.instance() or QApplication([])


def make_editor():
    editor = TranscriptEditor([], [], {
        'font': 'Tahoma', 'size': 14, 'fps': 23.976,
        'spell_check': False, 'smart_typing': {'enabled': False}})
    editor.resize(700, 220)
    editor.setPlainText(
        'first line\nsecond line\nthird line\nfourth line')
    editor.show()
    app.processEvents()
    return editor


class CurrentLineTests(unittest.TestCase):
    def test_background_selection_uses_low_alpha(self):
        editor = make_editor()
        editor.set_current_line_indicator(
            True, 'background', 12, '#378ADD', False, True)
        editor._refresh_extra_selections()
        selections = editor.extraSelections()
        self.assertEqual(len(selections), 1)
        color = selections[0].format.background().color()
        self.assertEqual(color.alpha(), 30)

    def test_rail_only_has_no_extra_selection(self):
        editor = make_editor()
        editor.set_current_line_indicator(
            True, 'rail', 12, '#378ADD', False, True)
        editor._refresh_extra_selections()
        self.assertEqual(editor.extraSelections(), [])

    def test_both_mode_keeps_spell_highlight(self):
        editor = make_editor()
        editor.set_current_line_indicator(
            True, 'both', 12, '', False, True, 'arrow', 4, '#185FA5')
        editor.highlight_word(0, 5)
        self.assertEqual(len(editor.extraSelections()), 2)

    def test_focus_only_hides_without_focus(self):
        editor = make_editor()
        editor.clearFocus()
        editor.set_current_line_indicator(
            True, 'background', 12, '', True, True)
        editor._refresh_extra_selections()
        self.assertEqual(editor.extraSelections(), [])

    def test_selection_hides_band_unless_transient_highlight(self):
        editor = make_editor()
        editor.set_current_line_indicator(
            True, 'background', 12, '', False, True)
        cursor = editor.textCursor()
        cursor.setPosition(0)
        cursor.setPosition(5, cursor.MoveMode.KeepAnchor)
        editor.setTextCursor(cursor)
        editor._refresh_extra_selections()
        self.assertEqual(editor.extraSelections(), [])

    def test_line_scope_selects_visual_line_not_paragraph(self):
        editor = TranscriptEditor([], [], {
            'font': 'Tahoma', 'size': 14, 'fps': 23.976,
            'spell_check': False, 'smart_typing': {'enabled': False}})
        editor.resize(400, 160)
        editor.setPlainText(
            'short line\n' +
            'This is a very long paragraph that should wrap over several '
            'visual lines because the editor is narrow. ' * 3)
        editor.show()
        app.processEvents()
        cursor = editor.textCursor()
        cursor.setPosition(120)
        editor.setTextCursor(cursor)
        editor.set_current_line_indicator(
            True, 'background', 12, '', False, True, scope='line')
        editor._refresh_extra_selections()
        line_text = editor.extraSelections()[0].cursor.selectedText()
        editor.set_current_line_indicator(
            True, 'background', 12, '', False, True, scope='paragraph')
        editor._refresh_extra_selections()
        paragraph_text = editor.extraSelections()[0].cursor.selectedText()
        self.assertLess(len(line_text), len(paragraph_text))
        self.assertLess(len(line_text), 60)

    def test_rail_reserves_left_gutter(self):
        editor = make_editor()
        editor.set_current_line_indicator(
            True, 'rail', 12, '', False, True, 'solid', 5)
        margin = editor.document().rootFrame().frameFormat().leftMargin()
        self.assertGreaterEqual(margin, 18)
        editor.set_current_line_indicator(
            True, 'background', 12, '', False, True)
        self.assertEqual(
            editor.document().rootFrame().frameFormat().leftMargin(), 4)

    def test_rail_shape_and_color_settings(self):
        editor = make_editor()
        editor.set_current_line_indicator(
            True, 'rail', 12, '', False, True, 'arrow', 5, '#ff0000')
        self.assertEqual(editor.current_line_rail_shape, 'arrow')
        self.assertEqual(editor.current_line_rail_width, 5)
        self.assertEqual(editor._rail_color().name(), '#ff0000')

    def test_editor_tab_exists_and_saves_settings(self):
        config = {
            'settings': {'font': 'Tahoma', 'size': 14, 'fps': 23.976},
            'backup_interval': 1,
            'default_window_size': [1024, 800],
            'ui': {'player_scaling_behavior': 'proportional'},
            'ocr': {}, 'playback': {}
        }
        dialog = PreferencesDialog(None, config, {})
        labels = [dialog.tabs.tabText(i) for i in range(dialog.tabs.count())]
        self.assertIn('Editor', labels)
        self.assertLess(labels.index('Editor'), labels.index('New Document'))
        saved = dialog.get_config()['ui']
        self.assertIn('current_line_style', saved)
        self.assertIn('current_line_scope', saved)
        self.assertIn('current_line_rail_shape', saved)
        self.assertIn('bottom_padding_percent', saved)


if __name__ == '__main__':
    unittest.main()

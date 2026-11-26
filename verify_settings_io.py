"""Checks for Preferences settings export and import."""

import json
import os
import tempfile
import unittest
from unittest.mock import patch

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

from PyQt6.QtWidgets import QApplication, QDialog, QMessageBox

import dialogs


app = QApplication.instance() or QApplication([])


def sample_config():
    return {
        'settings': {
            'font': 'Tahoma', 'size': 14, 'fps': 23.976,
            'spell_check': True, 'spell_check_lang': 'en_GB',
            'smart_typing': {'enabled': True, 'custom_triggers': '--, :'}
        },
        'shortcuts': [{'command': 'Toggle Pause and Play', 'trigger': 'Space'}],
        'snippets': [{'name': 'Interviewer', 'trigger': 'Alt+I', 'text': 'INTERVIEWER: '}],
        'recent_files': ['D:/work/one.tflow'],
        'backup_interval': 5,
        'default_window_size': [1400, 900],
        'ui': {
            'theme': 'Sepia', 'bottom_padding_percent': 33,
            'current_line_style': 'both', 'current_line_scope': 'paragraph',
            'player_scaling_behavior': 'island'
        },
        'playback': {'pitch_lock': False, 'preferred_player': 'mpv'},
        'ocr': {'engine': 'windows', 'shortcut': 'Ctrl+Shift+O'}
    }


class SettingsIoTests(unittest.TestCase):
    def test_export_includes_shortcuts_snippets_and_settings(self):
        config = sample_config()
        dialog = dialogs.PreferencesDialog(None, config, {})
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, 'settings.json')
            with patch.object(dialogs.QFileDialog, 'getSaveFileName',
                              return_value=(path, '')), \
                 patch.object(dialogs.QMessageBox, 'information'):
                dialog.export_settings()
            with open(path, 'r', encoding='utf-8') as handle:
                exported = json.load(handle)
        self.assertEqual(exported['shortcuts'], config['shortcuts'])
        self.assertEqual(exported['snippets'], config['snippets'])
        self.assertEqual(exported['settings']['smart_typing']['custom_triggers'],
                         '--, :')
        self.assertEqual(exported['ui']['current_line_scope'], 'paragraph')

    def test_import_replaces_settings_and_accepts_dialog(self):
        incoming = sample_config()
        incoming['shortcuts'] = [{'command': 'Imported', 'trigger': 'F9'}]
        incoming['snippets'] = [{'name': 'Imported snippet', 'trigger': 'F10',
                                 'text': 'HELLO'}]
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, 'settings.json')
            with open(path, 'w', encoding='utf-8') as handle:
                json.dump(incoming, handle)
            dialog = dialogs.PreferencesDialog(None, sample_config(), {})
            with patch.object(dialogs.QFileDialog, 'getOpenFileName',
                              return_value=(path, '')), \
                 patch.object(dialogs.QMessageBox, 'question',
                              return_value=QMessageBox.StandardButton.Yes), \
                 patch.object(dialogs.QMessageBox, 'information'), \
                 patch.object(dialogs, 'get_config_path',
                              return_value=os.path.join(folder, 'missing.json')):
                dialog.import_settings()
            result = dialog.get_config()
        self.assertEqual(dialog.result(), QDialog.DialogCode.Accepted)
        self.assertEqual(result['shortcuts'], incoming['shortcuts'])
        self.assertEqual(result['snippets'], incoming['snippets'])
        self.assertEqual(result['ui']['bottom_padding_percent'], 33)
        self.assertEqual(result['settings']['smart_typing']['custom_triggers'],
                         '--, :')

    def test_invalid_import_does_not_accept_dialog(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, 'bad.json')
            with open(path, 'w', encoding='utf-8') as handle:
                handle.write('{ not json')
            dialog = dialogs.PreferencesDialog(None, sample_config(), {})
            with patch.object(dialogs.QFileDialog, 'getOpenFileName',
                              return_value=(path, '')), \
                 patch.object(dialogs.QMessageBox, 'warning'):
                dialog.import_settings()
        self.assertNotEqual(dialog.result(), QDialog.DialogCode.Accepted)


if __name__ == '__main__':
    unittest.main()

"""Checks for session time, per-file time, and status-bar word count."""

import os
import time
import unittest
from unittest.mock import patch

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QApplication

import main


app = QApplication.instance() or QApplication([])
window = main.MainWindow()
window.sm.save = lambda config: True
window.stats_timer.stop()
window.engine.is_playing = lambda: False


class SessionStatsTests(unittest.TestCase):
    def setUp(self):
        window.session_seconds = 0.0
        window._file_seconds = {}
        window._last_activity_at = time.monotonic()
        window._last_stats_tick = time.monotonic()
        window._last_stats_tick_saved = window._last_stats_tick

    def tick(self, elapsed):
        window._last_activity_at = time.monotonic()
        window._last_stats_tick = time.monotonic() - elapsed
        with patch.object(QApplication, 'applicationState',
                          return_value=Qt.ApplicationState.ApplicationActive):
            window._stats_tick()

    def test_duration_format(self):
        self.assertEqual(window._format_duration(0), '00:00:00')
        self.assertEqual(window._format_duration(3661), '01:01:01')

    def test_active_tick_accumulates_session_and_file_time(self):
        window._set_current_file('D:/tmp/a.tflow')
        self.tick(5)
        self.assertGreaterEqual(window.session_seconds, 4.5)
        key = window._time_key('D:/tmp/a.tflow')
        self.assertGreaterEqual(window._file_seconds.get(key, 0), 4.5)
        self.assertTrue(window.lbl_session_time.text().startswith('Session 00:00:0'))

    def test_file_times_are_separate(self):
        window._set_current_file('D:/tmp/a.tflow')
        self.tick(3)
        a_key = window._time_key('D:/tmp/a.tflow')
        a_seconds = window._file_seconds[a_key]

        window._set_current_file('D:/tmp/b.tflow')
        self.tick(2)
        b_key = window._time_key('D:/tmp/b.tflow')
        self.assertAlmostEqual(window._file_seconds[a_key], a_seconds, places=1)
        self.assertGreaterEqual(window._file_seconds.get(b_key, 0), 1.5)
        self.assertGreaterEqual(window.session_seconds, 4.5)

    def test_idle_time_is_not_counted(self):
        window._set_current_file('D:/tmp/a.tflow')
        self.tick(2)
        before = window.session_seconds
        window._last_activity_at = time.monotonic() - 120
        window._last_stats_tick = time.monotonic() - 5
        with patch.object(QApplication, 'applicationState',
                          return_value=Qt.ApplicationState.ApplicationActive):
            window._stats_tick()
        self.assertAlmostEqual(window.session_seconds, before, places=1)

    def test_word_count_label(self):
        window.editor.setPlainText('one two\nthree')
        window._update_status_word_count()
        self.assertEqual(window.lbl_status_words.text(), '3 words')

    def test_file_times_are_persisted_to_config(self):
        window._set_current_file('D:/tmp/a.tflow')
        self.tick(2)
        window._persist_file_times()
        key = window._time_key('D:/tmp/a.tflow')
        self.assertIn(key, window.config['file_time'])


if __name__ == '__main__':
    unittest.main()

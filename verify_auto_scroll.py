"""Checks for the single Auto Scroll follow mode."""

import os
import time
import unittest
from unittest.mock import patch

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import QPropertyAnimation, QEasingCurve, Qt
from PyQt6.QtTest import QTest

from editor import TranscriptEditor


app = QApplication.instance() or QApplication([])


def arm_mouse(editor):
    """Simulate a live drag: flag plus fresh watchdog timestamps."""
    editor._mouse_selection_down = True
    editor._mouse_dragging = True
    now = time.monotonic()
    editor._mouse_down_at = now
    editor._mouse_last_activity_at = now
    editor._pause_auto_scroll_for_selection()


def arm_modifier(editor):
    """Simulate a held Shift/Ctrl: flag plus a fresh watchdog timestamp."""
    editor._modifier_selection_held = True
    editor._modifier_held_at = time.monotonic()
    editor._pause_auto_scroll_for_selection()


def make_editor(lines=6, pause_manual=False, type_only=False, padding=25,
                pause_selecting=True):
    editor = TranscriptEditor([], [], {
        'font': 'Tahoma', 'size': 14, 'fps': 23.976,
        'spell_check': False, 'smart_typing': {'enabled': False}})
    editor.resize(700, 300)
    editor.setPlainText('\n'.join(f'line {i}' for i in range(60)))
    editor.show()
    app.processEvents()
    editor.set_auto_scroll(True, lines, pause_manual, type_only, padding,
                           pause_selecting)
    cursor = editor.textCursor()
    cursor.setPosition(300)
    editor.setTextCursor(cursor)
    app.processEvents()
    return editor


class AutoScrollTests(unittest.TestCase):
    def test_lines_ahead_max_is_85_percent_of_visible_lines(self):
        editor = make_editor()
        maximum = editor.max_lines_ahead()
        visible = editor.viewport().height() / editor._effective_line_height()
        self.assertGreaterEqual(maximum, 3)
        self.assertLessEqual(maximum, int(visible * 0.85) + 1)

    def test_follow_scroll_updates_scrollbar_immediately(self):
        editor = make_editor()
        bar = editor.verticalScrollBar()
        before = bar.value()
        editor._follow_scroll_to(before + 20)
        self.assertEqual(bar.value(), before + 20)

    def test_follow_does_not_touch_the_wheel_animation(self):
        """The follow has no easing of its own; only the wheel animation is
        allowed to run, and a follow must cancel it rather than piggy-back."""
        editor = make_editor()
        bar = editor.verticalScrollBar()
        editor._scroll_anim.setStartValue(bar.value())
        editor._scroll_anim.setEndValue(bar.value() + 400)
        editor._scroll_anim.start()
        editor._follow_scroll_to(bar.value() + 20)
        self.assertEqual(editor._scroll_anim.state(),
                         QPropertyAnimation.State.Stopped)
        self.assertIsNone(editor._scroll_target)

    def test_follow_is_not_animated(self):
        """Regression guard: the auto-scroll animation was removed because a
        one-line follow cannot be made perceptible by easing, and a running
        animation during reflow flashed the view to the top on Enter."""
        editor = make_editor()
        self.assertFalse(hasattr(editor, '_auto_scroll_anim'))
        self.assertFalse(hasattr(editor, 'auto_scroll_animated'))

    def test_guard_on_and_off_scroll_identically(self):
        """The whole point of the option: pausing while selecting must not
        change how far the view follows."""
        results = {}
        for guard in (True, False):
            editor = make_editor()
            editor.set_auto_scroll(True, 6, False, False, 25, guard)
            editor.verticalScrollBar().setValue(0)
            cursor = editor.textCursor()
            cursor.setPosition(0)
            editor.setTextCursor(cursor)
            app.processEvents()
            for _ in range(30):
                cursor = editor.textCursor()
                cursor.insertText('\n')
                editor.setTextCursor(cursor)
                editor._handle_auto_scroll(from_typing=True,
                                           content_changed=True)
            results[guard] = editor.verticalScrollBar().value()
            self.assertEqual(editor._last_follow_skip, '')
        self.assertEqual(results[True], results[False])
        self.assertGreater(results[True], 0)

    def test_pause_after_manual_scroll_blocks_auto_scroll(self):
        editor = make_editor(pause_manual=True)
        editor.verticalScrollBar().setValue(0)
        app.processEvents()
        calls = []
        editor._follow_scroll_to = lambda target: calls.append(target)
        editor._last_manual_scroll_at = time.monotonic()
        editor._handle_auto_scroll()
        self.assertEqual(calls, [])
        editor._last_manual_scroll_at = 0.0
        editor._handle_auto_scroll()
        self.assertTrue(calls)

    def test_type_only_trigger_blocks_cursor_moves(self):
        editor = make_editor(type_only=True)
        calls = []
        editor._follow_scroll_to = lambda target: calls.append(target)
        editor._handle_auto_scroll()
        self.assertEqual(calls, [])
        editor._handle_auto_scroll(from_typing=True)
        self.assertTrue(calls)

    def test_active_selection_gesture_blocks_auto_scroll(self):
        editor = make_editor()
        editor.verticalScrollBar().setValue(0)
        calls = []
        editor._follow_scroll_to = lambda target: calls.append(target)
        # A live drag: selection is being made right now. Set the flag before
        # moving the cursor - mousePressEvent does exactly that in real use.
        arm_mouse(editor)
        cursor = editor.textCursor()
        cursor.setPosition(280)
        cursor.setPosition(320, cursor.MoveMode.KeepAnchor)
        editor.setTextCursor(cursor)
        editor._handle_auto_scroll()
        self.assertEqual(calls, [])

    def test_leftover_selection_does_not_block_auto_scroll(self):
        """A command that selects text (e.g. Go To Next Time Code) must not
        freeze auto-scroll for the rest of the document."""
        editor = make_editor()
        editor.verticalScrollBar().setValue(0)
        calls = []
        editor._follow_scroll_to = lambda target: calls.append(target)
        cursor = editor.textCursor()
        cursor.setPosition(280)
        cursor.setPosition(320, cursor.MoveMode.KeepAnchor)
        editor.setTextCursor(cursor)
        self.assertTrue(editor.textCursor().hasSelection())
        self.assertFalse(editor._selection_guard_active())
        editor._handle_auto_scroll()
        self.assertTrue(calls)

    def test_mouse_down_blocks_auto_scroll(self):
        editor = make_editor()
        editor.verticalScrollBar().setValue(0)
        calls = []
        editor._follow_scroll_to = lambda target: calls.append(target)
        arm_mouse(editor)
        with patch.object(QApplication, 'mouseButtons',
                          return_value=Qt.MouseButton.LeftButton):
            editor._handle_auto_scroll()
        self.assertEqual(calls, [])
        editor._mouse_selection_down = False
        editor._selection_pause_until = 0.0
        editor._handle_auto_scroll()
        self.assertTrue(calls)

    def test_stale_modifier_latch_clears_on_plain_keypress(self):
        """A dropped keyRelease must not strand the modifier latch forever."""
        editor = make_editor()
        arm_modifier(editor)
        self.assertTrue(editor._selection_guard_active())
        QTest.keyClick(editor, Qt.Key.Key_X)  # plain key, no modifiers
        app.processEvents()
        self.assertFalse(editor._modifier_selection_held)

    def test_content_insertion_clears_modifier_latch(self):
        """Ctrl+V latches Ctrl; the paste itself must still follow."""
        editor = make_editor()
        editor.verticalScrollBar().setValue(0)
        calls = []
        editor._follow_scroll_to = lambda target: calls.append(target)
        arm_modifier(editor)
        editor._handle_auto_scroll(from_typing=True, content_changed=True)
        self.assertTrue(calls)

    def test_content_insertion_during_mouse_drag_still_blocked(self):
        """A live drag wins over an insertion - never yank the selection."""
        editor = make_editor()
        editor.verticalScrollBar().setValue(0)
        calls = []
        editor._follow_scroll_to = lambda target: calls.append(target)
        arm_mouse(editor)
        editor._handle_auto_scroll(from_typing=True, content_changed=True)
        self.assertEqual(calls, [])

    def test_plain_click_does_not_block_auto_scroll(self):
        """Placing the caret with a mouse click must not latch the guard.

        This was the user-visible bug: click to place the caret -> auto scroll
        stopped completely, while typing/arrow placement kept working.
        """
        editor = make_editor()
        editor.verticalScrollBar().setValue(0)
        calls = []
        editor._follow_scroll_to = lambda target: calls.append(target)
        pos = editor.cursorRect().center()
        QTest.mouseClick(editor.viewport(), Qt.MouseButton.LeftButton, pos=pos)
        app.processEvents()
        self.assertFalse(editor._mouse_selection_down)
        self.assertFalse(editor._selection_guard_active())
        editor._handle_auto_scroll()
        self.assertTrue(calls)

    def test_mouse_watchdog_clears_missed_release(self):
        """A swallowed release must not latch the guard forever."""
        editor = make_editor()
        editor._mouse_selection_down = True
        editor._mouse_dragging = True
        editor._mouse_down_at = time.monotonic() - 60.0
        editor._mouse_last_activity_at = time.monotonic() - 60.0
        self.assertFalse(editor._selection_guard_active())
        self.assertFalse(editor._mouse_selection_down)

    def test_modifier_watchdog_clears_missed_release(self):
        editor = make_editor()
        editor._modifier_selection_held = True
        editor._modifier_held_at = time.monotonic() - 60.0
        self.assertFalse(editor._selection_guard_active())
        self.assertFalse(editor._modifier_selection_held)

    def test_set_auto_scroll_clears_selection_guard_state(self):
        editor = make_editor()
        editor._mouse_selection_down = True
        editor._selection_pause_until = 999999.0
        editor.set_auto_scroll(True, 6, False, False, 25, True)
        self.assertFalse(editor._mouse_selection_down)
        self.assertEqual(editor._selection_pause_until, 0.0)

    def test_shift_modifier_blocks_auto_scroll(self):
        editor = make_editor()
        editor.verticalScrollBar().setValue(0)
        calls = []
        editor._follow_scroll_to = lambda target: calls.append(target)
        arm_modifier(editor)
        editor._handle_auto_scroll()
        self.assertEqual(calls, [])

    def test_resume_delay_blocks_then_allows_auto_scroll(self):
        editor = make_editor()
        editor.verticalScrollBar().setValue(0)
        calls = []
        editor._follow_scroll_to = lambda target: calls.append(target)
        editor._selection_pause_until = time.monotonic() + 1
        editor._handle_auto_scroll()
        self.assertEqual(calls, [])
        editor._selection_pause_until = 0.0
        editor._handle_auto_scroll()
        self.assertTrue(calls)

    def test_option_off_allows_scroll_during_selection(self):
        editor = make_editor()
        editor.verticalScrollBar().setValue(0)
        editor.auto_scroll_pause_selecting = False
        calls = []
        editor._follow_scroll_to = lambda target: calls.append(target)
        cursor = editor.textCursor()
        cursor.setPosition(280)
        cursor.setPosition(320, cursor.MoveMode.KeepAnchor)
        editor.setTextCursor(cursor)
        editor._handle_auto_scroll()
        self.assertTrue(calls)

    def test_bottom_padding_percentage(self):
        editor = make_editor(padding=30)
        editor.update_bottom_margin()
        margin = editor.document().rootFrame().frameFormat().bottomMargin()
        self.assertGreater(margin, 0)
        editor.bottom_padding_percent = 0
        editor.update_bottom_margin()
        self.assertEqual(
            editor.document().rootFrame().frameFormat().bottomMargin(), 0.0)


if __name__ == '__main__':
    unittest.main()

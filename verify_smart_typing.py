"""Focused checks for the Smart Typing & Formatting rule engine."""

import unittest

import smart_typing


class SmartTypingRuleTests(unittest.TestCase):
    def setUp(self):
        self.rules = smart_typing.rules_from_settings(
            {'smart_typing': {'enabled': True}})

    def test_single_sentence_marks_arm_capitalization(self):
        for text in ('He spoke.', 'He spoke?', 'He spoke!',
                     'He spoke. ', 'He spoke." ', 'He spoke?!'):
            with self.subTest(text=text):
                self.assertTrue(smart_typing.should_capitalize(text, self.rules))

    def test_exactly_three_dots_suppresses_but_two_does_not(self):
        self.assertFalse(smart_typing.should_capitalize('He trailed...', self.rules))
        self.assertTrue(smart_typing.should_capitalize('He trailed..', self.rules))
        self.assertTrue(smart_typing.should_capitalize('He trailed....', self.rules))
        self.assertFalse(smart_typing.should_capitalize('He trailed\u2026', self.rules))

    def test_ellipsis_mode_can_capitalize(self):
        cap = dict(self.rules, capitalize_after_ellipsis=True)
        self.assertTrue(smart_typing.should_capitalize('He trailed...', cap))
        self.assertTrue(smart_typing.should_capitalize('He trailed\u2026', cap))
        self.assertTrue(smart_typing.should_capitalize('He trailed..', cap))

    def test_custom_triggers(self):
        rules = smart_typing.rules_from_settings({
            'smart_typing': {'enabled': True, 'custom_triggers': '--, :, ~'}})
        self.assertEqual(
            set(smart_typing.parse_custom_triggers(rules['custom_triggers'])),
            {'--', ':', '~'})
        for text in ('hello--', 'hello:', 'hello~'):
            with self.subTest(text=text):
                self.assertTrue(smart_typing.should_capitalize(text, rules))
        self.assertFalse(smart_typing.should_capitalize('hello-', rules))
        self.assertFalse(smart_typing.should_capitalize('hello#', rules))

    def test_abbreviations_and_timecodes_are_ignored(self):
        for text in ('Mr. ', 'Dr. ', 'e.g. ', 'U.S. ', 'J. ',
                     '[00:01:23.29] ', 'At 5 p.m. '):
            with self.subTest(text=text):
                self.assertFalse(smart_typing.should_capitalize(text, self.rules))
        self.assertFalse(smart_typing.should_capitalize('[00:01:23.', self.rules))

    def test_newline_rule(self):
        self.assertTrue(smart_typing.should_capitalize('first line\n', self.rules))
        no_newline = dict(self.rules, capitalize_after_newline=False)
        self.assertFalse(smart_typing.should_capitalize('first line\n', no_newline))

    def test_master_and_individual_switches(self):
        self.assertFalse(smart_typing.should_capitalize(
            'He spoke.', smart_typing.rules_from_settings({})))
        no_period = dict(self.rules, capitalize_after_period=False)
        self.assertFalse(smart_typing.should_capitalize('He spoke.', no_period))
        self.assertTrue(smart_typing.should_capitalize('He spoke?', no_period))

    def test_lookahead_finds_only_editable_letters(self):
        self.assertEqual(smart_typing.find_next_lowercase(' hello'), 1)
        self.assertEqual(smart_typing.find_next_lowercase('"hello'), 1)
        self.assertEqual(smart_typing.find_next_lowercase('\tHELLO'), -1)
        self.assertEqual(smart_typing.find_next_lowercase('[laughter]'), -1)
        self.assertEqual(smart_typing.find_next_lowercase('   '), -1)


if __name__ == '__main__':
    unittest.main()

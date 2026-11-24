"""Pure decision rules for Smart Typing & Formatting.

This module deliberately has no Qt dependency so the punctuation rules can be
verified independently of the editor widget.
"""

import re

DEFAULT_RULES = {
    'enabled': False,
    'capitalize_after_period': True,
    'capitalize_after_question': True,
    'capitalize_after_exclamation': True,
    'capitalize_after_newline': True,
    'capitalize_after_ellipsis': False,
    'custom_triggers': '',
}

_CLOSERS = ')]}"' + '\u201d\u2019\u00bb'
_OPEN_QUOTES = '"' + '\u201c\u2018'
_ABBREVIATIONS = {
    'mr', 'mrs', 'ms', 'dr', 'prof', 'sr', 'jr', 'st',
    'vs', 'etc', 'e.g', 'i.e', 'u.s', 'u.k', 'a.m', 'p.m',
}
_WORD_WITH_DOTS = re.compile(r'([A-Za-z]+(?:\.[A-Za-z]+)*)\Z')
_TIMECODE_TAIL = re.compile(r'\d{1,3}:\d{2}:\d{2}[.:]\d{1,3}\s*\Z')
_NUMERIC_TOKEN = re.compile(r'[\d:]+$')


def rules_from_settings(settings):
    """Return a complete rule set, filling missing keys with defaults."""
    stored = settings.get('smart_typing', {}) if isinstance(settings, dict) else {}
    if not isinstance(stored, dict):
        stored = {}
    rules = {}
    for key, default in DEFAULT_RULES.items():
        value = stored.get(key, default)
        if key == 'custom_triggers':
            rules[key] = value if isinstance(value, str) else ''
        else:
            rules[key] = bool(value)
    return rules


def parse_custom_triggers(value):
    """Return whole comma-separated trigger sequences, longest first."""
    if not isinstance(value, str):
        return ()
    tokens = {token.strip() for token in value.split(',') if token.strip()}
    return tuple(sorted(tokens, key=len, reverse=True))


def _meaningful_end(text_before):
    """Index of the last non-space, non-closer character before the cursor."""
    i = len(text_before) - 1
    while i >= 0 and text_before[i] in ' \t':
        i -= 1
    while i >= 0 and text_before[i] in _CLOSERS:
        i -= 1
        while i >= 0 and text_before[i] in ' \t':
            i -= 1
    return i


def ending_custom_trigger(text_before, rules):
    """Return the custom trigger sequence at the end of text_before, if any."""
    tokens = parse_custom_triggers(rules.get('custom_triggers', ''))
    if not tokens:
        return None
    i = _meaningful_end(text_before)
    if i < 0:
        return None
    for token in tokens:
        start = i - len(token) + 1
        if start >= 0 and text_before[start:i + 1] == token:
            return token
    return None


def _is_abbreviation(text_before_dot):
    match = _WORD_WITH_DOTS.search(text_before_dot)
    if not match:
        return False
    token = match.group(1)
    if token.lower() in _ABBREVIATIONS:
        return True
    # A lone capital letter before a dot is an initial, not a sentence end.
    return len(token) == 1 and token.isupper()


def should_capitalize(text_before, rules):
    """Should the next typed letter be capitalized?

    ``text_before`` is the document text immediately before that letter.
    Closing quotes/brackets and spaces after the trigger are ignored.
    """
    if not rules.get('enabled'):
        return False
    custom = parse_custom_triggers(rules.get('custom_triggers', ''))
    if not custom and not any(rules.get(key) for key in (
            'capitalize_after_period', 'capitalize_after_question',
            'capitalize_after_exclamation', 'capitalize_after_newline',
            'capitalize_after_ellipsis')):
        return False

    i = _meaningful_end(text_before)
    if i < 0:
        return False

    char = text_before[i]
    if char in '\r\n':
        return bool(rules.get('capitalize_after_newline'))
    if char == '\u2026':
        return bool(rules.get('capitalize_after_ellipsis'))

    if char == '.':
        # Exactly three dots is the ellipsis case controlled by the Options
        # checkbox. Two dots (or four or more) always behave like a plain period.
        run_end = i
        j = i
        while j >= 0 and text_before[j] == '.':
            j -= 1
        if run_end - j == 3:
            return bool(rules.get('capitalize_after_ellipsis'))

        line_start = max(text_before.rfind('\n', 0, i),
                         text_before.rfind('\r', 0, i)) + 1
        if _is_abbreviation(text_before[line_start:j + 1]):
            return False

        # A dot between digits belongs to a decimal number or a timecode.
        if j >= 0 and text_before[j].isdigit():
            token = _NUMERIC_TOKEN.search(text_before[line_start:j + 1])
            if token and ':' in token.group(0):
                return False
            after = text_before[i + 1:].lstrip(' \t').lstrip(_CLOSERS)
            if after[:1].isdigit():
                return False
            if _TIMECODE_TAIL.search(text_before[line_start:i + 1]):
                return False

        return bool(rules.get('capitalize_after_period'))

    if char == '?':
        return bool(rules.get('capitalize_after_question'))
    if char == '!':
        return bool(rules.get('capitalize_after_exclamation'))
    if ending_custom_trigger(text_before, rules):
        return True
    return False


def find_next_lowercase(text_after):
    """Index of the next editable lowercase letter, or -1 if there is none."""
    i = 0
    while i < len(text_after) and text_after[i] in ' \t':
        i += 1
    while i < len(text_after) and text_after[i] in _OPEN_QUOTES:
        i += 1
        while i < len(text_after) and text_after[i] in ' \t':
            i += 1
    if i >= len(text_after):
        return -1
    return i if 'a' <= text_after[i] <= 'z' else -1

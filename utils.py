import json
import re
import datetime
import os
import shutil
import csv
import io
from PyQt6.QtGui import QColor

class SubtitleEngine:
    """Robust engine to extract segments between timecodes regardless of position"""
    @staticmethod
    def get_segments(content, is_html=False):
        """Extract segments where each pair of timecodes defines start/end.
        
        Sequential pairing logic: First timecode is start, second is end,
        everything between them is the body, regardless of position in transcript.
        """
        # Match timecodes: bracketed or flexible unbracketed
        regex = TimecodeHelper.get_regex()
        
        # Split but keep the delimiters
        parts = re.split(regex, content)
        timecodes = re.findall(regex, content)
        
        segments = []
        # Index tracking: re.split with capturing group returns [non-match, match, non-match, match...]
        # parts[0] is text before 1st TC
        # parts[1] is 1st TC
        # parts[2] is text between 1st and 2nd TC
        # parts[3] is 2nd TC
        # parts[4] is text between 2nd and 3rd TC
        
        for i in range(len(timecodes)):
            start_tc = timecodes[i]
            
            # End TC is the next one, or None if it's the last one
            end_tc = timecodes[i+1] if i + 1 < len(timecodes) else None
            
            # Body is the text between this timecode and the next timecode
            # parts structure: [prefix, TC0, body0, TC1, body1, TC2, ...]
            # For timecode i, the body is at index (i * 2) + 2
            idx = (i * 2) + 2
            body = parts[idx] if idx < len(parts) else ""
            
            # Include empty subtitles as per user requirement
            segments.append({
                'start': start_tc,
                'end': end_tc,
                'body': body.strip()
            })
            
        return segments

def get_contrast_color(color):
    """Returns 'black' or 'white' based on luminance of the input QColor"""
    r, g, b = color.red(), color.green(), color.blue()
    # Perceptive luminance formula
    luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    return 'black' if luminance > 0.6 else 'white'

class TimecodeHelper:
    @staticmethod
    def get_regex(bracketed_only=False):
        """Returns a robust regex for matching timecodes with 1-3 digit fractions."""
        reg_unbracketed = r"(?:\b\d{1,2}:\d{2}:\d{2}(?:[\.,:;,]\d{1,3})?\b)"
        reg_bracketed = r"(?:[\[\({<]\d{1,2}:\d{2}:\d{2}(?:[\.,:;,]\d{1,3})?[\]\)}>])"
        if bracketed_only:
            return f"{reg_bracketed}"
        return f"({reg_bracketed}|{reg_unbracketed})"

    def __init__(self, fps=30.0, offset_ms=0):
        self.fps = float(fps)
        self.offset_ms = int(offset_ms)
        
        # Precision NTSC handling (1000/1001 ratio)
        self.is_ntsc = False
        # ONLY enable NTSC if specifically around 23.976 or 29.97 exactly
        # If user chooses 23.98 or 24, they might want real-time.
        # Broad NTSC check: covers 23.976, 23.98, 29.97, 29.98, 59.94
        if abs(self.fps - 23.976) < 0.05 or abs(self.fps - 23.98) < 0.05:
            self.base_fps = 24
            self.is_ntsc = True
        elif abs(self.fps - 29.97) < 0.05:
            self.base_fps = 30
            self.is_ntsc = True
        elif abs(self.fps - 59.94) < 0.05:
            self.base_fps = 60
            self.is_ntsc = True
        else:
            self.base_fps = int(round(self.fps)) if self.fps > 0 else 30
            
        if self.fps <= 0: self.fps = 30.0

    def timestamp_to_ms(self, tc_string, use_offset=True):
        if not tc_string: return 0
        # Cleans brackets and other delimiters
        clean = re.sub(r'[\[\]\(\)\{\}<>]', '', tc_string).strip()
        
        # Normalize separators: , ; and : (if count is 3) to a single standard for splitting
        # We need to handle SRT/VTT milliseconds (3 digits) which often use commas or dots
        clean = clean.replace(',', '.') 
        clean = clean.replace(';', '.') 
        
        if clean.count(':') == 3: # HH:MM:SS:FF
            parts = clean.rsplit(':', 1)
            clean = f"{parts[0]}.{parts[1]}"
            
        parts = clean.split('.')
        main = parts[0].split(':')
        
        # Detect if the decimal part is frames (2 digits) or ms (3 digits)
        decimal_str = parts[1] if len(parts) > 1 else ""
        val_after_sec = int(decimal_str) if decimal_str else 0
        is_millis = len(decimal_str) >= 3
        
        h, m, s = 0, 0, 0
        if len(main) == 3: h, m, s = map(int, main)
        elif len(main) == 2: m, s = map(int, main)
        elif len(main) == 1: s = int(main[0])
        
        total_sec = (h * 3600) + (m * 60) + s
        
        if is_millis:
            # Pure milliseconds (SRT/VTT style)
            total_ms = (total_sec * 1000) + val_after_sec
        else:
            # Frame-based
            if self.is_ntsc:
                # Precision NTSC: map SMPTE time to real wall-clock time
                # ms = (frames / base_fps + total_sec) * 1001
                total_ms = int(((val_after_sec / self.base_fps) + total_sec) * 1001)
            else:
                frame_ms = (val_after_sec / self.fps) * 1000
                total_ms = int((total_sec * 1000) + frame_ms)
        
        if use_offset:
            return total_ms - self.offset_ms
        return total_ms

    def ms_to_timestamp(self, ms, bracket="[]", omit_frames=False, use_frames_sep=":", use_offset=True):
        if use_offset:
            ms += self.offset_ms
        
        if ms < 0: ms = 0
        
        if self.is_ntsc:
            # Inverse of above: real time to SMPTE time
            # 1001ms real time = 1000ms TC time
            tc_seconds = ms / 1001.0
            h = int(tc_seconds // 3600)
            m = int((tc_seconds % 3600) // 60)
            s = int(tc_seconds % 60)
            f = int(round((tc_seconds % 1) * self.base_fps))
            if f >= self.base_fps:
                f = 0
                s += 1
                if s >= 60: s = 0; m += 1
                if m >= 60: m = 0; h += 1
        else:
            seconds = ms / 1000
            h = int(seconds // 3600)
            m = int((seconds % 3600) // 60)
            s = int(seconds % 60)
            f = int(round((seconds % 1) * self.fps))
            if f >= self.fps:
                f = 0
                s += 1
                if s >= 60: s = 0; m += 1
                if m >= 60: m = 0; h += 1
        
        if omit_frames:
            ts = f"{h:02d}:{m:02d}:{s:02d}"
        else:
            ts = f"{h:02d}:{m:02d}:{s:02d}{use_frames_sep}{f:02d}"
        if bracket and len(bracket) >= 2:
            return f"{bracket[0]}{ts}{bracket[1]}"
        return ts

    def shift_text_timecodes(self, text, shift_ms):
        """Finds all timecodes in text and shifts them by shift_ms."""
        regex = self.get_regex()
        
        def replace_tc(match):
            tc_str = match.group(0)
            # Detect formatting
            bracket = ""
            if tc_str.startswith('['): bracket = "[]"
            elif tc_str.startswith('('): bracket = "()"
            elif tc_str.startswith('{'): bracket = "{}"
            elif tc_str.startswith('<'): bracket = "<>"
            
            sep = ":"
            if "." in tc_str: sep = "."
            
            # Convert to ms, shift, convert back
            try:
                ms = self.timestamp_to_ms(tc_str, use_offset=False)
                new_ms = max(0, ms + shift_ms)
                return self.ms_to_timestamp(new_ms, bracket=bracket, use_frames_sep=sep, use_offset=False)
            except:
                return tc_str
                
        return re.sub(regex, replace_tc, text)

class TranscriptParser:
    """Parses transcript into structured entries: time, speaker, sfx, text"""
    @staticmethod
    def parse_text(text):
        # 0. Pre-clean HTML if present to remove style/head garbage
        is_html = '<html' in text.lower()
        if is_html:
            # Remove style and head content immediately
            text = re.sub(r'<(style|head)[^>]*>.*?</\1>', '', text, flags=re.IGNORECASE | re.DOTALL)
            # Remove the html/body wrappers but keep content
            text = re.sub(r'</?(html|body|doc-)[^>]*>', '', text, flags=re.IGNORECASE)

        # Flexible regex to find all variations of timecodes
        regex = TimecodeHelper.get_regex()
        # SFX regex: supports [], (), {}
        sfx_regex = r"[\[\(\{][^\]\)\}]*[\]\)\}]"
        
        all_parts = re.split(regex, text)
        entries = []
        current_time = "00:00:00:00"
        
        def extract_data(content, tc):
            # If it's HTML, we need to treat block tags as line breaks for parsing
            # But we keep it rich for the export
            parse_content = content
            if '<' in content and '>' in content:
                # Replace common block ends with newlines for parsing logic
                parse_content = re.sub(r'</p>|</div>|<br\s*/?>', '\n', content, flags=re.IGNORECASE)
                # Strip all other tags to get clean text for SFX/Speaker detection
                parse_content = re.sub(r'<[^>]+>', '', parse_content)

            # 1. Extract SFX
            sfx_matches = re.findall(sfx_regex, parse_content)
            rich_sfx_list = []
            clean_body = content
            
            for sfx in sfx_matches:
                esc_sfx = re.escape(sfx)
                # Look for the SFX with its potential surrounding tags
                rich_sfx_match = re.search(r'(<[^>]+>)*' + esc_sfx + r'(<[^>]+>)*', clean_body)
                if rich_sfx_match:
                    found_rich = rich_sfx_match.group(0)
                    rich_sfx_list.append(found_rich)
                    # REDACT from body to avoid duplication
                    clean_body = clean_body.replace(found_rich, "", 1)
            
            # 2. Extract Speaker
            # First, find the plain speaker from the plain version
            lines = parse_content.strip().split('\n')
            speaker = ""
            if lines:
                first_line = lines[0].strip()
                # Strip SFX from start of first line to see if there's a speaker
                temp_line = re.sub(r"^(\s*" + sfx_regex + r"\s*)*", "", first_line)
                delim_idx = temp_line.find(":")
                if delim_idx != -1:
                    pot_spk = temp_line[:delim_idx].strip()
                    if 0 < len(pot_spk) < 50 and "[" not in pot_spk and "(" not in pot_spk:
                        # We have a candidate. Now find it in the rich clean_body
                        # Look for the first colon in the rich content
                        rich_colon_idx = clean_body.find(':')
                        if rich_colon_idx != -1:
                            rich_prefix = clean_body[:rich_colon_idx].strip()
                            # Check if the text part of this prefix matches our speaker
                            plain_prefix = re.sub(r'<[^>]+>', '', rich_prefix).strip()
                            if plain_prefix == pot_spk:
                                speaker = rich_prefix
                                # Remove the speaker and the colon from the body
                                clean_body = clean_body[rich_colon_idx + 1:].strip()
            
            return {
                'timecode': tc,
                'speaker': speaker,
                'sfx': ", ".join(rich_sfx_list),
                'text': clean_body.strip()
            }

        i = 0
        while i < len(all_parts):
            p = all_parts[i]
            
            # Case 1: Text before the first timecode
            if i == 0:
                if p.strip() and not re.match(regex, p):
                    entries.append(extract_data(p, "00:00:00:00"))
            
            # Case 2: Timecode matched
            elif re.match(regex, p):
                current_time = p
                i += 1
                if i < len(all_parts):
                    content = all_parts[i]
                    entries.append(extract_data(content, current_time))
                else:
                    entries.append({'timecode': current_time, 'speaker': "", 'sfx': "", 'text': ""})
            i += 1
            
        return entries

class Exporter:
    @staticmethod
    def to_srt(content, fps=30.0, rich=False, settings=None):
        """Converts transcript to SubRip (.srt) using sequential timecode pairing"""
        helper = TimecodeHelper(fps)
        segments = SubtitleEngine.get_segments(content, is_html=rich)
        
        export_out_points = settings.get('export_out_points', True) if settings else True
        export_speakers = settings.get('export_speakers', False) if settings else False
        delim = settings.get('speaker_delimiter', ':') if settings else ':'
        
        srt_output = ""
        counter = 1
        
        for seg in segments:
            ms_start = helper.timestamp_to_ms(seg['start'])
            
            # End time logic
            if export_out_points and seg['end']:
                ms_end = helper.timestamp_to_ms(seg['end'])
            else:
                ms_end = ms_start + 2000 # Default 2s duration
                
            body = seg['body']
            
            # Speaker logic
            if not export_speakers:
                # Strip speaker if present
                body = re.sub(r'^[^:\[\(\n]+' + re.escape(delim), '', body).strip()
            
            if rich:
                body = Exporter._html_to_srt_rich(body)
            else:
                body = Exporter._html_to_srt_basic(body)
                
            if not body: body = " " # Ensure non-empty
            
            start_tc = Exporter._ms_to_srt_time(ms_start)
            end_tc = Exporter._ms_to_srt_time(ms_end)
            
            srt_output += f"{counter}\n{start_tc} --> {end_tc}\n{body}\n\n"
            counter += 1
            
        return srt_output

    @staticmethod
    def _ms_to_srt_time(ms):
        """Converts milliseconds to SRT format: HH:MM:SS,mmm"""
        if ms < 0: ms = 0
        h = int(ms // 3600000)
        m = int((ms % 3600000) // 60000)
        s = int((ms % 60000) // 1000)
        mmm = int(ms % 1000)
        return f"{h:02d}:{m:02d}:{s:02d},{mmm:03d}"

    @staticmethod
    def _html_to_srt_basic(html):
        """Convert HTML to plain text but preserve italic tags for basic SRT"""
        if not html: return ""
        res = html.replace('<em>', '<i>').replace('</em>', '</i>')
        def replace_italic_span(match):
            style = match.group(1).lower()
            content = match.group(2)
            if 'font-style:italic' in style: return f"<i>{content}</i>"
            return content
        res = re.sub(r'<span style="([^"]+)">(.+?)</span>', replace_italic_span, res, flags=re.IGNORECASE | re.DOTALL)
        res = re.sub(r'<(?!/?i[>\s])[^>]+>', '', res, flags=re.IGNORECASE)
        return res.strip()
    
    @staticmethod
    def _html_to_srt_rich(html):
        """High-fidelity HTML to SRT tag converter with color preservation"""
        if not html: return ""
        def replace_style(match):
            style = match.group(1).lower()
            content = match.group(2)
            color_match = re.search(r'color:\s*(#[0-9a-f]{6}|rgb\([^)]+\)|[a-z]+)', style, re.IGNORECASE)
            bold = any(x in style for x in ['font-weight:bold', 'font-weight:600', 'font-weight:700'])
            italic = 'font-style:italic' in style
            res = content
            if italic: res = f"<i>{res}</i>"
            if bold: res = f"<b>{res}</b>"
            if color_match: res = f'<font color="{color_match.group(1)}">{res}</font>'
            return res
        res = re.sub(r'<span style="([^"]+)">(.+?)</span>', replace_style, html, flags=re.IGNORECASE | re.DOTALL)
        res = res.replace('<strong>', '<b>').replace('</strong>', '</b>').replace('<em>', '<i>').replace('</em>', '</i>')
        res = re.sub(r'<br\s*/?>', '\n', res, flags=re.IGNORECASE)
        res = re.sub(r'<p[^>]*>', '', res, flags=re.IGNORECASE)
        res = res.replace('</p>', '\n').replace('</P>', '\n')
        res = re.sub(r'<(?!/?(?:b|i|u|font)[>\s])[^>]+>', '', res, flags=re.IGNORECASE)
        return res.strip()

    @staticmethod
    def to_html(content, fps=30.0, settings=None):
        """Converts transcript to a professional HTML table with resizable columns and plain text (black only)"""
        helper = TimecodeHelper(fps)
        entries = TranscriptParser.parse_text(content)
        
        # Determine columns
        export_out = settings.get('export_out_points', True) if settings else True
        export_dur = settings.get('export_durations', False) if settings else False
        export_spk = settings.get('export_speakers', False) if settings else False
        export_sfx = settings.get('export_sfx', True) if settings else True
        
        def strip_tags(text):
            if not text: return ""
            # Remove all HTML tags and trim
            return re.sub(r'<[^>]+>', '', text).strip()

        html = """
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
    body { font-family: 'Segoe UI', Tahoma, sans-serif; background: #fff; padding: 20px; }
    table { width: 100%; border-collapse: collapse; table-layout: fixed; border: 1px solid #ddd; }
    th { 
        background-color: #f8fafc; text-align: left; padding: 12px; border: 1px solid #ddd; 
        position: relative; color: #475569; font-weight: 600; font-size: 14px;
        overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
    }
    td { 
        padding: 10px; border: 1px solid #ddd; vertical-align: top; 
        color: #000; word-wrap: break-word; font-size: 13.5px; line-height: 1.5;
    }
    .time { white-space: nowrap; font-weight: 600; color: #1e293b; }
    .speaker { font-weight: 600; }
    .sfx { font-style: italic; color: #4b5563; }
    
    /* Resizer styling */
    .resizer {
        position: absolute;
        top: 0;
        right: 0;
        width: 6px;
        cursor: col-resize;
        user-select: none;
        height: 100%;
        background: transparent;
        transition: background 0.2s;
    }
    .resizer:hover { background: #3b82f6; }
    .resizing { border-right: 2px solid #3b82f6; }
</style>
</head>
<body>
<table id="resizeTable">
    <thead>
        <tr>
            <th style="width: 120px;">Start<div class="resizer"></div></th>
"""
        if export_out: html += '            <th style="width: 120px;">End<div class="resizer"></div></th>\n'
        if export_dur: html += '            <th style="width: 100px;">Duration<div class="resizer"></div></th>\n'
        if export_spk: html += '            <th style="width: 180px;">Speaker<div class="resizer"></div></th>\n'
        if export_sfx: html += '            <th style="width: 180px;">SFX<div class="resizer"></div></th>\n'
        html += '            <th>Transcript<div class="resizer"></div></th>\n        </tr>\n    </thead>\n    <tbody>\n'
        
        for i, e in enumerate(entries):
            # Calculate end time and duration
            ms_start = helper.timestamp_to_ms(strip_tags(e['timecode']))
            ms_end = 0
            duration_str = ""
            
            if i + 1 < len(entries):
                ms_end = helper.timestamp_to_ms(strip_tags(entries[i+1]['timecode']))
            else:
                ms_end = ms_start + 2000
            
            end_tc = helper.ms_to_timestamp(ms_end)
            
            if export_dur:
                diff = max(0, ms_end - ms_start)
                duration_str = helper.ms_to_timestamp(diff, bracket="", use_offset=False)

            html += "    <tr>\n"
            html += f'        <td class="time">{strip_tags(e["timecode"])}</td>\n'
            if export_out: html += f'        <td class="time">{strip_tags(end_tc)}</td>\n'
            if export_dur: html += f'        <td class="time">{strip_tags(duration_str)}</td>\n'
            
            # Use plain versions for speaker and sfx
            spk_plain = strip_tags(e["speaker"])
            sfx_plain = strip_tags(e["sfx"])
            
            if export_spk: html += f'        <td class="speaker">{spk_plain}</td>\n'
            if export_sfx: html += f'        <td class="sfx">{sfx_plain}</td>\n'
            
            # Plain text body with line breaks
            body = strip_tags(e['text']).replace('\n', '<br>')
            html += f'        <td>{body}</td>\n'
            html += "    </tr>\n"
            
        html += """
    </tbody>
</table>

<script>
document.addEventListener('DOMContentLoaded', function () {
    const table = document.getElementById('resizeTable');
    const resizers = table.querySelectorAll('.resizer');
    
    resizers.forEach(resizer => {
        resizer.addEventListener('mousedown', initDrag);
    });

    let startX, startWidth, col;

    function initDrag(e) {
        col = e.target.parentElement;
        startX = e.pageX;
        startWidth = col.offsetWidth;
        
        document.addEventListener('mousemove', drag);
        document.addEventListener('mouseup', stopDrag);
        col.classList.add('resizing');
    }

    function drag(e) {
        const width = startWidth + (e.pageX - startX);
        if (width > 40) {
            col.style.width = width + 'px';
        }
    }

    function stopDrag() {
        document.removeEventListener('mousemove', drag);
        document.removeEventListener('mouseup', stopDrag);
        if (col) col.classList.remove('resizing');
    }
});
</script>
</body>
</html>
"""
        return html

    @staticmethod
    def to_csv(content, settings=None):
        """Excel-friendly CSV with configurable columns"""
        output = io.StringIO()
        writer = csv.writer(output)
        
        export_out = settings.get('export_out_points', True) if settings else True
        export_dur = settings.get('export_durations', False) if settings else False
        export_spk = settings.get('export_speakers', False) if settings else False
        export_sfx = settings.get('export_sfx', True) if settings else True
        
        headers = ['Start']
        if export_out: headers.append('End')
        if export_dur: headers.append('Duration')
        if export_spk: headers.append('Speaker')
        if export_sfx: headers.append('SFX')
        headers.append('Transcript')
        writer.writerow(headers)
        
        entries = TranscriptParser.parse_text(content)
        fps = settings.get('fps', 30.0) if settings else 30.0
        helper = TimecodeHelper(fps)
        
        for i, e in enumerate(entries):
            ms_start = helper.timestamp_to_ms(e['timecode'])
            ms_end = 0
            if i + 1 < len(entries):
                ms_end = helper.timestamp_to_ms(entries[i+1]['timecode'])
            else:
                ms_end = ms_start + 2000
                
            row = [e['timecode']]
            if export_out:
                row.append(helper.ms_to_timestamp(ms_end))
            
            if export_dur:
                diff = max(0, ms_end - ms_start)
                row.append(helper.ms_to_timestamp(diff, bracket="", use_offset=False))
            
            if export_spk: row.append(e['speaker'])
            if export_sfx: row.append(e['sfx'])
            row.append(e['text'])
            writer.writerow(row)
            
        return output.getvalue()

    @staticmethod
    def to_tab(content, settings=None):
        """Tab-delimited text for database/spreadsheet import"""
        export_out = settings.get('export_out_points', True) if settings else True
        export_dur = settings.get('export_durations', False) if settings else False
        export_spk = settings.get('export_speakers', False) if settings else False
        export_sfx = settings.get('export_sfx', True) if settings else True
        
        entries = TranscriptParser.parse_text(content)
        fps = settings.get('fps', 30.0) if settings else 30.0
        helper = TimecodeHelper(fps)
        
        output = []
        for i, e in enumerate(entries):
            ms_start = helper.timestamp_to_ms(e['timecode'])
            ms_end = 0
            if i + 1 < len(entries):
                ms_end = helper.timestamp_to_ms(entries[i+1]['timecode'])
            else:
                ms_end = ms_start + 2000

            cols = [e['timecode']]
            if export_out:
                cols.append(helper.ms_to_timestamp(ms_end))
            
            if export_dur:
                diff = max(0, ms_end - ms_start)
                cols.append(helper.ms_to_timestamp(diff, bracket="", use_offset=False))
                
            if export_spk: cols.append(e['speaker'] or "")
            if export_sfx: cols.append(e['sfx'] or "")
            
            # Body text: replace tabs with spaces to preserve delimiter integrity
            body = e['text'].replace('\t', ' ').replace('\n', ' ')
            cols.append(body)
            
            output.append("\t".join(cols))
            
        return "\n".join(output)

    @staticmethod
    def to_scc(text, fps=29.97, settings=None):
        """Converts transcript to Scenarist Closed Caption (.scc) with proper body encoding"""
        helper = TimecodeHelper(fps)
        segments = SubtitleEngine.get_segments(text, is_html=False)
        scc_output = "Scenarist_SCC V1.0\n\n"
        
        for seg in segments:
            ms_start = helper.timestamp_to_ms(seg['start'])
            # SCC standard: use ; for drop-frame (NTSC)
            sep = ";" if helper.is_ntsc and abs(helper.fps - 29.97) < 0.05 else ":"
            scc_tc = helper.ms_to_timestamp(ms_start, bracket="", use_frames_sep=sep)
            
            body = re.sub(r'<[^>]+>', '', seg['body']).replace('\n', ' ').strip()
            if not body: continue
            
            # Convert body to parity-encoded hex pairs (simplified but more standard)
            # Standard SCC uses 2-byte pairs. We'll group chars into pairs.
            hex_data = "9420 9420 942c 942c " # Preamble: Resume Caption Loading + Erase
            
            # Simple parity/hex encoding for CC
            def char_to_hex(c):
                val = ord(c) & 0x7F
                # Add odd parity to bit 7
                parity = 0
                for i in range(7):
                    if (val >> i) & 1: parity += 1
                if parity % 2 == 0: val |= 0x80
                return hex(val)[2:].zfill(2).lower()

            words = []
            for i in range(0, len(body), 2):
                c1 = body[i]
                c2 = body[i+1] if i+1 < len(body) else ' '
                words.append(char_to_hex(c1) + char_to_hex(c2))
            
            hex_data += " ".join(words) + " 8080" # End with nulls/done
            scc_output += f"{scc_tc}\t{hex_data}\n"
                
        return scc_output

    @staticmethod
    def to_odf(text_edit, settings=None):
        from PyQt6.QtGui import QTextDocumentWriter
        from PyQt6.QtCore import QBuffer, QIODevice
        try:
            buffer = QBuffer()
            buffer.open(QIODevice.OpenModeFlag.WriteOnly)
            writer = QTextDocumentWriter(buffer, b"ODF")
            success = writer.write(text_edit.document())
            if not success: return None
            data = buffer.data()
            buffer.close()
            return bytes(data)
        except: return None

    @staticmethod
    def to_fcpxml(text, settings=None):
        entries = TranscriptParser.parse_text(text)
        xml = '<?xml version="1.0" encoding="UTF-8"?>\n<fcpxml version="1.9">\n<library>\n<event name="Transcript">\n'
        for e in entries:
            xml += f'  <clip name="{e["speaker"] or "Entry"}" start="{e["timecode"]}">\n'
            xml += f'    <note>{e["text"]}</note>\n'
            xml += f'  </clip>\n'
        xml += '</event>\n</library>\n</fcpxml>'
        return xml

    @staticmethod
    def to_fcp_markers(text, settings=None):
        entries = TranscriptParser.parse_text(text)
        output = "Name\tStart\tDuration\tColor\tNotes\n"
        for e in entries:
            output += f"{e['speaker'] or 'Marker'}\t{e['timecode']}\t00:00:01.00\tgreen\t{e['text']}\n"
        return output

    @staticmethod
    def to_stl(text, fps=30.0, settings=None):
        """Spruce STL format with sequential timecode pairing"""
        helper = TimecodeHelper(fps)
        segments = SubtitleEngine.get_segments(text, is_html=False)
        output = "$FontName = Arial\n$FontSize = 30\n\n"
        
        export_out_points = settings.get('export_out_points', True) if settings else True
        
        for seg in segments:
            ms_start = helper.timestamp_to_ms(seg['start'])
            if export_out_points and seg['end']:
                ms_end = helper.timestamp_to_ms(seg['end'])
            else:
                ms_end = ms_start + 2000
                
            s_tc = helper.ms_to_timestamp(ms_start, bracket="", use_frames_sep=":")
            e_tc = helper.ms_to_timestamp(ms_end, bracket="", use_frames_sep=":")
            body = seg['body'].replace('\n', ' ')
            output += f"{s_tc} , {e_tc} , {body}\n"
        return output


class Importer:
    @staticmethod
    def from_srt(content, fps=30.0):
        helper = TimecodeHelper(fps)
        # Simplistic SRT parser
        blocks = content.split('\n\n')
        transcript = ""
        for block in blocks:
            lines = block.split('\n')
            if len(lines) >= 3:
                tc_line = lines[1]
                if '-->' in tc_line:
                    raw_tc = tc_line.split('-->')[0].strip()
                    # Robust cleaning and conversion to frames
                    ms = helper.timestamp_to_ms(raw_tc)
                    norm_tc = helper.ms_to_timestamp(ms, bracket="[]")
                    
                    text = " ".join(lines[2:])
                    transcript += f"{norm_tc} {text}\n\n"
        return transcript

    @staticmethod
    def from_csv(content, fps=30.0):
        import csv
        import io
        helper = TimecodeHelper(fps)
        f = io.StringIO(content)
        reader = csv.DictReader(f)
        transcript = ""
        for row in reader:
            tc = row.get('Timecode', '') or row.get('Start', '')
            speaker = row.get('Speaker', '')
            sfx = row.get('SFX/Notes', '')
            text = row.get('Transcript', '')
            
            # Normalize TC
            ms = helper.timestamp_to_ms(tc)
            norm_tc = helper.ms_to_timestamp(ms, bracket="[]")
            
            line = f"{norm_tc} "
            if speaker: line += f"{speaker}: "
            if sfx: line += f"{sfx} "
            line += text
            transcript += line + "\n\n"
        return transcript

    @staticmethod
    def from_scc(content, fps=29.97):
        helper = TimecodeHelper(fps)
        transcript = ""
        lines = content.split('\n')
        for line in lines:
            if '\t' in line:
                tc, hex_part = line.split('\t', 1)
                
                # Normalize TC
                ms = helper.timestamp_to_ms(tc.strip())
                norm_tc = helper.ms_to_timestamp(ms, bracket="[]")
                
                # Decode hex words (4-digit hex strings)
                # SCC uses 2-byte pairs. Each byte has a parity bit in bit 7.
                hex_words = hex_part.replace("9420", "").replace("942c", "").split()
                chars = []
                for word in hex_words:
                    # Each word is 4 hex chars (2 bytes)
                    for i in range(0, len(word), 2):
                        hex_byte = word[i:i+2]
                        try:
                            val = int(hex_byte, 16) & 0x7F # Strip parity bit
                            if 32 <= val <= 126: # Only printable ASCII
                                chars.append(chr(val))
                        except:
                            pass
                text = "".join(chars).strip()
                transcript += f"{norm_tc} {text}\n\n"
        return transcript

    @staticmethod
    def from_stl(content, fps=30.0):
        # 00:00:01:00 , 00:00:03:00 , Text
        helper = TimecodeHelper(fps)
        transcript = ""
        for line in content.split('\n'):
            if ',' in line and ':' in line:
                parts = line.split(',')
                if len(parts) >= 3:
                    raw_tc = parts[0].strip()
                    ms = helper.timestamp_to_ms(raw_tc)
                    norm_tc = helper.ms_to_timestamp(ms, bracket="[]")
                    
                    text = parts[2].strip()
                    transcript += f"{norm_tc} {text}\n\n"
        return transcript

    @staticmethod
    def from_tab(content, fps=30.0):
        # Tcode \t Text
        helper = TimecodeHelper(fps)
        transcript = ""
        for line in content.split('\n'):
            if '\t' in line:
                parts = line.split('\t')
                ms = helper.timestamp_to_ms(parts[0])
                norm_tc = helper.ms_to_timestamp(ms, bracket="[]")
                transcript += f"{norm_tc} {parts[1]}\n\n"
        return transcript

class SettingsManager:
    """Handles persistent storage of all config in a single file"""
    def __init__(self):
        from path_manager import get_config_path
        self.config_path = get_config_path()

    def load(self, defaults):
        import os
        # Create a deep copy of defaults to work with
        import copy
        config = copy.deepcopy(defaults)
        
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    import json
                    saved = json.load(f)
                    
                    # Recursive merge function
                    def deep_merge(target, source):
                        for k, v in source.items():
                            if k in target and isinstance(target[k], dict) and isinstance(v, dict):
                                deep_merge(target[k], v)
                            else:
                                target[k] = v
                                
                    deep_merge(config, saved)
                    return config
            except Exception as e:
                print(f"Error loading config: {e}")
        return defaults

    def save(self, data):
        import json
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
            return True
        except Exception as e:
            print(f"Error saving config: {e}")
        return False

class FileManager:
    @staticmethod
    def save_tflow(path, data_dict):
        """
        Saves the transcript data to a JSON file (.tflow)
        data_dict should contain:
        - content (HTML/Text)
        - media_path
        - cursor_position
        - playback_position
        - timestamp
        """
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data_dict, f, indent=4)
            return True
        except Exception as e:
            print(f"Error saving .tflow: {e}")
            return False

    @staticmethod
    def load_tflow(path):
        """Loads .tflow JSON file"""
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading .tflow: {e}")
            return None

class BackupManager:
    def __init__(self, settings_manager=None):
        """
        Initialize BackupManager.
        settings_manager: Instance of SettingsManager to get/set backup directory config.
        """
        self.sm = settings_manager
        
        # Use AppData for backups instead of install directory
        from path_manager import get_backup_dir
        self.default_backup_dir = get_backup_dir()
        
    def get_backup_dir(self):
        """Get the current backup directory from config or default"""
        if self.sm:
            config = self.sm.load({})
            return config.get('backup_dir', self.default_backup_dir)
        return self.default_backup_dir
    
    def set_backup_dir(self, path):
        """Set a new backup directory and persist it"""
        if self.sm:
            config = self.sm.load({})
            config['backup_dir'] = path
            self.sm.save(config)
            
    def ensure_backup_dir(self):
        path = self.get_backup_dir()
        if not os.path.exists(path):
            os.makedirs(path)
        return path

    def save_backup(self, data_dict, prefix="autosave"):
        """Saves a timestamped backup"""
        backup_dir = self.ensure_backup_dir()
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"{prefix}_{timestamp}.tflow"
        filepath = os.path.join(backup_dir, filename)
        
        success = FileManager.save_tflow(filepath, data_dict)
        if success:
            # Enforce retention policies from config
            config = self.sm.load({})
            max_per_file = config.get('backups_per_file', 5)
            ret_months = config.get('backup_retention_months', 6)
            self.prune_backups(max_per_file=max_per_file, retention_months=ret_months)
            
        return success

    def clear_all_backups(self):
        """Deletes all .tflow files in the backup directory"""
        backup_dir = self.get_backup_dir()
        if not os.path.exists(backup_dir):
            return True
            
        success = True
        for f in os.listdir(backup_dir):
            if f.endswith('.tflow'):
                try:
                    os.remove(os.path.join(backup_dir, f))
                except Exception as e:
                    print(f"Error deleting {f}: {e}")
                    success = False
        return success

    def get_backups(self):
        """Returns list of backups sorted by modified time (newest first)"""
        backup_dir = self.get_backup_dir()
        if not os.path.exists(backup_dir):
            return []
            
        backups = []
        for f in os.listdir(backup_dir):
            if f.endswith('.tflow'):
                path = os.path.join(backup_dir, f)
                try:
                    stats = os.stat(path)
                    backups.append({
                        'filename': f,
                        'path': path,
                        'size': stats.st_size,
                        'mtime': stats.st_mtime,
                        'date': datetime.datetime.fromtimestamp(stats.st_mtime)
                    })
                except:
                    pass
        
        # Sort by mtime descending
        backups.sort(key=lambda x: x['mtime'], reverse=True)
        return backups

    def prune_backups(self, max_per_file=5, retention_months=6):
        """
        Customizable backup pruning.
        
        1. Age limit: Delete anything older than retention_months.
        2. Per-file limit: Keep only latest X backups for each unique prefix (e.g. 'autosave_')
        """
        backups = self.get_backups()
        if not backups:
            return

        now = datetime.datetime.now()
        
        # --- 1. Age-based Pruning ---
        if retention_months > 0:
            # Simple approximation: 30 days per month
            cutoff = now - datetime.timedelta(days=retention_months * 30)
            
            non_expired = []
            for b in backups:
                if b['date'] < cutoff:
                    try:
                        os.remove(b['path'])
                    except:
                        pass
                else:
                    non_expired.append(b)
            backups = non_expired

        # --- 2. Per-File Pruning ---
        # Group by prefix (part before the first '_' timestamp separator)
        if max_per_file > 0:
            groups = {}
            for b in backups:
                # Expecting format: prefix_YYYY-MM-DD_HH-MM-SS.tflow
                parts = b['filename'].rsplit('_', 2) # Get prefix
                prefix = parts[0] if len(parts) > 1 else 'unknown'
                
                if prefix not in groups:
                    groups[prefix] = []
                groups[prefix].append(b)
            
            for prefix, group in groups.items():
                if len(group) > max_per_file:
                    # group is already sorted newest first because it came from get_backups()
                    for b in group[max_per_file:]:
                        try:
                            os.remove(b['path'])
                        except:
                            pass
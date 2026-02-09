from utils import TimecodeHelper

def test_timecodes():
    print("Testing TimecodeHelper...")
    
    # PAL (25 fps)
    pal = TimecodeHelper(25)
    ms = pal.timestamp_to_ms("[00:00:01:00]", use_offset=False)
    print(f"PAL 25fps: 01:00 -> {ms}ms (Expected: 1000ms)")
    
    # NTSC (29.97 fps)
    ntsc = TimecodeHelper(29.97)
    ms_ntsc = ntsc.timestamp_to_ms("[00:00:01:00]", use_offset=False)
    print(f"NTSC 29.97fps: 01:00 -> {ms_ntsc}ms (Expected: ~1001ms)")
    
    back_tc = ntsc.ms_to_timestamp(ms_ntsc, bracket="", use_offset=False)
    print(f"NTSC Roundtrip: {ms_ntsc}ms -> {back_tc} (Expected: 00:00:01:00)")

    # Test 23.976
    film = TimecodeHelper(23.976)
    ms_film = film.timestamp_to_ms("[00:00:01:00]", use_offset=False)
    print(f"Film 23.976fps: 01:00 -> {ms_film}ms (Expected: ~1001ms)")

if __name__ == "__main__":
    test_timecodes()

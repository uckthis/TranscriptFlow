from utils import TimecodeHelper

def test_sync():
    print("Testing Sync/Shift Logic...")
    helper = TimecodeHelper(fps=24)
    
    # Simple transcript
    text = "[00:00:00:00] Start\n[00:01:00:00] Middle"
    
    # Shift by 2 seconds
    shift_ms = 2000
    new_text = helper.shift_text_timecodes(text, shift_ms)
    
    print(f"Original:\n{text}")
    print(f"Shifted (+2s):\n{new_text}")
    
    if "[00:00:02:00]" in new_text and "[00:01:02:00]" in new_text:
        print("PASS: Shift logic works for bracketed timecodes.")
    else:
        print("FAIL: Shift logic failed.")

    # Mixed formats
    text_mixed = "00:00:00,000 First\n<00:00:10.00> Second"
    new_mixed = helper.shift_text_timecodes(text_mixed, 500)
    print(f"Mixed Shifted (+0.5s):\n{new_mixed}")
    
    if "00:00:00.500" in new_mixed or "00:00:00.50" in new_mixed or "0:00:00:15" in new_mixed:
         print("PASS: Mixed format shifting handled.")
    else:
         print("FAIL: Mixed format shifting failed.")

if __name__ == "__main__":
    test_sync()

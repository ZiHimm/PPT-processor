# test_progress_fix.py
import queue

print("Testing progress message parsing...")
print("="*60)

test_messages = [
    "PROGRESS:1",
    "PROGRESS:25",
    "PROGRESS:50", 
    "PROGRESS:100",
    "PROGRESS:100%",
    "PROGRESS: 50 ",
    "PROGRESS:25% and some text",
    "DONE"
]

for msg in test_messages:
    print(f"\nTesting: '{msg}'")
    
    if "PROGRESS:" in msg:
        try:
            percent_str = msg.split("PROGRESS:")[1].strip()
            import re
            numbers = re.findall(r'\d+', percent_str)
            if numbers:
                percent = int(numbers[0])
                print(f"  ✓ Parsed: {percent}%")
            else:
                print(f"  ✗ No numbers found")
        except Exception as e:
            print(f"  ✗ Error: {e}")
    elif msg == "DONE":
        print(f"  ✓ DONE message")

print(f"\n{'='*60}")
print("If this works, your progress bar parsing is fixed!")
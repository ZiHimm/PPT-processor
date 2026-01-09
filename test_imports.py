# test_imports.py
import os
import sys

print("Current directory:", os.getcwd())
print("\nFiles in directory:")
for file in sorted(os.listdir()):
    print(f"  {file}")

print("\nTrying to import processor...")
try:
    from processor import process_presentations
    print("✓ Successfully imported process_presentations")
except ImportError as e:
    print(f"✗ Import failed: {e}")
    
print("\nPython path:")
for path in sys.path:
    print(f"  {path}")
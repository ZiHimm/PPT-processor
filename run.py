#!/usr/bin/env python
"""
Simple runner script
"""
import sys
import os

# Add current directory to path
sys.path.insert(0, os.path.dirname(__file__))

from main import main

if __name__ == "__main__":
    print("Starting PPT Post Extractor...")
    main()
    print("Done!")
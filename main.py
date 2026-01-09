# main.py
import argparse
import sys
import traceback
import logging
from processor import process_presentations

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%H:%M:%S'
)

def main():
    parser = argparse.ArgumentParser(
        description="Extract social media posts from PowerPoint files (multi-month support)"
    )
    parser.add_argument(
        "ppts",
        nargs="+",
        help="Input PowerPoint files (.pptx)"
    )
    parser.add_argument(
        "--output",
        "-o",
        required=True,
        help="Output Excel file (.xlsx)"
    )
    parser.add_argument(
        "--debug",
        "-d",
        action="store_true",
        help="Enable debug mode"
    )

    args = parser.parse_args()
    
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    
    try:
        count = process_presentations(args.ppts, args.output)
        print(f"\n{'='*60}")
        print(f"✅ PROCESSING COMPLETE")
        print(f"📊 Total Posts Extracted: {count:,}")
        print(f"📁 Output File: {args.output}")
        print(f"{'='*60}")
    except FileNotFoundError as e:
        print(f"❌ File not found: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Error: {e}")
        if args.debug:
            print(traceback.format_exc())
        sys.exit(1)

if __name__ == "__main__":
    main()
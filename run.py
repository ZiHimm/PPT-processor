#!/usr/bin/env python3
"""
Marketing Dashboard Automator Launcher
Run this file to start the application
"""

import os
import sys
import subprocess
import platform
from pathlib import Path

def check_python():
    """Check if Python 3.8+ is available"""
    if sys.version_info < (3, 8):
        print(f"❌ Python 3.8+ required. You have Python {sys.version_info.major}.{sys.version_info.minor}")
        print("Please install Python 3.8 or higher from https://www.python.org/downloads/")
        input("\nPress Enter to exit...")
        sys.exit(1)
    
    print(f"✓ Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro} detected")

def check_directories():
    """Create necessary directories"""
    directories = ["output", "logs", "cache", "backups", "config"]
    
    for dir_name in directories:
        dir_path = Path(dir_name)
        if not dir_path.exists():
            dir_path.mkdir(parents=True, exist_ok=True)
            print(f"✓ Created directory: {dir_name}")

def check_config():
    """Check and create default configuration"""
    config_file = Path("config/dashboard_config.yaml")
    
    if not config_file.exists():
        print("Creating default configuration...")
        
        default_config = """dashboard_config:
  social_media:
    keywords:
      - "TF Value-Mart FB Page Wallposts Performance"
      - "TF Value-Mart IG Page Wallposts Performance"
      - "TF Value-Mart Tiktok Video Performance"
    platforms:
      Facebook:
        metrics: ["Reach/Views", "Engagement", "Likes", "Shares", "Comments", "Saved"]
      Instagram:
        metrics: ["Reach/Views", "Engagement", "Likes", "Shares", "Comments", "Saved"]
      TikTok:
        metrics: ["Views", "Engagement", "Likes", "Shares", "Comments", "Saved"]
  
  community_marketing:
    keywords:
      - "Hive Marketing Wallposts Performance"
    metrics: ["Community Name", "Reach/Views", "Engagement", "Likes", "Shares", "Comments", "Saved"]
  
  kol_engagement:
    keywords:
      - "KOL Engagement Video Performance"
    metrics: ["KOL Name", "Views", "Likes", "Shares", "Comments", "Saved"]
  
  performance_marketing:
    keywords:
      - "FB Ads Page Likes Ad"
      - "IG Ads Profile Visits Ad"
      - "TikTok Ads Followers Ads"
    metrics: ["Ads", "Impression", "Page Likes", "Profile Visits", "Follows"]
  
  promotion_posts:
    keywords:
      - "Promotion Posts Ads Performance TFVM Brand Page"
    metrics: ["Reach", "Engagement", "Likes", "Shares", "Comments", "Saved"]
"""
        
        config_file.write_text(default_config, encoding='utf-8')
        print("✓ Created default configuration")

def check_dependencies():
    """Check and install required packages"""
    required_packages = [
        "python-pptx",
        "pandas",
        "openpyxl", 
        "plotly",
        "pyyaml",
        "numpy"
    ]
    
    print("\nChecking dependencies...")
    
    try:
        # Try to import each package
        import pandas
        import pptx
        import plotly
        import yaml
        import numpy
        import openpyxl
        print("✓ All dependencies are installed")
        return True
    except ImportError as e:
        print(f"⚠ Missing dependency: {e}")
        
        response = input("\nDo you want to install missing packages? (y/n): ").lower()
        if response in ['y', 'yes']:
            print("Installing required packages...")
            
            # Install packages using pip
            import subprocess
            import sys
            
            packages_to_install = []
            for package in required_packages:
                packages_to_install.append(package)
            
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pip"])
                subprocess.check_call([sys.executable, "-m", "pip", "install"] + packages_to_install)
                print("✓ Packages installed successfully")
                return True
            except subprocess.CalledProcessError:
                print("✗ Failed to install packages")
                print("\nPlease install manually:")
                print("pip install python-pptx pandas openpyxl plotly pyyaml numpy")
                input("\nPress Enter to exit...")
                return False
        else:
            print("\nPlease install required packages manually:")
            print("pip install python-pptx pandas openpyxl plotly pyyaml numpy")
            input("\nPress Enter to exit...")
            return False

def main():
    """Main launcher function"""
    print("\n" + "="*60)
    print("Marketing Dashboard Automator v2.0")
    print("="*60 + "\n")
    
    # Check Python version
    check_python()
    
    # Check/create directories
    check_directories()
    
    # Check/create configuration
    check_config()
    
    # Check dependencies
    if not check_dependencies():
        return
    
    # Start the application
    print("\n" + "="*60)
    print("Starting application...")
    print("="*60 + "\n")
    
    try:
        # Import and run the main application
        sys.path.insert(0, str(Path(__file__).parent))
        
        # Check if main.py exists
        main_file = Path("src/main.py")
        if not main_file.exists():
            print(f"❌ Main application file not found: {main_file}")
            print("Please ensure src/main.py exists")
            input("\nPress Enter to exit...")
            return
        
        # Run the application
        from src.main import main as app_main
        app_main()
        
    except Exception as e:
        print(f"\n❌ Error starting application: {e}")
        print("\nTroubleshooting steps:")
        print("1. Make sure all Python files are in the src/ directory")
        print("2. Check that main.py exists in src/")
        print("3. Verify Python version is 3.8 or higher")
        print("4. Install required packages: pip install python-pptx pandas openpyxl plotly pyyaml numpy")
        input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()
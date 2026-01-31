#!/usr/bin/env python3
"""
μLearn Email Sender GUI
A professional email marketing tool with HTML formatting support

To run: python run_email_gui.py
"""

import sys
import os

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from email_gui import main
    print("🚀 Starting μLearn Email Sender GUI...")
    main()
except ImportError as e:
    print(f"❌ Import error: {e}")
    print("📦 Installing required packages...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas"])
    
    # Try again
    from email_gui import main
    print("🚀 Starting μLearn Email Sender GUI...")
    main()
except Exception as e:
    print(f"❌ Error starting application: {e}")
    input("Press Enter to exit...")

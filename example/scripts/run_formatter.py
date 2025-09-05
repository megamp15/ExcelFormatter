#!/usr/bin/env python3
"""
Simple runner script for the Excel Formatter.
Run this script to process all files in the input_files directory.
"""

import sys
from pathlib import Path

# Add current directory to path
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

from excel_formatter import main

if __name__ == "__main__":
    main()

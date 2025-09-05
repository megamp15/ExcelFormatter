#!/usr/bin/env python3
"""
Configuration settings for Excel Formatter application.

This module contains all configuration settings, default values,
and application constants.
"""

import os
from pathlib import Path

# Application Information
APP_NAME = "Excel Formatter"
APP_VERSION = "1.0.0"
APP_DESCRIPTION = "Advanced Excel file processing and formatting tool"

# Directory Paths
BASE_DIR = Path(__file__).parent.parent
INPUT_DIR = BASE_DIR / "input_files"

# Default to Windows Downloads folder for output
try:
    import os
    DOWNLOADS_DIR = Path(os.path.expanduser("~")) / "Downloads"
    OUTPUT_DIR = DOWNLOADS_DIR if DOWNLOADS_DIR.exists() else BASE_DIR / "output_files"
except Exception:
    # Fallback to local directory if there's any issue
    OUTPUT_DIR = BASE_DIR / "output_files"

CONFIG_DIR = BASE_DIR / "config"
LOG_DIR = BASE_DIR / "logs"

# Ensure directories exist
INPUT_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
CONFIG_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)

# Logging Configuration
LOG_LEVEL = "INFO"
LOG_FILE = LOG_DIR / "excel_formatter.log"
MAX_LOG_SIZE = 10 * 1024 * 1024  # 10MB
LOG_BACKUP_COUNT = 5

# GUI Configuration
WINDOW_TITLE = f"{APP_NAME} v{APP_VERSION}"
WINDOW_SIZE = "900x700"
WINDOW_MIN_SIZE = (800, 600)

# File Processing Configuration
SUPPORTED_INPUT_FORMATS = [".xlsx", ".xls", ".xlsm"]
OUTPUT_FORMAT = ".xlsx"
DEFAULT_ENCODING = "utf-8"

# Column Mapping Configuration
DEFAULT_CONFIG_FILE = CONFIG_DIR / "mapping_config.json"

# Excel Processing Settings
MAX_ROWS_PREVIEW = 100
DEFAULT_SHEET_INDEX = 0
FORMULA_PREFIX = "="

# Default Column Alignments
COLUMN_ALIGNMENTS = ["left", "center", "right"]

# Default Number Formats
NUMBER_FORMATS = {
    "General": "General",
    "Number": "#,##0.00",
    "Currency": "$#,##0.00",
    "Percentage": "0.00%",
    "Date": "MM/DD/YYYY",
    "Time": "h:mm AM/PM",
    "Custom": "Custom"
}

# Default Header Colors
HEADER_COLORS = {
    "Blue": "366092",
    "Green": "4CAF50",
    "Red": "F44336",
    "Orange": "FF9800",
    "Purple": "9C27B0",
    "Teal": "009688",
    "Custom": "Custom"
}

# GUI Theme Colors
COLORS = {
    "primary": "#366092",
    "secondary": "#4CAF50", 
    "background": "#f5f5f5",
    "surface": "#ffffff",
    "error": "#F44336",
    "warning": "#FF9800",
    "success": "#4CAF50",
    "text_primary": "#212121",
    "text_secondary": "#757575",
    "border": "#e0e0e0"
}

# GUI Fonts
FONTS = {
    "default": ("Segoe UI", 9),
    "heading": ("Segoe UI", 12, "bold"),
    "button": ("Segoe UI", 9),
    "monospace": ("Consolas", 9)
}

# Progress Dialog Settings
PROGRESS_UPDATE_INTERVAL = 100  # milliseconds

# Error Messages
ERROR_MESSAGES = {
    "no_input_file": "Please select an input Excel file.",
    "no_output_dir": "Please select an output directory.",
    "invalid_file_format": "Unsupported file format. Please select an Excel file (.xlsx, .xls, .xlsm).",
    "file_not_found": "Selected file does not exist.",
    "no_columns_mapped": "Please map at least one column before processing.",
    "processing_error": "An error occurred while processing the file.",
    "config_save_error": "Failed to save configuration.",
    "config_load_error": "Failed to load configuration."
}

# Success Messages  
SUCCESS_MESSAGES = {
    "file_processed": "File processed successfully!",
    "config_saved": "Configuration saved successfully!",
    "config_loaded": "Configuration loaded successfully!"
}

# Default Mapping Configuration Template
DEFAULT_MAPPING_CONFIG = {
    "output_columns": [
        {
            "name": "Column 1",
            "source_column": "",
            "alignment": "left",
            "width": 15,
            "formatting": {}
        }
    ],
    "header_formatting": {
        "bold": True,
        "background_color": "366092",
        "font_color": "FFFFFF", 
        "alignment": "center"
    },
    "column_name_alignment": {},
    "general_settings": {
        "auto_fit_columns": True
    },
    "void": {
        "enabled": False,
        "zero_columns": []
    }
}

# File Dialog Settings
FILE_DIALOG_OPTIONS = {
    "input_filetypes": [
        ("Excel files", "*.xlsx *.xls *.xlsm"),
        ("Excel Workbook", "*.xlsx"),
        ("Excel 97-2003", "*.xls"), 
        ("Excel Macro-Enabled", "*.xlsm"),
        ("All files", "*.*")
    ],
    "config_filetypes": [
        ("JSON files", "*.json"),
        ("All files", "*.*")
    ],
    "output_filetypes": [
        ("Excel Workbook", "*.xlsx"),
        ("All files", "*.*")
    ]
}
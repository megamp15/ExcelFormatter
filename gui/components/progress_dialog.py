#!/usr/bin/env python3
"""
Progress dialog component for Excel Formatter application.

This module provides a progress dialog for showing processing status
during file operations.
"""

import tkinter as tk
from tkinter import ttk
from typing import Optional

from config.settings import *


class ProgressDialog:
    """Progress dialog for showing processing status."""
    
    def __init__(self, parent, title: str = "Processing", message: str = "Please wait..."):
        """
        Initialize the progress dialog.
        
        Args:
            parent: Parent tkinter widget
            title: Dialog title
            message: Progress message
        """
        self.parent = parent
        
        # Create dialog window
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("400x150")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center dialog
        self.center_dialog()
        
        # Prevent closing
        self.dialog.protocol("WM_DELETE_WINDOW", lambda: None)
        
        self.create_widgets(message)
        
    def center_dialog(self):
        """Center dialog on parent window."""
        self.dialog.update_idletasks()
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = self.parent.winfo_rootx() + (self.parent.winfo_width() // 2) - (width // 2)
        y = self.parent.winfo_rooty() + (self.parent.winfo_height() // 2) - (height // 2)
        self.dialog.geometry(f"{width}x{height}+{x}+{y}")
        
    def create_widgets(self, message: str):
        """Create dialog widgets."""
        main_frame = ttk.Frame(self.dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Message label
        self.message_label = ttk.Label(
            main_frame,
            text=message,
            font=FONTS["default"],
            justify=tk.CENTER
        )
        self.message_label.pack(pady=(0, 20))
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(
            main_frame,
            mode="indeterminate",
            length=300
        )
        self.progress_bar.pack(pady=(0, 10))
        
        # Start animation
        self.progress_bar.start(10)
        
    def update_message(self, message: str):
        """Update the progress message."""
        self.message_label.config(text=message)
        self.dialog.update()
        
    def destroy(self):
        """Destroy the progress dialog."""
        self.progress_bar.stop()
        self.dialog.destroy()
#!/usr/bin/env python3
"""
File selector component for Excel Formatter application.

This module provides a GUI component for selecting input files
and output directories.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from typing import Callable, Optional

from config.settings import *


class FileSelector(ttk.Frame):
    """File selection component for input files and output directory."""
    
    def __init__(self, parent, input_file_var: tk.StringVar, output_dir_var: tk.StringVar, 
                 on_file_selected: Optional[Callable[[str], None]] = None):
        """
        Initialize the file selector component.
        
        Args:
            parent: Parent tkinter widget
            input_file_var: StringVar to store input file path
            output_dir_var: StringVar to store output directory path
            on_file_selected: Callback function when file is selected
        """
        super().__init__(parent)
        
        self.input_file_var = input_file_var
        self.output_dir_var = output_dir_var
        self.on_file_selected = on_file_selected
        
        # Trace variable changes
        self.input_file_var.trace('w', self._on_input_file_changed)
        
        self.create_widgets()
        
    def create_widgets(self):
        """Create and arrange the file selector widgets."""
        # Configure grid
        self.grid_columnconfigure(1, weight=1)
        
        # Input selection type
        self.input_type = tk.StringVar(value="file")
        type_frame = ttk.Frame(self)
        type_frame.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))
        
        ttk.Radiobutton(type_frame, text="Single File", variable=self.input_type, value="file",
                       command=self._on_input_type_changed).pack(side=tk.LEFT, padx=(0, 20))
        ttk.Radiobutton(type_frame, text="Folder (Batch Processing)", variable=self.input_type, value="folder",
                       command=self._on_input_type_changed).pack(side=tk.LEFT)
        
        # Input file/folder selection
        self.input_label = ttk.Label(self, text="Input Excel File:", font=FONTS["default"])
        self.input_label.grid(row=1, column=0, sticky="w", pady=(0, 5))
        
        input_frame = ttk.Frame(self)
        input_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(0, 15))
        input_frame.grid_columnconfigure(0, weight=1)
        
        self.input_entry = ttk.Entry(
            input_frame,
            textvariable=self.input_file_var,
            font=FONTS["default"],
            state="readonly"
        )
        self.input_entry.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        
        self.browse_input_btn = ttk.Button(
            input_frame,
            text="Browse...",
            command=self.browse_input,
            width=10
        )
        self.browse_input_btn.grid(row=0, column=1)
        
        self.clear_input_btn = ttk.Button(
            input_frame,
            text="Clear",
            command=self.clear_input,
            width=8
        )
        self.clear_input_btn.grid(row=0, column=2, padx=(5, 0))
        
        # File info display
        self.info_frame = ttk.LabelFrame(self, text="File Information", padding=10)
        self.info_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(0, 15))
        self.info_frame.grid_columnconfigure(1, weight=1)
        
        # File details labels
        ttk.Label(self.info_frame, text="File Name:").grid(row=0, column=0, sticky="w", pady=2)
        self.file_name_label = ttk.Label(self.info_frame, text="", foreground=COLORS["text_secondary"])
        self.file_name_label.grid(row=0, column=1, sticky="w", padx=(10, 0), pady=2)
        
        ttk.Label(self.info_frame, text="File Size:").grid(row=1, column=0, sticky="w", pady=2)
        self.file_size_label = ttk.Label(self.info_frame, text="", foreground=COLORS["text_secondary"])
        self.file_size_label.grid(row=1, column=1, sticky="w", padx=(10, 0), pady=2)
        
        ttk.Label(self.info_frame, text="File Type:").grid(row=2, column=0, sticky="w", pady=2)
        self.file_type_label = ttk.Label(self.info_frame, text="", foreground=COLORS["text_secondary"])
        self.file_type_label.grid(row=2, column=1, sticky="w", padx=(10, 0), pady=2)
        
        ttk.Label(self.info_frame, text="Sheets:").grid(row=3, column=0, sticky="w", pady=2)
        self.sheets_label = ttk.Label(self.info_frame, text="", foreground=COLORS["text_secondary"])
        self.sheets_label.grid(row=3, column=1, sticky="w", padx=(10, 0), pady=2)
        
        # Output directory selection
        output_label = ttk.Label(self, text="Output Directory:", font=FONTS["default"])
        output_label.grid(row=4, column=0, sticky="w", pady=(15, 5))
        
        output_frame = ttk.Frame(self)
        output_frame.grid(row=5, column=0, columnspan=3, sticky="ew")
        output_frame.grid_columnconfigure(0, weight=1)
        
        self.output_entry = ttk.Entry(
            output_frame,
            textvariable=self.output_dir_var,
            font=FONTS["default"]
        )
        self.output_entry.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        
        self.browse_output_btn = ttk.Button(
            output_frame,
            text="Browse...",
            command=self.browse_output_directory,
            width=10
        )
        self.browse_output_btn.grid(row=0, column=1)
        
        self.open_output_btn = ttk.Button(
            output_frame,
            text="Open",
            command=self.open_output_directory,
            width=8,
            state=tk.DISABLED
        )
        self.open_output_btn.grid(row=0, column=2, padx=(5, 0))
        
        # Update open button state
        self.output_dir_var.trace('w', self._on_output_dir_changed)
        self._on_output_dir_changed()
        
        # Initially hide file info
        self.update_file_info(None)
        
    def _on_input_type_changed(self):
        """Handle input type change (file vs folder)."""
        if self.input_type.get() == "folder":
            self.input_label.config(text="Input Folder (Batch Processing):")
            # Clear current selection when switching types
            self.input_file_var.set("")
        else:
            self.input_label.config(text="Input Excel File:")
            # Clear current selection when switching types
            self.input_file_var.set("")
            
    def browse_input(self):
        """Open dialog to select input file or folder."""
        if self.input_type.get() == "folder":
            folder_path = filedialog.askdirectory(
                title="Select Folder with Excel Files",
                initialdir=str(INPUT_DIR)
            )
            if folder_path:
                # Check if folder contains Excel files
                excel_files = self._get_excel_files_in_folder(folder_path)
                if excel_files:
                    self.input_file_var.set(folder_path)
                else:
                    messagebox.showwarning(
                        "No Excel Files", 
                        f"No Excel files found in the selected folder.\\n\\n"
                        f"Supported formats: {', '.join(SUPPORTED_INPUT_FORMATS)}"
                    )
        else:
            file_path = filedialog.askopenfilename(
                title="Select Input Excel File",
                filetypes=FILE_DIALOG_OPTIONS["input_filetypes"],
                initialdir=str(INPUT_DIR)
            )
            if file_path:
                self.input_file_var.set(file_path)
                
    def _get_excel_files_in_folder(self, folder_path):
        """Get list of Excel files in folder."""
        folder = Path(folder_path)
        excel_files = []
        for ext in SUPPORTED_INPUT_FORMATS:
            excel_files.extend(folder.glob(f"*{ext}"))
        return excel_files
        
    def clear_input(self):
        """Clear the selected input file or folder."""
        self.input_file_var.set("")
        
    def is_folder_mode(self):
        """Check if currently in folder mode."""
        return self.input_type.get() == "folder"
        
    def is_folder_selected(self):
        """Check if a folder is currently selected."""
        path = self.input_file_var.get()
        return path and Path(path).is_dir()
        
    def browse_output_directory(self):
        """Open dialog to select output directory."""
        directory = filedialog.askdirectory(
            title="Select Output Directory",
            initialdir=self.output_dir_var.get() or str(OUTPUT_DIR)
        )
        
        if directory:
            self.output_dir_var.set(directory)
            
    def open_output_directory(self):
        """Open the output directory in file explorer."""
        output_dir = Path(self.output_dir_var.get())
        if output_dir.exists():
            import subprocess
            import sys
            
            try:
                if sys.platform == "win32":
                    subprocess.run(["explorer", str(output_dir)], check=True)
                elif sys.platform == "darwin":
                    subprocess.run(["open", str(output_dir)], check=True)
                else:
                    subprocess.run(["xdg-open", str(output_dir)], check=True)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open directory: {str(e)}")
        else:
            messagebox.showwarning("Warning", "Output directory does not exist.")
            
    def _on_input_file_changed(self, *args):
        """Handle input file variable changes."""
        file_path = self.input_file_var.get()
        
        if file_path and Path(file_path).exists():
            self.update_file_info(file_path)
        else:
            self.update_file_info(None)
            
        # Call callback if provided
        if self.on_file_selected:
            self.on_file_selected(file_path)
            
    def _on_output_dir_changed(self, *args):
        """Handle output directory variable changes."""
        output_dir = self.output_dir_var.get()
        self.open_output_btn.config(
            state=tk.NORMAL if output_dir and Path(output_dir).exists() else tk.DISABLED
        )
        
    def update_file_info(self, file_path: Optional[str]):
        """Update the file information display."""
        if not file_path or not Path(file_path).exists():
            # Hide file info
            self.info_frame.grid_remove()
            self.file_name_label.config(text="")
            self.file_size_label.config(text="")
            self.file_type_label.config(text="")
            self.sheets_label.config(text="")
            return
            
        try:
            file_path_obj = Path(file_path)
            
            # Show file info
            self.info_frame.grid()
            
            if file_path_obj.is_dir():
                # Handle folder
                self.info_frame.config(text="Folder Information")
                self.file_name_label.config(text=file_path_obj.name)
                
                # Count Excel files
                excel_files = self._get_excel_files_in_folder(file_path)
                self.file_size_label.config(text=f"{len(excel_files)} Excel files")
                
                # File type
                self.file_type_label.config(text="Folder (Batch Processing)")
                
                # Show file list
                if excel_files:
                    file_names = [f.name for f in excel_files[:5]]  # Show first 5
                    if len(excel_files) > 5:
                        file_names.append(f"... and {len(excel_files) - 5} more")
                    self.sheets_label.config(text=", ".join(file_names))
                else:
                    self.sheets_label.config(text="No Excel files found")
            else:
                # Handle single file
                self.info_frame.config(text="File Information")
                
                # File name
                self.file_name_label.config(text=file_path_obj.name)
                
                # File size
                file_size = file_path_obj.stat().st_size
                size_str = self.format_file_size(file_size)
                self.file_size_label.config(text=size_str)
                
                # File type
                file_type = file_path_obj.suffix.upper()
                if file_type == '.XLSX':
                    type_desc = "Excel Workbook"
                elif file_type == '.XLS':
                    type_desc = "Excel 97-2003"
                elif file_type == '.XLSM':
                    type_desc = "Excel Macro-Enabled"
                else:
                    type_desc = f"Unknown ({file_type})"
                    
                self.file_type_label.config(text=type_desc)
                
                # Sheet count (this would need to be provided by controller)
                # For now, just show a placeholder
                self.sheets_label.config(text="Loading...")
            
            # Update sheets info asynchronously if needed (only for single files)
            if not file_path_obj.is_dir():
                self.after_idle(lambda: self._update_sheet_info(file_path))
            
        except Exception as e:
            # Hide file info on error
            self.info_frame.grid_remove()
            
    def _update_sheet_info(self, file_path: str):
        """Update sheet information for the file."""
        try:
            # This would typically be provided by the controller
            # For now, just show a default
            self.sheets_label.config(text="1 (Sheet1)")
        except Exception:
            self.sheets_label.config(text="Unknown")
            
    def format_file_size(self, size_bytes: int) -> str:
        """Format file size in human-readable format."""
        if size_bytes == 0:
            return "0 B"
            
        size_names = ["B", "KB", "MB", "GB"]
        i = 0
        while size_bytes >= 1024 and i < len(size_names) - 1:
            size_bytes /= 1024.0
            i += 1
            
        return f"{size_bytes:.1f} {size_names[i]}"
        
    def get_input_file(self) -> str:
        """Get the selected input file path."""
        return self.input_file_var.get()
        
    def get_output_directory(self) -> str:
        """Get the selected output directory path."""
        return self.output_dir_var.get()
        
    def validate(self) -> tuple[bool, str]:
        """
        Validate the file selector inputs.
        
        Returns:
            Tuple of (is_valid, error_message)
        """
        input_file = self.input_file_var.get()
        output_dir = self.output_dir_var.get()
        
        if not input_file:
            return False, ERROR_MESSAGES["no_input_file"]
            
        if not Path(input_file).exists():
            return False, ERROR_MESSAGES["file_not_found"]
            
        if Path(input_file).suffix.lower() not in SUPPORTED_INPUT_FORMATS:
            return False, ERROR_MESSAGES["invalid_file_format"]
            
        if not output_dir:
            return False, ERROR_MESSAGES["no_output_dir"]
            
        return True, ""
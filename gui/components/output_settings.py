#!/usr/bin/env python3
"""
Output settings component for Excel Formatter application.

This module provides a GUI component for configuring output
formatting settings like headers, general settings, and void filtering.
"""

import tkinter as tk
from tkinter import ttk, colorchooser
from typing import Dict, Any

from config.settings import *


class OutputSettings(ttk.Frame):
    """Component for configuring output formatting settings."""
    
    def __init__(self, parent, on_change_callback=None):
        """
        Initialize the output settings component.
        
        Args:
            parent: Parent tkinter widget
            on_change_callback: Optional callback function to call when settings change
        """
        super().__init__(parent)
        
        # Initialize variables
        self.available_columns = []
        self.output_columns = []
        self.on_change_callback = on_change_callback
        self.init_variables()
        self.create_widgets()
        
    def _notify_change(self):
        """Notify parent that settings have changed."""
        if self.on_change_callback:
            self.on_change_callback()
        
    def init_variables(self):
        """Initialize tkinter variables for settings."""
        # Header formatting
        self.header_bold = tk.BooleanVar(value=True)
        self.header_bg_color = tk.StringVar(value="366092")
        self.header_font_color = tk.StringVar(value="FFFFFF")
        self.header_alignment = tk.StringVar(value="center")
        
        # General settings
        self.auto_fit_columns = tk.BooleanVar(value=True)
        self.freeze_header_row = tk.BooleanVar(value=True)
        self.selected_freeze_columns = []
        
        # Void filtering
        self.void_enabled = tk.BooleanVar(value=False)
        self.selected_void_columns = []
        
        # Set up change notifications for all variables
        self._setup_change_notifications()
        
    def _setup_change_notifications(self):
        """Set up change notifications for all tkinter variables."""
        # Header formatting variables
        self.header_bold.trace('w', lambda *args: self._notify_change())
        self.header_bg_color.trace('w', lambda *args: self._notify_change())
        self.header_font_color.trace('w', lambda *args: self._notify_change())
        self.header_alignment.trace('w', lambda *args: self._notify_change())
        
        # General settings variables
        self.auto_fit_columns.trace('w', lambda *args: self._notify_change())
        self.freeze_header_row.trace('w', lambda *args: self._notify_change())
        
        # Void filtering variables
        self.void_enabled.trace('w', lambda *args: self._notify_change())
        
    def create_widgets(self):
        """Create and arrange the settings widgets."""
        # Configure grid
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # Create notebook for different setting categories
        notebook = ttk.Notebook(self)
        notebook.grid(row=0, column=0, sticky="nsew", pady=10)
        
        # Header formatting tab
        self.create_header_tab(notebook)
        
        # General settings tab
        self.create_general_tab(notebook)
        
        # Void filtering tab
        self.create_void_tab(notebook)
        
    def create_header_tab(self, notebook):
        """Create header formatting settings tab."""
        header_frame = ttk.Frame(notebook)
        notebook.add(header_frame, text="Header Formatting")
        
        # Configure grid
        header_frame.grid_columnconfigure(1, weight=1)
        
        # Bold checkbox
        ttk.Checkbutton(
            header_frame,
            text="Bold headers",
            variable=self.header_bold
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=5)
        
        # Background color
        ttk.Label(header_frame, text="Background Color:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        
        bg_frame = ttk.Frame(header_frame)
        bg_frame.grid(row=1, column=1, sticky="w", padx=10, pady=5)
        
        bg_entry = ttk.Entry(bg_frame, textvariable=self.header_bg_color, width=10)
        bg_entry.pack(side=tk.LEFT)
        
        ttk.Button(
            bg_frame,
            text="Choose...",
            command=lambda: self.choose_color(self.header_bg_color),
            width=8
        ).pack(side=tk.LEFT, padx=(5, 0))
        
        # Font color
        ttk.Label(header_frame, text="Font Color:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        
        font_frame = ttk.Frame(header_frame)
        font_frame.grid(row=2, column=1, sticky="w", padx=10, pady=5)
        
        font_entry = ttk.Entry(font_frame, textvariable=self.header_font_color, width=10)
        font_entry.pack(side=tk.LEFT)
        
        ttk.Button(
            font_frame,
            text="Choose...",
            command=lambda: self.choose_color(self.header_font_color),
            width=8
        ).pack(side=tk.LEFT, padx=(5, 0))
        
        # Alignment
        ttk.Label(header_frame, text="Alignment:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        
        align_combo = ttk.Combobox(
            header_frame,
            textvariable=self.header_alignment,
            values=COLUMN_ALIGNMENTS,
            state="readonly",
            width=10
        )
        align_combo.grid(row=3, column=1, sticky="w", padx=10, pady=5)
        
    def create_general_tab(self, notebook):
        """Create general settings tab."""
        general_frame = ttk.Frame(notebook)
        notebook.add(general_frame, text="General Settings")
        
        # Configure grid
        general_frame.grid_columnconfigure(0, weight=1)
        general_frame.grid_rowconfigure(2, weight=1)
        
        # Auto-fit columns
        ttk.Checkbutton(
            general_frame,
            text="Auto-fit column widths",
            variable=self.auto_fit_columns
        ).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        
        # Freeze panes section
        freeze_label = ttk.Label(general_frame, text="Freeze Panes:", font=FONTS["heading"])
        freeze_label.grid(row=1, column=0, sticky="w", padx=10, pady=(15, 5))
        
        # Freeze header row
        ttk.Checkbutton(
            general_frame,
            text="Freeze header row (keep column names visible when scrolling)",
            variable=self.freeze_header_row
        ).grid(row=2, column=0, sticky="w", padx=20, pady=5)
        
        # Instructions
        instruction_text = (
            "Select OUTPUT columns to freeze (keep visible when scrolling horizontally).\\n"
            "Commonly used to keep key identifier columns like Name, ID, etc. always visible.\\n"
            "Note: These are your mapped output columns, not the original input columns."
        )
        
        instruction_label = ttk.Label(
            general_frame,
            text=instruction_text,
            foreground=COLORS["text_secondary"],
            justify=tk.LEFT
        )
        instruction_label.grid(row=3, column=0, sticky="w", padx=20, pady=5)
        
        # Column selection area
        columns_label = ttk.Label(general_frame, text="Output columns to freeze:")
        columns_label.grid(row=4, column=0, sticky="w", padx=20, pady=(10, 5))
        
        # Frame for column checkboxes with scrollbar
        freeze_checkbox_frame = ttk.Frame(general_frame)
        freeze_checkbox_frame.grid(row=5, column=0, sticky="nsew", padx=20, pady=5)
        freeze_checkbox_frame.grid_columnconfigure(0, weight=1)
        freeze_checkbox_frame.grid_rowconfigure(0, weight=1)
        
        # Canvas and scrollbar for checkboxes
        self.freeze_canvas = tk.Canvas(freeze_checkbox_frame, height=150)
        freeze_scrollbar = ttk.Scrollbar(freeze_checkbox_frame, orient="vertical", command=self.freeze_canvas.yview)
        self.freeze_scrollable_frame = ttk.Frame(self.freeze_canvas)
        
        self.freeze_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.freeze_canvas.configure(scrollregion=self.freeze_canvas.bbox("all"))
        )
        
        self.freeze_canvas.create_window((0, 0), window=self.freeze_scrollable_frame, anchor="nw")
        self.freeze_canvas.configure(yscrollcommand=freeze_scrollbar.set)
        
        self.freeze_canvas.grid(row=0, column=0, sticky="nsew")
        freeze_scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Initialize empty checkboxes container
        self.freeze_checkboxes = {}
        
    def create_void_tab(self, notebook):
        """Create void filtering settings tab."""
        void_frame = ttk.Frame(notebook)
        notebook.add(void_frame, text="Void Filtering")
        
        # Configure grid
        void_frame.grid_columnconfigure(0, weight=1)
        void_frame.grid_rowconfigure(3, weight=1)
        
        # Enable void filtering
        ttk.Checkbutton(
            void_frame,
            text="Enable void filtering (remove rows where specified columns are all zero)",
            variable=self.void_enabled
        ).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        
        # Instructions
        instruction_text = (
            "Select columns to check for zero values.\n"
            "Rows where ALL selected columns are zero will be removed."
        )
        
        instruction_label = ttk.Label(
            void_frame,
            text=instruction_text,
            foreground=COLORS["text_secondary"],
            justify=tk.LEFT
        )
        instruction_label.grid(row=1, column=0, sticky="w", padx=10, pady=5)
        
        # Column selection area
        columns_label = ttk.Label(void_frame, text="Columns to check:", font=FONTS["heading"])
        columns_label.grid(row=2, column=0, sticky="w", padx=10, pady=(10, 5))
        
        # Frame for column checkboxes with scrollbar
        checkbox_frame = ttk.Frame(void_frame)
        checkbox_frame.grid(row=3, column=0, sticky="nsew", padx=10, pady=5)
        checkbox_frame.grid_columnconfigure(0, weight=1)
        checkbox_frame.grid_rowconfigure(0, weight=1)
        
        # Canvas and scrollbar for checkboxes
        self.void_canvas = tk.Canvas(checkbox_frame, height=200)
        void_scrollbar = ttk.Scrollbar(checkbox_frame, orient="vertical", command=self.void_canvas.yview)
        self.void_scrollable_frame = ttk.Frame(self.void_canvas)
        
        self.void_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.void_canvas.configure(scrollregion=self.void_canvas.bbox("all"))
        )
        
        self.void_canvas.create_window((0, 0), window=self.void_scrollable_frame, anchor="nw")
        self.void_canvas.configure(yscrollcommand=void_scrollbar.set)
        
        self.void_canvas.grid(row=0, column=0, sticky="nsew")
        void_scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Initialize empty checkboxes container
        self.void_checkboxes = {}
        
        # Save button for void filter settings
        save_button_frame = ttk.Frame(void_frame)
        save_button_frame.grid(row=4, column=0, sticky="w", padx=10, pady=10)
        
        ttk.Button(
            save_button_frame,
            text="Save Void Filter Settings",
            command=self._save_void_settings,
            style="Accent.TButton"
        ).pack(side=tk.LEFT)
        
    def choose_color(self, color_var: tk.StringVar):
        """Open color chooser dialog."""
        current_color = color_var.get()
        
        # Convert hex to RGB for color chooser
        try:
            if len(current_color) == 6:
                rgb = tuple(int(current_color[i:i+2], 16) for i in (0, 2, 4))
            else:
                rgb = None
        except ValueError:
            rgb = None
            
        color = colorchooser.askcolor(color=rgb, title="Choose Color")
        
        if color[1]:  # If a color was selected
            # Remove # from hex color
            hex_color = color[1].lstrip('#').upper()
            color_var.set(hex_color)
            
    def set_available_columns(self, columns):
        """Set available input columns for void filtering."""
        self.available_columns = columns
        self._update_void_checkboxes()
        
    def set_output_columns(self, output_columns):
        """Set output columns for freeze panes selection."""
        self.output_columns = output_columns if output_columns else []
        self._update_freeze_columns()
        
        
    def _update_freeze_columns(self):
        """Update freeze panes column checkboxes."""
        # Clear existing checkboxes
        for widget in self.freeze_scrollable_frame.winfo_children():
            widget.destroy()
        self.freeze_checkboxes.clear()
        
        # Create checkboxes for each output column
        for i, column in enumerate(self.output_columns):
            var = tk.BooleanVar()
            checkbox = ttk.Checkbutton(
                self.freeze_scrollable_frame,
                text=column,
                variable=var,
                command=self._on_freeze_selection_changed
            )
            checkbox.grid(row=i, column=0, sticky="w", padx=5, pady=2)
            self.freeze_checkboxes[column] = var
                
    def _update_void_checkboxes(self):
        """Update void filtering column checkboxes."""
        # Clear existing checkboxes
        for widget in self.void_scrollable_frame.winfo_children():
            widget.destroy()
        self.void_checkboxes.clear()
        
        # Create checkboxes for each column
        for i, column in enumerate(self.available_columns):
            var = tk.BooleanVar()
            checkbox = ttk.Checkbutton(
                self.void_scrollable_frame,
                text=column,
                variable=var,
                command=self._on_void_selection_changed
            )
            checkbox.grid(row=i, column=0, sticky="w", padx=5, pady=2)
            self.void_checkboxes[column] = var
            
            # Restore selected state if this column was previously selected
            if column in self.selected_void_columns:
                var.set(True)
                
        # Apply pending void columns if they exist (from config loading)
        if hasattr(self, '_pending_void_columns') and self._pending_void_columns:
            for column, var in self.void_checkboxes.items():
                var.set(column in self._pending_void_columns)
            # Clear pending columns
            delattr(self, '_pending_void_columns')
            
    def _on_freeze_selection_changed(self):
        """Handle changes to freeze column selection."""
        self.selected_freeze_columns = [
            col for col, var in self.freeze_checkboxes.items() if var.get()
        ]
        self._notify_change()
        
    def _on_void_selection_changed(self):
        """Handle changes to void column selection."""
        self.selected_void_columns = [
            col for col, var in self.void_checkboxes.items() if var.get()
        ]
        self._notify_change()
        
    def _save_void_settings(self):
        """Save void filter settings and show confirmation."""
        # Update the selected columns
        self._on_void_selection_changed()
        
        # Show confirmation message
        if self.void_enabled.get() and self.selected_void_columns:
            message = f"Void filter enabled for columns: {', '.join(self.selected_void_columns)}"
        elif self.void_enabled.get():
            message = "Void filter enabled but no columns selected"
        else:
            message = "Void filter disabled"
            
        # Show confirmation (you might want to use a more subtle notification)
        import tkinter.messagebox as msgbox
        msgbox.showinfo("Void Filter Settings", f"Settings saved!\n\n{message}")
        
    def get_configuration(self) -> Dict[str, Any]:
        """Get current output settings configuration."""
        config = {
            "header_formatting": {
                "bold": self.header_bold.get(),
                "background_color": self.header_bg_color.get(),
                "font_color": self.header_font_color.get(),
                "alignment": self.header_alignment.get()
            },
            "general_settings": {
                "auto_fit_columns": self.auto_fit_columns.get()
            },
            "void": {
                "enabled": self.void_enabled.get(),
                "zero_columns": self.selected_void_columns.copy()
            }
        }
        
        # Add freeze panes if specified
        freeze_config = {}
        if self.freeze_header_row.get():
            freeze_config["freeze_header"] = True
            
        if self.selected_freeze_columns:
            freeze_config["freeze_columns"] = self.selected_freeze_columns.copy()
            
        if freeze_config:
            config["general_settings"]["freeze_panes"] = freeze_config
                
        return config
        
    def set_configuration(self, config: Dict[str, Any]):
        """Set configuration from config dict."""
        # Header formatting
        header_config = config.get("header_formatting", {})
        self.header_bold.set(header_config.get("bold", True))
        self.header_bg_color.set(header_config.get("background_color", "366092"))
        self.header_font_color.set(header_config.get("font_color", "FFFFFF"))
        self.header_alignment.set(header_config.get("alignment", "center"))
        
        # General settings
        general_config = config.get("general_settings", {})
        self.auto_fit_columns.set(general_config.get("auto_fit_columns", True))
        
        # Parse freeze panes
        freeze_panes = general_config.get("freeze_panes", {})
        if isinstance(freeze_panes, dict):
            # New format with checkboxes
            self.freeze_header_row.set(freeze_panes.get("freeze_header", True))
            freeze_columns = freeze_panes.get("freeze_columns", [])
            self.selected_freeze_columns = freeze_columns.copy()
        elif isinstance(freeze_panes, str) and freeze_panes:
            # Legacy format (e.g., "A2") - convert to new format
            self.freeze_header_row.set(True)  # Default to freezing header
            self.selected_freeze_columns = []  # No columns selected from legacy format
        else:
            # No freeze panes
            self.freeze_header_row.set(True)
            self.selected_freeze_columns = []
        
        # Void filtering
        void_config = config.get("void", {})
        self.void_enabled.set(void_config.get("enabled", False))
        void_columns = void_config.get("zero_columns", [])
        self.selected_void_columns = void_columns.copy()
        
        # Update checkbox states if checkboxes exist
        if hasattr(self, 'void_checkboxes') and self.void_checkboxes:
            for column, var in self.void_checkboxes.items():
                var.set(column in void_columns)
        else:
            # Store the void columns to be applied when checkboxes are created
            self._pending_void_columns = void_columns.copy()
            
        # Update freeze panes checkbox states
        for column, var in self.freeze_checkboxes.items():
            var.set(column in self.selected_freeze_columns)
            

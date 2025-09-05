#!/usr/bin/env python3
"""
Column mapper component for Excel Formatter application.

This module provides a GUI component for mapping input columns
to output columns with various transformation options.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, List, Any, Optional, Callable
import json

from config.settings import *


class ColumnMapper(ttk.Frame):
    """Component for mapping input columns to output columns."""
    
    def __init__(self, parent, on_mapping_changed: Optional[Callable[[Dict[str, Any]], None]] = None):
        """
        Initialize the column mapper component.
        
        Args:
            parent: Parent tkinter widget
            on_mapping_changed: Callback when mapping configuration changes
        """
        super().__init__(parent)
        
        self.on_mapping_changed = on_mapping_changed
        self.input_columns = []
        self.mapping_rows = []
        self.format_dialog_open = False  # Track if format dialog is open
        
        self.create_widgets()
        self.add_initial_row()
        
    def create_widgets(self):
        """Create and arrange the column mapper widgets."""
        # Configure grid
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # Header frame
        header_frame = ttk.Frame(self)
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header_frame.grid_columnconfigure(1, weight=1)
        
        title_label = ttk.Label(
            header_frame,
            text="Column Mapping",
            font=FONTS["heading"]
        )
        title_label.grid(row=0, column=0, sticky="w")
        
        # Control buttons
        button_frame = ttk.Frame(header_frame)
        button_frame.grid(row=0, column=1, sticky="e")
        
        self.add_row_btn = ttk.Button(
            button_frame,
            text="Add Column",
            command=self.add_mapping_row,
            width=12
        )
        self.add_row_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.remove_row_btn = ttk.Button(
            button_frame,
            text="Remove Last",
            command=self.remove_last_row,
            width=12,
            state=tk.DISABLED
        )
        self.remove_row_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.examples_btn = ttk.Button(
            button_frame,
            text="Examples",
            command=self.show_examples,
            width=12
        )
        self.examples_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.clear_all_btn = ttk.Button(
            button_frame,
            text="Clear All",
            command=self.clear_all_mappings,
            width=12
        )
        self.clear_all_btn.pack(side=tk.LEFT)
        
        # Scrollable frame for mapping rows
        self.create_scrollable_frame()
        
        # Instructions
        instruction_text = (
            "Column Mapping Instructions:\n"
            "• Select from dropdown OR type manually for advanced features:\n"
            "• Direct mapping: Select column from dropdown\n"
            "• Empty columns: Leave blank or select empty option\n"
            "• Formulas: Type '=Chk Amt - Gross - Fica' (use output column names)\n"
            "• Click '...' for advanced formatting options"
        )
        
        instruction_label = ttk.Label(
            self,
            text=instruction_text,
            font=FONTS["default"],
            foreground=COLORS["text_secondary"],
            justify=tk.LEFT
        )
        instruction_label.grid(row=2, column=0, sticky="w", pady=(10, 0))
        
    def create_scrollable_frame(self):
        """Create scrollable frame for mapping rows."""
        # Canvas and scrollbar setup
        canvas_frame = ttk.Frame(self)
        canvas_frame.grid(row=1, column=0, sticky="nsew")
        canvas_frame.grid_rowconfigure(0, weight=1)
        canvas_frame.grid_columnconfigure(0, weight=1)
        
        self.canvas = tk.Canvas(canvas_frame, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self._on_frame_configure()
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.grid(row=0, column=0, sticky="nsew")
        # Don't grid the scrollbar initially - it will be shown when needed
        
        # Bind canvas resize to update scrollbar visibility
        self.canvas.bind("<Configure>", lambda e: self._on_canvas_configure())
        
        # Bind mousewheel to canvas and frame
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.scrollable_frame.bind("<MouseWheel>", self._on_mousewheel)
        
        # Bind mouse enter/leave for scrolling focus
        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)
        
        # Configure scrollable frame grid
        self.scrollable_frame.grid_columnconfigure(1, weight=1)  # Source column dropdown
        self.scrollable_frame.grid_columnconfigure(3, weight=1)  # Output column entry
        
        # Create header row
        self.create_header_row()
        
    def create_header_row(self):
        """Create the header row for the mapping table."""
        headers = ["#", "Input Column", "→", "Output Column", "Align", "Width", "Format", "Actions"]
        
        for col, header in enumerate(headers):
            label = ttk.Label(
                self.scrollable_frame,
                text=header,
                font=FONTS["heading"],
                foreground=COLORS["primary"]
            )
            
            if col == 1 or col == 3:  # Input and Output columns
                label.grid(row=0, column=col, sticky="ew", padx=5, pady=5)
            else:
                label.grid(row=0, column=col, padx=5, pady=5)
                
        # Separator line
        separator = ttk.Separator(self.scrollable_frame, orient="horizontal")
        separator.grid(row=1, column=0, columnspan=7, sticky="ew", pady=5)
        
    def _on_mousewheel(self, event):
        """Handle mouse wheel scrolling."""
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
    def _bind_mousewheel(self, event):
        """Bind mousewheel when mouse enters the canvas."""
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
    def _unbind_mousewheel(self, event):
        """Unbind mousewheel when mouse leaves the canvas."""
        self.canvas.unbind_all("<MouseWheel>")
        
    def _on_frame_configure(self):
        """Handle frame resize and auto-hide/show scrollbar."""
        # Update scroll region
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self._update_scrollbar_visibility()
        
    def _on_canvas_configure(self):
        """Handle canvas resize and update scrollbar visibility."""
        self._update_scrollbar_visibility()
        
    def _update_scrollbar_visibility(self):
        """Update scrollbar visibility based on content size."""
        try:
            # Get dimensions
            canvas_height = self.canvas.winfo_height()
            content_height = self.scrollable_frame.winfo_reqheight()
            
            # Only check if both dimensions are valid
            if canvas_height > 1 and content_height > 1:
                if content_height > canvas_height:
                    # Content is taller than canvas, show scrollbar
                    self.scrollbar.grid(row=0, column=1, sticky="ns")
                else:
                    # Content fits in canvas, hide scrollbar
                    self.scrollbar.grid_remove()
        except tk.TclError:
            # Widget might not be ready yet, ignore
            pass
        
    def add_initial_row(self):
        """Add the first mapping row."""
        self.add_mapping_row()
        
    def add_mapping_row(self):
        """Add a new column mapping row."""
        row_index = len(self.mapping_rows)
        row_num = row_index + 2  # +2 because of header and separator
        
        # Row data
        row_data = {
            'index': row_index,
            'widgets': {}
        }
        
        # Row number
        row_label = ttk.Label(
            self.scrollable_frame,
            text=str(row_index + 1),
            foreground=COLORS["text_secondary"]
        )
        row_label.grid(row=row_num, column=0, padx=5, pady=2)
        row_data['widgets']['row_label'] = row_label
        
        # Input mapping frame - contains the mapping builder
        input_frame = ttk.Frame(self.scrollable_frame)
        input_frame.grid(row=row_num, column=1, sticky="ew", padx=5, pady=2)
        input_frame.grid_columnconfigure(1, weight=1)
        
        # Build Expression button (on the left)
        build_btn = ttk.Button(
            input_frame,
            text="Build...",
            command=lambda idx=row_index: self.show_expression_builder(idx),
            width=8
        )
        build_btn.grid(row=0, column=0, padx=(0, 5))
        
        # Simple display for the current mapping (on the right)
        input_var = tk.StringVar()
        input_display = ttk.Entry(
            input_frame,
            textvariable=input_var,
            state="readonly",
            width=25,
            font=("Consolas", 9)
        )
        input_display.grid(row=0, column=1, sticky="ew")
        
        row_data['widgets']['input_frame'] = input_frame
        row_data['widgets']['input_display'] = input_display
        row_data['widgets']['build_btn'] = build_btn
        row_data['input_var'] = input_var
        row_data['expression_parts'] = []  # Store the structured expression
        
        # Arrow
        arrow_label = ttk.Label(self.scrollable_frame, text="→")
        arrow_label.grid(row=row_num, column=2, padx=5, pady=2)
        row_data['widgets']['arrow_label'] = arrow_label
        
        # Output column name
        output_var = tk.StringVar(value=f"Column {row_index + 1}")
        output_entry = ttk.Entry(
            self.scrollable_frame,
            textvariable=output_var,
            width=20
        )
        output_entry.grid(row=row_num, column=3, sticky="ew", padx=5, pady=2)
        output_entry.bind("<KeyRelease>", lambda e, idx=row_index: self._on_mapping_changed(idx))
        row_data['widgets']['output_entry'] = output_entry
        row_data['output_var'] = output_var
        
        # Alignment dropdown
        align_var = tk.StringVar(value="left")
        align_combo = ttk.Combobox(
            self.scrollable_frame,
            textvariable=align_var,
            values=COLUMN_ALIGNMENTS,
            state="readonly",
            width=8
        )
        align_combo.grid(row=row_num, column=4, padx=5, pady=2)
        align_combo.bind("<<ComboboxSelected>>", lambda e, idx=row_index: self._on_mapping_changed(idx))
        row_data['widgets']['align_combo'] = align_combo
        row_data['align_var'] = align_var
        
        # Width entry
        width_var = tk.StringVar(value="15")
        width_entry = ttk.Entry(
            self.scrollable_frame,
            textvariable=width_var,
            width=6
        )
        width_entry.grid(row=row_num, column=5, padx=5, pady=2)
        width_entry.bind("<KeyRelease>", lambda e, idx=row_index: self._on_mapping_changed(idx))
        row_data['widgets']['width_entry'] = width_entry
        row_data['width_var'] = width_var
        
        # Format entry
        format_var = tk.StringVar(value="General")
        format_entry = ttk.Entry(
            self.scrollable_frame,
            textvariable=format_var,
            width=12
        )
        format_entry.grid(row=row_num, column=6, padx=5, pady=2)
        format_entry.bind("<KeyRelease>", lambda e, idx=row_index: self._on_mapping_changed(idx))
        format_entry.bind("<Button-1>", lambda e, idx=row_index: self.show_format_dialog(idx))
        row_data['widgets']['format_entry'] = format_entry
        row_data['format_var'] = format_var
        
        # Action buttons frame
        action_frame = ttk.Frame(self.scrollable_frame)
        action_frame.grid(row=row_num, column=7, padx=5, pady=2)
        
        # Advanced settings button
        advanced_btn = ttk.Button(
            action_frame,
            text="...",
            command=lambda idx=row_index: self.show_advanced_settings(idx),
            width=3
        )
        advanced_btn.pack(side=tk.LEFT, padx=(0, 2))
        row_data['widgets']['advanced_btn'] = advanced_btn
        
        # Delete button
        delete_btn = ttk.Button(
            action_frame,
            text="×",
            command=lambda idx=row_index: self.remove_mapping_row(idx),
            width=3
        )
        delete_btn.pack(side=tk.LEFT)
        row_data['widgets']['delete_btn'] = delete_btn
        row_data['widgets']['action_frame'] = action_frame
        
        # Advanced settings (initially empty)
        row_data['advanced_settings'] = {}
        
        self.mapping_rows.append(row_data)
        self.update_button_states()
        self.update_canvas_scroll()
        
        # Trigger callback
        self._on_mapping_changed(row_index)
        
    def remove_mapping_row(self, index: int):
        """Remove a specific mapping row."""
        if 0 <= index < len(self.mapping_rows):
            # Remove widgets
            row_data = self.mapping_rows[index]
            for widget in row_data['widgets'].values():
                if hasattr(widget, 'destroy'):
                    widget.destroy()
                    
            # Remove from list
            del self.mapping_rows[index]
            
            # Update indices and row numbers
            self.refresh_row_display()
            self.update_button_states()
            self.update_canvas_scroll()
            
            # Trigger callback
            if self.on_mapping_changed:
                self.on_mapping_changed(self.get_configuration())
                
    def remove_last_row(self):
        """Remove the last mapping row."""
        if self.mapping_rows:
            self.remove_mapping_row(len(self.mapping_rows) - 1)
            
    def clear_all_mappings(self):
        """Clear all mapping rows."""
        if messagebox.askyesno("Confirm", "Are you sure you want to clear all column mappings?"):
            # Remove all rows
            for row_data in self.mapping_rows:
                for widget in row_data['widgets'].values():
                    if hasattr(widget, 'destroy'):
                        widget.destroy()
                        
            self.mapping_rows.clear()
            self.update_button_states()
            self.update_canvas_scroll()
            
            # Add one empty row
            self.add_mapping_row()
            
    def refresh_row_display(self):
        """Refresh the display of all rows after deletion."""
        for i, row_data in enumerate(self.mapping_rows):
            row_data['index'] = i
            
            # Update row number
            row_data['widgets']['row_label'].config(text=str(i + 1))
            
            # Update grid positions
            row_num = i + 2  # +2 for header and separator
            for widget in row_data['widgets'].values():
                if hasattr(widget, 'grid_info') and widget.grid_info():
                    info = widget.grid_info()
                    widget.grid(row=row_num, column=info['column'], 
                              sticky=info.get('sticky', ''),
                              padx=info.get('padx', 0),
                              pady=info.get('pady', 0))
                              
    def show_advanced_settings(self, index: int):
        """Show advanced settings dialog for a column."""
        if 0 <= index < len(self.mapping_rows):
            AdvancedSettingsDialog(
                self,
                self.mapping_rows[index],
                self._on_advanced_settings_changed
            )
            
    def show_format_dialog(self, row_index: int):
        """Show format dialog for the specified row."""
        if 0 <= row_index < len(self.mapping_rows) and not self.format_dialog_open:
            self.format_dialog_open = True
            row_data = self.mapping_rows[row_index]
            FormatDialog(self, row_data, self._on_format_dialog_closed)
            
    def _on_format_dialog_closed(self, row_index: int):
        """Handle format dialog closure."""
        self.format_dialog_open = False
        self._on_mapping_changed(row_index)
            
    def show_expression_builder(self, index: int):
        """Show expression builder dialog for a column."""
        if 0 <= index < len(self.mapping_rows):
            # Get current output columns for formula building
            output_columns = [row['output_var'].get() for row in self.mapping_rows 
                            if row['output_var'].get().strip() and row != self.mapping_rows[index]]
            
            ExpressionBuilderDialog(
                self,
                self.mapping_rows[index],
                self.input_columns,
                self._on_expression_changed,
                output_columns
            )
            
    def _on_expression_changed(self, index: int):
        """Handle changes from expression builder."""
        self._on_mapping_changed(index)

    def show_examples(self):
        """Show examples of different mapping types."""
        examples_window = tk.Toplevel(self)
        examples_window.title("Column Mapping Examples")
        examples_window.geometry("600x400")
        examples_window.resizable(True, True)
        examples_window.transient(self)
        
        # Create main frame with scrollbar
        main_frame = ttk.Frame(examples_window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="Column Mapping Examples", font=("Segoe UI", 14, "bold"))
        title_label.pack(pady=(0, 15))
        
        # Create text widget with scrollbar
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        text_widget = tk.Text(text_frame, wrap=tk.WORD, font=("Consolas", 10), padx=10, pady=10)
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Examples content
        examples_text = """
COLUMN MAPPING EXAMPLES

1. DIRECT MAPPING
   Input Column: Name
   → Maps the "Name" column directly to output
   
2. EMPTY/BLANK COLUMNS  
   Input Column: (leave blank or empty)
   → Creates empty column in output
   
3. FORMULAS (using output column names)
   Input Column: =Chk Amt - Gross - Fica
   → Calculates: Check Amount minus Gross minus FICA
   
   Input Column: =Gross * 0.15
   → Calculates: Gross times 0.15 (15%)
   

5. PRACTICAL EXAMPLES:

   For Payroll Processing:
   • Employee Name → Name (direct mapping)
   • Check # → (blank for manual entry)  
   • Net Pay → Chk Amt (direct mapping)
   • Adjusted Gross → Gross (direct mapping)
   • Employee taxes - SS + Employee taxes - Med → Fica E/R
   • =Chk Amt - Gross - Fica → Liab (calculated liability)
   • Pay Date → Date (direct mapping)
   • Time Period → Period (direct mapping)

TIPS:
• For formulas (=), use OUTPUT column names as they appear in your mapping
• Leave input blank to create empty columns for manual data entry
• Use the "..." button for advanced formatting (colors, number formats, etc.)
• Preview your output before processing to verify mappings
"""
        
        text_widget.insert(1.0, examples_text.strip())
        text_widget.config(state=tk.DISABLED)
        
        # Close button
        ttk.Button(main_frame, text="Close", command=examples_window.destroy).pack(pady=(10, 0))
            
    def _on_advanced_settings_changed(self, index: int):
        """Handle changes in advanced settings."""
        self._on_mapping_changed(index)
        
    def update_button_states(self):
        """Update the state of control buttons."""
        has_rows = len(self.mapping_rows) > 0
        self.remove_row_btn.config(state=tk.NORMAL if has_rows else tk.DISABLED)
        
    def update_canvas_scroll(self):
        """Update canvas scroll region."""
        self.canvas.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
    def set_input_columns(self, columns: List[str]):
        """Set available input columns."""
        self.input_columns = columns
        
        # No need to update individual widgets since the new interface
        # uses the expression builder dialog which gets columns dynamically
            
    def clear_input_columns(self):
        """Clear available input columns."""
        self.set_input_columns([])
        
    def has_valid_mapping(self) -> bool:
        """Check if there is at least one valid column mapping."""
        for row_data in self.mapping_rows:
            output_name = row_data['output_var'].get().strip()
            if output_name:
                return True
        return False
        
    def _on_mapping_changed(self, index: int):
        """Handle mapping changes."""
        if self.on_mapping_changed:
            config = self.get_configuration()
            self.on_mapping_changed(config)
            
    def get_configuration(self) -> Dict[str, Any]:
        """Get current mapping configuration."""
        output_columns = []
        
        for row_data in self.mapping_rows:
            output_name = row_data['output_var'].get().strip()
            if not output_name:
                continue
                
            # Get the actual source column for processing (not display text)
            if 'expression_parts' in row_data and 'expression' in row_data['expression_parts']:
                # Use stored expression (empty for blank columns)
                source_column = row_data['expression_parts']['expression']
            else:
                # Fallback to display text (for backward compatibility)
                display_text = row_data['input_var'].get()
                source_column = "" if display_text == "(empty column)" else display_text
                
            alignment = row_data['align_var'].get()
            
            try:
                width = int(row_data['width_var'].get() or 15)
            except ValueError:
                width = 15
                
            # Get format value
            format_value = row_data['format_var'].get() or "General"
            
            # Update formatting with number_format
            formatting = row_data.get('advanced_settings', {}).copy()
            formatting['number_format'] = format_value
            
            column_config = {
                "name": output_name,
                "source_column": source_column,
                "alignment": alignment,
                "width": width,
                "formatting": formatting
            }
            
            output_columns.append(column_config)
            
        return {"output_columns": output_columns}
        
    def set_configuration(self, config: Dict[str, Any]):
        """Set mapping configuration from config dict."""
        # Clear existing mappings
        for row_data in self.mapping_rows:
            for widget in row_data['widgets'].values():
                if hasattr(widget, 'destroy'):
                    widget.destroy()
                    
        self.mapping_rows.clear()
        
        # Add rows from configuration
        output_columns = config.get("output_columns", [])
        
        if not output_columns:
            # Add one empty row if no configuration
            self.add_mapping_row()
            return
            
        for col_config in output_columns:
            self.add_mapping_row()
            row_data = self.mapping_rows[-1]
            
            # Set values
            source_column = col_config.get("source_column", "")
            # Handle empty columns - display "(empty column)" for empty source columns
            if not source_column or source_column == "(empty column)":
                row_data['input_var'].set("(empty column)")
            else:
                row_data['input_var'].set(source_column)
            row_data['output_var'].set(col_config.get("name", ""))
            row_data['align_var'].set(col_config.get("alignment", "left"))
            row_data['width_var'].set(str(col_config.get("width", 15)))
            
            # Handle format - get from formatting.number_format or default to General
            formatting = col_config.get("formatting", {})
            format_value = formatting.get("number_format", "General")
            row_data['format_var'].set(format_value)
            
            row_data['advanced_settings'] = formatting
            
        self.update_button_states()
        self.update_canvas_scroll()


class FormatDialog:
    """Dialog for selecting Excel-style number formats."""
    
    def __init__(self, parent, row_data: Dict[str, Any], callback: Callable[[int], None]):
        """
        Initialize format dialog.
        
        Args:
            parent: Parent widget
            row_data: Row data containing format information
            callback: Callback function when format changes
        """
        self.parent = parent
        self.row_data = row_data
        self.callback = callback
        
        # Create dialog window
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Number Format")
        self.dialog.geometry("700x600")
        self.dialog.resizable(True, True)
        
        # Center the dialog
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Handle window close
        self.dialog.protocol("WM_DELETE_WINDOW", self.on_cancel)
        
        # Create main frame
        main_frame = ttk.Frame(self.dialog, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Current format display
        current_frame = ttk.LabelFrame(main_frame, text="Current Format", padding=10)
        current_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.current_format_var = tk.StringVar(value=row_data['format_var'].get())
        current_entry = ttk.Entry(current_frame, textvariable=self.current_format_var, 
                                 font=("Consolas", 10), width=50)
        current_entry.pack(fill=tk.X)
        
        # Format categories
        categories_frame = ttk.LabelFrame(main_frame, text="Format Categories", padding=10)
        categories_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Create notebook for categories
        self.notebook = ttk.Notebook(categories_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Predefined formats
        self.create_predefined_formats()
        
        # Custom format
        self.create_custom_format()
        
        # Preview section
        preview_frame = ttk.LabelFrame(main_frame, text="Preview", padding=10)
        preview_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.preview_var = tk.StringVar(value="1234.5678")
        preview_entry = ttk.Entry(preview_frame, textvariable=self.preview_var, 
                                 font=("Consolas", 10), width=50, state="readonly")
        preview_entry.pack(fill=tk.X)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="OK", command=self.apply_format).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Cancel", command=self.on_cancel).pack(side=tk.RIGHT)
        
        # Bind events
        self.current_format_var.trace('w', self.update_preview)
        
    def create_predefined_formats(self):
        """Create predefined format categories."""
        # General
        general_frame = ttk.Frame(self.notebook)
        self.notebook.add(general_frame, text="General")
        
        general_formats = [
            ("General", "General"),
            ("Number", "#,##0"),
            ("Number (2 decimals)", "#,##0.00"),
            ("Currency", '"$"#,##0.00'),
            ("Percentage", "0%"),
            ("Percentage (2 decimals)", "0.00%"),
            ("Date", "mm/dd/yyyy"),
            ("Time", "h:mm AM/PM"),
            ("Text", "@")
        ]
        
        self.create_format_buttons(general_frame, general_formats)
        
        # Number formats
        number_frame = ttk.Frame(self.notebook)
        self.notebook.add(number_frame, text="Number")
        
        number_formats = [
            ("0", "0"),
            ("0.0", "0.0"),
            ("0.00", "0.00"),
            ("#,##0", "#,##0"),
            ("#,##0.0", "#,##0.0"),
            ("#,##0.00", "#,##0.00"),
            ("0.0%", "0.0%"),
            ("0.00%", "0.00%"),
            ("0.000%", "0.000%")
        ]
        
        self.create_format_buttons(number_frame, number_formats)
        
        # Currency formats
        currency_frame = ttk.Frame(self.notebook)
        self.notebook.add(currency_frame, text="Currency")
        
        currency_formats = [
            ('$0', '"$"0'),
            ('$0.00', '"$"0.00'),
            ('$#,##0', '"$"#,##0'),
            ('$#,##0.00', '"$"#,##0.00'),
            ('$#,##0.00_);($#,##0.00)', '"$"#,##0.00_);("$"#,##0.00)'),
            ('$0.00_);($0.00)', '"$"0.00_);("$"0.00)'),
            ('(#,##0.00)', '(#,##0.00) - Accounting Format'),
            ('-(#,##0.00)', '-(#,##0.00) - Negative Outside Parentheses'),
            ('(#,##0)', '(#,##0) - Accounting Format (no decimals)')
        ]
        
        self.create_format_buttons(currency_frame, currency_formats)
        
        # Date formats
        date_frame = ttk.Frame(self.notebook)
        self.notebook.add(date_frame, text="Date")
        
        date_formats = [
            ("mm/dd/yyyy", "mm/dd/yyyy"),
            ("m/d/yy", "m/d/yy"),
            ("mm-dd-yyyy", "mm-dd-yyyy"),
            ("d-mmm-yy", "d-mmm-yy"),
            ("d-mmm", "d-mmm"),
            ("mmm-yy", "mmm-yy"),
            ("h:mm AM/PM", "h:mm AM/PM"),
            ("h:mm:ss AM/PM", "h:mm:ss AM/PM"),
            ("h:mm", "h:mm"),
            ("h:mm:ss", "h:mm:ss")
        ]
        
        self.create_format_buttons(date_frame, date_formats)
        
    def create_custom_format(self):
        """Create custom format tab."""
        custom_frame = ttk.Frame(self.notebook)
        self.notebook.add(custom_frame, text="Custom")
        
        # Instructions
        instructions = ttk.Label(custom_frame, 
            text="Enter a custom Excel format code:\n\n"
                 "Examples:\n"
                 "• #,##0.00 - Number with thousands separator and 2 decimals\n"
                 "• \"$\"#,##0.00 - Currency with dollar sign\n"
                 "• 0.00% - Percentage with 2 decimals\n"
                 "• mm/dd/yyyy - Date format\n"
                 "• @ - Text format\n\n"
                 "For more help, see Excel's Format Cells dialog.",
            justify=tk.LEFT, font=("Segoe UI", 9))
        instructions.pack(anchor=tk.W, pady=(0, 10))
        
        # Custom format entry
        custom_frame_inner = ttk.Frame(custom_frame)
        custom_frame_inner.pack(fill=tk.X, pady=10)
        
        ttk.Label(custom_frame_inner, text="Custom Format:").pack(anchor=tk.W)
        self.custom_format_var = tk.StringVar()
        custom_entry = ttk.Entry(custom_frame_inner, textvariable=self.custom_format_var,
                                font=("Consolas", 10), width=50)
        custom_entry.pack(fill=tk.X, pady=(5, 0))
        custom_entry.bind('<KeyRelease>', lambda e: self.current_format_var.set(self.custom_format_var.get()))
        
    def create_format_buttons(self, parent, formats):
        """Create format selection buttons."""
        # Create scrollable frame
        canvas = tk.Canvas(parent, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Create buttons
        for i, (display_name, format_code) in enumerate(formats):
            btn = ttk.Button(scrollable_frame, text=display_name,
                           command=lambda code=format_code: self.select_format(code))
            btn.pack(fill=tk.X, padx=5, pady=2)
            
    def select_format(self, format_code):
        """Select a format code."""
        self.current_format_var.set(format_code)
        
    def update_preview(self, *args):
        """Update the preview display."""
        try:
            # This is a simplified preview - in a real implementation,
            # you'd want to use a library that can actually format numbers
            # according to Excel format codes
            format_code = self.current_format_var.get()
            if format_code:
                # Basic preview logic
                if "0.00" in format_code:
                    self.preview_var.set("1,234.56")
                elif "0.0" in format_code:
                    self.preview_var.set("1,234.5")
                elif "0" in format_code and "#" in format_code:
                    self.preview_var.set("1,234")
                elif "%" in format_code:
                    self.preview_var.set("12.34%")
                elif "$" in format_code:
                    self.preview_var.set("$1,234.56")
                elif "mm" in format_code or "dd" in format_code:
                    self.preview_var.set("12/25/2023")
                else:
                    self.preview_var.set("1234.5678")
            else:
                self.preview_var.set("1234.5678")
        except:
            self.preview_var.set("1234.5678")
            
    def apply_format(self):
        """Apply the selected format."""
        format_code = self.current_format_var.get()
        self.row_data['format_var'].set(format_code)
        self.callback(self.row_data['index'])
        self.dialog.destroy()
        
    def on_cancel(self):
        """Handle cancel or close."""
        self.dialog.destroy()


class AdvancedSettingsDialog:
    """Dialog for advanced column formatting settings."""
    
    def __init__(self, parent, row_data: Dict[str, Any], callback: Callable[[int], None]):
        """
        Initialize advanced settings dialog.
        
        Args:
            parent: Parent widget
            row_data: Row data dictionary
            callback: Callback function when settings change
        """
        self.parent = parent
        self.row_data = row_data
        self.callback = callback
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Advanced Settings - {row_data['output_var'].get()}")
        self.dialog.geometry("400x300")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center dialog on parent
        self.center_dialog()
        
        # Current settings
        self.settings = row_data.get('advanced_settings', {}).copy()
        
        self.create_widgets()
        
    def center_dialog(self):
        """Center dialog on parent window."""
        self.dialog.update_idletasks()
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = self.parent.winfo_rootx() + (self.parent.winfo_width() // 2) - (width // 2)
        y = self.parent.winfo_rooty() + (self.parent.winfo_height() // 2) - (height // 2)
        self.dialog.geometry(f"{width}x{height}+{x}+{y}")
        
    def create_widgets(self):
        """Create dialog widgets."""
        main_frame = ttk.Frame(self.dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Number format
        ttk.Label(main_frame, text="Number Format:").grid(row=0, column=0, sticky="w", pady=5)
        
        self.number_format_var = tk.StringVar(value=self.settings.get("number_format", ""))
        number_combo = ttk.Combobox(
            main_frame,
            textvariable=self.number_format_var,
            values=list(NUMBER_FORMATS.values()),
            width=30
        )
        number_combo.grid(row=0, column=1, sticky="w", padx=(10, 0), pady=5)
        
        # Bold checkbox
        self.bold_var = tk.BooleanVar(value=self.settings.get("bold", False))
        bold_check = ttk.Checkbutton(main_frame, text="Bold", variable=self.bold_var)
        bold_check.grid(row=1, column=0, columnspan=2, sticky="w", pady=5)
        
        # Background color
        ttk.Label(main_frame, text="Background Color:").grid(row=2, column=0, sticky="w", pady=5)
        
        self.bg_color_var = tk.StringVar(value=self.settings.get("background_color", ""))
        bg_color_entry = ttk.Entry(main_frame, textvariable=self.bg_color_var, width=10)
        bg_color_entry.grid(row=2, column=1, sticky="w", padx=(10, 0), pady=5)
        
        # Font color
        ttk.Label(main_frame, text="Font Color:").grid(row=3, column=0, sticky="w", pady=5)
        
        self.font_color_var = tk.StringVar(value=self.settings.get("font_color", ""))
        font_color_entry = ttk.Entry(main_frame, textvariable=self.font_color_var, width=10)
        font_color_entry.grid(row=3, column=1, sticky="w", padx=(10, 0), pady=5)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=(20, 0))
        
        ttk.Button(button_frame, text="OK", command=self.save_settings).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Cancel", command=self.dialog.destroy).pack(side=tk.LEFT)
        
    def save_settings(self):
        """Save advanced settings."""
        # Update settings
        self.settings["number_format"] = self.number_format_var.get()
        self.settings["bold"] = self.bold_var.get()
        self.settings["background_color"] = self.bg_color_var.get()
        self.settings["font_color"] = self.font_color_var.get()
        
        # Remove empty values
        self.settings = {k: v for k, v in self.settings.items() if v}
        
        # Update row data
        self.row_data['advanced_settings'] = self.settings
        
        # Call callback
        self.callback(self.row_data['index'])
        
        self.dialog.destroy()
        
    def _update_dialog_scrollbar(self, canvas, scrollbar):
        """Update dialog scrollbar visibility."""
        try:
            # Update scroll region
            canvas.configure(scrollregion=canvas.bbox("all"))
            
            # Get dimensions
            canvas_height = canvas.winfo_height()
            content_height = self.scrollable_main_frame.winfo_reqheight()
            
            # Only check if both dimensions are valid
            if canvas_height > 1 and content_height > 1:
                if content_height > canvas_height:
                    # Content is taller than canvas, show scrollbar
                    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                else:
                    # Content fits in canvas, hide scrollbar
                    scrollbar.pack_forget()
        except tk.TclError:
            # Widget might not be ready yet, ignore
            pass


class ExpressionBuilderDialog:
    """Dialog for building column mapping expressions step-by-step."""
    
    def __init__(self, parent, row_data: Dict[str, Any], input_columns: List[str], callback: Callable[[int], None], output_columns: List[str] = None):
        """
        Initialize expression builder dialog.
        
        Args:
            parent: Parent widget
            row_data: Row data dictionary
            input_columns: List of available input columns
            callback: Callback function when expression changes
            output_columns: List of available output columns (optional)
        """
        self.parent = parent
        self.row_data = row_data
        self.input_columns = input_columns
        self.output_columns = output_columns or []
        self.callback = callback
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Build Expression - {row_data['output_var'].get()}")
        self.dialog.geometry("700x500")
        self.dialog.resizable(True, True)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center dialog on parent
        self.center_dialog()
        
        # Current expression parts
        self.expression_parts = row_data.get('expression_parts', []).copy()
        
        self.create_widgets()
        self.update_expression_display()
        
    def center_dialog(self):
        """Center dialog on parent window."""
        self.dialog.update_idletasks()
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = self.parent.winfo_rootx() + (self.parent.winfo_width() // 2) - (width // 2)
        y = self.parent.winfo_rooty() + (self.parent.winfo_height() // 2) - (height // 2)
        self.dialog.geometry(f"{width}x{height}+{x}+{y}")
        
    def create_widgets(self):
        """Create dialog widgets."""
        # Create scrollable main frame
        canvas = tk.Canvas(self.dialog, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.dialog, orient="vertical", command=canvas.yview)
        self.scrollable_main_frame = ttk.Frame(canvas, padding=10)
        
        self.scrollable_main_frame.bind(
            "<Configure>",
            lambda e: self._update_dialog_scrollbar(canvas, scrollbar)
        )
        
        canvas.bind("<Configure>", lambda e: self._update_dialog_scrollbar(canvas, scrollbar))
        
        canvas.create_window((0, 0), window=self.scrollable_main_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        # Don't pack scrollbar initially - will be shown when needed
        
        # Bind mousewheel
        canvas.bind("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))
        self.scrollable_main_frame.bind("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))
        
        main_frame = self.scrollable_main_frame
        
        # Title and mode selection
        title_label = ttk.Label(main_frame, text="Expression Builder", font=("Segoe UI", 14, "bold"))
        title_label.pack(pady=(0, 15))
        
        # Mode selection
        mode_frame = ttk.LabelFrame(main_frame, text="Mapping Type", padding=10)
        mode_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.mode_var = tk.StringVar(value="direct")
        
        ttk.Radiobutton(mode_frame, text="Direct Column Mapping", variable=self.mode_var, 
                       value="direct", command=self.on_mode_change).pack(anchor=tk.W)
        ttk.Radiobutton(mode_frame, text="Empty/Blank Column", variable=self.mode_var, 
                       value="blank", command=self.on_mode_change).pack(anchor=tk.W)
        ttk.Radiobutton(mode_frame, text="Formula Expression", variable=self.mode_var, 
                       value="formula", command=self.on_mode_change).pack(anchor=tk.W)
        
        # Content frame that changes based on mode
        self.content_frame = ttk.Frame(main_frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Expression preview
        preview_frame = ttk.LabelFrame(main_frame, text="Expression Preview", padding=10)
        preview_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.expression_display = tk.Text(preview_frame, height=3, font=("Consolas", 10), 
                                        wrap=tk.WORD, state=tk.DISABLED)
        self.expression_display.pack(fill=tk.X)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="OK", command=self.save_expression).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Cancel", command=self.dialog.destroy).pack(side=tk.RIGHT)
        ttk.Button(button_frame, text="Clear", command=self.clear_expression).pack(side=tk.LEFT)
        
        # Initialize content based on current mode
        self.detect_current_mode()
        self.on_mode_change()
        
    def detect_current_mode(self):
        """Detect current mode from existing expression."""
        current_value = self.row_data['input_var'].get()
        
        if not current_value:
            self.mode_var.set("direct")  # Default to direct mapping instead of blank
        elif current_value.startswith("="):
            self.mode_var.set("formula")
        else:
            self.mode_var.set("direct")
            
    def on_mode_change(self):
        """Handle mode selection change."""
        # Clear content frame
        for widget in self.content_frame.winfo_children():
            widget.destroy()
            
        mode = self.mode_var.get()
        
        if mode == "direct":
            self.create_direct_mode()
        elif mode == "blank":
            self.create_blank_mode()
        elif mode == "formula":
            self.create_formula_mode()
            
    def create_direct_mode(self):
        """Create interface for direct column mapping."""
        ttk.Label(self.content_frame, text="Select input column to map directly:").pack(anchor=tk.W, pady=(0, 10))
        
        # Calculate optimal width for input columns
        if self.input_columns:
            max_width = min(max(len(col) for col in self.input_columns) + 5, 60)  # +5 for padding, max 60
            optimal_width = max(max_width, 20)  # minimum 20 chars
        else:
            optimal_width = 30
        
        self.direct_var = tk.StringVar()
        direct_combo = ttk.Combobox(self.content_frame, textvariable=self.direct_var, 
                                   values=self.input_columns, state="readonly", width=optimal_width)
        direct_combo.pack(anchor=tk.W, pady=(0, 10))
        direct_combo.bind("<<ComboboxSelected>>", lambda e: self.update_expression_display())
        
        # Set current value if it's a direct mapping
        current = self.row_data['input_var'].get()
        if current in self.input_columns:
            self.direct_var.set(current)
            
    def create_blank_mode(self):
        """Create interface for blank columns."""
        ttk.Label(self.content_frame, 
                 text="This will create an empty column in the output.\nUseful for manual data entry fields like Check #.",
                 justify=tk.LEFT).pack(anchor=tk.W, pady=10)
        self.update_expression_display()
        
    def create_formula_mode(self):
        """Create interface for formula expressions."""
        ttk.Label(self.content_frame, text="Build Formula Expression:").pack(anchor=tk.W, pady=(0, 10))
        
        # Formula builder frame
        formula_frame = ttk.Frame(self.content_frame)
        formula_frame.pack(fill=tk.BOTH, expand=True)
        
        # Instructions
        # Column source selection
        source_frame = ttk.LabelFrame(formula_frame, text="Column Source", padding=5)
        source_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.column_source = tk.StringVar(value="input")
        
        # Create radio buttons in a horizontal layout
        radio_frame = ttk.Frame(source_frame)
        radio_frame.pack(fill=tk.X)
        
        ttk.Radiobutton(radio_frame, text="Use Input Columns (original file columns)", 
                       variable=self.column_source, value="input",
                       command=self.update_column_source).pack(anchor=tk.W)
        
        if self.output_columns:
            ttk.Radiobutton(radio_frame, text="Use Output Columns (your mapped columns)", 
                           variable=self.column_source, value="output",
                           command=self.update_column_source).pack(anchor=tk.W)
        
        # Info label that updates based on selection
        self.column_info_label = ttk.Label(source_frame, font=("Segoe UI", 9, "italic"))
        self.column_info_label.pack(anchor=tk.W, pady=(5, 0))
        
        self.update_column_source()  # Initialize the info label
        
        # Formula builder type selection
        builder_frame = ttk.LabelFrame(formula_frame, text="Formula Builder", padding=10)
        builder_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.formula_type = tk.StringVar(value="visual")
        ttk.Radiobutton(builder_frame, text="Visual Builder (Recommended)", 
                       variable=self.formula_type, value="visual",
                       command=self.update_formula_interface).pack(anchor=tk.W)
        ttk.Radiobutton(builder_frame, text="Manual Entry (Advanced)", 
                       variable=self.formula_type, value="manual",
                       command=self.update_formula_interface).pack(anchor=tk.W)
        
        # Content frame that changes based on formula type
        self.formula_content_frame = ttk.Frame(formula_frame)
        self.formula_content_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # Set current value if it's a formula
        current = self.row_data['input_var'].get()
        if current.startswith("="):
            self.current_formula = current[1:]  # Remove =
        else:
            self.current_formula = ""
            
        self.update_formula_interface()
        
    def update_column_source(self):
        """Update the column source information and refresh interface."""
        if self.column_source.get() == "output":
            info_text = f"Using {len(self.output_columns)} output columns. Great for referencing already-mapped columns!"
        else:
            info_text = f"Using {len(self.input_columns)} input columns from your source file."
        
        self.column_info_label.config(text=info_text)
        
        # Refresh formula interface to use new column source
        if hasattr(self, 'formula_content_frame') and self.formula_content_frame.winfo_exists():
            self.update_formula_interface()
    
    def get_current_columns(self):
        """Get the current column list based on selected source."""
        if self.column_source.get() == "output" and self.output_columns:
            return self.output_columns
        return self.input_columns
        
    def update_formula_interface(self):
        """Update formula interface based on selected type."""
        # Clear content frame safely
        try:
            for widget in self.formula_content_frame.winfo_children():
                widget.destroy()
        except tk.TclError:
            # Widget already destroyed, ignore
            pass
            
        if self.formula_type.get() == "visual":
            self.create_visual_formula_builder()
        else:
            self.create_manual_formula_entry()
            
    def create_visual_formula_builder(self):
        """Create visual formula builder with dropdowns."""
        # Visual builder
        visual_frame = ttk.Frame(self.formula_content_frame)
        visual_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(visual_frame, text="Build your formula step by step:").pack(anchor=tk.W, pady=(0, 5))
        
        # Formula parts container with scrolling
        self.formula_parts = []
        
        # Create scrollable frame for formula parts
        parts_canvas_frame = ttk.Frame(visual_frame)
        parts_canvas_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        parts_canvas_frame.grid_rowconfigure(0, weight=1)
        parts_canvas_frame.grid_columnconfigure(0, weight=1)
        
        parts_canvas = tk.Canvas(parts_canvas_frame, height=200, highlightthickness=0)
        parts_scrollbar = ttk.Scrollbar(parts_canvas_frame, orient="vertical", command=parts_canvas.yview)
        self.formula_parts_frame = ttk.Frame(parts_canvas)
        
        self.formula_parts_frame.bind(
            "<Configure>",
            lambda e: parts_canvas.configure(scrollregion=parts_canvas.bbox("all"))
        )
        
        parts_canvas.create_window((0, 0), window=self.formula_parts_frame, anchor="nw")
        parts_canvas.configure(yscrollcommand=parts_scrollbar.set)
        
        parts_canvas.grid(row=0, column=0, sticky="nsew")
        parts_scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Bind mousewheel for formula parts
        parts_canvas.bind("<MouseWheel>", lambda e: parts_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))
        self.formula_parts_frame.bind("<MouseWheel>", lambda e: parts_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))
        
        # Add first part
        self.add_formula_part()
        
        # Add part button
        ttk.Button(visual_frame, text="Add More", 
                  command=self.add_formula_part).pack(anchor=tk.W, pady=(5, 0))
        
        # Parse existing formula if it exists
        if self.current_formula:
            self.parse_existing_formula()
            
    def add_formula_part(self, initial_column="", initial_operator="+", initial_value=""):
        """Add a new part to the visual formula builder."""
        # Check if the frame still exists
        if not hasattr(self, 'formula_parts_frame') or not self.formula_parts_frame.winfo_exists():
            return
            
        part_frame = ttk.Frame(self.formula_parts_frame)
        part_frame.pack(fill=tk.X, pady=2)
        
        part_data = {}
        
        # If not the first part, add operator
        if self.formula_parts:
            op_var = tk.StringVar(value=initial_operator if initial_operator else "+")
            op_combo = ttk.Combobox(part_frame, textvariable=op_var, 
                                   values=["+", "-", "*", "/"], width=3, state="readonly")
            op_combo.pack(side=tk.LEFT, padx=(0, 5))
            op_combo.bind("<<ComboboxSelected>>", lambda e: self.update_expression_display())
            part_data['operator'] = op_var
            part_data['op_combo'] = op_combo
        
        # Column or value selection
        type_var = tk.StringVar(value="column" if initial_column else "column")
        ttk.Radiobutton(part_frame, text="Column:", variable=type_var, value="column",
                       command=lambda: self.toggle_part_type(part_data)).pack(side=tk.LEFT)
        
        # Column dropdown - dynamically sized and using current column source
        col_var = tk.StringVar(value=initial_column)
        current_columns = self.get_current_columns()
        
        # Calculate optimal width based on longest column name, with reasonable limits
        if current_columns:
            max_width = min(max(len(col) for col in current_columns) + 2, 40)  # +2 for padding, max 40
            optimal_width = max(max_width, 15)  # minimum 15 chars
        else:
            optimal_width = 20
        
        col_combo = ttk.Combobox(part_frame, textvariable=col_var, 
                                values=current_columns, width=optimal_width)
        col_combo.pack(side=tk.LEFT, padx=(5, 10))
        col_combo.bind("<<ComboboxSelected>>", lambda e: self.update_expression_display())
        col_combo.bind("<KeyRelease>", lambda e: self.update_expression_display())
        
        ttk.Radiobutton(part_frame, text="Number:", variable=type_var, value="number",
                       command=lambda: self.toggle_part_type(part_data)).pack(side=tk.LEFT)
        
        # Number entry
        num_var = tk.StringVar(value=initial_value if initial_value and not initial_column else "")
        num_entry = ttk.Entry(part_frame, textvariable=num_var, width=10)
        num_entry.pack(side=tk.LEFT, padx=(5, 10))
        num_entry.bind("<KeyRelease>", lambda e: self.update_expression_display())
        
        # Remove button (not for first part)
        if len(self.formula_parts) > 0:
            ttk.Button(part_frame, text="×", width=3,
                      command=lambda: self.remove_formula_part(part_frame, part_data)).pack(side=tk.LEFT)
        
        part_data.update({
            'frame': part_frame,
            'type_var': type_var,
            'col_var': col_var,
            'col_combo': col_combo,
            'num_var': num_var,
            'num_entry': num_entry
        })
        
        self.formula_parts.append(part_data)
        self.toggle_part_type(part_data)  # Set initial state
        self.update_expression_display()
        
    def toggle_part_type(self, part_data):
        """Toggle between column and number input."""
        part_type = part_data['type_var'].get()
        
        if part_type == "column":
            part_data['col_combo'].config(state="normal")
            part_data['num_entry'].config(state="disabled")
        else:
            part_data['col_combo'].config(state="disabled")
            part_data['num_entry'].config(state="normal")
            
        self.update_expression_display()
        
    def remove_formula_part(self, frame, part_data):
        """Remove a formula part."""
        frame.destroy()
        self.formula_parts = [p for p in self.formula_parts if p != part_data]
        self.update_expression_display()
        
    def parse_existing_formula(self):
        """Parse existing formula into visual parts."""
        # This is a simple parser - could be enhanced
        # For now, just put the formula in manual mode
        self.formula_type.set("manual")
        self.update_formula_interface()
        
    def create_manual_formula_entry(self):
        """Create manual formula entry interface."""
        manual_frame = ttk.Frame(self.formula_content_frame)
        manual_frame.pack(fill=tk.BOTH, expand=True)
        
        # Instructions
        ttk.Label(manual_frame, 
                 text="Enter formula using INPUT column names (put column names in quotes if they contain spaces)",
                 font=("Segoe UI", 9, "italic")).pack(anchor=tk.W, pady=(0, 10))
        
        # Formula entry
        ttk.Label(manual_frame, text="Formula (without =):").pack(anchor=tk.W)
        self.manual_formula_var = tk.StringVar(value=self.current_formula)
        formula_entry = ttk.Entry(manual_frame, textvariable=self.manual_formula_var, width=60)
        formula_entry.pack(fill=tk.X, pady=(5, 10))
        formula_entry.bind("<KeyRelease>", lambda e: self.update_expression_display())
        
        # Quick reference
        ref_frame = ttk.LabelFrame(manual_frame, text="Quick Reference", padding=5)
        ref_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Label(ref_frame, text="Available Input Columns:").pack(anchor=tk.W)
        
        # Create a scrollable list of columns
        columns_text = tk.Text(ref_frame, height=4, wrap=tk.WORD, font=("Consolas", 9))
        columns_scrollbar = ttk.Scrollbar(ref_frame, orient="vertical", command=columns_text.yview)
        columns_text.configure(yscrollcommand=columns_scrollbar.set)
        
        # Add columns to text widget
        columns_str = ", ".join([f'"{col}"' if ' ' in col else col for col in self.input_columns])
        columns_text.insert(1.0, columns_str)
        columns_text.config(state=tk.DISABLED)
        
        columns_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        columns_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Examples
        ttk.Label(ref_frame, text="\nExamples:").pack(anchor=tk.W, pady=(10, 0))
        examples = [
            '"Net Pay" - "Adjusted Gross" - "Employee Taxes"',
            '"Hours" * "Rate"',
            '"Gross Pay" * 0.15'
        ]
        
        for example in examples:
            ttk.Label(ref_frame, text=f"  • {example}", font=("Consolas", 9)).pack(anchor=tk.W)
            
        
    def set_formula(self, formula):
        """Set a predefined formula."""
        if hasattr(self, 'formula_var'):
            self.formula_var.set(formula)
            self.update_expression_display()
            
    def update_expression_display(self):
        """Update the expression preview display."""
        # Check if the dialog still exists
        if not hasattr(self, 'mode_var') or not self.mode_var:
            return
            
        mode = self.mode_var.get()
        expression = ""
        
        if mode == "direct":
            if hasattr(self, 'direct_var'):
                expression = self.direct_var.get()
        elif mode == "blank":
            expression = "(empty column)"
        elif mode == "formula":
            if hasattr(self, 'formula_type') and self.formula_type.get() == "visual":
                # Build expression from visual parts
                if hasattr(self, 'formula_parts'):
                    formula_parts = []
                    for i, part in enumerate(self.formula_parts):
                        part_str = ""
                        
                        # Add operator (except for first part)
                        if i > 0 and 'operator' in part:
                            part_str += f" {part['operator'].get()} "
                        
                        # Add value (column or number)
                        part_type = part['type_var'].get()
                        if part_type == "column":
                            col_name = part['col_var'].get()
                            if col_name:
                                # Quote column names with spaces
                                if ' ' in col_name:
                                    part_str += f'"{col_name}"'
                                else:
                                    part_str += col_name
                        else:  # number
                            num_val = part['num_var'].get()
                            if num_val:
                                part_str += num_val
                        
                        if part_str.strip():
                            formula_parts.append(part_str)
                    
                    formula = "".join(formula_parts)
                    expression = f"={formula}" if formula else ""
            else:
                # Manual formula entry
                if hasattr(self, 'manual_formula_var'):
                    formula = self.manual_formula_var.get()
                    expression = f"={formula}" if formula else ""
            
        # Update display
        self.expression_display.config(state=tk.NORMAL)
        self.expression_display.delete(1.0, tk.END)
        self.expression_display.insert(1.0, expression)
        self.expression_display.config(state=tk.DISABLED)
        
    def clear_expression(self):
        """Clear the current expression."""
        self.mode_var.set("blank")
        self.on_mode_change()
        
    def save_expression(self):
        """Save the built expression."""
        mode = self.mode_var.get()
        expression = ""
        
        if mode == "direct":
            if hasattr(self, 'direct_var'):
                expression = self.direct_var.get()
        elif mode == "blank":
            expression = "(empty column)"
        elif mode == "formula":
            if hasattr(self, 'formula_type') and self.formula_type.get() == "visual":
                # Build expression from visual parts
                if hasattr(self, 'formula_parts'):
                    formula_parts = []
                    for i, part in enumerate(self.formula_parts):
                        part_str = ""
                        
                        # Add operator (except for first part)
                        if i > 0 and 'operator' in part:
                            part_str += f" {part['operator'].get()} "
                        
                        # Add value (column or number)
                        part_type = part['type_var'].get()
                        if part_type == "column":
                            col_name = part['col_var'].get().strip()
                            if col_name:
                                # Quote column names with spaces
                                if ' ' in col_name:
                                    part_str += f'"{col_name}"'
                                else:
                                    part_str += col_name
                        else:  # number
                            num_val = part['num_var'].get().strip()
                            if num_val:
                                part_str += num_val
                        
                        if part_str.strip():
                            formula_parts.append(part_str)
                    
                    formula = "".join(formula_parts).strip()
                    expression = f"={formula}" if formula else ""
            else:
                # Manual formula entry
                if hasattr(self, 'manual_formula_var'):
                    formula = self.manual_formula_var.get().strip()
                    expression = f"={formula}" if formula else ""
            
        # Update the row data
        # For blank mode, store empty string but display "(empty column)"
        stored_expression = "" if mode == "blank" else expression
        display_expression = "(empty column)" if mode == "blank" else expression
        
        self.row_data['input_var'].set(display_expression)
        self.row_data['expression_parts'] = {
            'mode': mode,
            'expression': stored_expression,
            'display_expression': display_expression
        }
        
        # Call callback
        self.callback(self.row_data['index'])
        
        self.dialog.destroy()
        
    def _update_dialog_scrollbar(self, canvas, scrollbar):
        """Update dialog scrollbar visibility."""
        try:
            # Update scroll region
            canvas.configure(scrollregion=canvas.bbox("all"))
            
            # Get dimensions
            canvas_height = canvas.winfo_height()
            content_height = self.scrollable_main_frame.winfo_reqheight()
            
            # Only check if both dimensions are valid
            if canvas_height > 1 and content_height > 1:
                if content_height > canvas_height:
                    # Content is taller than canvas, show scrollbar
                    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                else:
                    # Content fits in canvas, hide scrollbar
                    scrollbar.pack_forget()
        except tk.TclError:
            # Widget might not be ready yet, ignore
            pass
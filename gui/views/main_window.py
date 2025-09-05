#!/usr/bin/env python3
"""
Main window for Excel Formatter application.

This module contains the main GUI window that orchestrates all the
different components and provides the primary user interface.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
from pathlib import Path
from typing import Dict, List, Any, Optional

from config.settings import *
from gui.components.file_selector import FileSelector
from gui.components.column_mapper import ColumnMapper
from gui.components.output_settings import OutputSettings
from gui.components.progress_dialog import ProgressDialog


class MainWindow(ttk.Frame):
    """Main application window containing all GUI components."""
    
    def __init__(self, parent, controller):
        """
        Initialize the main window.
        
        Args:
            parent: Parent tkinter widget
            controller: Main controller instance
        """
        super().__init__(parent)
        self.parent = parent
        self.controller = controller
        
        # Initialize variables
        self.input_file_path = tk.StringVar()
        self.output_directory = tk.StringVar(value=str(OUTPUT_DIR))
        self.processing_thread = None
        self.current_tab = 0  # Track current tab (0=File Selection, 1=Column Mapping, 2=Output Settings)
        self.mapping_has_changes = False  # Track if mappings have been changed from defaults
        self.is_initializing = True  # Track if we're in initialization phase
        
        # Set up the GUI
        self.setup_window()
        self.create_widgets()
        self.bind_events()
        
    def setup_window(self):
        """Configure the main window properties."""
        self.parent.title(WINDOW_TITLE)
        self.parent.geometry(WINDOW_SIZE)
        self.parent.minsize(*WINDOW_MIN_SIZE)
        
        # Configure grid weights
        self.parent.grid_rowconfigure(0, weight=1)
        self.parent.grid_columnconfigure(0, weight=1)
        
        # Configure main frame
        self.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.grid_rowconfigure(1, weight=1)  # Column mapper gets most space
        self.grid_columnconfigure(0, weight=1)
        
    def create_widgets(self):
        """Create and arrange all GUI widgets."""
        # Title Label
        title_label = ttk.Label(
            self,
            text=APP_NAME,
            font=FONTS["heading"]
        )
        title_label.grid(row=0, column=0, pady=(0, 20), sticky="w")
        
        # Create notebook for tabbed interface
        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        
        # File Selection Tab
        self.file_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.file_frame, text="File Selection")
        self.create_file_selection_tab()
        
        # Column Mapping Tab  
        self.mapping_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.mapping_frame, text="Column Mapping")
        self.create_column_mapping_tab()
        
        # Output Settings Tab
        self.output_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.output_frame, text="Output Settings")
        self.create_output_settings_tab()
        
        # Action Buttons Frame
        self.create_action_buttons()
        
        # Initially disable tabs except file selection (after all widgets are created)
        self.after_idle(self._initialize_tab_states)
        # Mark initialization as complete
        self.is_initializing = False
        
    def _initialize_tab_states(self):
        """Initialize tab states after all widgets are created."""
        try:
            # Disable tabs that require input file
            self.notebook.tab(1, state="disabled")  # Column Mapping tab
            self.notebook.tab(2, state="disabled")  # Output Settings tab
        except Exception as e:
            # If there's still an issue, just log it and continue
            print(f"Warning: Could not initialize tab states: {e}")
        
    def create_file_selection_tab(self):
        """Create the file selection tab."""
        # Configure grid
        self.file_frame.grid_rowconfigure(1, weight=1)
        self.file_frame.grid_columnconfigure(0, weight=1)
        
        # File selector component
        self.file_selector = FileSelector(
            self.file_frame,
            input_file_var=self.input_file_path,
            output_dir_var=self.output_directory,
            on_file_selected=self.on_input_file_selected
        )
        self.file_selector.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        # File preview area
        preview_label = ttk.Label(self.file_frame, text="File Preview:", font=FONTS["heading"])
        preview_label.grid(row=1, column=0, sticky="w", padx=10, pady=(10, 5))
        
        # Preview text widget
        preview_frame = ttk.Frame(self.file_frame)
        preview_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)
        
        self.preview_text = tk.Text(
            preview_frame,
            height=15,
            wrap=tk.NONE,
            font=FONTS["monospace"],
            state=tk.DISABLED
        )
        
        # Scrollbars for preview
        v_scroll = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_text.yview)
        h_scroll = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.preview_text.xview)
        self.preview_text.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        self.preview_text.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        
    def create_column_mapping_tab(self):
        """Create the column mapping tab."""
        # Configure grid
        self.mapping_frame.grid_rowconfigure(0, weight=1)
        self.mapping_frame.grid_columnconfigure(0, weight=1)
        
        # Column mapper component
        self.column_mapper = ColumnMapper(
            self.mapping_frame,
            on_mapping_changed=self.on_mapping_changed
        )
        self.column_mapper.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
    def create_output_settings_tab(self):
        """Create the output settings tab."""
        # Configure grid
        self.output_frame.grid_rowconfigure(0, weight=1)
        self.output_frame.grid_columnconfigure(0, weight=1)
        
        # Output settings component
        self.output_settings = OutputSettings(self.output_frame)
        self.output_settings.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
    def create_action_buttons(self):
        """Create the action buttons at the bottom."""
        button_frame = ttk.Frame(self)
        button_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        
        # Configure button frame
        for i in range(5):
            button_frame.grid_columnconfigure(i, weight=1)
            
        # Load Config Button
        self.load_config_btn = ttk.Button(
            button_frame,
            text="Load Config",
            command=self.load_configuration
        )
        self.load_config_btn.grid(row=0, column=0, padx=(0, 5), sticky="ew")
        
        # Save Config Button
        self.save_config_btn = ttk.Button(
            button_frame,
            text="Save Config", 
            command=self.save_configuration
        )
        self.save_config_btn.grid(row=0, column=1, padx=5, sticky="ew")
        
        # Preview Button
        self.preview_btn = ttk.Button(
            button_frame,
            text="Preview Output",
            command=self.preview_output,
            state=tk.DISABLED
        )
        self.preview_btn.grid(row=0, column=2, padx=5, sticky="ew")
        
        # Process Button
        self.process_btn = ttk.Button(
            button_frame,
            text="Process File",
            command=self.process_file,
            state=tk.DISABLED
        )
        self.process_btn.grid(row=0, column=3, padx=5, sticky="ew")
        
        # Exit Button
        self.exit_btn = ttk.Button(
            button_frame,
            text="Exit",
            command=self.parent.quit
        )
        self.exit_btn.grid(row=0, column=4, padx=(5, 0), sticky="ew")
        
    def bind_events(self):
        """Bind events and keyboard shortcuts."""
        self.parent.bind('<Control-o>', lambda e: self.file_selector.browse_input_file())
        self.parent.bind('<Control-s>', lambda e: self.save_configuration())
        self.parent.bind('<Control-Return>', lambda e: self.process_file())
        self.parent.bind('<F5>', lambda e: self.preview_output())
        
        # Bind tab change event
        self.notebook.bind('<<NotebookTabChanged>>', self.on_tab_changed)
        
    def on_tab_changed(self, event):
        """Handle tab change events."""
        try:
            # Get the currently selected tab index
            self.current_tab = self.notebook.index(self.notebook.select())
            self.update_button_states()
        except Exception as e:
            # If there's an error getting the tab, just continue
            pass
        
    def on_input_file_selected(self, file_path: str):
        """Handle input file selection."""
        try:
            if not file_path:
                self.clear_preview()
                if hasattr(self, 'column_mapper'):
                    self.column_mapper.clear_input_columns()
                if hasattr(self, 'output_settings'):
                    self.output_settings.set_available_columns([])
                # Reset mapping changes flag when clearing input
                self.mapping_has_changes = False
                self.update_button_states()
                return
                
            # Load file preview
            preview_data = self.controller.get_file_preview(file_path)
            self.update_preview(preview_data)
            
            # Handle both files and folders
            if Path(file_path).is_dir():
                # For folders, get columns from the first Excel file
                folder_path = Path(file_path)
                excel_files = []
                for ext in SUPPORTED_INPUT_FORMATS:
                    excel_files.extend(folder_path.glob(f"*{ext}"))
                
                if excel_files:
                    # Use first file to determine column structure
                    first_file = str(excel_files[0])
                    columns = self.controller.get_file_columns(first_file)
                    
                    if hasattr(self, 'column_mapper'):
                        self.column_mapper.set_input_columns(columns)
                    
                    if hasattr(self, 'output_settings'):
                        self.output_settings.set_available_columns(columns)
                else:
                    # No Excel files in folder
                    if hasattr(self, 'column_mapper'):
                        self.column_mapper.clear_input_columns()
                    if hasattr(self, 'output_settings'):
                        self.output_settings.set_available_columns([])
            else:
                # Single file
                columns = self.controller.get_file_columns(file_path)
                if hasattr(self, 'column_mapper'):
                    self.column_mapper.set_input_columns(columns)
                
                if hasattr(self, 'output_settings'):
                    self.output_settings.set_available_columns(columns)
            
            # Update button states
            self.update_button_states()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {str(e)}")
            
    def on_mapping_changed(self, mapping_config: Dict[str, Any]):
        """Handle column mapping changes."""
        # Mark that mappings have been changed from defaults (but not during initialization)
        if not self.is_initializing:
            self.mapping_has_changes = True
        
        # Update output settings with current output column names
        if hasattr(self, 'output_settings'):
            output_columns = [col.get("name", "") for col in mapping_config.get("output_columns", []) if col.get("name", "")]
            self.output_settings.set_output_columns(output_columns)
        
        self.update_button_states()
        
    def update_preview(self, preview_data: str):
        """Update the file preview display."""
        self.preview_text.config(state=tk.NORMAL)
        self.preview_text.delete(1.0, tk.END)
        self.preview_text.insert(1.0, preview_data)
        self.preview_text.config(state=tk.DISABLED)
        
    def clear_preview(self):
        """Clear the file preview display."""
        self.preview_text.config(state=tk.NORMAL)
        self.preview_text.delete(1.0, tk.END)
        self.preview_text.config(state=tk.DISABLED)
        
    def update_button_states(self):
        """Update the state of action buttons and tabs based on current state."""
        has_input_file = bool(self.input_file_path.get())
        has_mapping = self.column_mapper.has_valid_mapping() if hasattr(self, 'column_mapper') else False
        is_on_mapping_tab = (self.current_tab == 1)  # Column Mapping tab
        
        # Enable/disable tabs based on input file (with safety checks)
        try:
            if has_input_file:
                self.notebook.tab(1, state="normal")   # Column Mapping tab
                self.notebook.tab(2, state="normal")   # Output Settings tab
            else:
                self.notebook.tab(1, state="disabled") # Column Mapping tab
                self.notebook.tab(2, state="disabled") # Output Settings tab
                # Switch back to file selection tab if no input file
                self.notebook.select(0)
        except Exception as e:
            # If tabs aren't ready yet, just continue
            pass
        
        # Load config button - only enabled if we have input file and are on mapping tab
        if hasattr(self, 'load_config_btn'):
            self.load_config_btn.config(state=tk.NORMAL if has_input_file and is_on_mapping_tab else tk.DISABLED)
        
        # Save config button - only enabled if we have input file, are on mapping tab, and have changes
        if hasattr(self, 'save_config_btn'):
            self.save_config_btn.config(state=tk.NORMAL if has_input_file and is_on_mapping_tab and self.mapping_has_changes else tk.DISABLED)
        
        # Preview button - enabled if we have input file and mapping and are on mapping tab
        if hasattr(self, 'preview_btn'):
            self.preview_btn.config(state=tk.NORMAL if has_input_file and has_mapping and is_on_mapping_tab else tk.DISABLED)
        
        # Process button - enabled if we have input file and mapping and are on mapping tab
        if hasattr(self, 'process_btn'):
            self.process_btn.config(state=tk.NORMAL if has_input_file and has_mapping and is_on_mapping_tab else tk.DISABLED)
        
    def load_configuration(self):
        """Load mapping configuration from file."""
        try:
            file_path = filedialog.askopenfilename(
                title="Load Configuration",
                filetypes=FILE_DIALOG_OPTIONS["config_filetypes"],
                defaultextension=".json"
            )
            
            if file_path:
                config = self.controller.load_configuration(file_path)
                
                # Update components with loaded configuration
                if hasattr(self, 'column_mapper'):
                    self.column_mapper.set_configuration(config)
                if hasattr(self, 'output_settings'):
                    self.output_settings.set_configuration(config)
                    
                    # Update output columns for freeze panes
                    output_columns = [col.get("name", "") for col in config.get("output_columns", []) if col.get("name", "")]
                    self.output_settings.set_output_columns(output_columns)
                
                # Reset mapping changes flag since we loaded a config
                self.mapping_has_changes = False
                self.update_button_states()
                
        except Exception as e:
            messagebox.showerror("Error", f"{ERROR_MESSAGES['config_load_error']}: {str(e)}")
            
    def save_configuration(self):
        """Save current mapping configuration to file."""
        try:
            file_path = filedialog.asksaveasfilename(
                title="Save Configuration",
                filetypes=FILE_DIALOG_OPTIONS["config_filetypes"],
                defaultextension=".json"
            )
            
            if file_path:
                # Get configuration from components
                config = self.get_current_configuration()
                
                self.controller.save_configuration(config, file_path)
                messagebox.showinfo("Success", SUCCESS_MESSAGES["config_saved"])
                # Reset mapping changes flag since we saved the config
                self.mapping_has_changes = False
                self.update_button_states()
                
        except Exception as e:
            messagebox.showerror("Error", f"{ERROR_MESSAGES['config_save_error']}: {str(e)}")
            
    def get_current_configuration(self) -> Dict[str, Any]:
        """Get current configuration from all components."""
        config = {}
        
        # Get mapping configuration from column mapper
        if hasattr(self, 'column_mapper'):
            config.update(self.column_mapper.get_configuration())
            
        # Get output settings configuration
        if hasattr(self, 'output_settings'):
            config.update(self.output_settings.get_configuration())
            
        return config
        
    def preview_output(self):
        """Show preview of processed output."""
        if not self.validate_inputs():
            return
            
        try:
            input_file = self.input_file_path.get()
            config = self.get_current_configuration()
            
            preview_data = self.controller.preview_output(input_file, config)
            
            # Show preview dialog
            self.show_preview_dialog(preview_data)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate preview: {str(e)}")
            
    def show_preview_dialog(self, preview_data: str):
        """Show preview data in a dialog window."""
        preview_window = tk.Toplevel(self.parent)
        preview_window.title("Output Preview")
        preview_window.geometry("800x600")
        
        # Create text widget with scrollbars
        frame = ttk.Frame(preview_window)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        text_widget = tk.Text(
            frame,
            wrap=tk.NONE,
            font=FONTS["monospace"]
        )
        
        v_scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=text_widget.yview)
        h_scroll = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=text_widget.xview)
        text_widget.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        
        text_widget.insert(1.0, preview_data)
        text_widget.config(state=tk.DISABLED)
        
        # Close button
        ttk.Button(
            preview_window,
            text="Close",
            command=preview_window.destroy
        ).pack(pady=10)
        
    def process_file(self):
        """Process the input file with current configuration."""
        if not self.validate_inputs():
            return
            
        # Disable buttons during processing
        self.set_processing_state(True)
        
        # Start processing in separate thread
        self.processing_thread = threading.Thread(
            target=self._process_file_thread,
            daemon=True
        )
        self.processing_thread.start()
        
    def _process_file_thread(self):
        """Process file(s) in separate thread."""
        try:
            input_path = self.input_file_path.get()
            output_dir = self.output_directory.get()
            config = self.get_current_configuration()
            
            # Show progress dialog
            self.after(0, self.show_progress_dialog)
            
            # Check if processing folder or single file
            if Path(input_path).is_dir():
                # Batch processing
                output_files = self.controller.process_folder(input_path, output_dir, config)
                self.after(0, lambda: self.on_batch_processing_complete(output_files))
            else:
                # Single file processing
                output_file = self.controller.process_file(input_path, output_dir, config)
                self.after(0, lambda: self.on_processing_complete(output_file))
            
        except Exception as e:
            # Show error message
            error_msg = str(e)
            self.after(0, lambda: self.on_processing_error(error_msg))
            
    def show_progress_dialog(self):
        """Show progress dialog during processing."""
        self.progress_dialog = ProgressDialog(
            self.parent,
            title="Processing File",
            message="Processing Excel file, please wait..."
        )
        
    def on_processing_complete(self, output_file: str):
        """Handle successful processing completion."""
        if hasattr(self, 'progress_dialog'):
            self.progress_dialog.destroy()
            
        self.set_processing_state(False)
        
        messagebox.showinfo(
            "Success", 
            f"{SUCCESS_MESSAGES['file_processed']}\n\nOutput file: {Path(output_file).name}"
        )
        
    def on_batch_processing_complete(self, output_files: list):
        """Handle successful batch processing completion."""
        if hasattr(self, 'progress_dialog'):
            self.progress_dialog.destroy()
            
        self.set_processing_state(False)
        
        if output_files:
            file_count = len(output_files)
            file_list = "\\n".join([Path(f).name for f in output_files[:10]])  # Show first 10
            if len(output_files) > 10:
                file_list += f"\\n... and {len(output_files) - 10} more files"
                
            messagebox.showinfo(
                "Batch Processing Complete", 
                f"Successfully processed {file_count} file(s)!\\n\\nOutput files:\\n{file_list}"
            )
        else:
            messagebox.showwarning(
                "No Files Processed",
                "No files were processed. Please check that the selected folder contains valid Excel files with the required columns."
            )
        
    def on_processing_error(self, error_message: str):
        """Handle processing error."""
        if hasattr(self, 'progress_dialog'):
            self.progress_dialog.destroy()
            
        self.set_processing_state(False)
        
        messagebox.showerror(
            "Error",
            f"{ERROR_MESSAGES['processing_error']}\n\n{error_message}"
        )
        
    def set_processing_state(self, processing: bool):
        """Enable/disable UI during processing."""
        state = tk.DISABLED if processing else tk.NORMAL
        
        if hasattr(self, 'process_btn'):
            self.process_btn.config(state=state)
        if hasattr(self, 'preview_btn'):
            self.preview_btn.config(state=state) 
        if hasattr(self, 'load_config_btn'):
            self.load_config_btn.config(state=state)
        if hasattr(self, 'save_config_btn'):
            self.save_config_btn.config(state=state)
        
        if not processing:
            self.update_button_states()
            
    def validate_inputs(self) -> bool:
        """Validate all inputs before processing."""
        if not self.input_file_path.get():
            messagebox.showerror("Error", ERROR_MESSAGES["no_input_file"])
            return False
            
        if not Path(self.input_file_path.get()).exists():
            messagebox.showerror("Error", ERROR_MESSAGES["file_not_found"])
            return False
            
        if not self.output_directory.get():
            messagebox.showerror("Error", ERROR_MESSAGES["no_output_dir"])
            return False
            
        if not self.column_mapper.has_valid_mapping():
            messagebox.showerror("Error", ERROR_MESSAGES["no_columns_mapped"])
            return False
            
        return True
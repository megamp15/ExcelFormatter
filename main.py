#!/usr/bin/env python3
"""
Excel Formatter Application - Main Entry Point

A professional GUI application for processing and formatting Excel files
with configurable column mappings and advanced formatting options.

Usage:
    python main.py [--gui] [--help]
    
    --gui     Run in GUI mode (default)
    --help    Show help information
"""

import tkinter as tk
from tkinter import ttk, messagebox
import argparse
import logging
import sys
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from config.settings import *
from gui.views.main_window import MainWindow
from gui.controllers.main_controller import MainController


class ExcelFormatterApp:
    """Main application class."""
    
    def __init__(self):
        """Initialize the application."""
        self.root = None
        self.controller = None
        self.main_window = None
        
        # Setup logging
        self.setup_logging()
        self.logger = logging.getLogger(__name__)
        
    def setup_logging(self):
        """Configure application logging."""
        # Ensure log directory exists
        LOG_DIR.mkdir(exist_ok=True)
        
        # Configure logging
        logging.basicConfig(
            level=getattr(logging, LOG_LEVEL),
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(LOG_FILE, encoding='utf-8'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        
        # Set log file rotation if needed
        if LOG_FILE.exists() and LOG_FILE.stat().st_size > MAX_LOG_SIZE:
            # Simple log rotation - keep last backup
            backup_file = LOG_FILE.with_suffix('.log.bak')
            if backup_file.exists():
                backup_file.unlink()
            LOG_FILE.rename(backup_file)
            
    def create_gui(self):
        """Create and setup the GUI."""
        try:
            self.logger.info("Starting Excel Formatter GUI application")
            
            # Create root window
            self.root = tk.Tk()
            self.root.withdraw()  # Hide until fully loaded
            
            # Configure root window
            self.setup_root_window()
            
            # Create controller
            self.controller = MainController(self.root)
            
            # Create main window
            self.main_window = MainWindow(self.root, self.controller)
            
            # Bind cleanup events
            self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
            
            # Show window
            self.root.deiconify()
            self.center_window()
            
            self.logger.info("GUI application initialized successfully")
            
        except Exception as e:
            error_msg = f"Failed to initialize GUI: {str(e)}"
            self.logger.error(error_msg)
            messagebox.showerror("Initialization Error", error_msg)
            sys.exit(1)
            
    def setup_root_window(self):
        """Configure the root window properties."""
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        
        # Set window icon if available
        try:
            icon_path = project_root / "icon.ico"
            if icon_path.exists():
                self.root.iconbitmap(str(icon_path))
        except Exception:
            pass  # Icon is optional
            
        # Configure style
        self.setup_styles()
        
    def setup_styles(self):
        """Configure ttk styles for consistent appearance."""
        try:
            style = ttk.Style()
            
            # Configure notebook style
            style.configure(
                "TNotebook.Tab",
                padding=[20, 10],
                font=FONTS["default"]
            )
            
            # Configure button style
            style.configure(
                "TButton",
                font=FONTS["button"],
                padding=[10, 5]
            )
            
        except Exception as e:
            self.logger.warning(f"Could not configure styles: {str(e)}")
            
    def center_window(self):
        """Center the window on screen."""
        try:
            self.root.update_idletasks()
            
            # Get window dimensions
            width = self.root.winfo_width()
            height = self.root.winfo_height()
            
            # Get screen dimensions
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            
            # Calculate center position
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2
            
            # Set window position
            self.root.geometry(f"{width}x{height}+{x}+{y}")
            
        except Exception as e:
            self.logger.warning(f"Could not center window: {str(e)}")
            
    def run_gui(self):
        """Start the GUI event loop."""
        try:
            self.create_gui()
            
            # Start main loop
            self.logger.info("Starting GUI main loop")
            self.root.mainloop()
            
        except KeyboardInterrupt:
            self.logger.info("Application interrupted by user")
            self.cleanup()
            
        except Exception as e:
            error_msg = f"Unexpected error in GUI: {str(e)}"
            self.logger.error(error_msg)
            messagebox.showerror("Application Error", error_msg)
            
        finally:
            self.cleanup()
            
    def on_closing(self):
        """Handle application closing."""
        try:
            # Ask for confirmation if needed
            if messagebox.askokcancel("Exit", "Do you want to exit Excel Formatter?"):
                self.cleanup()
                self.root.destroy()
                
        except Exception as e:
            self.logger.error(f"Error during shutdown: {str(e)}")
            self.root.destroy()
            
    def cleanup(self):
        """Perform cleanup operations."""
        try:
            self.logger.info("Cleaning up application resources")
            
            if self.controller:
                self.controller.shutdown()
                
        except Exception as e:
            self.logger.error(f"Error during cleanup: {str(e)}")


def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description="Excel Formatter - Professional Excel file processing tool",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"""
Examples:
    python main.py              # Run in GUI mode
    python main.py --gui         # Run in GUI mode (explicit)
    
Version: {APP_VERSION}
"""
    )
    
    parser.add_argument(
        "--gui",
        action="store_true",
        default=True,
        help="Run in GUI mode (default)"
    )
    
    parser.add_argument(
        "--version",
        action="version",
        version=f"{APP_NAME} v{APP_VERSION}"
    )
    
    return parser.parse_args()


def show_startup_info():
    """Show startup information."""
    print(f"{APP_NAME} v{APP_VERSION}")
    print(f"{APP_DESCRIPTION}")
    print(f"Starting application...")
    print("-" * 50)


def main():
    """Main application entry point."""
    try:
        # Parse arguments
        args = parse_arguments()
        
        # Show startup info
        show_startup_info()
        
        # Create and run application
        app = ExcelFormatterApp()
        
        if args.gui:
            app.run_gui()
        else:
            # Future: Could add command-line processing mode
            print("Only GUI mode is currently supported.")
            sys.exit(1)
            
    except KeyboardInterrupt:
        print("\nApplication interrupted by user.")
        sys.exit(0)
        
    except Exception as e:
        print(f"Fatal error: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
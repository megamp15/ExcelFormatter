#!/usr/bin/env python3
"""
Main controller for Excel Formatter application.

This module contains the main controller that handles business logic,
coordinates between GUI and core processing, and manages application state.
"""

import json
import logging
from pathlib import Path
from typing import Dict, List, Any, Optional
import pandas as pd

from config.settings import *
from core.excel_processor import ExcelProcessor
from core.config_manager import ConfigManager


class MainController:
    """Main controller for coordinating application logic."""
    
    def __init__(self, root):
        """
        Initialize the main controller.
        
        Args:
            root: Root tkinter widget
        """
        self.root = root
        
        # Initialize core components
        self.excel_processor = ExcelProcessor()
        self.config_manager = ConfigManager()
        
        # Initialize logging
        self.setup_logging()
        
        self.logger = logging.getLogger(__name__)
        self.logger.info("Main controller initialized")
        
    def setup_logging(self):
        """Set up logging configuration."""
        logging.basicConfig(
            level=getattr(logging, LOG_LEVEL),
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(LOG_FILE),
                logging.StreamHandler()
            ]
        )
        
    def get_file_preview(self, file_path: str, max_rows: int = MAX_ROWS_PREVIEW) -> str:
        """
        Get a preview of the Excel file or folder contents.
        
        Args:
            file_path: Path to the Excel file or folder
            max_rows: Maximum number of rows to preview
            
        Returns:
            String representation of file preview
            
        Raises:
            Exception: If file/folder cannot be read
        """
        try:
            self.logger.info(f"Getting preview for file: {file_path}")
            
            file_path_obj = Path(file_path)
            
            if file_path_obj.is_dir():
                # Handle folder preview
                return self._get_folder_preview(file_path_obj)
            else:
                # Handle single file preview
                return self._get_file_preview(file_path_obj, max_rows)
            
        except Exception as e:
            self.logger.error(f"Error generating file preview: {str(e)}")
            raise
            
    def _get_file_preview(self, file_path: Path, max_rows: int) -> str:
        """Get preview for a single Excel file."""
        df = self.excel_processor.read_excel_file(file_path)
        
        # Limit rows for preview
        preview_df = df.head(max_rows)
        
        # Format as string
        preview_text = f"File: {file_path.name}\n"
        preview_text += f"Shape: {df.shape[0]} rows × {df.shape[1]} columns\n"
        preview_text += f"Columns: {', '.join(df.columns.tolist())}\n\n"
        preview_text += "Preview (first {} rows):\n".format(min(max_rows, len(df)))
        preview_text += "=" * 80 + "\n"
        preview_text += preview_df.to_string(max_rows=max_rows, max_cols=10)
        
        if len(df) > max_rows:
            preview_text += f"\n\n... and {len(df) - max_rows} more rows"
            
        return preview_text
        
    def _get_folder_preview(self, folder_path: Path) -> str:
        """Get preview for a folder containing Excel files."""
        # Find Excel files in the folder
        excel_files = []
        for ext in ['.xlsx', '.xls', '.xlsm']:
            excel_files.extend(folder_path.glob(f"*{ext}"))
        
        if not excel_files:
            return f"Folder: {folder_path.name}\n\nNo Excel files found in this folder.\n\nSupported formats: .xlsx, .xls, .xlsm"
        
        # Sort files by name
        excel_files.sort(key=lambda f: f.name)
        
        # Get preview from first file
        try:
            first_file_preview = self._get_file_preview(excel_files[0], 50)  # Smaller preview for folders
        except Exception as e:
            first_file_preview = f"Error reading first file: {str(e)}"
        
        # Create folder summary
        preview_text = f"Folder: {folder_path.name}\n"
        preview_text += f"Excel Files Found: {len(excel_files)}\n\n"
        
        # List files
        preview_text += "Files to be processed:\n"
        preview_text += "-" * 40 + "\n"
        for i, file_path in enumerate(excel_files[:10]):  # Show first 10 files
            preview_text += f"{i+1:2d}. {file_path.name}\n"
        
        if len(excel_files) > 10:
            preview_text += f"    ... and {len(excel_files) - 10} more files\n"
        
        preview_text += "\n" + "=" * 80 + "\n"
        preview_text += f"Sample preview from first file ({excel_files[0].name}):\n"
        preview_text += "=" * 80 + "\n"
        preview_text += first_file_preview.split("Preview (first")[1] if "Preview (first" in first_file_preview else first_file_preview
        
        return preview_text
            
    def get_file_columns(self, file_path: str) -> List[str]:
        """
        Get column names from Excel file.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            List of column names
            
        Raises:
            Exception: If file cannot be read
        """
        try:
            self.logger.info(f"Getting columns for file: {file_path}")
            
            df = self.excel_processor.read_excel_file(Path(file_path))
            columns = df.columns.tolist()
            
            self.logger.info(f"Found {len(columns)} columns")
            return columns
            
        except Exception as e:
            self.logger.error(f"Error getting file columns: {str(e)}")
            raise
            
    def preview_output(self, input_file: str, config: Dict[str, Any], max_rows: int = None) -> str:
        """
        Generate a preview of the processed output.
        
        Args:
            input_file: Path to input Excel file or folder
            config: Processing configuration
            max_rows: Maximum rows to preview
            
        Returns:
            String representation of output preview
            
        Raises:
            Exception: If preview cannot be generated
        """
        try:
            self.logger.info(f"Generating output preview for: {input_file}")
            
            file_path_obj = Path(input_file)
            
            if file_path_obj.is_dir():
                # Handle folder preview
                return self._get_folder_output_preview(file_path_obj, config, max_rows)
            else:
                # Handle single file preview
                return self._get_file_output_preview(file_path_obj, config, max_rows)
            
        except Exception as e:
            self.logger.error(f"Error generating output preview: {str(e)}")
            raise
            
    def _get_file_output_preview(self, file_path: Path, config: Dict[str, Any], max_rows: int = None) -> str:
        """Get output preview for a single Excel file."""
        # Process file with configuration
        input_df = self.excel_processor.read_excel_file(file_path)
        output_df = self.excel_processor.apply_mapping(input_df, config)
        
        # Generate preview - show all rows if max_rows is None
        if max_rows is None:
            preview_df = output_df
            preview_text = f"Output Preview for: {file_path.name}\n"
            preview_text += f"Output Shape: {output_df.shape[0]} rows × {output_df.shape[1]} columns\n"
            preview_text += f"Output Columns: {', '.join(output_df.columns.tolist())}\n\n"
            preview_text += f"Complete Output ({len(output_df)} rows):\n"
            preview_text += "=" * 80 + "\n"
            preview_text += preview_df.to_string()
        else:
            preview_df = output_df.head(max_rows)
            preview_text = f"Output Preview for: {file_path.name}\n"
            preview_text += f"Output Shape: {output_df.shape[0]} rows × {output_df.shape[1]} columns\n"
            preview_text += f"Output Columns: {', '.join(output_df.columns.tolist())}\n\n"
            preview_text += f"Preview (first {min(max_rows, len(output_df))} rows):\n"
            preview_text += "=" * 80 + "\n"
            preview_text += preview_df.to_string(max_rows=max_rows)
            
            if len(output_df) > max_rows:
                preview_text += f"\n\n... and {len(output_df) - max_rows} more rows"
            
        self.logger.info("Output preview generated successfully")
        return preview_text
        
    def _get_folder_output_preview(self, folder_path: Path, config: Dict[str, Any], max_rows: int = None) -> str:
        """Get output preview for a folder containing Excel files."""
        # Find Excel files in the folder
        excel_files = []
        for ext in ['.xlsx', '.xls', '.xlsm']:
            excel_files.extend(folder_path.glob(f"*{ext}"))
        
        if not excel_files:
            return f"Folder: {folder_path.name}\n\nNo Excel files found in this folder.\n\nSupported formats: .xlsx, .xls, .xlsm"
        
        # Sort files by name
        excel_files.sort(key=lambda f: f.name)
        
        # Get preview from first file - show all rows if max_rows is None
        try:
            first_file_preview = self._get_file_output_preview(excel_files[0], config, max_rows)
        except Exception as e:
            first_file_preview = f"Error reading first file: {str(e)}"
        
        # Create folder summary
        preview_text = f"Folder: {folder_path.name}\n"
        preview_text += f"Excel Files Found: {len(excel_files)}\n\n"
        
        # List files
        preview_text += "Files to be processed:\n"
        preview_text += "-" * 40 + "\n"
        for i, file_path in enumerate(excel_files[:10]):  # Show first 10 files
            preview_text += f"{i+1:2d}. {file_path.name}\n"
        
        if len(excel_files) > 10:
            preview_text += f"    ... and {len(excel_files) - 10} more files\n"
        
        preview_text += "\n" + "=" * 80 + "\n"
        if max_rows is None:
            preview_text += f"Complete output preview from first file ({excel_files[0].name}):\n"
        else:
            preview_text += f"Sample output preview from first file ({excel_files[0].name}):\n"
        preview_text += "=" * 80 + "\n"
        
        # Extract the preview content (skip the header info)
        if "Complete Output" in first_file_preview:
            preview_text += first_file_preview.split("Complete Output")[1]
        elif "Preview (first" in first_file_preview:
            preview_text += first_file_preview.split("Preview (first")[1]
        else:
            preview_text += first_file_preview
        
        return preview_text
            
    def process_file(self, input_file: str, output_dir: str, config: Dict[str, Any]) -> str:
        """
        Process Excel file with the given configuration.
        
        Args:
            input_file: Path to input Excel file
            output_dir: Output directory path
            config: Processing configuration
            
        Returns:
            Path to the generated output file
            
        Raises:
            Exception: If processing fails
        """
        try:
            self.logger.info(f"Processing file: {input_file}")
            
            # Create output directory if it doesn't exist
            output_dir_path = Path(output_dir)
            output_dir_path.mkdir(parents=True, exist_ok=True)
            
            # Process the file
            output_file_path = self.excel_processor.process_file(
                Path(input_file), 
                output_dir_path, 
                config
            )
            
            self.logger.info(f"File processed successfully: {output_file_path}")
            return str(output_file_path)
            
        except Exception as e:
            self.logger.error(f"Error processing file: {str(e)}")
            raise
            
    def process_folder(self, input_folder: str, output_dir: str, config: Dict[str, Any]) -> list:
        """
        Process all Excel files in a folder with the given configuration.
        
        Args:
            input_folder: Path to folder containing Excel files
            output_dir: Output directory path
            config: Processing configuration
            
        Returns:
            List of paths to generated output files
            
        Raises:
            Exception: If processing fails
        """
        try:
            self.logger.info(f"Processing folder: {input_folder}")
            
            folder_path = Path(input_folder)
            if not folder_path.exists() or not folder_path.is_dir():
                raise ValueError(f"Invalid folder path: {input_folder}")
            
            # Create output directory if it doesn't exist
            output_dir_path = Path(output_dir)
            output_dir_path.mkdir(parents=True, exist_ok=True)
            
            # Find all Excel files in the folder
            excel_files = []
            for ext in ['.xlsx', '.xls', '.xlsm']:
                excel_files.extend(folder_path.glob(f"*{ext}"))
            
            if not excel_files:
                self.logger.warning(f"No Excel files found in folder: {input_folder}")
                return []
            
            self.logger.info(f"Found {len(excel_files)} Excel files to process")
            
            # Process each file
            output_files = []
            for excel_file in excel_files:
                try:
                    self.logger.info(f"Processing file: {excel_file.name}")
                    
                    # Process the file
                    output_file_path = self.excel_processor.process_file(
                        excel_file, 
                        output_dir_path, 
                        config
                    )
                    
                    output_files.append(str(output_file_path))
                    self.logger.info(f"Successfully processed: {excel_file.name}")
                    
                except Exception as e:
                    self.logger.error(f"Error processing file {excel_file.name}: {str(e)}")
                    # Continue processing other files even if one fails
                    continue
            
            self.logger.info(f"Batch processing complete: {len(output_files)} files processed successfully")
            return output_files
            
        except Exception as e:
            self.logger.error(f"Error processing folder: {str(e)}")
            raise
            
    def load_configuration(self, config_file: str) -> Dict[str, Any]:
        """
        Load configuration from JSON file.
        
        Args:
            config_file: Path to configuration file
            
        Returns:
            Configuration dictionary
            
        Raises:
            Exception: If configuration cannot be loaded
        """
        try:
            self.logger.info(f"Loading configuration from: {config_file}")
            
            config = self.config_manager.load_config(Path(config_file))
            
            self.logger.info("Configuration loaded successfully")
            return config
            
        except Exception as e:
            self.logger.error(f"Error loading configuration: {str(e)}")
            raise
            
    def save_configuration(self, config: Dict[str, Any], config_file: str):
        """
        Save configuration to JSON file.
        
        Args:
            config: Configuration dictionary
            config_file: Path to save configuration file
            
        Raises:
            Exception: If configuration cannot be saved
        """
        try:
            self.logger.info(f"Saving configuration to: {config_file}")
            
            self.config_manager.save_config(config, Path(config_file))
            
            self.logger.info("Configuration saved successfully")
            
        except Exception as e:
            self.logger.error(f"Error saving configuration: {str(e)}")
            raise
            
    def get_default_configuration(self) -> Dict[str, Any]:
        """
        Get default configuration.
        
        Returns:
            Default configuration dictionary
        """
        return DEFAULT_MAPPING_CONFIG.copy()
        
    def validate_configuration(self, config: Dict[str, Any]) -> tuple[bool, str]:
        """
        Validate configuration dictionary.
        
        Args:
            config: Configuration to validate
            
        Returns:
            Tuple of (is_valid, error_message)
        """
        try:
            # Check required fields
            if "output_columns" not in config:
                return False, "Missing 'output_columns' in configuration"
                
            output_columns = config["output_columns"]
            if not isinstance(output_columns, list):
                return False, "'output_columns' must be a list"
                
            if not output_columns:
                return False, "At least one output column must be defined"
                
            # Validate each column configuration
            for i, col_config in enumerate(output_columns):
                if not isinstance(col_config, dict):
                    return False, f"Column {i+1} configuration must be a dictionary"
                    
                if "name" not in col_config:
                    return False, f"Column {i+1} missing required 'name' field"
                    
                if not col_config["name"].strip():
                    return False, f"Column {i+1} name cannot be empty"
                    
                # Check alignment if specified
                alignment = col_config.get("alignment", "left")
                if alignment not in COLUMN_ALIGNMENTS:
                    return False, f"Column {i+1} has invalid alignment: {alignment}"
                    
            self.logger.info("Configuration validation successful")
            return True, ""
            
        except Exception as e:
            error_msg = f"Configuration validation error: {str(e)}"
            self.logger.error(error_msg)
            return False, error_msg
            
    def get_recent_files(self, max_files: int = 5) -> List[str]:
        """
        Get list of recently processed files.
        
        Args:
            max_files: Maximum number of files to return
            
        Returns:
            List of recent file paths
        """
        # This would typically be stored in settings or a database
        # For now, return empty list
        return []
        
    def cleanup_temp_files(self):
        """Clean up temporary files if any."""
        try:
            # Clean up any temporary files created during processing
            pass
        except Exception as e:
            self.logger.warning(f"Error during cleanup: {str(e)}")
            
    def shutdown(self):
        """Perform cleanup when application is closing."""
        self.logger.info("Shutting down main controller")
        self.cleanup_temp_files()
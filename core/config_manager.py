#!/usr/bin/env python3
"""
Configuration manager for Excel Formatter application.

This module handles loading, saving, and validating configuration files
for the Excel Formatter application.
"""

import json
import logging
from pathlib import Path
from typing import Dict, Any, Optional

from config.settings import DEFAULT_MAPPING_CONFIG, DEFAULT_CONFIG_FILE


class ConfigManager:
    """Manages configuration loading, saving, and validation."""
    
    def __init__(self):
        """Initialize the configuration manager."""
        self.logger = logging.getLogger(__name__)
        
    def load_config(self, config_path: Path) -> Dict[str, Any]:
        """
        Load configuration from JSON file.
        
        Args:
            config_path: Path to configuration file
            
        Returns:
            Configuration dictionary
            
        Raises:
            FileNotFoundError: If configuration file doesn't exist
            json.JSONDecodeError: If configuration file is invalid JSON
            ValueError: If configuration is invalid
        """
        try:
            self.logger.info(f"Loading configuration from: {config_path}")
            
            if not config_path.exists():
                raise FileNotFoundError(f"Configuration file not found: {config_path}")
                
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                
            # Validate configuration
            self.validate_config(config)
            
            self.logger.info("Configuration loaded and validated successfully")
            return config
            
        except json.JSONDecodeError as e:
            error_msg = f"Invalid JSON in configuration file: {str(e)}"
            self.logger.error(error_msg)
            raise json.JSONDecodeError(error_msg, e.doc, e.pos)
            
        except Exception as e:
            self.logger.error(f"Error loading configuration: {str(e)}")
            raise
            
    def save_config(self, config: Dict[str, Any], config_path: Path):
        """
        Save configuration to JSON file.
        
        Args:
            config: Configuration dictionary to save
            config_path: Path where to save configuration
            
        Raises:
            ValueError: If configuration is invalid
            OSError: If file cannot be written
        """
        try:
            self.logger.info(f"Saving configuration to: {config_path}")
            
            # Validate configuration before saving
            self.validate_config(config)
            
            # Ensure directory exists
            config_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Save with pretty formatting
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
                
            self.logger.info("Configuration saved successfully")
            
        except Exception as e:
            self.logger.error(f"Error saving configuration: {str(e)}")
            raise
            
    def validate_config(self, config: Dict[str, Any]):
        """
        Validate configuration dictionary.
        
        Args:
            config: Configuration to validate
            
        Raises:
            ValueError: If configuration is invalid
        """
        try:
            # Check that config is a dictionary
            if not isinstance(config, dict):
                raise ValueError("Configuration must be a dictionary")
                
            # Validate output columns
            self._validate_output_columns(config)
            
            # Validate header formatting if present
            if "header_formatting" in config:
                self._validate_header_formatting(config["header_formatting"])
                
            # Validate general settings if present
            if "general_settings" in config:
                self._validate_general_settings(config["general_settings"])
                
            # Validate void settings if present
            if "void" in config:
                self._validate_void_settings(config["void"])
                
            self.logger.debug("Configuration validation passed")
            
        except Exception as e:
            self.logger.error(f"Configuration validation failed: {str(e)}")
            raise ValueError(f"Invalid configuration: {str(e)}")
            
    def _validate_output_columns(self, config: Dict[str, Any]):
        """Validate output columns configuration."""
        if "output_columns" not in config:
            raise ValueError("Missing required 'output_columns' field")
            
        output_columns = config["output_columns"]
        
        if not isinstance(output_columns, list):
            raise ValueError("'output_columns' must be a list")
            
        if not output_columns:
            raise ValueError("'output_columns' cannot be empty")
            
        # Validate each column
        for i, col_config in enumerate(output_columns):
            self._validate_column_config(col_config, i + 1)
            
    def _validate_column_config(self, col_config: Dict[str, Any], col_num: int):
        """Validate individual column configuration."""
        if not isinstance(col_config, dict):
            raise ValueError(f"Column {col_num} must be a dictionary")
            
        # Required fields
        if "name" not in col_config:
            raise ValueError(f"Column {col_num} missing required 'name' field")
            
        if not isinstance(col_config["name"], str):
            raise ValueError(f"Column {col_num} 'name' must be a string")
            
        if not col_config["name"].strip():
            raise ValueError(f"Column {col_num} 'name' cannot be empty")
            
        # Optional fields validation
        if "source_column" in col_config and not isinstance(col_config["source_column"], str):
            raise ValueError(f"Column {col_num} 'source_column' must be a string")
            
        if "alignment" in col_config:
            alignment = col_config["alignment"]
            if alignment not in ["left", "center", "right"]:
                raise ValueError(f"Column {col_num} invalid alignment: {alignment}")
                
        if "width" in col_config:
            try:
                width = float(col_config["width"])
                if width <= 0:
                    raise ValueError(f"Column {col_num} width must be positive")
            except (TypeError, ValueError):
                raise ValueError(f"Column {col_num} width must be a number")
                
        if "formatting" in col_config and not isinstance(col_config["formatting"], dict):
            raise ValueError(f"Column {col_num} 'formatting' must be a dictionary")
            
    def _validate_header_formatting(self, header_config: Dict[str, Any]):
        """Validate header formatting configuration."""
        if not isinstance(header_config, dict):
            raise ValueError("'header_formatting' must be a dictionary")
            
        # Validate boolean fields
        for field in ["bold"]:
            if field in header_config and not isinstance(header_config[field], bool):
                raise ValueError(f"Header formatting '{field}' must be boolean")
                
        # Validate color fields (should be hex colors without #)
        for field in ["background_color", "font_color"]:
            if field in header_config:
                color = header_config[field]
                if not isinstance(color, str):
                    raise ValueError(f"Header formatting '{field}' must be a string")
                if color and not self._is_valid_hex_color(color):
                    raise ValueError(f"Header formatting '{field}' must be a valid hex color")
                    
        # Validate alignment
        if "alignment" in header_config:
            alignment = header_config["alignment"]
            if alignment not in ["left", "center", "right"]:
                raise ValueError(f"Header formatting invalid alignment: {alignment}")
                
    def _validate_general_settings(self, general_config: Dict[str, Any]):
        """Validate general settings configuration."""
        if not isinstance(general_config, dict):
            raise ValueError("'general_settings' must be a dictionary")
            
        # Validate boolean fields
        for field in ["auto_fit_columns"]:
            if field in general_config and not isinstance(general_config[field], bool):
                raise ValueError(f"General settings '{field}' must be boolean")
                
        # Validate freeze_panes format (string for legacy, dict for new format)
        if "freeze_panes" in general_config:
            freeze_panes = general_config["freeze_panes"]
            if freeze_panes:
                if isinstance(freeze_panes, str):
                    # Legacy format (e.g., "A2", "B3") - validate Excel cell reference
                    import re
                    if not re.match(r'^[A-Z]+\d+$', freeze_panes.upper()):
                        raise ValueError("General settings 'freeze_panes' string must be a valid Excel cell reference (e.g., 'A2', 'B3')")
                elif isinstance(freeze_panes, dict):
                    # New format with freeze_header and freeze_columns
                    if "freeze_header" in freeze_panes and not isinstance(freeze_panes["freeze_header"], bool):
                        raise ValueError("General settings freeze_panes 'freeze_header' must be boolean")
                    if "freeze_columns" in freeze_panes:
                        freeze_columns = freeze_panes["freeze_columns"]
                        if not isinstance(freeze_columns, list):
                            raise ValueError("General settings freeze_panes 'freeze_columns' must be a list")
                        for i, col_name in enumerate(freeze_columns):
                            if not isinstance(col_name, str):
                                raise ValueError(f"General settings freeze_panes 'freeze_columns[{i}]' must be a string")
                else:
                    raise ValueError("General settings 'freeze_panes' must be a string or dictionary")
                
    def _validate_void_settings(self, void_config: Dict[str, Any]):
        """Validate void filtering configuration."""
        if not isinstance(void_config, dict):
            raise ValueError("'void' must be a dictionary")
            
        # Validate enabled field
        if "enabled" in void_config and not isinstance(void_config["enabled"], bool):
            raise ValueError("Void settings 'enabled' must be boolean")
            
        # Validate zero_columns field
        if "zero_columns" in void_config:
            zero_columns = void_config["zero_columns"]
            if not isinstance(zero_columns, list):
                raise ValueError("Void settings 'zero_columns' must be a list")
            for i, col_name in enumerate(zero_columns):
                if not isinstance(col_name, str):
                    raise ValueError(f"Void settings 'zero_columns[{i}]' must be a string")
                    
    def _is_valid_hex_color(self, color: str) -> bool:
        """Check if string is a valid hex color (without #)."""
        if not color:
            return True  # Empty string is allowed
            
        # Should be 6 characters, all hex digits
        if len(color) != 6:
            return False
            
        try:
            int(color, 16)
            return True
        except ValueError:
            return False
            
    def get_default_config(self) -> Dict[str, Any]:
        """
        Get default configuration.
        
        Returns:
            Default configuration dictionary
        """
        return DEFAULT_MAPPING_CONFIG.copy()
        
    def load_default_config(self) -> Dict[str, Any]:
        """
        Load default configuration from file or return built-in default.
        
        Returns:
            Default configuration dictionary
        """
        try:
            if DEFAULT_CONFIG_FILE.exists():
                return self.load_config(DEFAULT_CONFIG_FILE)
            else:
                self.logger.info("No default config file found, using built-in defaults")
                return self.get_default_config()
        except Exception as e:
            self.logger.warning(f"Error loading default config: {str(e)}, using built-in defaults")
            return self.get_default_config()
            
    def create_sample_config(self, config_path: Path):
        """
        Create a sample configuration file.
        
        Args:
            config_path: Path where to create sample configuration
        """
        try:
            sample_config = {
                "output_columns": [
                    {
                        "name": "Employee Name",
                        "source_column": "Name",
                        "alignment": "left",
                        "width": 20,
                        "formatting": {}
                    },
                    {
                        "name": "Amount",
                        "source_column": "Net Pay",
                        "alignment": "right",
                        "width": 12,
                        "formatting": {
                            "number_format": "#,##0.00"
                        }
                    },
                    {
                        "name": "Date",
                        "source_column": "Pay Date",
                        "alignment": "center",
                        "width": 12,
                        "formatting": {
                            "date_format": "MM/DD/YYYY"
                        }
                    }
                ],
                "header_formatting": {
                    "bold": True,
                    "background_color": "366092",
                    "font_color": "FFFFFF",
                    "alignment": "center"
                },
                "general_settings": {
                    "auto_fit_columns": True
                },
                "void": {
                    "enabled": False,
                    "zero_columns": []
                }
            }
            
            self.save_config(sample_config, config_path)
            self.logger.info(f"Sample configuration created at: {config_path}")
            
        except Exception as e:
            self.logger.error(f"Error creating sample configuration: {str(e)}")
            raise
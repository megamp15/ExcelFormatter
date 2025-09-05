# Excel Formatter

A professional GUI application for processing and formatting Excel files with configurable column mappings and advanced formatting options.

## Features

- **Professional GUI Interface**: Tabbed interface with file selection, column mapping, and output settings
- **Flexible Column Mapping**: Direct mapping, formulas, and blank columns
- **Advanced Formatting**: Custom number formats, colors, alignment, and styling
- **Void Filtering**: Remove rows where specified columns are all zero
- **Preview Functionality**: Preview output before processing
- **Configuration Management**: Save and load mapping configurations
- **Multiple File Format Support**: .xlsx, .xls, .xlsm files

## Installation

1. Ensure you have Python 3.8+ installed
2. Install required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### GUI Mode (Default)

```bash
python main.py
```

### Command Line Options

```bash
python main.py --help          # Show help
python main.py --version       # Show version
```

## GUI Interface

The application features a tabbed interface with three main sections:

### 1. File Selection

- Select input Excel files
- View file information and preview
- Choose output directory

### 2. Column Mapping

- Map input columns to output columns
- Support for different mapping types:
  - **Direct mapping**: `Input Column → Output Column`
  - **Blank columns**: Leave source empty for blank columns
  - **Formulas**: `=Chk Amt - Gross - Fica` for calculations
- Configure column alignment and width
- Advanced formatting options per column

### 3. Output Settings

- **Header Formatting**: Colors, fonts, alignment
- **General Settings**: Auto-fit columns, freeze panes
- **Void Filtering**: Remove rows with zero values in specified columns

## Configuration Files

### Mapping Configuration (JSON)

Save and load column mapping configurations for reuse:

```json
{
  "output_columns": [
    {
      "name": "Employee Name",
      "source_column": "Name",
      "alignment": "left",
      "width": 20,
      "formatting": {}
    }
  ],
  "header_formatting": {
    "bold": true,
    "background_color": "366092",
    "font_color": "FFFFFF",
    "alignment": "center"
  },
  "void": {
    "enabled": true,
    "zero_columns": ["Amount", "Hours"]
  }
}
```

## Column Mapping Types

### Direct Mapping

Maps input column directly to output column:

- **Source**: `Employee Name`
- **Output**: Direct copy of values

### Blank Columns

Creates empty columns in output:

- **Source**: `` (empty)
- **Output**: Blank column

### Formula Columns

Calculates values using Excel-style formulas:

- **Source**: `=Chk Amt - Gross - Fica`
- **Output**: Calculated values

## File Structure

```
ExcelFormatter/
├── main.py                 # Main application entry point
├── requirements.txt        # Python dependencies
├── config/
│   ├── settings.py        # Application settings
│   └── mapping_config.json # Default mapping configuration
├── gui/
│   ├── views/
│   │   └── main_window.py # Main GUI window
│   ├── controllers/
│   │   └── main_controller.py # Business logic controller
│   └── components/
│       ├── file_selector.py # File selection component
│       ├── column_mapper.py # Column mapping component
│       ├── output_settings.py # Output settings component
│       └── progress_dialog.py # Progress dialog
├── core/
│   ├── excel_processor.py # Core Excel processing logic
│   └── config_manager.py  # Configuration management
├── utils/              # Utility modules (future)
├── input_files/       # Input Excel files directory
├── output_files/      # Generated output files directory
└── logs/             # Application logs
```

## Development

The application follows the Model-View-Controller (MVC) pattern:

- **Views**: GUI components in `gui/views/` and `gui/components/`
- **Controllers**: Business logic in `gui/controllers/`
- **Models**: Core processing in `core/`

### Key Components

- **MainWindow**: Primary GUI interface with tabbed layout
- **FileSelector**: File selection and preview component
- **ColumnMapper**: Dynamic column mapping interface
- **OutputSettings**: Formatting and void filtering settings
- **ExcelProcessor**: Core Excel file processing logic
- **ConfigManager**: Configuration loading and validation

## Logging

Application logs are saved to `logs/excel_formatter.log` with automatic rotation when files exceed 10MB.

## Error Handling

The application includes comprehensive error handling:

- File validation and format checking
- Configuration validation
- Processing error recovery
- User-friendly error messages

## Future Enhancements

- Command-line processing mode
- Batch file processing
- Custom formula functions
- Template management
- Multi-sheet support
- Export to other formats

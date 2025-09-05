# Excel Formatter Configuration Guide

## Overview

The Excel Formatter uses a JSON configuration file (`mapping_config.json`) to define how input data is mapped and formatted in the output Excel file.

## Configuration Structure

### 1. Output Columns

Each column in the output file is defined with the following properties:

```json
{
  "name": "Column Name in Output",
  "source_column": "Source Column Name or Formula",
  "alignment": "left|center|right",
  "width": 15,
  "formatting": {
    "number_format": "#,##0.00",
    "negative_format": "(#,##0.00)",
    "date_format": "MM/DD/YYYY",
    "date_range_format": "MM/DD/YYYY - MM/DD/YYYY"
  }
}
```

#### Column Properties:

- **name**: The header name that appears in the output Excel file
- **source_column**:
  - Empty string `""` for blank columns
  - Column name from input file for direct mapping
  - `"Column1 + Column2"` for combining two columns
  - `"=Formula"` for calculated columns (e.g., `"=Chk Amt - Gross - Fica"`)
- **alignment**: How data is aligned in the column (`left`, `center`, `right`)
- **width**: Column width in Excel units (optional)
- **formatting**: Number and date formatting options

### 2. Header Formatting

Controls the appearance of the header row:

```json
"header_formatting": {
  "bold": true,
  "background_color": "366092",
  "font_color": "FFFFFF",
  "alignment": "center"
}
```

### 3. Column Name Alignment

Allows individual alignment for each column header (overrides default header alignment):

```json
"column_name_alignment": {
  "Employee Name": "left",
  "Check #": "center",
  "Chk Amt": "right",
  "Gross": "right",
  "Fica E/R": "right",
  "Liab": "right",
  "Date": "center",
  "Period": "center"
}
```

### 4. General Settings

Worksheet-level settings:

```json
"general_settings": {
  "auto_fit_columns": true
}
```

### 5. Void Filtering

Remove rows where specified columns are all zero:

```json
"void": {
  "enabled": true,
  "zero_columns": ["Chk Amt", "Gross", "Fica E/R"]
}
```

#### Void Filtering Properties:

- **enabled**: `true` to enable void filtering, `false` to disable
- **zero_columns**: Array of column names to check for zero values
- **Behavior**: Rows where ALL specified columns are zero (or empty) will be removed

## Example Mappings

### Direct Column Mapping

```json
{
  "name": "Employee Name",
  "source_column": "Name",
  "alignment": "left"
}
```

### Blank Column

```json
{
  "name": "Check #",
  "source_column": "",
  "alignment": "center"
}
```

```json
{
  "name": "Fica E/R",
  "source_column": "Employee taxes - SS + Employee taxes - Med",
  "alignment": "right"
}
```

### Formula Column

```json
{
  "name": "Liab",
  "source_column": "=Chk Amt - Gross - Fica",
  "alignment": "right",
  "formatting": {
    "negative_format": "(#,##0.00)"
  }
}
```

## Number Formatting Options

- `"#,##0.00"` - Currency with thousands separator
- `"0.00"` - Decimal with 2 places
- `"0"` - Whole numbers
- `"mm/dd/yyyy"` - Date format
- `"(#,##0.00)"` - Negative numbers in parentheses

## Usage

1. Modify `mapping_config.json` to change column mappings and formatting
2. Run the formatter: `python run_formatter.py`
3. Output files will be generated in the `output_files` directory with timestamps

## Tips

- Column names in `column_name_alignment` must match exactly with the `name` field in `output_columns`
- Use empty string `""` for `source_column` to create blank columns
- Formulas use Excel-style syntax with `=` prefix
- All alignment options: `left`, `center`, `right`

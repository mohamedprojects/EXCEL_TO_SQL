# EXCEL_TO_SQL
A Python script that reads Excel files (.xlsx or .xls) and generates SQL INSERT INTO VALUES statements for database insertion.

## Features

- Automatically detects columns with data
- Handles various data types (strings, numbers, dates)
- Supports custom column selection
- Configurable input and output folders
- Proper SQL string escaping
- Command-line interface with flexible options

## Installation

1. Clone or download this repository
2. Install the required dependencies:

```bash
pip install -r requirements_excel_to_sql.txt
```

### Dependencies

- pandas >= 2.0.0
- openpyxl >= 3.0.0

## Usage

### Basic Usage

```bash
python excel_to_sql.py <excel_file> -t <table_name>
```

### Examples

#### Use all columns from Excel file
```bash
python excel_to_sql.py data.xlsx -t users
```

#### Specify specific columns
```bash
python excel_to_sql.py data.xlsx -t users -c name email age
```

#### Specify sheet name
```bash
python excel_to_sql.py data.xlsx -t users -s Sheet1
```

#### Output to file
```bash
python excel_to_sql.py data.xlsx -t users -o output.sql
```

#### Use input folder
```bash
python excel_to_sql.py data.xlsx -t users -i C:/excel_files
```

#### Use output folder (auto-generates filename)
```bash
python excel_to_sql.py data.xlsx -t users -d C:/sql_output
```

#### Use both input and output folders
```bash
python excel_to_sql.py data.xlsx -t users -i C:/excel_files -d C:/sql_output
```

#### Use output folder with custom filename
```bash
python excel_to_sql.py data.xlsx -t users -d C:/sql_output -o custom_output.sql
```

### Command Line Options

- `excel_file`: Path to the Excel file (.xlsx or .xls) or filename if --input-folder is used
- `-t, --table`: SQL table name (required)
- `-c, --columns`: Column names to include (default: auto-detect columns with data)
- `-s, --sheet`: Sheet name or index (default: first sheet)
- `-o, --output`: Output SQL file name (default: print to stdout or auto-generate if --output-folder is used)
- `-i, --input-folder`: Folder path where Excel files are located
- `-d, --output-folder`: Folder path where output SQL files will be saved
- `--header`: Row to use as column names (0-indexed, default: 0)
- `--skip-rows`: Number of rows to skip at the start

## Configuration

Edit the script to set default input and output folders:

```python
DEFAULT_INPUT_FOLDER = r"C:\path\to\excel\files"
DEFAULT_OUTPUT_FOLDER = r"C:\path\to\sql\output"
```

## Data Handling

- **NULL values**: NaN, None, empty strings, and '?' are converted to NULL
- **Dates**: Converted to 'YYYY-MM-DD' format 
- **Strings**: Properly escaped with single quotes
- **Numbers**: Converted to strings for SQL compatibility

## Output Format

Generates standard SQL INSERT statements:

```sql
INSERT INTO table_name (col1, col2, col3) VALUES ('value1', 'value2', 'value3');
```


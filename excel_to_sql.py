#!/usr/bin/env python3
"""
Excel to SQL INSERT Generator
Reads an Excel file and generates SQL INSERT INTO VALUES statements.
"""

import pandas as pd
import argparse
import sys
from pathlib import Path
from datetime import datetime

# ============================================================================
# CONFIGURATION: Set your default input and output folders here
# ============================================================================
DEFAULT_INPUT_FOLDER = r"C:\path\to\excel\files"  # Change this to your Excel files folder
DEFAULT_OUTPUT_FOLDER = r"C:\path\to\sql\output"   # Change this to your SQL output folder
# ============================================================================


def escape_sql_string(value):
    """Escape single quotes in SQL strings."""
    if value is None:
        return 'NULL'
    if isinstance(value, str):
        # Replace single quotes with two single quotes for SQL
        return f"'{value.replace(chr(39), chr(39) + chr(39))}'"
    return str(value)


def format_sql_value(value):
    """Format a value for SQL INSERT statement."""
    # Treat literal '?' as NULL
    if isinstance(value, str) and value.strip() == '?':
        return 'NULL'

    # pandas NaN / NA / None -> NULL
    if pd.isna(value) or value is None:
        return 'NULL'

    # Dates already parsed by pandas -> 'YYYY-MM-DD'
    if isinstance(value, pd.Timestamp):
        return f"'{value.strftime('%Y-%m-%d')}'"

    # Strings that look like dates in 'DD-MM-YYYY' -> convert to 'YYYY-MM-DD'
    if isinstance(value, str):
        s = value.strip()
        if s:
            try:
                dt = datetime.strptime(s, '%d-%m-%Y')
                return f"'{dt.strftime('%Y-%m-%d')}'"
            except ValueError:
                # Not a DD-MM-YYYY date, treat as normal string below
                pass

    # For this use-case: everything else is treated as a string and quoted
    # (numbers, booleans, text, etc.), with proper escaping.
    return escape_sql_string(str(value))


def _detect_columns_with_data(df):
    """
    Detect columns that actually contain data.
    A column is considered empty if all its values are:
      - NaN / None
      - empty string
      - literal '?'
    """
    cols_with_data = []
    for col in df.columns:
        col_has_data = False
        for v in df[col]:
            # Skip NaN / None
            if pd.isna(v) or v is None:
                continue
            s = str(v).strip()
            if s == '' or s == '?':
                continue
            col_has_data = True
            break
        if col_has_data:
            cols_with_data.append(col)
    return cols_with_data


def generate_insert_statements(df, table_name, columns=None):
    """
    Generate SQL INSERT statements from DataFrame.
    
    Args:
        df: pandas DataFrame
        table_name: Name of the SQL table
        columns: List of column names to use. If None, automatically
                 choose only the columns that actually contain data.
    
    Returns:
        List of SQL INSERT statements
    """
    if columns is None:
        # Auto-detect only columns that actually have data
        columns = _detect_columns_with_data(df)
    else:
        # Validate that all specified columns exist
        missing_cols = [col for col in columns if col not in df.columns]
        if missing_cols:
            raise ValueError(f"Columns not found in Excel file: {missing_cols}")
    
    insert_statements = []
    
    for _, row in df.iterrows():
        values = [format_sql_value(row[col]) for col in columns]
        columns_str = ', '.join(columns)
        values_str = ', '.join(values)
        insert_stmt = f"INSERT INTO {table_name} ({columns_str}) VALUES ({values_str});"
        insert_statements.append(insert_stmt)
    
    return insert_statements


def main():
    parser = argparse.ArgumentParser(
        description='Generate SQL INSERT statements from an Excel file',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Use all columns from Excel file
  python excel_to_sql.py data.xlsx -t users
  
  # Specify specific columns
  python excel_to_sql.py data.xlsx -t users -c name email age
  
  # Specify sheet name
  python excel_to_sql.py data.xlsx -t users -s Sheet1
  
  # Output to file
  python excel_to_sql.py data.xlsx -t users -o output.sql
  
  # Use input folder
  python excel_to_sql.py data.xlsx -t users -i C:/excel_files
  
  # Use output folder (auto-generates filename)
  python excel_to_sql.py data.xlsx -t users -d C:/sql_output
  
  # Use both input and output folders
  python excel_to_sql.py data.xlsx -t users -i C:/excel_files -d C:/sql_output
  
  # Use output folder with custom filename
  python excel_to_sql.py data.xlsx -t users -d C:/sql_output -o custom_output.sql
        """
    )
    
    parser.add_argument('excel_file', help='Path to the Excel file (.xlsx or .xls) or filename if --input-folder is used')
    parser.add_argument('-t', '--table', required=True, help='SQL table name')
    parser.add_argument('-c', '--columns', nargs='+', help='Column names to include (default: auto-detect columns with data)')
    parser.add_argument('-s', '--sheet', default=0, help='Sheet name or index (default: first sheet)')
    parser.add_argument('-o', '--output', help='Output SQL file name (default: print to stdout or auto-generate if --output-folder is used)')
    parser.add_argument('-i', '--input-folder', default=None, help='Folder path where Excel files are located')
    parser.add_argument('-d', '--output-folder', default=None, help='Folder path where output SQL files will be saved')
    parser.add_argument('--header', type=int, default=0, help='Row to use as column names (0-indexed, default: 0)')
    parser.add_argument('--skip-rows', type=int, default=0, help='Number of rows to skip at the start')
    
    args = parser.parse_args()
    
    # Determine input file path
    # Use command-line argument if provided, otherwise use default from configuration
    input_folder_path = getattr(args, 'input_folder', None) or DEFAULT_INPUT_FOLDER
    
    if input_folder_path:
        input_folder = Path(input_folder_path)
        if not input_folder.exists() or not input_folder.is_dir():
            print(f"Error: Input folder '{input_folder_path}' does not exist or is not a directory.", file=sys.stderr)
            sys.exit(1)
        excel_path = input_folder / args.excel_file
    else:
        excel_path = Path(args.excel_file)
    
    # Check if file exists
    if not excel_path.exists():
        print(f"Error: File '{excel_path}' not found.", file=sys.stderr)
        sys.exit(1)
    
    # Determine output file path
    # Use command-line argument if provided, otherwise use default from configuration
    output_folder_path = getattr(args, 'output_folder', None) or DEFAULT_OUTPUT_FOLDER
    
    output_path = None
    if output_folder_path:
        output_folder = Path(output_folder_path)
        if not output_folder.exists():
            # Create output folder if it doesn't exist
            try:
                output_folder.mkdir(parents=True, exist_ok=True)
                print(f"Created output folder: {output_folder}", file=sys.stderr)
            except Exception as e:
                print(f"Error: Cannot create output folder '{output_folder_path}': {e}", file=sys.stderr)
                sys.exit(1)
        
        if args.output:
            # Use specified output filename in the output folder
            output_path = output_folder / args.output
        else:
            # Auto-generate output filename based on Excel file name
            excel_stem = excel_path.stem
            output_path = output_folder / f"{excel_stem}_{args.table}.sql"
            print(f"Auto-generated output filename: {output_path.name}", file=sys.stderr)
    elif args.output:
        # Use specified output path (relative or absolute)
        output_path = Path(args.output)
    
    try:
        # Read Excel file
        print(f"Reading Excel file: {excel_path}", file=sys.stderr)
        df = pd.read_excel(
            excel_path,
            sheet_name=args.sheet,
            header=args.header,
            skiprows=args.skip_rows
        )
        
        if df.empty:
            print("Warning: Excel file is empty or contains no data.", file=sys.stderr)
            sys.exit(0)
        
        print(f"Found {len(df)} rows and {len(df.columns)} columns", file=sys.stderr)
        print(f"Columns: {', '.join(df.columns.tolist())}", file=sys.stderr)
        
        # Generate INSERT statements
        insert_statements = generate_insert_statements(df, args.table, args.columns)
        
        # Output results
        output_lines = insert_statements
        
        if output_path:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(output_lines))
            print(f"\nGenerated {len(insert_statements)} INSERT statements", file=sys.stderr)
            print(f"Output saved to: {output_path}", file=sys.stderr)
        else:
            # Print to stdout
            print('\n'.join(output_lines))
    
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()

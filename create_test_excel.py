#!/usr/bin/env python3
"""
Create a sample Excel file for testing the excel_to_sql.py script
"""

import pandas as pd
from datetime import datetime

# Create sample data
data = {
    'id': [1, 2, 3, 4, 5],
    'name': ['John Doe', 'Jane Smith', 'Bob Johnson', "Mary O'Connor", 'Alice Brown'],
    'email': ['john@example.com', 'jane@example.com', 'bob@example.com', 'mary@example.com', 'alice@example.com'],
    'age': [30, 25, 35, 28, 32],
    'salary': [50000.50, 60000.00, 55000.75, 65000.00, 70000.00],
    'is_active': [True, True, False, True, True],
    'created_at': [datetime(2023, 1, 15), datetime(2023, 2, 20), datetime(2023, 3, 10), datetime(2023, 4, 5), datetime(2023, 5, 12)],
    'notes': ['Manager', 'Developer', None, 'Designer', 'Analyst']  # One NULL value
}

df = pd.DataFrame(data)

# Save to Excel file
output_file = 'test_data.xlsx'
df.to_excel(output_file, index=False, sheet_name='Users')
print(f"Created test Excel file: {output_file}")
print(f"\nSample data preview:")
print(df.to_string())
print(f"\nTotal rows: {len(df)}")
print(f"Columns: {', '.join(df.columns.tolist())}")

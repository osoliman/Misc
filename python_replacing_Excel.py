import pandas as pd
from google.cloud import bigquery
from openpyxl import load_workbook

# Initialize BigQuery client
client = bigquery.Client()

# First, get the new values from BigQuery
unique_identifiers = ['A', 'B']  # Get these from your actual data
query = f"""
SELECT 
    identifier,
    new_value
FROM your_project.your_dataset.your_table
WHERE identifier IN ({','.join([f"'{id}'" for id in unique_identifiers])})
"""

# Execute query and convert to DataFrame
bq_df = client.query(query).to_dataframe()
bq_df.columns = ['identifier', 'new_value']

# Create a dictionary for quick lookup
replacement_dict = dict(zip(bq_df['identifier'], bq_df['new_value']))

# Load the workbook and select the active sheet
wb = load_workbook('input.xlsx', data_only=False)  # data_only=False preserves formulas
ws = wb.active

# Define the columns where identifier and value are located
id_col = 'A'  # Change these to match your Excel structure
value_col = 'B'

# Update only the specific cells while preserving everything else
for row in range(2, ws.max_row + 1):  # Assuming header is in row 1
    identifier = ws[f'{id_col}{row}'].value
    if identifier in replacement_dict:
        ws[f'{value_col}{row}'].value = replacement_dict[identifier]

# Save the workbook
wb.save('output.xlsx')

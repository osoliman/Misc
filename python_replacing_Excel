import pandas as pd
from google.cloud import bigquery

# Initialize BigQuery client
client = bigquery.Client()

# Read the Excel file
excel_df = pd.read_excel('input.xlsx', names=['identifier', 'value'])

# Get unique identifiers from Excel
unique_identifiers = excel_df['identifier'].unique().tolist()

# Query BigQuery for the new values
query = f"""
SELECT 
    identifier,
    new_value
FROM your_project.your_dataset.your_table
WHERE identifier IN ({','.join([f"'{id}'" for id in unique_identifiers])})
"""

# Execute query and convert to pandas DataFrame
bq_df = client.query(query).to_dataframe()
bq_df.columns = ['identifier', 'new_value']

# Merge Excel data with BigQuery data based on identifier
merged_df = pd.merge(
    excel_df,
    bq_df,
    on='identifier',
    how='left'
)

# Replace old values with new values
merged_df['value'] = merged_df['new_value']

# Drop the new_value column and keep original structure
final_df = merged_df[['identifier', 'value']]

# Save the modified DataFrame to a new Excel file
final_df.to_excel('output.xlsx', index=False)

# Check for any unmatched identifiers (optional)
unmatched = excel_df[~excel_df['identifier'].isin(bq_df['identifier'])]['identifier'].tolist()
if unmatched:
    print(f"Warning: No matching values found for identifiers: {unmatched}")

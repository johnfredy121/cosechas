import pandas as pd
import json

# Load the Excel files
ref_file = 'chia_compa_cursos_09032026_CierreFebrero2026 (1).xlsx'
src_file = 'chia_compa_cursos_CierreFeb2026.xlsx'

# Read the Excel files
ref_df = pd.read_excel(ref_file)
src_df = pd.read_excel(src_file)

# Initialize report
report = {'column_structure': {}, 'formulas': {}, 'hidden_columns': []}

# Compare column structures
ref_columns = set(ref_df.columns)
src_columns = set(src_df.columns)

# Identify missing columns in source and extra columns
missing_columns = ref_columns - src_columns
extra_columns = src_columns - ref_columns

# Report column differences
report['column_structure']['missing_columns'] = list(missing_columns)
report['column_structure']['extra_columns'] = list(extra_columns)

# Check for formulas and hidden columns
for col in ref_columns:
    if col in src_df.columns:
        # Check if the column is hidden (can be done through some logic based on actual data)
        hidden = src_df[col].isnull().all()  # Example condition for hidden column
        if hidden:
            report['hidden_columns'].append(col)

# Generate JSON report
json_report = json.dumps(report, indent=4)

# Save the JSON report
with open('transformation_report.json', 'w') as f:
    f.write(json_report)

print('Transformation report generated successfully!')

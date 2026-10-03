from datetime import datetime
import os
import re
import sys
import pandas as pd


# Helper to automatically resolve input paths into ../downloads/ if not found locally
def resolve_input_path(user_input, default_dir='../downloads'):
  if not user_input:
    return ''
  cleaned = str(user_input).strip().strip('"\'')
  if os.path.exists(cleaned):
    return cleaned
  download_path = os.path.join(default_dir, cleaned)
  if os.path.exists(download_path):
    return download_path
  return cleaned


# Get CSV filename from command-line arguments or interactive prompt as fallback
if len(sys.argv) < 2:
  csv_file = input('Enter path to El Paso spatial properties CSV file: ')
else:
  csv_file = sys.argv[1]

csv_file = resolve_input_path(csv_file)

if not os.path.exists(csv_file):
  raise FileNotFoundError(f'CSV file not found at: {csv_file}')

# 1. Load the CSV file
df = pd.read_csv(csv_file)

# 2. Define the base URL
base_url = 'https://property.spatialest.com/co/elpaso/#/property/'


# Helper function to format the unit column cleanly without trailing .0
def format_unit(val):
  if pd.isna(val):
    return ''
  try:
    f_val = float(val)
    if f_val.is_integer():
      return str(int(f_val))
  except ValueError:
    pass
  return str(val)


# Helper function to remove the leading street number to get just the street name for grouping
def extract_street_name(prop_name):
  if pd.isna(prop_name):
    return ''
  # Removes leading digits and spaces (e.g., "14025 Gleneagle Dr" -> "Gleneagle Dr")
  return re.sub(r'^\d+\s+', '', str(prop_name)).strip()


# 3. Create a helper column for the street name (without street number)
col_b_name = df.columns[1]
col_c_name = df.columns[2]
df['Street_Group'] = df[col_b_name].apply(extract_street_name)

# Sort by the extracted street name, then by the full property name/number
df = df.sort_values(by=['Street_Group', col_b_name])

# 4. Generate separate tables grouped by street name
tables_html = ''
grouped = df.groupby('Street_Group', sort=False)

for street_name, group in grouped:
  # Create a 2-column DataFrame for this specific street group
  group_df = pd.DataFrame()
  # Column 1: Full property name formatted as the link (opening in a new tab)
  group_df[col_b_name] = group.apply(
      lambda row: f'<a href="{base_url}{row.iloc[2]}" target="_blank">{row[col_b_name]}</a>',
      axis=1,
  )
  # Column 2: Cleaned unit text
  group_df[col_c_name] = group.iloc[:, 2].apply(format_unit)

  # Convert group dataframe to HTML table
  group_table = group_df.to_html(
      classes='table table-striped thin-border', index=False, escape=False
  )

  # Append a header for the street name followed by its table
  tables_html += f"""
    <h3>{street_name}</h3>
    {group_table}
    <br>
    """

# 5. Wrap everything in a div with margin-top: 20px and add CSS for thin borders
wrapped_html = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Committee Pages - El Paso Links</title>
    <style>
        table.thin-border {{
            border-collapse: collapse;
            width: 100%;
            margin-bottom: 20px;
        }}
        table.thin-border th, table.thin-border td {{
            border: 1px solid #ccc;
            padding: 8px;
        }}
        h3 {{
            color: #336488;
            margin-top: 20px;
            margin-bottom: 10px;
        }}
    </style>
</head>
<body>
    <div style="margin-top: 20px;">
        {tables_html}
    </div>
</body>
</html>
"""

# 6. Ensure output directory exists and save timestamped output file
output_dir = '../output'
os.makedirs(output_dir, exist_ok=True)

timestamp_str = datetime.now().strftime('_%Y%m%d_%H%M%S')
output_file = os.path.join(
    output_dir, f'committeePage-elpaso-links{timestamp_str}.html'
)

with open(output_file, 'w', encoding='utf-8') as f:
  f.write(wrapped_html)

print(
    f'Success! Converted {csv_file} and saved grouped tables to {output_file}'
)
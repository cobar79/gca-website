import argparse
from datetime import datetime
import os
import numpy as np
import pandas as pd
import re

# Set up command-line argument parsing with parameters
parser = argparse.ArgumentParser(
    description='Reconcile Balance Bookeeping master file with GCA import template.'
)
parser.add_argument(
    '-m',
    '--master',
    type=str,
    help='Path to the Balance Bookeeping master bookkeeping file (.xls/.xlsx)',
)
parser.add_argument(
    '-g', '--gca', type=str, help='Path to the GCA export file (.csv)'
)
parser.add_argument(
    '-t',
    '--template',
    type=str,
    default='../downloads/importusers.csv',
    help='Path to the official import template CSV',
)
parser.add_argument(
    '-D',
    '--debug',
    action='store_true',
    help='Enable address matching debug prints',
)
args = parser.parse_args()
DEBUG_MODE = args.debug


# Helper to automatically resolve paths into ../downloads/ if not found locally
def resolve_input_path(user_input, default_dir='../downloads'):
  if not user_input:
    return ''
  cleaned = str(user_input).strip().strip('"\'')
  # If path exists as-is, use it
  if os.path.exists(cleaned):
    return cleaned
  # Otherwise, check inside the downloads directory
  download_path = os.path.join(default_dir, cleaned)
  if os.path.exists(download_path):
    return download_path
  # Fallback to original user input so standard FileNotFoundError triggers cleanly
  return cleaned


# Fallback to interactive prompts if parameters are not provided via CLI
master_file = args.master
if not master_file:
  master_file = input(
      'Enter path to Balance Bookkeeping master bookkeeping file: '
  )
master_file = resolve_input_path(master_file)

gca_file = args.gca
if not gca_file:
  gca_file = input('Enter path to GCA export file (.csv): ')
gca_file = resolve_input_path(gca_file)

import_template_file = resolve_input_path(args.template)
if not os.path.exists(import_template_file):
  user_template = input(
      f'Default template not found at {import_template_file}. Enter path to'
      ' import template CSV: '
  )
  import_template_file = resolve_input_path(user_template)

# Validate files exist before proceeding
for f_path, f_label in [
    (master_file, 'Bookkeeping master file'),
    (gca_file, 'GCA export file'),
    (import_template_file, 'Import template file'),
]:
  if not os.path.exists(f_path):
    raise FileNotFoundError(f'{f_label} not found at: {f_path}')

# Load files
xl = pd.ExcelFile(master_file)
master_df = pd.read_excel(xl, sheet_name=xl.sheet_names[0])
gca_df = pd.read_csv(gca_file)
template_df = pd.read_csv(import_template_file)

# Get official import columns from the template
import_columns = list(template_df.columns)


# Helper function to find columns flexibly
def find_column(df, keyword):
  for col in df.columns:
    if keyword.lower() in str(col).lower():
      return col
  return None


# Identify the correct Primary User column name in the template
col_primary_user = find_column(template_df, 'primary user')
if not col_primary_user:
  col_primary_user = 'Primary User'
  if col_primary_user not in import_columns:
    import_columns.append(col_primary_user)

# Get current date formatted for new records
today_str = datetime.today().strftime('%m/%d/%Y')

col_status = find_column(gca_df, 'status')
col_level = find_column(gca_df, 'user level') or find_column(gca_df, 'level')

if len(gca_df.columns) > 12:
  col_gca_mailing = gca_df.columns[12]
else:
  raise IndexError(
      'GCA file does not have enough columns to access Column M (index 12).'
  )

# Filter out Expired and Level 3 users from GCA active consideration
gca_active = gca_df[
    (gca_df[col_status].astype(str).str.lower() == 'active')
    & (gca_df[col_level] != 3)
].copy()


# Helper function to clean street numbers
def clean_street(val):
  if pd.isna(val):
    return ''
  try:
    return str(int(float(val)))
  except:
    return str(val).strip()


master_df['Street_Num'] = master_df['Street #'].apply(clean_street)
master_df['Full_Address'] = (
    master_df['Street_Num'] + ' ' + master_df['Address 1'].astype(str)
).str.strip()

# Balance Bookeeping MAILING ADDRESS FROM COLUMNS F & G (Indices 5 and 6)
col_f = master_df.columns[5] if len(master_df.columns) > 5 else ''
col_g = master_df.columns[6] if len(master_df.columns) > 6 else ''

master_df['CK_Street_Num'] = master_df[col_f].apply(clean_street)
master_df['Balance Bookeeping_Mailing'] = (
    master_df['CK_Street_Num']
    + ' '
    + master_df[col_g].astype(str).str.strip()
).str.strip()


# Robust Address Normalization
def normalize_base_address(addr):
  if pd.isna(addr):
    return ''
  addr_str = str(addr)
  first_part = re.split(r'[,^\n\r]+', addr_str)[0].strip().lower()
  first_part = re.sub(r'[^\w\s]', '', first_part)

  suffixes = [
      'drive',
      'dr',
      'way',
      'wy',
      'street',
      'st',
      'court',
      'ct',
      'place',
      'pl',
      'avenue',
      'ave',
      'circle',
      'cir',
      'boulevard',
      'blvd',
      'road',
      'rd',
      'lane',
      'ln',
  ]

  words = first_part.split()
  words = [w for w in words if w not in suffixes]
  return ' '.join(words)


master_df['clean_base_address'] = master_df['Balance Bookeeping_Mailing'].apply(
    normalize_base_address
)
gca_active['clean_base_address'] = gca_active[col_gca_mailing].apply(
    normalize_base_address
)

delete_list = []
add_list = []
update_list = []
change_log = []

gca_groups = dict(list(gca_active.groupby('clean_base_address')))


def format_delimited_address(m_row):
  full_addr = m_row['Full_Address']
  city = str(m_row.get('City', '')).strip()
  state = str(m_row.get('State', 'Color')).strip()
  zip_val = int(m_row['Zip Code']) if pd.notna(m_row['Zip Code']) else ''

  raw_lines = f'{full_addr}\n{city}\n{state}\n{zip_val}'
  lines = [l.strip() for l in raw_lines.split('\n') if l.strip()]

  if not lines:
    return ''
  if len(lines) == 1:
    return lines[0] + '^'

  res = lines[0] + '^^' + lines[1]
  for l in lines[2:]:
    res += '^' + l

  if not res.endswith('^'):
    res += '^'
  return res


# Helper to build a clean record matching the official import template
def build_import_record(
    first_name,
    last_name,
    email,
    bookkeeping_acct,
    user_level,
    address_str,
    primary_user_name=None,
    status='Active',
    access_level='MEMBER',
    join_date=today_str,
    membership_type_id=36742,
):
  rec = {col: np.nan for col in import_columns}

  rec['First Name'] = first_name
  rec['Last Name'] = last_name
  rec['Email Address'] = email if pd.notna(email) and email else np.nan
  rec['Bookkeeping Account'] = bookkeeping_acct
  rec['User Level'] = user_level
  rec['Address'] = address_str
  rec['Status'] = status
  rec['Access Level'] = access_level
  rec['Join Date'] = join_date
  rec['Membership Type ID'] = membership_type_id

  if user_level == 2 and primary_user_name:
    rec[col_primary_user] = primary_user_name

  # Username rule: Email if it exists, otherwise firstname.lastname
  if pd.notna(email) and str(email).strip():
    rec['Username'] = str(email).strip()
  else:
    f_clean = str(first_name).strip().lower()
    l_clean = str(last_name).strip().lower()
    rec['Username'] = f'{f_clean}.{l_clean}'

  return rec


for idx, m_row in master_df.iterrows():
  c_addr = m_row['clean_base_address']
  if not c_addr:
    continue

  m_f1 = str(m_row.get('First Name', '')).strip()
  m_l1 = str(m_row.get('Last Name', '')).strip()

  raw_email = str(m_row.get('Email', '')).strip()
  m_email = (
      re.split(r'[;,]', raw_email)[0].strip() if pd.notna(raw_email) else ''
  )
  if m_email.lower() == 'nan':
    m_email = ''

  m_acct = str(m_row.get('Account #', m_row.get('Account', ''))).strip()
  if m_acct.lower() == 'nan':
    m_acct = ''

  m_f2 = str(m_row.get('2nd First Name', '')).strip()
  m_l2 = str(m_row.get('2nd Last Name', '')).strip()
  has_p2 = bool(
      m_f2 and m_f2.lower() != 'nan' and m_l2 and m_l2.lower() != 'nan'
  )

  gca_sub = gca_groups.get(c_addr, pd.DataFrame())
  delimited_addr = format_delimited_address(m_row)

  # Find Level 1 Full Name ("First Last") from GCA if it already exists for this address group
  p1_name = None
  if len(gca_sub) > 0:
    g_lvl1_check = gca_sub[gca_sub[col_level] == 1]
    if len(g_lvl1_check) > 0:
      g1_row = g_lvl1_check.iloc[0]
      p1_name = (
          f"{str(g1_row.get('First Name', '')).strip()} {str(g1_row.get('Last Name', '')).strip()}"
          .strip()
      )

  # If not in GCA yet, the Level 1 person we are adding in this run will be m_f1 m_l1
  if not p1_name:
    p1_name = f'{m_f1} {m_l1}'.strip()

  if len(gca_sub) == 0:
    change_log.append({'Reason': 'Address not in GCA'})

    add_list.append(
        build_import_record(
            m_f1,
            m_l1,
            m_email,
            m_acct,
            1,
            delimited_addr,
            primary_user_name=None,
            join_date=today_str,
        )
    )
    if has_p2:
      add_list.append(
          build_import_record(
              m_f2,
              m_l2,
              '',
              m_acct,
              2,
              delimited_addr,
              primary_user_name=p1_name,
              join_date=today_str,
          )
      )
  else:
    gca_lvl1 = gca_sub[gca_sub[col_level] == 1]
    gca_lvl2 = gca_sub[gca_sub[col_level] == 2]

    has_g1 = len(gca_lvl1) > 0
    has_g2 = len(gca_lvl2) > 0

    g_f1 = (
        str(gca_lvl1.iloc[0].get('First Name', '')).strip() if has_g1 else ''
    )
    g_l1 = str(gca_lvl1.iloc[0].get('Last Name', '')).strip() if has_g1 else ''
    g_email1 = (
        str(gca_lvl1.iloc[0].get('Email Address', '')).strip()
        if has_g1
        else ''
    )

    g_f2 = (
        str(gca_lvl2.iloc[0].get('First Name', '')).strip() if has_g2 else ''
    )
    g_l2 = str(gca_lvl2.iloc[0].get('Last Name', '')).strip() if has_g2 else ''

    match_p1 = has_g1 and (
        (g_f1.lower() == m_f1.lower() and g_l1.lower() == m_l1.lower())
        or (m_email and g_email1.lower() == m_email.lower())
    )

    if has_p2:
      match_p2 = (
          has_g2
          and g_f2.lower() == m_f2.lower()
          and g_l2.lower() == m_l2.lower()
      )
    else:
      match_p2 = not has_g2

    if match_p1 and match_p2:
      change_log.append({'Reason': 'Exact match'})
      for _, r in gca_sub.iterrows():
        r_dict = r.to_dict()
        upd = {col: r_dict.get(col, np.nan) for col in import_columns}
        upd['Bookkeeping Account'] = m_acct
        upd['Address'] = delimited_addr
        if upd.get('User Level') == 2:
          upd[col_primary_user] = p1_name
        update_list.append(upd)
    else:
      reasons = []
      if not match_p1:
        reasons.append('Level 1 Details Change')
      if not match_p2:
        if has_p2 and not has_g2:
          reasons.append('Added Level 2 Secondary Person')
        elif not has_p2 and has_g2:
          reasons.append('Removed Level 2 Secondary Person')
        else:
          reasons.append('Level 2 Name Change')

      change_log.append({
          'Reason': (
              ', '.join(reasons)
              if reasons
              else 'Property Match with Updates'
          )
      })

      if has_g1:
        r1_dict = gca_lvl1.iloc[0].to_dict()
        upd1 = {col: r1_dict.get(col, np.nan) for col in import_columns}
        upd1['First Name'] = m_f1
        upd1['Last Name'] = m_l1
        upd1['Email Address'] = m_email
        upd1['Bookkeeping Account'] = m_acct
        upd1['User Level'] = 1
        upd1['Address'] = delimited_addr
        if col_primary_user in upd1:
          upd1[col_primary_user] = np.nan
        if not upd1.get('Username'):
          upd1['Username'] = m_email if m_email else f'{m_f1.lower()}.{m_l1.lower()}'
        update_list.append(upd1)
      else:
        add_list.append(
            build_import_record(
                m_f1,
                m_l1,
                m_email,
                m_acct,
                1,
                delimited_addr,
                primary_user_name=None,
                join_date=today_str,
            )
        )

      if has_p2:
        if has_g2:
          r2_dict = gca_lvl2.iloc[0].to_dict()
          upd2 = {col: r2_dict.get(col, np.nan) for col in import_columns}
          upd2['First Name'] = m_f2
          upd2['Last Name'] = m_l2
          upd2['Email Address'] = np.nan
          upd2['Bookkeeping Account'] = m_acct
          upd2['User Level'] = 2
          upd2['Address'] = delimited_addr
          upd2[col_primary_user] = p1_name
          if not upd2.get('Username'):
            upd2['Username'] = f'{m_f2.lower()}.{m_l2.lower()}'
          update_list.append(upd2)
        else:
          add_list.append(
              build_import_record(
                  m_f2,
                  m_l2,
                  '',
                  m_acct,
                  2,
                  delimited_addr,
                  primary_user_name=p1_name,
                  join_date=today_str,
              )
          )
      else:
        if has_g2:
          for _, r in gca_lvl2.iterrows():
            delete_list.append(r.to_dict())

# Convert lists to DataFrames strictly formatted to import template columns
delete_df = pd.DataFrame(delete_list, columns=import_columns)
update_df = pd.DataFrame(update_list, columns=import_columns)
add_df = pd.DataFrame(add_list, columns=import_columns)

# Ensure the output directory exists
output_dir = '../output'
os.makedirs(output_dir, exist_ok=True)

# Generate timestamp string (e.g., _20261003_155005)
timestamp_str = datetime.now().strftime('_%Y%m%d_%H%M%S')

# Save outputs as CSV (.csv) files with timestamp in the output folder
delete_df.to_csv(
    os.path.join(output_dir, f'gca_records_to_delete{timestamp_str}.csv'),
    index=False,
)
update_df.to_csv(
    os.path.join(output_dir, f'gca_records_to_update{timestamp_str}.csv'),
    index=False,
)
add_df.to_csv(
    os.path.join(output_dir, f'gca_records_to_add{timestamp_str}.csv'),
    index=False,
)

# Print Summary Breakdown
df_log = pd.DataFrame(change_log)
print('=== CHANGE REASONS BREAKDOWN ===')
if not df_log.empty and 'Reason' in df_log.columns:
  print(df_log['Reason'].value_counts())
else:
  print('No changes recorded.')

print('\n=== FILE TOTALS ===')
print(f'Add:    {len(add_df)}')
print(f'Update: {len(update_df)}')
print(f'Delete: {len(delete_df)}')
print(f'\nOutput files saved successfully to directory: {output_dir}')
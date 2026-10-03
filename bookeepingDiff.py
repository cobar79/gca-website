import argparse
from datetime import datetime
import numpy as np
import pandas as pd
import re

# Set up command-line argument parsing
parser = argparse.ArgumentParser(
    description='Reconcile CompKeeper master file with GCA export.'
)
parser.add_argument(
    '-D',
    '--debug',
    action='store_true',
    help='Enable address matching debug prints',
)
args = parser.parse_args()
DEBUG_MODE = args.debug

# Load files
master_file = 'comp-keep-2006-9-25.xls'
gca_file = 'gca-export-all-users.csv'

xl = pd.ExcelFile(master_file)
master_df = pd.read_excel(xl, sheet_name=xl.sheet_names[0])
gca_df = pd.read_csv(gca_file)

# Get current date formatted for new records
today_str = datetime.today().strftime('%m/%d/%Y')


# Helper function to find columns flexibly
def find_column(df, keyword):
  for col in df.columns:
    if keyword.lower() in str(col).lower():
      return col
  return None


col_status = find_column(gca_df, 'status')
col_level = find_column(gca_df, 'user level') or find_column(gca_df, 'level')

# Explicitly use Column M (index 12) for GCA Mailing Address per user confirmation
if len(gca_df.columns) > 12:
  col_gca_mailing = gca_df.columns[12]
else:
  raise IndexError(
      'GCA file does not have enough columns to access Column M (index 12).'
  )

col_prop = (
    find_column(gca_df, 'property name')
    or find_column(gca_df, 'property')
    or find_column(gca_df, 'address')
)

# --- STARTUP DIAGNOSTIC PRINT ---
print('=== COLUMN DETECTION DIAGNOSTICS ===')
print(f'GCA Status Column     : {col_status}')
print(f'GCA User Level Column : {col_level}')
print(f'GCA Mailing Address   : {col_gca_mailing} (Column M)')
print(f'GCA Total Rows Loaded : {len(gca_df)}')
print(f'Master Total Rows     : {len(master_df)}')
print('====================================\n')

if not col_status or not col_level or not col_gca_mailing:
  raise KeyError(
      'Could not automatically locate required GCA columns (Status, User Level).'
  )

# Filter out Expired and Level 3 users from GCA active consideration
gca_active = gca_df[
    (gca_df[col_status].astype(str).str.lower() == 'active')
    & (gca_df[col_level] != 3)
].copy()


# Helper function to clean street numbers (handles Excel float conversion like 105.0 -> 105)
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

# --- COMPKEEPER MAILING ADDRESS FROM COLUMNS F & G (Indices 5 and 6) ---
col_f = master_df.columns[5] if len(master_df.columns) > 5 else ''
col_g = master_df.columns[6] if len(master_df.columns) > 6 else ''

# Clean column F street number to prevent floating-point ".0" issues
master_df['CK_Street_Num'] = master_df[col_f].apply(clean_street)
master_df['CompKeeper_Mailing'] = (
    master_df['CK_Street_Num']
    + ' '
    + master_df[col_g].astype(str).str.strip()
).str.strip()


# Robust Address Normalization that splits on commas, carets, or newlines to isolate the street address
def normalize_base_address(addr):
  if pd.isna(addr):
    return ''
  addr_str = str(addr)
  # Split by comma, caret (^), or newline to get just the first line (street address)
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


master_df['clean_base_address'] = master_df['CompKeeper_Mailing'].apply(
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

# Diagnostic check on group dictionary size
print(f'Unique GCA base address groups created: {len(gca_groups)}')
if len(gca_groups) > 0:
  sample_keys = list(gca_groups.keys())[:5]
  print(f'Sample GCA normalized keys: {sample_keys}\n')


# Caret-delimited Address formatter for the 'Address' column
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

  p1_name_full = f'{m_f1} {m_l1}'.strip()
  gca_sub = gca_groups.get(c_addr, pd.DataFrame())

  # --- GATED DEBUG PRINT (ASCII Safe) ---
  if DEBUG_MODE:
    raw_ck_mailing = m_row['CompKeeper_Mailing']
    match_status = 'MATCH FOUND' if len(gca_sub) > 0 else 'MISS (Added)'
    print(
        f'[CK Mailing]: "{raw_ck_mailing}" (Cleaned: "{c_addr}") ---> GCA'
        f' Lookup: {match_status}'
    )
    if len(gca_sub) > 0:
      sample_gca_raw = str(gca_sub.iloc[0].get(col_gca_mailing, ''))
      print(
          f'   -> Matched GCA Raw: "{sample_gca_raw}" | Group Count:'
          f' {len(gca_sub)}'
      )
  # ------------------------------------

  zip_val = int(m_row['Zip Code']) if pd.notna(m_row['Zip Code']) else ''
  city = str(m_row.get('City', '')).strip()
  state = str(m_row.get('State', 'Color')).strip()
  full_addr = m_row['Full_Address']

  delimited_addr = format_delimited_address(m_row)
  prop_addr_display = (
      f'{full_addr}, {city}, {state} {zip_val}'
      if zip_val
      else f'{full_addr}, {city}, {state}'
  )

  if len(gca_sub) == 0:
    change_log.append({'Reason': 'Address not in GCA'})

    template = gca_df.iloc[0].copy()
    template['First Name'] = m_f1
    template['Last Name'] = m_l1
    template['Email Address'] = m_email
    template['Bookeeping Acct'] = m_acct
    template[col_level] = 1
    template['Primary User'] = np.nan
    template[col_gca_mailing] = prop_addr_display
    if col_prop and col_prop in template:
      template[col_prop] = prop_addr_display
    template['Address'] = delimited_addr
    template['Property Group'] = 'Property in good standing'
    template['Date Added'] = today_str
    add_list.append(template.to_dict())

    if has_p2:
      template2 = template.copy()
      template2['First Name'] = m_f2
      template2['Last Name'] = m_l2
      template2['Email Address'] = np.nan
      template2['Bookeeping Acct'] = m_acct
      template2[col_level] = 2
      template2['Primary User'] = p1_name_full
      template2['Date Added'] = today_str
      add_list.append(template2.to_dict())
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
        r_dict['Bookeeping Acct'] = m_acct
        r_dict['Address'] = delimited_addr
        r_dict[col_gca_mailing] = prop_addr_display
        update_list.append(r_dict)
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
        r1 = gca_lvl1.iloc[0].copy()
        r1['First Name'] = m_f1
        r1['Last Name'] = m_l1
        r1['Email Address'] = m_email
        r1['Bookeeping Acct'] = m_acct
        r1[col_level] = 1
        r1['Primary User'] = np.nan
        r1['Address'] = delimited_addr
        r1[col_gca_mailing] = prop_addr_display
        update_list.append(r1.to_dict())
      else:
        template = gca_sub.iloc[0].copy()
        template['First Name'] = m_f1
        template['Last Name'] = m_l1
        template['Email Address'] = m_email
        template['Bookeeping Acct'] = m_acct
        template[col_level] = 1
        template['Primary User'] = np.nan
        template[col_gca_mailing] = prop_addr_display
        template['Address'] = delimited_addr
        template['Date Added'] = today_str
        add_list.append(template.to_dict())

      if has_p2:
        if has_g2:
          r2 = gca_lvl2.iloc[0].copy()
          r2['First Name'] = m_f2
          r2['Last Name'] = m_l2
          r2['Email Address'] = np.nan
          r2['Bookeeping Acct'] = m_acct
          r2[col_level] = 2
          r2['Primary User'] = p1_name_full
          r2['Address'] = delimited_addr
          r2[col_gca_mailing] = prop_addr_display
          update_list.append(r2.to_dict())
        else:
          template2 = gca_sub.iloc[0].copy()
          template2['First Name'] = m_f2
          template2['Last Name'] = m_l2
          template2['Email Address'] = np.nan
          template2['Bookeeping Acct'] = m_acct
          template2[col_level] = 2
          template2['Primary User'] = p1_name_full
          template2[col_gca_mailing] = prop_addr_display
          template2['Address'] = delimited_addr
          template2['Date Added'] = today_str
          add_list.append(template2.to_dict())
      else:
        if has_g2:
          for _, r in gca_lvl2.iterrows():
            delete_list.append(r.to_dict())

# Save outputs as Excel (.xlsx) files with Column L (index 11) safety stamp
delete_df = pd.DataFrame(delete_list)
if not delete_df.empty and delete_df.shape[1] > 11:
  delete_df.iloc[:, 11] = today_str
delete_df.to_excel('gca_records_to_delete.xlsx', index=False)

update_df = pd.DataFrame(update_list)
if not update_df.empty and update_df.shape[1] > 11:
  update_df.iloc[:, 11] = today_str
update_df.to_excel('gca_records_to_update.xlsx', index=False)

add_df = pd.DataFrame(add_list)
if not add_df.empty and add_df.shape[1] > 11:
  add_df.iloc[:, 11] = today_str
add_df.to_excel('gca_records_to_add.xlsx', index=False)

# Calculate Add counts broken down by Level 1 and Level 2
add_l1 = (
    len(add_df[add_df[col_level] == 1])
    if not add_df.empty and col_level in add_df.columns
    else 0
)
add_l2 = (
    len(add_df[add_df[col_level] == 2])
    if not add_df.empty and col_level in add_df.columns
    else 0
)

# Print Summary Breakdown
df_log = pd.DataFrame(change_log)
print('=== CHANGE REASONS BREAKDOWN ===')
print(df_log['Reason'].value_counts())
print('\n=== FILE TOTALS ===')
print(
    f'Add:    {len(add_list)} (Level 1: {add_l1}, Level 2: {add_l2})'
)
print(f'Update: {len(update_list)}')
print(f'Delete: {len(delete_list)}')

from datetime import datetime
import numpy as np
import pandas as pd
import re

# Load files (Update master filename as needed)
master_file = 'comp-keep-2006-9-25.xls'
gca_file = 'gca-export-all-users.csv'

xl = pd.ExcelFile(master_file)
master_df = pd.read_excel(xl, sheet_name=xl.sheet_names[0])
gca_df = pd.read_csv(gca_file)

# Get current date formatted as MM/DD/YYYY for new records
today_str = datetime.today().strftime('%m/%d/%Y')

# Filter out Expired and Level 3 users from GCA active consideration
gca_active = gca_df[
    (gca_df['Status'].astype(str).str.lower() == 'active')
    & (gca_df['User Level'] != 3)
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


# Robust Address Normalization
def normalize_base_address(addr):
  if pd.isna(addr):
    return ''
  addr_first = str(addr).split('\r\n')[0].split('\n')[0].strip().lower()
  addr_first = re.sub(r'[^\w\s]', '', addr_first)

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
  ]

  words = addr_first.split()
  words = [w for w in words if w not in suffixes]
  return ' '.join(words)


master_df['clean_base_address'] = master_df['Full_Address'].apply(
    normalize_base_address
)
gca_active['clean_base_address'] = gca_active['Property Name'].apply(
    normalize_base_address
)

delete_list = []
add_list = []
update_list = []
change_log = []

gca_groups = dict(list(gca_active.groupby('clean_base_address')))

for idx, m_row in master_df.iterrows():
  c_addr = m_row['clean_base_address']
  if not c_addr:
    continue

  m_f1 = str(m_row.get('First Name', '')).strip()
  m_l1 = str(m_row.get('Last Name', '')).strip()

  # Email Fix: Take primary email if semicolon-separated
  raw_email = str(m_row.get('Email', '')).strip()
  m_email = (
      re.split(r'[;,]', raw_email)[0].strip() if pd.notna(raw_email) else ''
  )
  if m_email.lower() == 'nan':
    m_email = ''

  # --- ACCOUNT FIX: Grab directly from 'Account #' and preserve alphanumeric codes ---
  m_acct = str(m_row.get('Account #', m_row.get('Account', ''))).strip()
  if m_acct.lower() == 'nan':
    m_acct = ''
  # ---------------------------------------------------------------------------------

  m_f2 = str(m_row.get('2nd First Name', '')).strip()
  m_l2 = str(m_row.get('2nd Last Name', '')).strip()
  has_p2 = bool(
      m_f2 and m_f2.lower() != 'nan' and m_l2 and m_l2.lower() != 'nan'
  )

  p1_name_full = f'{m_f1} {m_l1}'.strip()

  gca_sub = gca_groups.get(c_addr, pd.DataFrame())

  if len(gca_sub) == 0:
    change_log.append({'Reason': 'Address not in GCA'})

    template = gca_df.iloc[0].copy()
    template['First Name'] = m_f1
    template['Last Name'] = m_l1
    template['Email Address'] = m_email
    template['Bookeeping Acct'] = m_acct
    template['User Level'] = 1
    template['Primary User'] = np.nan
    template['Property Address'] = (
        f"{m_row['Full_Address']}\n{m_row['City']}, {m_row['State']}"
        f" {int(m_row['Zip Code']) if pd.notna(m_row['Zip Code']) else ''}"
    )
    template['Address'] = (
        f"{m_row['Full_Address']}^^{m_row['City']}^{m_row['State']}^{int(m_row['Zip Code']) if pd.notna(m_row['Zip Code']) else ''}^"
    )
    template['Property Group'] = 'Property in good standing'
    template['Date Added'] = today_str
    add_list.append(template.to_dict())

    if has_p2:
      template2 = template.copy()
      template2['First Name'] = m_f2
      template2['Last Name'] = m_l2
      template2['Email Address'] = np.nan
      template2['Bookeeping Acct'] = m_acct
      template2['User Level'] = 2
      template2['Primary User'] = p1_name_full
      template2['Date Added'] = today_str
      add_list.append(template2.to_dict())
  else:
    gca_lvl1 = gca_sub[gca_sub['User Level'] == 1]
    gca_lvl2 = gca_sub[gca_sub['User Level'] == 2]

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
        update_list.append(r_dict)
    else:
      reasons = []
      if not match_p1:
        if (
            g_f1.lower() != m_f1.lower()
            or g_l1.lower() != m_l1.lower()
            or g_email1.lower() != m_email.lower()
        ):
          reasons.append('Level 1 Details Change')
      if not match_p2:
        if has_p2 and not has_g2:
          reasons.append('Added Level 2 Secondary Person')
        elif not has_p2 and has_g2:
          reasons.append('Removed Level 2 Secondary Person')
        elif g_f2.lower() != m_f2.lower() or g_l2.lower() != m_l2.lower():
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
        r1['User Level'] = 1
        r1['Primary User'] = np.nan
        update_list.append(r1.to_dict())
      else:
        template = gca_sub.iloc[0].copy()
        template['First Name'] = m_f1
        template['Last Name'] = m_l1
        template['Email Address'] = m_email
        template['Bookeeping Acct'] = m_acct
        template['User Level'] = 1
        template['Primary User'] = np.nan
        template['Date Added'] = today_str
        add_list.append(template.to_dict())

      if has_p2:
        if has_g2:
          r2 = gca_lvl2.iloc[0].copy()
          r2['First Name'] = m_f2
          r2['Last Name'] = m_l2
          r2['Email Address'] = np.nan
          r2['Bookeeping Acct'] = m_acct
          r2['User Level'] = 2
          r2['Primary User'] = p1_name_full
          update_list.append(r2.to_dict())
        else:
          template2 = gca_sub.iloc[0].copy()
          template2['First Name'] = m_f2
          template2['Last Name'] = m_l2
          template2['Email Address'] = np.nan
          template2['Bookeeping Acct'] = m_acct
          template2['User Level'] = 2
          template2['Primary User'] = p1_name_full
          template2['Date Added'] = today_str
          add_list.append(template2.to_dict())
      else:
        if has_g2:
          for _, r in gca_lvl2.iterrows():
            delete_list.append(r.to_dict())

# Save outputs with Column L (index 11) safety stamp in MM/DD/YYYY format
delete_df = pd.DataFrame(delete_list)
if not delete_df.empty and delete_df.shape[1] > 11:
  delete_df.iloc[:, 11] = today_str
delete_df.to_csv('gca_records_to_delete.csv', index=False)

add_df = pd.DataFrame(add_list)
if not add_df.empty and add_df.shape[1] > 11:
  add_df.iloc[:, 11] = today_str
add_df.to_csv('gca_records_to_add.csv', index=False)

update_df = pd.DataFrame(update_list)
if not update_df.empty and update_df.shape[1] > 11:
  update_df.iloc[:, 11] = today_str
update_df.to_csv('gca_records_to_update.csv', index=False)

# Print Summary Breakdown
df_log = pd.DataFrame(change_log)
print('=== CHANGE REASONS BREAKDOWN ===')
print(df_log['Reason'].value_counts())
print('\n=== FILE TOTALS ===')
print(f'Add:    {len(add_list)}')
print(f'Update: {len(update_list)}')
print(f'Delete: {len(delete_list)}')

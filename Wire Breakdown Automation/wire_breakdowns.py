import pandas as pd
import os
import base64
from tkinter import filedialog, Tk
from email.message import EmailMessage

# google API libraries
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# excel styling libraries
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side

# define security permissions needed to make drafts
SCOPES = ['https://www.googleapis.com/auth/gmail.compose', 'https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/spreadsheets.readonly' ]

# paths and authentication
script_dir = os.path.dirname(os.path.abspath(__file__))
credentials_path = os.path.join(script_dir, 'credentials.json')
token_path = os.path.join(script_dir, 'token.json')

creds = None
# check if logged in previously
if os.path.exists(token_path):
    creds = Credentials.from_authorized_user_file(token_path, SCOPES)

# if no token, open the browser to log in
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
        creds = flow.run_local_server(port=0)
    # save token for next time
    with open(token_path, 'w') as token:
        token.write(creds.to_json())

# build official gmail service
gmail_service = build('gmail', 'v1', credentials=creds)

# grab email of current user
user_profile = gmail_service.users().getProfile(userId='me').execute()
user_email = user_profile.get('emailAddress', '').lower().strip()

# admin team emails
MARCOS_EMAIL = 'marcos.soares@simplicitygroup.com'
DAVID_EMAIL = 'david.campbell@simplicitygroup.com'
GABRIELLE_EMAIL = 'gabrielle.williams@simplicitygroup.com'

# Dynamically construct the CC list so users don't CC themselves
cc_set = set()

if user_email == MARCOS_EMAIL:
    cc_set.update([DAVID_EMAIL, GABRIELLE_EMAIL])
elif user_email == GABRIELLE_EMAIL:
    cc_set.update([DAVID_EMAIL, MARCOS_EMAIL])
elif user_email == DAVID_EMAIL:
    cc_set.update([GABRIELLE_EMAIL, MARCOS_EMAIL])
else:
    # fallback default if another coworker runs script
    cc_set.update([DAVID_EMAIL, GABRIELLE_EMAIL, MARCOS_EMAIL])

# convert emails into a clean comma separated string for email header
TEAM_CC = ", ".join(cc_set)

# user inputs
org_name = input("Enter the Organization: ")
process_date = input("Enter the Date (MMDDYYYY): ")

formatted_date = f"{process_date[0:2]}/{process_date[2:4]}/{process_date[4:8]}"

# dynamic loan tracking through google sheets
LOAN_SPREADSHEET_ID = "1LYzEnZK7qB6W6M_3qSTGuW0O5gNimMVss0olgvdCVOo"

COLUMN_ALIASES = {
    'Total Cont': 'Total Cont', 'Total Contribution': 'Total Cont', 'Contribution': 'Total Cont',
    'Policy Number': 'Policy Number', 'Policy': 'Policy Number',
    'Carrier': 'Carrier',
    'Carrier Email': 'Carrier Email', 'Email': 'Carrier Email',
    'Premium Type':'Premium Type', 'Status': 'Premium Type', 'Payment Type': 'Premium Type',
}

REQUIRED_COLUMNS = ['Policy Number', 'Carrier', 'Total Cont', 'Carrier Email']

# select file
root = Tk()
root.withdraw()
file_path = filedialog.askopenfilename(title=f"Select invoice for {org_name}")

if not file_path:
    print("No invoice selected. Exiting...")
    exit()

# select folder to upload wire breakdowns for backup
output_path = filedialog.askdirectory(title="Select folder to save wire breakdowns")
if not output_path:
    print("No folder selected. Exiting...")
    exit()

print(f"Files will be saved to {output_path}")

# fetch dynamic loan sheet data
print("Connecting to Google Sheets & fetching active loan policies...")

# build sheets API service and gather data
sheets_service = build('sheets', 'v4', credentials=creds)
sheet_range = "'Active Loan Tracking'!A:Z"
sheet_result = sheets_service.spreadsheets().values().get(
    spreadsheetId = LOAN_SPREADSHEET_ID,
    range=sheet_range
).execute()

rows = sheet_result.get('values', [])

if not rows:
    print("Warning: No data found in 'Active Loan Tracking' sheet.")
    loan_policies = set()
else:
    headers = [str(cell).strip() for cell in rows[0]]

    if 'Policy Number' not in headers:
        print("Error: Could not find 'Policy Number' column header in 'Loan Tracking' sheet.")
        exit()

    policy_col_idx = headers.index('Policy Number')

    # extract policies under column into a set for fast lookup
    loan_policies = set(
        str(row[policy_col_idx]).strip()
        for row in rows[1:]
        if len(row) > policy_col_idx and row[policy_col_idx]
    )

print(f"Successfully loaded {len(loan_policies)} active loan policies from Google Sheets.")

# process excel
df = pd.read_excel(file_path, sheet_name='Policy Recon', usecols= "A:H")
df.columns = df.columns.str.strip()

# rename alt headers to standard names and check for missing columns
df = df.rename(columns=COLUMN_ALIASES)
missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
if missing_cols:
    print("\n" + "="*60)
    print("Error: Missing required columns in 'Policy Recon' tab")
    print(f"The script could not find the following required column(s):")
    for col in missing_cols:
        print(f" - {col}")
    print("\nFound columns in sheet:", list(df.columns))
    print("="*60 + "\n")
    # stop script so invoice can be fixed
    input("Press any key to exit and fix the invoice file...")
          

df = df.dropna(subset=['Carrier', 'Total Cont'])
df['Policy Number'] = df['Policy Number'].astype(str).str.strip()


# new three step logic for premium type hierarchy
def determine_premium_type(row):
    status_val = str(row.get('Premium Type', '')).strip().lower()
    if status_val == 'new business premium':
        return 'New Business Premium'
    policy_num = str(row['Policy Number']).strip()
    if policy_num in loan_policies:
        return 'Loan Repayment'
    return 'Inforce Premium'
df['Premium Type'] = df.apply(determine_premium_type, axis=1)

carriers = df['Carrier'].unique()

for carrier in carriers:
    carrier_data = df[df['Carrier'] == carrier].copy()
    
    # calculate total on wire breakdown
    total_amount = carrier_data['Total Cont'].sum()
    formatted_total = f"${total_amount:,.2f}"
    
    recipient_email = str(carrier_data['Carrier Email'].iloc[0]).strip()

    # check if email is valid
    if "@" not in recipient_email or "." not in recipient_email:
        print(f"Skipped Gmail draft for {carrier}: Found an invalid email address for ('{recipient_email})'. Please check 'Carrier Directory' sheet in invoice file.")
        continue
    
    # naming file
    attachment_name = f"{org_name} - {carrier} Wire Breakdown {process_date}.xlsx"

    # create row to hold sum text and raw sum
    summary_row = pd.DataFrame([{'Policy Owner': 'TOTAL', 'Amount': total_amount}])

    carrier_data = carrier_data.rename(columns={'Total Cont': 'Amount'})
    carrier_data = carrier_data.rename(columns={'Participant': 'Policy Owner'})

    # create carrier specific breakdown file
    columns_to_export = ['Policy Owner', 'Carrier', 'Policy Number', 'Premium Type', 'Amount']
    export_df = pd.concat([carrier_data[columns_to_export], summary_row], ignore_index=True)
    
    # defining file name for wire breakdown folder path
    file_name = f"{org_name} - {carrier} Wire Breakdown {process_date}.xlsx"
    full_save_path = os.path.join(output_path, file_name)

    export_df.to_excel(full_save_path, index=False)
    
    wb = load_workbook(full_save_path)
    ws = wb.active
    
    # finds exactly where last row is
    total_row_idx = ws.max_row
    
    # define bold font & accounting style underline border
    bold_font = Font(name='Calibri', size=11, bold=True)
    
    accounting_border = Border(
        top=Side(style='double'),
        bottom=Side(style='thin')
    )
    
    # formatter configuration for Column D (amount)
    for row in range(2, total_row_idx):
        ws[f'E{row}'].number_format = '$#,##0.00'

    # styles cell D (last row) to be bold and underlined
    total_amount_cell = ws[f'E{total_row_idx}']
    total_amount_cell.font = bold_font
    total_amount_cell.border = accounting_border
    total_amount_cell.number_format = '$#,##0.00'
    
    # styles cell A[last row] ('TOTAL:') as bold
    ws[f'A{total_row_idx}'].font = bold_font
    
    # save workbook changes
    wb.save(full_save_path)
    wb.close()
    
    # create email
    msg = EmailMessage()
    msg['Subject'] = f"{org_name} - {carrier} {formatted_date} Wire Breakdown"
    msg['To'] = recipient_email
    msg['Cc'] = TEAM_CC
    
    body_lines= [
        f"Hi team,",
        f"Attached is the breakdown of the wire remitted on {formatted_date} by {org_name} in the amount of {formatted_total}.",
        f"Please apply premiums according to the attached breakdown. Let me know if you have any questions or need any additional information. Thank you!"
    ]

    plain_text_body = "\n\n".join(body_lines)
    msg.set_content(plain_text_body)

    # converts plain text to HTML to avoid paragraph breaks
    html_body = "".join([f"<p>{line}</p>" for line in body_lines])
    msg.add_alternative(html_body, subtype='html')
    
    # attach excel workbook securely
    with open(full_save_path, 'rb') as f:
        file_data = f.read()
        msg.add_attachment(
            file_data, 
            maintype='application', 
            subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
            filename=file_name
        )
    
    # encode email for Google's server
    raw_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    
    # gmail draft
    try:
        gmail_service.users().drafts().create(
            userId='me',
            body={'message': {'raw': raw_message}}
        ).execute()
        print(f"Draft created: {attachment_name} | Total: {formatted_total}")
    except Exception as e:
        print(f"Failed to create draft for {carrier}: {e}")

print(f"\nSuccess! All {org_name} breakdowns are officially waiting in your Gmail drafts.")
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
SCOPES = ['https://www.googleapis.com/auth/gmail.compose']

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

# user inputs
org_name = input("Enter the Organization: ")
process_date = input("Enter the Date (MMDDYYYY): ")

formatted_date = f"{process_date[0:2]}/{process_date[2:4]}/{process_date[4:8]}"
TEAM_CC = 'david.campbell@simplicitygroup.com, gabrielle.williams@simplicitygroup.com'

# select file
root = Tk()
root.withdraw()
file_path = filedialog.askopenfilename(title=f"Select invoice for {org_name}")

if not file_path:
    print("No invoice selected. Exiting...")
    exit()

loan_path = filedialog.askopenfilename(title="Select loan workbook")

if not loan_path:
    print("No loan workbook selected. Exiting...")
    exit()

# load sheets from loan workbook and extract policy numbers
loan_excel = pd.ExcelFile(loan_path)
sheet_names = loan_excel.sheet_names

loan_policies = set()
for sheet in sheet_names[:2]:
    loan_df = pd.read_excel(loan_path, sheet_name=sheet)
    # grabs policies with loans and matches accordingly with payment type
    if 'Policy Number' in loan_df.columns:
        policies= loan_df['Policy Number'].dropna().astype(str).str.strip().tolist()
        loan_policies.update(policies)

# process excel
df = pd.read_excel(file_path, sheet_name='Wire_Breakdown_Export')
df['Policy Number'] = df['Policy Number'].astype(str).str.strip()
df['Payment Type'] = df['Policy Number'].apply(lambda x: 'Loan Repayment' if x in loan_policies else 'Premium')
carriers = df['Carrier'].unique()

for carrier in carriers:
    carrier_data = df[df['Carrier'] == carrier].copy()
    
    # calculate total on wire breakdown
    total_amount = carrier_data['Amount'].sum()
    formatted_total = f"${total_amount:,.2f}"
    
    recipient_email = carrier_data['Carrier Email'].iloc[0]
    
    # naming file
    attachment_name = f"{org_name} - {carrier} Wire Breakdown {process_date}.xlsx"

    # create row to hold sum text and raw sum
    summary_row = pd.DataFrame([{'Policy Owner': 'TOTAL', 'Amount': total_amount}])
    
    # create carrier specific breakdown file
    columns_to_export = ['Policy Owner', 'Carrier', 'Policy Number', 'Payment Type', 'Amount']
    export_df = pd.concat([carrier_data[columns_to_export], summary_row], ignore_index=True)
    
    # 3. write basic data to Excel
    export_df.to_excel(attachment_name, index=False)
    
    wb = load_workbook(attachment_name)
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
    wb.save(attachment_name)
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
    with open(attachment_name, 'rb') as f:
        file_data = f.read()
        msg.add_attachment(
            file_data, 
            maintype='application', 
            subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
            filename=attachment_name
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
        
    # cleanup temp file
    os.remove(attachment_name)

print(f"\nSuccess! All {org_name} breakdowns are officially waiting in your Gmail drafts.")
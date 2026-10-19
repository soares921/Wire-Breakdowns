import streamlit as st
from streamlit_oauth import OAuth2Component
import pandas as pd
import os
import base64
from email.message import EmailMessage
# google api libraries
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
# excel styling libraries
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side

### STREAMLIT PAGE SETUP
st.set_page_config(page_title="Wire Breakdown", layout="wide")

st.title("Wire Breakdown & Gmail Draft Generator")
st.caption("Version 3.0: Cloud Edition")

### CONSTANTS & CONFIGURATIONS
SCOPES = [
    'https://www.googleapis.com/auth/gmail.compose', 
    'https://www.googleapis.com/auth/gmail.readonly', 
    'https://www.googleapis.com/auth/spreadsheets.readonly'
]

LOAN_SPREADSHEET_ID = "1DNaUyxTbtLCaTNSb2Az4TnmeSlw_JMrmpSGKrCPcyoY"


DEFAULT_COLUMN_ALIASES = {
    'Total Cont': 'Total Cont', 'Total Contribution': 'Total Cont', 'Contribution': 'Total Cont', 'Amount': 'Total Cont',
    'Policy Number': 'Policy Number', 'Policy': 'Policy Number', 'Policy #': 'Policy Number',
    'Carrier': 'Carrier', 'Company': 'Carrier',
    'Carrier Email': 'Carrier Email', 'Email': 'Carrier Email', 'Recipient Email': 'Carrier Email',
    'Premium Type': 'Premium Type', 'Status': 'Premium Type', 'Payment Type': 'Premium Type',
}

REQUIRED_CORE_FIELDS = ['Policy Number', 'Carrier', 'Total Cont']

CLIENT_ID = st.secrets["google"]["client_id"]
CLIENT_SECRET = st.secrets["google"]["client_secret"]
REDIRECT_URI = st.secrets["google"]["redirect_uri"]

AUTHORIZE_URL = "https://accounts.google.com/o/oauth2/v2/auth"
TOKEN_URL = "https://oauth2.googleapis.com/token"
REFRESH_TOKEN_URL = TOKEN_URL
REVOKE_TOKEN_URL = "https://oauth2.googleapis.com/revoke"

if "user_token" not in st.session_state:
    st.session_state.user_token = None

if 'generated_files_data' not in st.session_state:
    st.session_state.generated_files_data = []

# gmail authentication gateway (login check)
oauth2 = OAuth2Component(CLIENT_ID, CLIENT_SECRET, AUTHORIZE_URL, TOKEN_URL, REFRESH_TOKEN_URL, REVOKE_TOKEN_URL)

if not st.session_state.get("user_token"):
    st.info("Please log in with your Google account to generate Gmail drafts in your personal inbox.")
    
    result = oauth2.authorize_button(
        name="Log in with Google",
        icon="https://www.google.com/favicon.ico",
        redirect_uri=REDIRECT_URI,
        scope=" ".join(SCOPES),
        key="google_login",
        extras_params={"prompt": "consent", "access_type": "offline"}
    )
    
    if result and "token" in result:
        st.session_state.user_token = result["token"]
        st.rerun()
    else:
        st.stop()  # stops execution here until user logs in


# helper functions for user authenticated services
def get_user_credentials():
    # converts session_state oauth token dictionary into google credentials object
    token_data = st.session_state.user_token
    creds = Credentials(
        token=token_data.get("access_token"),
        refresh_token=token_data.get("refresh_token"),
        token_uri=TOKEN_URL,
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET,
        scopes=SCOPES
    )
    return creds

def get_user_gmail_service():
    """Builds Gmail service for the currently logged-in user."""
    return build('gmail', 'v1', credentials=get_user_credentials())

def get_shared_sheets_service():
    """Builds Google Sheets service using user's token."""
    return build('sheets', 'v4', credentials=get_user_credentials())


### SIDEBAR CONTROL PANEL
st.sidebar.title("Operations Panel")

# section 1: required execution input
st.sidebar.subheader("Processing Details")
org_name = st.sidebar.text_input("Organization Name", placeholder="e.g. Ohio Health")
process_date = st.sidebar.text_input("Process Date (MMDDYYYY)", max_chars=8, placeholder="")

st.sidebar.divider()

# section 2: admin quick links
st.sidebar.subheader("Admin Resources")

# section 3: mini instructions expander
with st.sidebar.expander("How to Use Automation"):
    st.write("""
    1. Enter **Organization Name** and **Process Date** above.
    2. Upload the invoice Excel file on the main screen.
    3. Confirm tab selection and column mappings.
    4. Click **Process Invoice** to generate Excel breakdowns & Gmail drafts.
    5. Download completed files from the bottom section.
    """)

# link directly to your loan tracking/email directory sheet
sheet_url = f"https://docs.google.com/spreadsheets/d/{LOAN_SPREADSHEET_ID}/edit"
st.sidebar.link_button("Open Loan Tracking Sheet", sheet_url)

### MAIN INTERFACE: FILE UPLOAD & SHEET SELECTOR
st.subheader("Upload & Select Workbook Tab")
uploaded_file = st.file_uploader(
    "Upload Invoice File", 
    type=['xlsx', 'xls']
)

selected_sheet = None
df_raw = None

if uploaded_file is not None:
    if st.session_state.get('last_uploaded') != uploaded_file.name:
        st.session_state.generated_files_data = []
        st.session_state.last_uploaded = uploaded_file.name
    excel_file = pd.ExcelFile(uploaded_file)
    sheet_names = excel_file.sheet_names
    
    default_index = sheet_names.index('Policy Recon') if 'Policy Recon' in sheet_names else 0
    selected_sheet = st.selectbox("Select Excel Sheet to Process:", sheet_names, index=default_index)
    
    # read raw header preview
    df_raw = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
    st.info(f"Loaded tab **'{selected_sheet}'** with **{len(df_raw)}** rows and **{len(df_raw.columns)}** columns.")


    ### DYNAMIC COLUMN MAPPING
    st.subheader("Confirm Column Mapping")
    st.caption("Match the columns from your Excel file to the required internal fields:")

    raw_columns = [str(c).strip() for c in df_raw.columns]
    
    # precalculate best guessed column matches
    col_mapping = {}
    col_cols = st.columns(4)
    
    target_fields = ['Policy Number', 'Carrier', 'Total Cont']
    
    for idx, field in enumerate(target_fields):
        guessed_col = None
        for col in raw_columns:
            if DEFAULT_COLUMN_ALIASES.get(col, col) == field:
                guessed_col = col
                break
        
        options = ["-- Select Column --"] + raw_columns
        selected_index = options.index(guessed_col) if guessed_col in raw_columns else 0
        
        with col_cols[idx]:
            chosen = st.selectbox(f"Field: **{field}**", options, index=selected_index)
            if chosen != "-- Select Column --":
                col_mapping[chosen] = field

def find_carrier_contact(carrier_name, directory):
    
    # helper function that looks up contact info for a carrier using exact matching first, then falls back to substring matching
    clean_name = str(carrier_name).strip().lower()
    
    if clean_name in directory:
        return directory[clean_name]
    
    for dir_carrier, contact_info in directory.items():
        if dir_carrier in clean_name or clean_name in dir_carrier:
            return contact_info
            
    # strip common suffix words as a final fallback
    suffixes = [' life', ' insurance', ' co', ' company', ' group', ' (lsw)', ' financial']
    stripped_name = clean_name
    for suffix in suffixes:
        stripped_name = stripped_name.replace(suffix, '').strip()
        
    for dir_carrier, contact_info in directory.items():
        dir_clean = dir_carrier
        for suffix in suffixes:
            dir_clean = dir_clean.replace(suffix, '').strip()
        if dir_clean and (dir_clean in stripped_name or stripped_name in dir_clean):
            return contact_info
            
    return {}

st.divider()

### EXECUTION LOGIC
st.subheader("Run Automation")

if st.button("Process Invoice & Queue Gmail Drafts", type="primary"):

    import tempfile
    # create a temporary directory on the cloud server to store breakdown workbooks
    temp_export_dir = tempfile.mkdtemp()
    
    # validation checks
    inputs_valid = bool(org_name and len(process_date) == 8 and process_date.isdigit())
    
    if not uploaded_file:
        st.warning("Please upload an invoice file before proceeding.")
        st.stop()
    elif not inputs_valid:
        st.warning("Please enter a valid Organization Name and an 8-digit Process Date (MMDDYYYY) in the sidebar.")
        st.stop()

    formatted_date = f"{process_date[0:2]}/{process_date[2:4]}/{process_date[4:8]}"
    status_box = st.status("Connecting to Google API services...", expanded=True)
    
    try:
        # authenticate google services
        gmail_service = get_user_gmail_service()
        sheets_service = get_shared_sheets_service()
        
        # get current authenticated user email
        user_profile = gmail_service.users().getProfile(userId='me').execute()
        user_email = user_profile.get('emailAddress', '').lower().strip()
        status_box.write(f"Logged in as: **{user_email}**")


        ### FETCH ACTIVE LOANS & EMAIL DIRECTORY FROM GOOGLE SHEETS
        status_box.write("Syncing loan tracking & email directory from Google Sheets...")
        
        # 1. fetch active loan policies
        loans_result = sheets_service.spreadsheets().values().get(
            spreadsheetId=LOAN_SPREADSHEET_ID,
            range="'Active Loan Tracking'!A:Z"
        ).execute()
        
        loan_rows = loans_result.get('values', [])
        loan_policies = set()
        if loan_rows:
            headers = [str(c).strip() for c in loan_rows[0]]
            if 'Policy Number' in headers:
                p_idx = headers.index('Policy Number')
                loan_policies = set(str(r[p_idx]).strip() for r in loan_rows[1:] if len(r) > p_idx and r[p_idx])
        
        status_box.write(f"Loaded **{len(loan_policies)}** active loan policies.")

        # 2. fetch email directory tab
        directory_result = sheets_service.spreadsheets().values().get(
            spreadsheetId=LOAN_SPREADSHEET_ID,
            range="'Email Directory'!A:Z"
        ).execute()
        
        dir_rows = directory_result.get('values', [])
        email_directory = {}
        if dir_rows:
            dir_headers = [str(c).strip() for c in dir_rows[0]]
            if 'Carrier' in dir_headers and 'Carrier Email' in dir_headers:
                c_idx = dir_headers.index('Carrier')
                e_idx = dir_headers.index('Carrier Email')
                cc_idx = dir_headers.index('Default CC List') if 'Default CC List' in dir_headers else -1
                
                for r in dir_rows[1:]:
                    if len(r) > c_idx and r[c_idx]:
                        carrier_key = str(r[c_idx]).strip().lower()
                        c_email = str(r[e_idx]).strip() if len(r) > e_idx else ''
                        c_cc = str(r[cc_idx]).strip() if cc_idx != -1 and len(r) > cc_idx else ''
                        email_directory[carrier_key] = {'email': c_email, 'cc': c_cc}

        status_box.write(f"Loaded **{len(email_directory)}** carrier contacts from Google Sheets Directory.")

        ### PARSE & PREPROCESS EXCEL DATA
        status_box.write("Applying dynamic column mapping...")
        
        # rename mapped columns
        df = df_raw.rename(columns=col_mapping)
        
        # make sure core fields exist
        missing = [f for f in REQUIRED_CORE_FIELDS if f not in df.columns]
        if missing:
            st.error(f"Missing required mapped fields: {missing}. Please adjust the column selectors above.")
            st.stop()

        df['Total Cont'] = (
        df['Total Cont']
        .astype(str)
        .str.replace('$', '', regex=False)
        .str.replace(',', '', regex=False)
        .str.replace('-', '', regex=False)
        .str.strip()
        )
        df['Total Cont'] = pd.to_numeric(df['Total Cont'], errors='coerce').fillna(0.0)

        df = df.dropna(subset=['Carrier'])
        df['Policy Number'] = df['Policy Number'].astype(str).str.strip()

        df = df.dropna(subset=['Carrier', 'Total Cont'])
        df['Policy Number'] = df['Policy Number'].astype(str).str.strip()

        # premium type hierarchy logic
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
        status_box.write(f"Processing **{len(carriers)}** carriers...")

        os.makedirs(temp_export_dir, exist_ok=True)
        generated_files = []

        ### GENERATE BREAKDOWNS & EMAIL DRAFTS
        for carrier in carriers:
            carrier_data = df[df['Carrier'] == carrier].copy()
            total_amount = carrier_data['Total Cont'].sum()
            formatted_total = f"${total_amount:,.2f}"

            # determine recipient & cc emails (fallback to google sheet directory)
            carrier_key = str(carrier).strip().lower()
            directory_info = find_carrier_contact(carrier, email_directory)

            # pull directly from google sheet email directory
            recipient_email = directory_info.get('email', '')
            raw_cc = directory_info.get('cc', '')
            cc_list = [addr.strip() for addr in raw_cc.replace(';', ',').split(',') if addr.strip()]
            filtered_cc_list = [addr for addr in cc_list if addr.lower() != user_email]
            team_cc = ", ".join(filtered_cc_list)

            if "@" not in recipient_email or "." not in recipient_email:
                st.warning(f"Skipped draft for **{carrier}**: No valid email found in uploaded sheet or Google Directory.")
                continue

            summary_row = pd.DataFrame([{'Policy Owner': 'TOTAL', 'Amount': total_amount}])
            if 'Participant' in carrier_data.columns:
                carrier_data = carrier_data.rename(columns={'Participant': 'Policy Owner'})
            elif 'Policy Owner' not in carrier_data.columns:
                carrier_data['Policy Owner'] = 'N/A'

            carrier_data = carrier_data.rename(columns={'Total Cont': 'Amount'})

            # select columns & ensure column index is unique
            columns_to_export = ['Policy Owner', 'Carrier', 'Policy Number', 'Premium Type', 'Amount']
            # filter columns and drop any accidentally duplicated columns
            export_df = carrier_data.loc[:, ~carrier_data.columns.duplicated()][columns_to_export].copy()
            export_df = export_df.reset_index(drop=True)

            # create summary row with explicit column order matching export_df
            summary_data = {
                'Policy Owner': 'TOTAL',
                'Carrier': '',
                'Policy Number': '',
                'Premium Type': '',
                'Amount': total_amount
            }
            summary_df = pd.DataFrame([summary_data])

            # concatnate safely
            final_export_df = pd.concat([export_df, summary_df], ignore_index=True)

            file_name = f"{org_name} - {carrier} Wire Breakdown {process_date}.xlsx"
            full_save_path = os.path.join(temp_export_dir, file_name)

            final_export_df.to_excel(full_save_path, index=False)

            ### EXCEL STYLING & AUTO-FIT COLUMNS
            wb = load_workbook(full_save_path)
            ws = wb.active
            total_row_idx = ws.max_row

            bold_font = Font(name='Calibri', size=11, bold=True)
            accounting_border = Border(top=Side(style='double'), bottom=Side(style='thin'))

            for row in range(2, total_row_idx):
                ws[f'E{row}'].number_format = '$#,##0.00'

            total_amount_cell = ws[f'E{total_row_idx}']
            total_amount_cell.font = bold_font
            total_amount_cell.border = accounting_border
            total_amount_cell.number_format = '$#,##0.00'
            ws[f'A{total_row_idx}'].font = bold_font

            # auto fit column widths
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    if cell.value is not None:
                        val = cell.value
                        if isinstance(val, (int, float)):
                            cell_len = len(f"${val:,.2f}")
                        else:
                            try:
                                # string-to-float conversion if formatted as text
                                num_val = float(str(val).replace('$', '').replace(',', '').strip())
                                cell_len = len(f"${num_val:,.2f}")
                            except ValueError:
                                cell_len = len(str(val))
                        if cell_len > max_len:
                            max_len = cell_len
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

            wb.save(full_save_path)
            wb.close()

            ### CONSTRUCT GMAIL DRAFT
            msg = EmailMessage()
            msg['Subject'] = f"{org_name} - {carrier} {formatted_date} Wire Breakdown"
            msg['To'] = recipient_email
            if team_cc:
                msg['Cc'] = team_cc

            body_lines = [
                "Hi team,",
                f"Attached is the breakdown for the {formatted_total} wire remitted by {org_name} on {formatted_date}.",
                "Please apply premiums according to the breakdown. Let me know if you have any questions or need any additional information. Thank you!"
            ]

            msg.set_content("\n\n".join(body_lines))
            html_body = "".join([f"<p>{line}</p>" for line in body_lines])
            msg.add_alternative(html_body, subtype='html')

            with open(full_save_path, 'rb') as f:
                file_data = f.read()
                msg.add_attachment(
                    file_data, 
                    maintype='application', 
                    subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                    filename=file_name
                )

            raw_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()

            gmail_service.users().drafts().create(
                userId='me',
                body={'message': {'raw': raw_message}}
            ).execute()
            
            status_box.write(f"Draft created: **{carrier}** | Total: **{formatted_total}** | Sent to: `{recipient_email}`")

            # Ssore file bytes for persistent download buttons
            with open(full_save_path, "rb") as f:
                file_bytes = f.read()

            generated_files.append((file_name, file_bytes))

        # Save to session state
        st.session_state.generated_files_data = generated_files

        status_box.update(label="Process complete!", state="complete", expanded=False)
        st.success(f"Success! All {org_name} breakdowns have been exported and drafts created.")


    except Exception as e:
        if 'status_box' in locals() and status_box:
            status_box.update(label="An error occurred", state="error")
        st.error(f"An unexpected error occurred during execution: {e}")

### DOWNLOAD BREAKDOWNS
if st.session_state.get('generated_files_data'):
    st.divider()
    st.subheader("Download Generated Breakdowns")
    
    st.info(
        "**Tip:** To choose a specific folder on your PC for each file, "
        "make sure **'Ask where to save each file before downloading'** is enabled in your browser settings."
    )

    for f_name, f_bytes in st.session_state.generated_files_data:
        st.download_button(
            label=f"Download {f_name}",
            data=f_bytes,
            file_name=f_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"download_{f_name}"
        )
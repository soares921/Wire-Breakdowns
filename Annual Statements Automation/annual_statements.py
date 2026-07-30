import datetime
from pathlib import Path
import os
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException, WebDriverException
import time

SCRIPT_DIR = Path(__file__).resolve().parent
EXCEL_FILE = SCRIPT_DIR / "Annual Statements Tracking List.xlsx"
SHEET_NAME = ["Q1", "Q2", "Q3", "Q4"]

load_dotenv(SCRIPT_DIR / ".env")

print(f"DEBUG: Looking for .env file at {SCRIPT_DIR / '.env'}")
print(f"DEBUG: Does the .env file exist there? {(SCRIPT_DIR / '.env').exists()}")
print(f"DEBUG: ACCORDIA_USER found? {bool(os.getenv('ACCORDIA_USER'))}")

# scrolls to element so that it's centered and clickable
def safe_click(driver, element):
    driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", element)
    time.sleep(0.5)
    try:
        element.click()
    except Exception:
        # fallback to pure javascript click if an overlay/header blocks standard click
        driver.execute_script("arguments[0].click();", element)

# function to rename downloaded statements to following convention:
# Last Name, First Name - (carrier) (policy number) Annual Statement (policy target date).pdf
def rename_statement(download_dir, item_data):
    participant = str(item_data['Participant']).strip()
    carrier = str(item_data['Carrier']).strip()
    policy_id = str(item_data['Policy Number']).strip()
    target_date_obj = pd.to_datetime(item_data['2026 Target Date'])
    target_date = target_date_obj.strftime("%m%d%Y")

    clean_filename = f"{participant} - {carrier} {policy_id} Annual Statement {target_date}"
    for char in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']:
        clean_filename = clean_filename.replace(char,'')

    target_path = download_dir / f"{clean_filename}.pdf"

    timeout = 15
    start_time = time.time()

    while time.time() - start_time < timeout:
        crdownloads = list(download_dir.glob("*.crdownload"))
        if not crdownloads:
            pdf_files = list(download_dir.glob("*.pdf"))
            if pdf_files:
                latest_pdf = max(pdf_files, key=os.path.getmtime)
                if time.time() - os.path.getmtime(latest_pdf) < 30:
                    if target_path.exists():
                        target_path.unlink()

                    latest_pdf.rename(target_path)
                    print(f"Renamed file to: {target_path.name}")
                    return True
        time.sleep(1)

    print("Could not locate file to rename.")
    return False


# reads excel file
def get_download(file_path, sheet):
    sheets_dict = pd.read_excel(file_path, sheet_name=sheet, header=0)

    combined_records = []
    today = datetime.date.today()

    # loop through each sheet's df
    for name, df in sheets_dict.items():
        df.columns = df.columns.str.strip()

        df["2026 Target Date"] = pd.to_datetime(df["2026 Target Date"]).dt.date

        # filter conditions for target date less than to today & upload status is false, NAN or 0
        past_due = df["2026 Target Date"] <= today
        not_uploaded = (
            df["Upload Status"].isin([False, 0]) | df["Upload Status"].isna()
        )
        queue_df = df[past_due & not_uploaded]

        # required columns for auto loop
        required_columns = [
            "Participant",
            "Organization",
            "Carrier",
            "Policy Number",
            "2026 Target Date",
        ]
        queue_df = queue_df[required_columns]

        # convert to dictionaries and add to our master list
        sheet_records = queue_df.to_dict(orient="records")
        combined_records.extend(sheet_records)

    return combined_records

# grab credentials saved in .env file
def download_accordia_statement(policy_items):
    username = os.getenv("ACCORDIA_USER")
    password = os.getenv("ACCORDIA_PASS")

    if not username or not password:
        print("Error: crendentials missing from .env file.")
        return False
    
    print(f"Launching browser to fetch {len(policy_items)} policies from Accordia")

    options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": str(SCRIPT_DIR),
        "plugins.always_open_pdf_externally": True,
        "download.prompt_for_download": False
    }
    options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=options)


    try:
        driver.get("https://www.globalatlanticlink.com/AgentWeb/pages/login.xhtml")
        time.sleep(3)

        user_field = driver.find_element(By.ID, "j_username")
        pass_field = driver.find_element(By.ID, "j_password")

        user_field.send_keys(username)
        pass_field.send_keys(password)
        pass_field.send_keys(Keys.RETURN)
        time.sleep(5)

        print("MFA security triggered. Check email for Accordia verification code.")
        mfa_code = input("Enter 6-digit MFA code here and press Enter: ")
        mfa_field = driver.find_element(By.ID, "oneTimePassword")

        mfa_field.send_keys(mfa_code)
        mfa_field.send_keys(Keys.RETURN)

        print("MFA Code sumbitted. Proceeding to portal.")
        time.sleep(5)

        ####### ACCORDIA PORTAL NAVIGATION #######

        for item in policy_items:
            policy_id = str(item['Policy Number']).strip()
            participant = item['Participant']

            try:
                driver.switch_to.default_content()
            except WebDriverException:
                print(f"Browser session lost prior to Policy{policy_id}. Ending queue.")
                break
        
            try:
                # 1. click "searches" drop down
                searches_menu = driver.find_element(By.LINK_TEXT, "Searches")
                safe_click(driver, searches_menu)
            except NoSuchElementException:
                driver.get("https://www.globalatlanticlink.com/AgentWeb/secured/policybobSearch.xhtml")
                time.sleep(3)
                searches_menu = driver.find_element(By.LINK_TEXT, "Searches")
                safe_click(driver, searches_menu)
                
            time.sleep(1.5)

            #2. click "policy search"
            policy_search_btn = driver.find_element(By.LINK_TEXT, "Policy Search")
            safe_click(driver, policy_search_btn)
            time.sleep(3)

            #3. Enter policy number
            search_field = driver.find_element(By.ID, "policybobSearchFormId:policyNumber")
            search_field.clear()
            search_field.send_keys(policy_id)
            search_field.send_keys(Keys.RETURN)
            time.sleep(3)

            #4. click policy number search result link; if no result found, skip
            try:
                result_cell = driver.find_element(By.XPATH, f"//td[contains(text(), '{policy_id}')]")
                safe_click(driver, result_cell)
                time.sleep(4)
            except NoSuchElementException:
                try:
                    result_row =driver.find_element(By.XPATH, f"//tr[contains(., '{policy_id}')]")
                    safe_click(driver, result_row)
                    time.sleep(4)
                except NoSuchElementException:
                    print(f"Policy {policy_id} not found in results table. Moving onto next policy...")
                    continue # skip to next policy

            # safeguard for when a policy portal data is blocked
            try:
                popup_back_btn = driver.find_element(
                    By.XPATH, 
                    "//*[contains(text(), 'not available') or contains(text(), 'Not Available')]/ancestor::div//input[@value='Back'] | //input[@value='Back'] | //button[contains(text(), 'Back')]"
                )
                if popup_back_btn.is_displayed():
                    print(f"'Portal Not Available' popup encountered for Policy {policy_id}. Clicking 'Back' and skipping...\n")
                    popup_back_btn.click()
                    time.sleep(3)
                    continue
            except NoSuchElementException:
                pass        

            #5. click "view documents"
            view_docs_btn = driver.find_element(By.LINK_TEXT, "View Documents")
            safe_click(driver, view_docs_btn)
            time.sleep(3)

            #6. select "annual statements" in dropdown
            try:
                dropdown_element = driver.find_element(By.ID, "j_idt96")
                doc_select = Select(dropdown_element)
                try:
                    doc_select.select_by_visible_text("Annual Statement")
                except NoSuchElementException:
                    # Fallback in case it's named slightly differently in the dropdown list
                    doc_select.select_by_visible_text("Annual Statements")
                time.sleep(2)
            except Exception as dropdown_err:
                print(f"Unable to select Annual Statement: {dropdown_err}")

            # 7. check first checkbox using javaScript
            try:
                first_checkbox = driver.find_element(By.XPATH, "//table//input[@type='checkbox']")
                driver.execute_script("arguments[0].click();", first_checkbox)
                time.sleep(1)
            except NoSuchElementException:
                print(f"No statement checkboxes found for policy {policy_id}. Skipping...\n")
                continue
                
            # 8. click "view" button to trigger PDF download
            try:
                view_button = driver.find_element(By.ID, "stmtPDFButton")
                
                safe_click(driver, view_button)

                print(f"Download triggered for Policy {policy_id}!")
                time.sleep(5)

                # Rename downloaded PDF according to naming convention
                rename_statement(SCRIPT_DIR, item)

            except NoSuchElementException:
                try:
                    view_button = driver.find_element(By.XPATH, "//button[contains(@id, 'stmtPDFButton')]")
                    safe_click(driver, view_button)
                    
                    print(f"Download triggered for Policy {policy_id}!")
                    time.sleep(5)
                    rename_statement(SCRIPT_DIR, item)
                except NoSuchElementException:
                    print(f"Could not locate the 'View' button for policy {policy_id}.")

            except Exception as item_error:
                print (f"Error processing policy {policy_id}; {item_error}\n")

        return True
    
    except Exception as e:
        print(f"Selenium error: {e}")
        return False
    
    finally:
        print("Closing browser session...")
        driver.quit()

if __name__ == "__main__":
    print("Reading Excel sheet and generating queue...\n")
    download_queue = get_download(EXCEL_FILE, SHEET_NAME)

    print(
        f"\nFound {len(download_queue)} annual statements ready for download today:")
    for item in download_queue:
        print(f"- {item['Participant']} | {item['Carrier']} | Policy: {item['Policy Number']}")

    # filter accordia policies
    accordia_item = [item for item in download_queue if str(item['Carrier']).strip().lower() == 'accordia']

    if accordia_item:
        print(f"\nMatch found! Processing {len(accordia_item)} Accordia policies")
        download_accordia_statement(accordia_item)
    else:
        print("\nNo Accordia policies found in today's queue.")
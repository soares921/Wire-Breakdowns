# Wire Breakdown Automation (v2.0)

An automated Python tool designed to read Excel invoice reports, create .CSV wire breakdown summaries, sanitize & categorize premium entries, track loan repayments via Google Sheets API, and generate email drafts in Gmail.

---

## Features
* **Invoice-to-Breakdowns Transformation:** Takes invoice information and filters the data by carriers for processing.
* **Google API Integration:** Automatically cross-references dynamic tracking sheets and draft emails using Google Sheets & Gmail APIs.
* **Three-Tier Categorization Logic:** Priorities are mapped as:
  1. *New Business Premium* (Manual flag overrides)
  2. *Loan Repayment* (Validated against Google Sheets)
  3. *Inforce Premium*
* **Dynamic Column Mapping:** Uses alias dictionaries to handle varying Excel header names automatically.
* **Error Handling:** Validates required columns and email address formats prior to execution.

---

## Prerequisites & Installation

### 1. Python Dependencies
Install the required packages using pip:
```bash
pip install pandas google-api-python-client google-auth-httplib2 google-auth-oauthlib

from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from email.mime.text import MIMEText
from datetime import datetime as dt
import base64
import os
import pickle
import json

from DocFunctions import \
    authenticate, display_document, update_checklist

# Define the API scopes
scope_docs = 'https://www.googleapis.com/auth/documents'
scope_gmail  = 'https://www.googleapis.com/auth/gmail.send'
scope_sheets = 'https://www.googleapis.com/auth/spreadsheets'
SCOPES_ALL = [scope_docs, scope_gmail, scope_sheets]

if __name__ == '__main__':
    # Authenticate for Google Docs & Gmail
    creds = authenticate(SCOPES_ALL)
    docService = build('docs', 'v1', credentials=creds)
    sheetService = build('sheets', 'v4', credentials=creds)

    print("\n---- FUNCTION DESCRIPTIONS ---- ")
    print("\n1 - Concentrate Checklists -----------")
    print("   - Get all documents in info/checklistDocIDs.json ")
    print("   - Iterate through all tabs in these documents ")
    print("   - Search for tables with first row, first cell == 'Checklist:...'")
    print("   - Add info to info/failedRows.json where invalid data is present")
    print("     - Validity requires three cells where the last two are of the form mm-dd-yyyy-hh-MM")
    print("   - Add all valid rows where the second column is more recent than the timestamp in info/runtimes.json's 'update_checklist' key")
    print("   - Ignore the columns that are less recent")

    print("\n2 - Retry Failed Checklist Rows ------")
    print("   - Perform 'Concentrate Checklists' for all rows in info/failedRows.json")

    print("\n\n---- AVAILABLE FUNCTIONS ---- ")
    print("1 - Concentrate Checklists")
    print("2 - Retry Failed Checklist Rows")
    print("------------------------------------- ")
    functionID = int(input("\nPlease enter the function ID you'd like to run: "))
    if functionID == 1:
        print("\nConcentrating Checklists...")
        update_checklist(docService, sheetService, False)
    elif functionID == 2:
        print("Retrying Failed Checklist Rows...")
        update_checklist(docService, sheetService, True)
    else:
        print(f"Invalid input: {functionID}")
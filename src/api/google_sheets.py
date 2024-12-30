from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

import os
print("Current working directory:", os.getcwd())

# Define the scope for Google Sheets and Drive
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

# This is the specific Google Sheet ID for the current project.
# Update this line to use a different spreadsheet for other projects.
SPREADSHEET_ID = '1byKP-xqB3mmTncWTefvyUVMwHjaAlT4XienGR3Lj7tY'

def get_service():
    """Authenticate and return the Google Sheets service."""
    creds = Credentials.from_service_account_file('src/api/service_account.json', scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)

def list_sheets(spreadsheet_id):
    """List all sheet names in the spreadsheet."""
    service = get_service()
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', [])
    return [sheet['properties']['title'] for sheet in sheets]

def get_sheet_metadata(spreadsheet_id, sheet_name=None):
    """Retrieve metadata for a spreadsheet or a specific sheet."""
    service = get_service()
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    if sheet_name:
        for sheet in sheet_metadata.get('sheets', []):
            if sheet['properties']['title'] == sheet_name:
                return sheet['properties']
    return sheet_metadata

def read_sheet(spreadsheet_id, sheet_name='Sheet1', range_name=None):
    """
    Read data from a Google Sheets spreadsheet.
    - If range_name is None, fetches the entire sheet.
    """
    service = get_service()
    range_to_read = sheet_name if not range_name else f"{sheet_name}!{range_name}"
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_to_read).execute()
    return result.get('values', [])

def write_sheet(spreadsheet_id, sheet_name='Sheet1', range_name='A1', values=[]):
    """
    Write data to a Google Sheets spreadsheet.
    - Writes starting from range_name.
    """
    service = get_service()
    range_to_write = f"{sheet_name}!{range_name}"
    body = {'values': values}
    result = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_to_write,
        valueInputOption='RAW',
        body=body
    ).execute()
    return result

if __name__ == "__main__":
    # NOTE: The SPREADSHEET_ID is currently set for the project "RadarSheets".
    # To use this script for another project, replace the value of SPREADSHEET_ID above.

    try:
        # List all sheets in the spreadsheet
        sheets = list_sheets(SPREADSHEET_ID)
        print("Sheets in the spreadsheet:", sheets)

        # Get metadata for the spreadsheet
        metadata = get_sheet_metadata(SPREADSHEET_ID)
        print("Spreadsheet metadata:", metadata)

        # Read all data from the first sheet
        if sheets:
            data = read_sheet(SPREADSHEET_ID, sheet_name=sheets[0])
            print(f"Data from sheet '{sheets[0]}':")
            for row in data:
                print(row)

        # Example: Write data to the first sheet
        example_values = [["Name", "Age"], ["Alice", 30], ["Bob", 25]]
        write_sheet(SPREADSHEET_ID, sheet_name=sheets[0], range_name='A1', values=example_values)
        print("Data written successfully.")

    except Exception as e:
        print(f"Error: {e}")

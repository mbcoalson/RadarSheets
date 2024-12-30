from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Define the scope for Google Sheets and Drive
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def get_service():
    """Authenticate and return the Google Sheets service."""
    creds = Credentials.from_service_account_file('src/api/service_account.json', scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)

def read_sheet(spreadsheet_id, range_name):
    """Read data from a Google Sheets spreadsheet."""
    service = get_service()
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    return result.get('values', [])

if __name__ == "__main__":
    # Replace with your actual spreadsheet ID and range
    SPREADSHEET_ID = 'your_spreadsheet_id_here'
    RANGE_NAME = 'Sheet1!A1:D5'

    try:
        data = read_sheet(SPREADSHEET_ID, RANGE_NAME)
        print("Data from sheet:")
        for row in data:
            print(row)
    except Exception as e:
        print(f"Error reading sheet: {e}")

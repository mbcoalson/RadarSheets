import os
import subprocess
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Google API setup
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
SPREADSHEET_ID = '1byKP-xqB3mmTncWTefvyUVMwHjaAlT4XienGR3Lj7tY'
SERVICE_ACCOUNT_FILE = 'C:\\Users\\mattc\\OneDrive\\Desktop\\Code\\RadarSheets\\src\\api\\service_account.json'

RADAR_PLOT_SCRIPT = 'C:\\Users\\mattc\\OneDrive\\Desktop\\Code\\RadarSheets\\src\\api\\google_sheets_radar_plot.1.py'
RADAR_PLOTS_DIR = 'C:\\Users\\mattc\\OneDrive\\Desktop\\Code\\RadarSheets\\radar_plots'

def get_google_services():
    """Authenticate and return the Google Sheets and Drive services."""
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    sheets_service = build('sheets', 'v4', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)
    return sheets_service, drive_service

def run_radar_plot_script():
    """Call the radar plot generation script."""
    print("Running radar plot script...")
    import sys
    subprocess.run([sys.executable, RADAR_PLOT_SCRIPT], check=True)
    print("Radar plot script completed.")

def list_sheets(sheets_service):
    """List all sheet names in the spreadsheet."""
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheets = sheet_metadata.get('sheets', [])
    return [sheet['properties']['title'] for sheet in sheets]

def upload_image_to_drive(file_path, drive_service):
    """Upload a PNG file to Google Drive and return its public URL."""
    try:
        file_name = os.path.basename(file_path)
        file_metadata = {'name': file_name, 'mimeType': 'image/png'}
        media = MediaFileUpload(file_path, mimetype='image/png')

        # Upload the file
        uploaded_file = drive_service.files().create(
            body=file_metadata, media_body=media, fields='id'
        ).execute()
        file_id = uploaded_file.get('id')

        # Make the file publicly accessible
        drive_service.permissions().create(fileId=file_id, body={'role': 'reader', 'type': 'anyone'}).execute()

        # Return the file's public URL
        file_url = f"https://drive.google.com/uc?id={file_id}"
        return file_url
    except Exception as e:
        print(f"Error uploading image to Drive: {e}")
        return None

def insert_image_url_to_sheet(sheets_service, sheet_name, row, column, image_url, width=300, height=300):
    """Insert the URL of the image into the specified cell in Google Sheets with custom size."""
    try:
        image_formula = f'=IMAGE("{image_url}", 4, {width}, {height})'
        sheets_service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!{chr(65 + column)}{row + 1}",
            valueInputOption="USER_ENTERED",
            body={"values": [[image_formula]]}
        ).execute()
        print(f"Inserted image into {sheet_name} at row {row + 1}, column {column + 1} with size {width}x{height}")
    except Exception as e:
        print(f"Error inserting image into sheet {sheet_name}: {e}")

def resize_sheet_cells(sheets_service, sheet_name, row, column, row_height=300, column_width=300):
    """Resize the row height and column width for better display of the image."""
    try:
        sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        sheet_id = None
        for sheet in sheet_metadata['sheets']:
            if sheet['properties']['title'] == sheet_name:
                sheet_id = sheet['properties']['sheetId']
                break

        if sheet_id is None:
            print(f"Sheet {sheet_name} not found.")
            return

        requests = [
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": row,
                        "endIndex": row + 1
                    },
                    "properties": {"pixelSize": row_height},
                    "fields": "pixelSize"
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": column,
                        "endIndex": column + 1
                    },
                    "properties": {"pixelSize": column_width},
                    "fields": "pixelSize"
                }
            }
        ]

        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": requests}
        ).execute()

        print(f"Resized row {row + 1} and column {column + 1} in {sheet_name} to {row_height}x{column_width}.")
    except Exception as e:
        print(f"Error resizing cells in sheet {sheet_name}: {e}")

if __name__ == "__main__":
    sheets_service, drive_service = get_google_services()

    # Step 1: Run the radar plot generation script
    run_radar_plot_script()

    # Step 2: Process all sheets and insert radar plots
    failed_inserts = []
    invalid_sheets = []

    sheet_names = list_sheets(sheets_service)
    print(f"Found sheets: {sheet_names}")

    for sheet_name in sheet_names:
        print(f"Processing sheet: {sheet_name}")
        try:
            sheet_data = sheets_service.spreadsheets().values().get(
                spreadsheetId=SPREADSHEET_ID, range=f"{sheet_name}!A1:Z").execute()
            rows = sheet_data.get('values', [])

            if not rows or len(rows[0]) < 8:
                print(f"Invalid data structure in sheet: {sheet_name}")
                invalid_sheets.append(sheet_name)
                continue

            headers = rows[0]  # Assume the first row contains headers
            categories = headers[3:8]  # Radar plot categories ('Innovate' to 'Risk')

            grouped_data = {}
            for row in rows[1:]:  # Skip header row
                if len(row) < 8:
                    continue
                company_name = row[0]
                if company_name not in grouped_data:
                    grouped_data[company_name] = []
                grouped_data[company_name].append(row)

            for company_name, company_rows in grouped_data.items():
                try:
                    valid_rows = []
                    for row in company_rows:
                        try:
                            values = list(map(float, row[3:8]))  # Extract radar plot values
                            valid_rows.append((row[1], row[2], values))  # (Evaluator, Lab, Values)
                        except ValueError:
                            continue

                    if not valid_rows:
                        print(f"No valid data for company: {company_name} in sheet {sheet_name}")
                        continue

                    output_png_path = os.path.join(RADAR_PLOTS_DIR, f"{company_name}_{sheet_name}_radar.png")
                    # Upload PNG to Drive
                    image_url = upload_image_to_drive(output_png_path, drive_service)
                    if not image_url:
                        failed_inserts.append(company_name)
                        continue

                    # Insert the image URL into the sheet
                    insert_image_url_to_sheet(sheets_service, sheet_name, rows.index(company_rows[0]), 10, image_url)  # Column K (index 10)
                    resize_sheet_cells(sheets_service, sheet_name, rows.index(company_rows[0]), 10, row_height=300, column_width=300)

                except Exception as e:
                    print(f"Error processing company {company_name} in sheet {sheet_name}: {e}")
                    failed_inserts.append(company_name)
        except Exception as e:
            print(f"Error processing sheet {sheet_name}: {e}")
            invalid_sheets.append(sheet_name)

    # Step 3: Output summary
    if failed_inserts:
        print(f"Failed to insert images for the following companies: {failed_inserts}")
    if invalid_sheets:
        print(f"Sheets with invalid data structure: {invalid_sheets}")
    if not failed_inserts and not invalid_sheets:
        print("All radar plots inserted into Google Sheets successfully.")


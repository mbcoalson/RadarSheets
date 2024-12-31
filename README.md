RadarSheets: Automating Radar Plot Integration with Google Sheets
Overview
RadarSheets is a Python-based project designed to automate radar plot generation for Google Sheets. This is a one-off project specifically for the GPG FY25 initiative but is adaptable for future GPG years, provided the data format remains consistent.

The radar plots are created as PNG images, uploaded to Google Drive, and inserted into corresponding Google Sheets for visualization. Additionally, interactive HTML versions of the radar plots are saved locally on the machine running these scripts.

Key Features
Automated Radar Plot Generation:

Generates radar plots categorized by company, layering three lab reviews for each plot.
Saves plots as PNG and HTML files.
Google Drive Integration:

Uploads radar plots to Google Drive with public access.
Generates URLs for the uploaded images.
Google Sheets Integration:

Embeds radar plot images in Google Sheets using the IMAGE function.
Automatically resizes rows and columns to fit the images.
Error Handling:

Validates sheet data structure and logs invalid entries.
Logs failed uploads and inserts for debugging.
Data Cleaning:

Employs dataclean.gs, a Google Apps Script, to convert formulas or data pulls into a sheet with standardized data types (e.g., strings, floats).
Prerequisites
Software Requirements:

Python 3.12 or higher
Google Sheets and Google Drive APIs enabled
Python Libraries:

google-auth
google-auth-oauthlib
google-auth-httplib2
google-api-python-client
plotly
Service Account Credentials:

A JSON key file for Google API authentication.
Spreadsheet:

A Google Spreadsheet containing structured evaluation data.
Directory Structure
plaintext
Copy code
RadarSheets/
│
└── src/
    ├── api/
    │   ├── google_sheets.py                 # Early iteration, not in use
    │   ├── google_sheets_radar_plot.1.py    # Script used by main_script.py
    │   ├── google_sheets_radar_plot.py      # Early iteration, not in use
    │   ├── main_script.py                   # Main execution script
    │   ├── service_account.json             # Google API credentials file
    │   └── __init__.py                      # Not in use
    │
    ├── plots/                               # Not in use
    │   ├── radar_plot.py
    │   └── __init__.py
    │
    ├── utils/                               # Not in use
    │   ├── data_cleaner.py
    │   └── __init__.py
    │
    ├── main.py                              # Not in use
    └── __init__.py                          # Not in use
Setup
Install Python Libraries:

bash
Copy code
pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client plotly

Service Account Setup:

Create a service account in Google Cloud Console.
Enable the Google Drive and Google Sheets APIs.
Download the JSON credentials file and place it in src/api as service_account.json.
Prepare Google Spreadsheet:

Ensure the spreadsheet matches the format of 2025 RFI Reviewer Scoring.
Use the dataclean.gs file:
Go to Extensions > Apps Script and create a new Apps Script file named dataclean.
Copy and paste the dataclean.gs script from this repository.
Run the script to validate that a sheet named "Cleaned Values Only Data" is created with standard data types (e.g., numbers without formulas).
Update Configuration:

Modify main_script.py:
Set SPREADSHEET_ID to your Google Spreadsheet ID.
Set SERVICE_ACCOUNT_FILE to the path of your credentials file.
Usage
Run the Main Script:

bash
Copy code
python main_script.py
Process Overview:

Authenticates with Google APIs.
Generates radar plots.
Uploads radar plots to Google Drive.
Embeds radar plots into Google Sheets and resizes cells.
Error Handling
Invalid Data Structures:

Logs invalid sheets for review.
Failed Uploads or Inserts:

Logs companies with failed operations for debugging.
Debugging:

Check the terminal output for error details.
Customization
Image Size:

Adjust row_height and column_width in the resize_sheet_cells function.
Radar Plot Appearance:

Edit google_sheets_radar_plot.1.py for custom styles and metrics.
Future Enhancements
Add support for additional chart types.
Provide a web-based interface for non-technical users.
Implement database integration for processed data storage.
License
Not applicable.

Support
For issues or questions, please contact the developer.


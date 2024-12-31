# **RadarSheets: Automating Radar Plot Integration with Google Sheets**

## **Overview**
RadarSheets is a Python-based project designed to automate radar plot generation for Google Sheets. This is a one-off project specifically for the GPG FY25 initiative but is adaptable for future GPG years, provided the data format remains consistent.

The radar plots are created as PNG images, uploaded to Google Drive, and inserted into corresponding Google Sheets for visualization. Additionally, interactive HTML versions of the radar plots are saved locally on the machine running these scripts.

---

## **Key Features**
- **Automated Radar Plot Generation**:
  - Generates radar plots categorized by company, layering three lab reviews for each plot.
  - Saves plots as PNG and HTML files.

- **Google Drive Integration**:
  - Uploads radar plots to Google Drive with public access.
  - Generates URLs for the uploaded images.

- **Google Sheets Integration**:
  - Embeds radar plot images in Google Sheets using the `IMAGE` function.
  - Automatically resizes rows and columns to fit the images.

- **Error Handling**:
  - Validates sheet data structure and logs invalid entries.
  - Logs failed uploads and inserts for debugging.

- **Data Cleaning**:
  - Employs `dataclean.gs`, a Google Apps Script, to convert formulas or data pulls into a sheet with standardized data types (e.g., strings, floats).

---

## **Prerequisites**
### **Software Requirements**:
- Python 3.12 or higher
- Google Sheets and Google Drive APIs enabled

### **Python Libraries**:
- `google-auth`
- `google-auth-oauthlib`
- `google-auth-httplib2`
- `google-api-python-client`
- `plotly`

### **Service Account Credentials**:
- A JSON key file for Google API authentication.

### **Spreadsheet**:
- A Google Spreadsheet containing structured evaluation data.

---

## **Directory Structure**
```plaintext
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


# RadarSheets: Automating Radar Plot Integration with Google Sheets

## Overview
RadarSheets is a Python-based project designed to automate the process of generating radar plots for companies based on evaluation data in Google Sheets. The radar plots are created as PNG images, uploaded to Google Drive, and inserted into corresponding Google Sheets for easy visualization. This project leverages Google APIs to ensure seamless integration and supports error handling for invalid data structures.

---

## Key Features
1. **Automated Radar Plot Generation**:
   - Generates radar plots using company evaluation data extracted from Google Sheets.
   - Plots are saved as PNG files.

2. **Google Drive Integration**:
   - Radar plots are uploaded to Google Drive with public access.
   - Generates URLs for uploaded images.

3. **Google Sheets Integration**:
   - Inserts radar plot images directly into Google Sheets using the `IMAGE` function.
   - Automatically resizes rows and columns to accommodate the images.

4. **Error Handling**:
   - Handles invalid data structures in sheets.
   - Logs failed uploads and invalid sheets for further debugging.

---

## Prerequisites
1. Python 3.7+
2. Required Python Libraries:
   - `google-auth`
   - `google-auth-oauthlib`
   - `google-auth-httplib2`
   - `google-api-python-client`
   - `plotly`
3. Service Account Credentials:
   - A JSON key file for Google API authentication.
4. A Google Spreadsheet with structured evaluation data.

---

## Directory Structure
```plaintext
RadarSheets/
├── src/
│   ├── api/
│   │   ├── google_sheets_radar_plot.1.py  # Script for radar plot generation
│   │   ├── service_account.json           # Google API credentials
├── radar_plots/                           # Directory for generated PNGs
└── main.py                                # Main script integrating all functionalities
```

---

## Setup
1. **Install Python Libraries**:
   ```bash
   pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client plotly
   ```

2. **Service Account Setup**:
   - Create a service account in Google Cloud Console.
   - Enable Google Drive and Google Sheets APIs.
   - Download the JSON credentials file and save it as `service_account.json` in the `src/api` directory.

3. **Prepare Google Spreadsheet**:
   - Ensure the spreadsheet follows this structure:
     - Column A: Company Name
     - Columns B-C: Evaluator and Lab
     - Columns D-H: Metrics for radar plots (e.g., Innovate, Impact, Savings, etc.)
   - Add more sheets as needed; the script processes all sheets in the spreadsheet.

4. **Update Configuration**:
   - In `main.py`, update:
     - `SPREADSHEET_ID`: The ID of your Google Spreadsheet.
     - `SERVICE_ACCOUNT_FILE`: Path to your credentials file.

---

## Usage
1. **Run the Main Script**:
   ```bash
   python main.py
   ```
2. **Process Flow**:
   - The script authenticates Google APIs.
   - Runs the radar plot generation script.
   - Uploads the radar plots to Google Drive.
   - Inserts radar plots into Google Sheets and resizes the cells.

---

## Error Handling
1. **Invalid Data Structure**:
   - If a sheet does not contain the required structure, it is logged as invalid.

2. **Failed Uploads or Inserts**:
   - Companies with failed uploads or inserts are logged for review.

3. **Debugging**:
   - Review the terminal output for details on errors.

---

## Customization
1. **Image Size**:
   - Modify `row_height` and `column_width` in the `resize_sheet_cells` function to change cell dimensions.

2. **Radar Plot Appearance**:
   - Customize the plotly script (`google_sheets_radar_plot.1.py`) for different styles and metrics.

---

## Future Enhancements
1. Add support for different chart types.
2. Provide a web-based interface for non-technical users.
3. Include email notifications for errors or completion statuses.
4. Implement database integration to store processed data.

---

## License
This project is licensed under the MIT License. See the LICENSE file for details.

---

## Support
For issues or questions, contact the developer or submit a GitHub issue in the repository.

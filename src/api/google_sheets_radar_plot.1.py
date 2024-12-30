import os
import plotly.graph_objects as go
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Google Sheets setup
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1byKP-xqB3mmTncWTefvyUVMwHjaAlT4XienGR3Lj7tY'


def get_service():
    """Authenticate and return the Google Sheets service."""
    creds_path = 'C:\\Users\\mattc\\OneDrive\\Desktop\\Code\\RadarSheets\\src\\api\\service_account.json'

    print("Using service account file:", os.path.abspath(creds_path))

    if not os.path.exists(creds_path):
        raise FileNotFoundError(f"Service account file not found: {os.path.abspath(creds_path)}")

    creds = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    print("Authentication successful.")
    return build('sheets', 'v4', credentials=creds)

def read_sheet(spreadsheet_id, sheet_name='Sheet1', range_name=None):
    """Read data from a specific sheet."""
    service = get_service()
    range_to_read = sheet_name if not range_name else f"{sheet_name}!{range_name}"
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_to_read).execute()
    return result.get('values', [])

def generate_company_radar_plot(valid_rows, categories, company_name, output_html_path, output_png_path):
    """
    Generates a radar plot for a single company with multiple evaluators and saves as HTML and PNG.
    - valid_rows: List of tuples [(Evaluator, Lab, [Scores]), ...].
    - categories: List of category names for radar plot (e.g., ['Innovate', 'Impact', 'Savings']).
    - company_name: Name of the company.
    - output_html_path: Path to save the radar plot HTML file.
    - output_png_path: Path to save the radar plot PNG file.
    """
    fig = go.Figure()

    for evaluator_name, lab_name, values in valid_rows:
        values.append(values[0])  # Closing the radar chart loop
        categories_loop = categories + [categories[0]]  # Repeat the first category to close the chart

        # Add evaluator trace
        fig.add_trace(go.Scatterpolar(
            r=values,
            theta=categories_loop,
            fill='toself',
            name=f"{evaluator_name}, {lab_name}"  # Append lab name to evaluator in legend
        ))

    # Update layout for the radar plot
    fig.update_layout(
        polar=dict(
            radialaxis=dict(
                visible=True,
                range=[0, 10]  # Adjust range as needed
            )
        ),
        title=f"Radar Plot for {company_name}",
        showlegend=True
    )

    # Save the plot as an HTML file
    fig.write_html(output_html_path)
    print(f"Radar plot HTML saved: {output_html_path}")

    # Save the plot as a PNG file
    fig.write_image(output_png_path)
    print(f"Radar plot PNG saved: {output_png_path}")

def list_sheets(spreadsheet_id):
    """List all sheet names in the spreadsheet."""
    service = get_service()
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', [])
    return [sheet['properties']['title'] for sheet in sheets]

if __name__ == "__main__":
    # Ensure the radar_plots directory exists
    output_dir = 'radar_plots'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Created directory: {os.path.abspath(output_dir)}")
    else:
        print(f"Output directory exists: {os.path.abspath(output_dir)}")

    # List all sheets in the spreadsheet
    sheet_names = list_sheets(SPREADSHEET_ID)
    print(f"Found sheets: {sheet_names}")

    # Keep track of sheets that cannot be processed
    unprocessed_sheets = []

    for sheet_name in sheet_names:
        print(f"Processing sheet: {sheet_name}")
        try:
            # Read data from the current sheet
            data = read_sheet(SPREADSHEET_ID, sheet_name=sheet_name)

            if not data:
                print(f"No data found in sheet: {sheet_name}")
                unprocessed_sheets.append(sheet_name)
                continue

            print(f"Fetched data from {sheet_name}:")
            for row in data:
                print(row)

            # Step 2: Process data grouped by company
            headers = data[0]  # Assume the first row contains headers
            categories = headers[3:8]  # Select radar plot categories ('Innovate' to 'Risk')

            # Group rows by company
            grouped_data = {}
            for row in data[1:]:  # Skip the header row
                company_name = row[0]  # First column
                if company_name not in grouped_data:
                    grouped_data[company_name] = []
                grouped_data[company_name].append(row)

            # Generate radar plots for each company
            for company_name, company_rows in grouped_data.items():
                try:
                    # Prepare radar plot data
                    valid_rows = []
                    for row in company_rows:
                        try:
                            # Extract only numeric values for the radar plot
                            values = list(map(float, row[3:8]))  # Select columns 'Innovate' to 'Risk'
                            valid_rows.append((row[1], row[2], values))  # (Evaluator, Lab, Values)
                        except ValueError as ve:
                            print(f"Skipping row due to invalid data: {row} - {ve}")
                            continue

                    if not valid_rows:
                        print(f"No valid data found for company: {company_name}")
                        continue

                    # Generate and save the radar plot
                    output_html_path = os.path.join(output_dir, f"{company_name}_{sheet_name}_radar.html")
                    output_png_path = os.path.join(output_dir, f"{company_name}_{sheet_name}_radar.png")
                    generate_company_radar_plot(valid_rows, categories, company_name, output_html_path, output_png_path)
                except Exception as e:
                    print(f"Error generating radar plot for company {company_name}: {e}")
        except Exception as e:
            print(f"Error processing sheet {sheet_name}: {e}")
            unprocessed_sheets.append(sheet_name)

    # Print summary of unprocessed sheets
    if unprocessed_sheets:
        print(f"Sheets that could not be processed: {unprocessed_sheets}")
    else:
        print("All sheets processed successfully.")

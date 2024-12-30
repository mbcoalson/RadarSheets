import plotly.graph_objects as go
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Google Sheets setup
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1byKP-xqB3mmTncWTefvyUVMwHjaAlT4XienGR3Lj7tY'

def get_service():
    """Authenticate and return the Google Sheets service."""
    creds = Credentials.from_service_account_file('src/api/service_account.json', scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)

def read_sheet(spreadsheet_id, sheet_name='Sheet1', range_name=None):
    """Read data from a specific sheet."""
    service = get_service()
    range_to_read = sheet_name if not range_name else f"{sheet_name}!{range_name}"
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_to_read).execute()
    return result.get('values', [])

def generate_radar_plot(categories, values, evaluator_name, output_path):
    """
    Generates a radar plot for a single evaluator.
    - categories: List of category names (e.g., ['Innovate', 'Impact', 'Savings']).
    - values: Corresponding scores for the categories.
    - evaluator_name: Name of the evaluator.
    - output_path: Path to save the radar plot HTML file.
    """
    # Ensure the radar plot is closed
    values.append(values[0])  # Closing the radar chart loop
    categories.append(categories[0])

    fig = go.Figure()

    fig.add_trace(go.Scatterpolar(
        r=values,
        theta=categories,
        fill='toself',
        name=evaluator_name
    ))

    fig.update_layout(
        polar=dict(
            radialaxis=dict(
                visible=True,
                range=[0, 10]  # Adjust range as needed
            )
        ),
        title=f"Radar Plot for {evaluator_name}",
        showlegend=False
    )

    # Save the plot as an HTML file
    fig.write_html(output_path)
    print(f"Radar plot saved: {output_path}")

if __name__ == "__main__":
    # Step 1: Read data from a single tab
    sheet_name = 'Sheet1'  # Replace with your sheet name
    data = read_sheet(SPREADSHEET_ID, sheet_name=sheet_name)
    
    if not data:
        print(f"No data found in sheet: {sheet_name}")
    else:
        # Step 2: Process data
        headers = data[0]  # Assume the first row contains headers
        categories = headers[1:]  # Skip the first column (evaluator names)

        for row in data[1:]:  # Skip the header row
            evaluator_name = row[0]
            values = list(map(float, row[1:]))  # Convert scores to float
            output_path = f"radar_plots/{evaluator_name}_radar.html"

            # Generate and save radar plot
            generate_radar_plot(categories[:], values[:], evaluator_name, output_path)

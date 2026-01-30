import os
import pandas as pd
import requests
from flask import Flask, render_template
from io import BytesIO
import traceback

# Flask Dashboard App


app = Flask(__name__)
EXCEL_URL = "https://sheet.zohopublic.in/sheet/published/zn2qy64c7d154d78141b8b07bfe9e9d793dc2?download=xlsx"

def download_and_process_excel():
    """
    Downloads the Excel file and extracts only the specified columns.
    Using the corrected indices provided.
    """
    try:
        # Download the file content
        print("Downloading Excel file...")
        response = requests.get(EXCEL_URL, timeout=30)
        response.raise_for_status()

        # Read the Excel file
        df = pd.read_excel(BytesIO(response.content), engine='openpyxl')
        
        # Print all available columns for debugging
        print("\n=== Available columns in the Excel file ===")
        for idx, col in enumerate(df.columns):
            print(f"Index {idx}: '{col}'")
        print("="*50 + "\n")
        
        # Define the indices based on your corrections
        column_indices = {
            'Name': 2,
            'Date': 0,
            'Highest Degree': 7,
            'Institution Name': 9,
            'CGPA/Percentage': 11,
            'Year of Graduation': 10,
            'Duration': 42,
            'Statement of Purpose': 15,
            'Select Institutes': 29,
            'Select Faculty': 30,
            'Upload Resume/CV': 17,
            'Reference Letter': 31
        }
        
        # Display which indices we're using
        print("Using the following column indices:")
        for col_name, idx in column_indices.items():
            if idx < len(df.columns):
                print(f"  {col_name}: Index {idx} -> '{df.columns[idx]}'")
            else:
                print(f"  {col_name}: Index {idx} -> ERROR: Index out of bounds!")
        
        # Extract only the columns we need
        indices = list(column_indices.values())
        
        # Check if any index is out of bounds
        max_index = len(df.columns) - 1
        invalid_indices = [idx for idx in indices if idx > max_index]
        
        if invalid_indices:
            print(f"\nERROR: Some indices are out of bounds! Max index is {max_index}")
            print(f"Invalid indices: {invalid_indices}")
            return []
        
        selected_data = df.iloc[:, indices]
        
        # Set column names for display
        selected_data.columns = [
            'Name',
            'Date',
            'Highest Degree Obtained',
            'Institution Name',
            'CGPA/Percentage',
            'Year of Graduation',
            'Duration',
            'Statement of Purpose',
            'Select Institutes',
            'Select Faculty',
            'Upload Resume/CV',
            'Reference Letter'
        ]
        
        # Clean the data - replace NaN with empty strings
        selected_data = selected_data.fillna('')
        
        # Convert date columns to string for better display
        if 'Date' in selected_data.columns:
            selected_data['Date'] = selected_data['Date'].astype(str)
        
        # Convert the DataFrame to a list of dictionaries
        records = selected_data.to_dict(orient='records')
        
        # REVERSE THE ORDER - Last entry first, first entry last
        records.reverse()
        
        print(f"\nSuccessfully processed {len(records)} records.")
        print("Note: Records are displayed in REVERSE order (newest first).")
        
        return records

    except requests.exceptions.RequestException as e:
        print(f"Error downloading file: {e}")
        return []
    except Exception as e:
        print(f"Error processing Excel file: {e}")
        traceback.print_exc()
        return []

@app.route('/')
def index():
    data = download_and_process_excel()
    return render_template('index.html', excel_data=data)

@app.route('/refresh')
def refresh():
    data = download_and_process_excel()
    return render_template('index.html', excel_data=data)

if __name__ == '__main__':
    # Use PORT environment variable for Render, default to 5000 for local
    port = int(os.environ.get('PORT', 5000))
    # Run with 0.0.0.0 to allow external connections on Render
    app.run(host='0.0.0.0', port=port, debug=False)
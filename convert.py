import pandas as pd
import json

def json_to_excel(json_file_path, excel_file_path):
    try:
        # Read the JSON file
        with open(json_file_path, 'r') as f:
            json_data = json.load(f)
        
        # Create a new Excel writer object
        with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
            # Iterate over each item in the JSON data
            for item in json_data:
                sheet_name = f"{item['id']}_{item['title']}"[:31]  # Combine id and title for the sheet name, truncated to 31 characters
                cards = item['cards']
                
                # Create a DataFrame from the cards data
                df = pd.DataFrame(cards)
                
                # Write the DataFrame to a specific sheet
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"Successfully converted {json_file_path} to {excel_file_path}")
    except Exception as e:
        print(f"An error occurred: {e}")

# File paths
json_file_path = 'files/xxxx.json'  # Update this path to your JSON file
excel_file_path = 'files/xxxx.xlsx'  # Update this path to your desired Excel file

# Convert JSON data to Excel file
json_to_excel(json_file_path, excel_file_path)

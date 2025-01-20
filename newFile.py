import os
import pandas as pd

# Define file paths
file_to_check = os.path.join(os.path.expanduser("~"), "Downloads", "Kams 40k 200125.xlsx")
master_folder = os.path.join(os.path.expanduser("~"), "Desktop", "MasterFileMTA")

# Read the main file
try:
    data_file = pd.read_excel(file_to_check, engine='openpyxl')
except Exception as e:
    raise ValueError(f"Error reading the main file: {e}")

# Ensure the 'Mobile' column exists in the main file
if 'Mobile' not in data_file.columns:
    raise ValueError("The 'Mobile' column is missing in the file to check.")

# Add Status and Date columns
data_file['Status'] = "Unique"
data_file['Date'] = ""

# Process each file in the MasterFileMTA folder
for master_file_name in os.listdir(master_folder):
    master_file_path = os.path.join(master_folder, master_file_name)

    # Skip non-Excel files
    if not master_file_name.endswith(('.xlsx', '.xls')):
        continue

    try:
        # Read the master file
        master_file = pd.read_excel(master_file_path, engine='openpyxl')
    except Exception as e:
        print(f"Skipping file {master_file_name}: {e}")
        continue

    # Ensure the required columns exist in the master file
    if 'Mobile' not in master_file.columns or 'Data date' not in master_file.columns:
        print(f"Skipping file {master_file_name}: Required columns missing.")
        continue

    # Normalize the 'Mobile' column for comparison
    master_file['Mobile'] = master_file['Mobile'].astype(str).str.strip()
    data_file['Mobile'] = data_file['Mobile'].astype(str).str.strip()

    # Identify duplicates and update the main file
    for index, row in data_file.iterrows():
        if row['Status'] == "Unique":
            matched_row = master_file[master_file['Mobile'] == row['Mobile']]
            if not matched_row.empty:
                data_file.at[index, 'Status'] = "Duplicate"
                data_file.at[index, 'Date'] = matched_row.iloc[0]['Data date']

# Save the updated file
output_file_path = os.path.join(os.path.expanduser("~"), "Downloads", "Kams 40k 200125_updated.xlsx")
data_file.to_excel(output_file_path, index=False)

print(f"File processed and saved to: {output_file_path}")
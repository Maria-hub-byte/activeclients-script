import pandas as pd
import os

# Define file and folder paths
main_file_path = os.path.expanduser("~/Desktop/Kams 50k 13012025.xlsx")
master_folder_path = os.path.expanduser("~/Desktop/MasterFileMTA")
output_file_path = os.path.expanduser("~/Desktop/Kams_50k_13012025_Processed.xlsx")

# Load the main file
main_df = pd.read_excel(main_file_path, engine='openpyxl')

# Ensure the necessary columns exist in the main file
if 'Mobile' not in main_df.columns:
    raise ValueError("The 'Mobile' column is missing from the main file.")

# Add new columns for Status and Date
main_df['Status'] = 'Unique'
main_df['Date'] = None

# Check if the master folder exists
if not os.path.exists(master_folder_path):
    raise FileNotFoundError(f"The folder {master_folder_path} does not exist.")

# Process each file in the master folder
for file_name in os.listdir(master_folder_path):
    file_path = os.path.join(master_folder_path, file_name)

    # Skip invalid or temporary files
    if not file_name.endswith(".xlsx") or file_name.startswith("~$"):
        print(f"Skipping invalid or temporary file: {file_name}")
        continue

    try:
        master_df = pd.read_excel(file_path, engine='openpyxl')

        # Ensure the necessary columns exist in the master file
        if 'Mobile' not in master_df.columns or 'Data Date' not in master_df.columns:
            print(f"Skipping file {file_name} due to missing columns.")
            continue

        # Normalize and check for duplicates
        main_df['Mobile'] = main_df['Mobile'].astype(str).str.strip()
        master_df['Mobile'] = master_df['Mobile'].astype(str).str.strip()

        duplicates = main_df['Mobile'].isin(master_df['Mobile'])
        main_df.loc[duplicates, 'Status'] = 'Duplicate'

        # Update the 'Date' column for duplicates
        for index, row in main_df[duplicates].iterrows():
            match = master_df[master_df['Mobile'] == row['Mobile']]
            if not match.empty:
                main_df.at[index, 'Date'] = match.iloc[0]['Data Date']

    except Exception as e:
        print(f"Skipping file {file_name} due to an error: {e}")
        continue

# Save the updated file
main_df.to_excel(output_file_path, index=False)

print(f"Processing complete. Updated file saved as: {output_file_path}")


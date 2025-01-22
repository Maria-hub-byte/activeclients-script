import pandas as pd
import os

# File and folder details
file_to_check = os.path.expanduser("~/Downloads/TEST3_001_WED 22-01-25 (1).xlsx")
master_folder = os.path.expanduser("~/Desktop/MasterFileMTA")

# Load the file to check
df_to_check = pd.read_excel(file_to_check, engine='openpyxl')

# Ensure necessary columns are present in the file to check
if "Mobile" not in df_to_check.columns:
    raise ValueError("The file to check must have a 'Mobile' column.")

# Normalize the Mobile column in the file to check
df_to_check["Mobile"] = df_to_check["Mobile"].astype(str).str.strip()

# Add 'Status' and 'Date' columns to the file to check
df_to_check["Status"] = "Unique"
df_to_check["Date"] = ""

# Process master files
for master_file in os.listdir(master_folder):
    master_file_path = os.path.join(master_folder, master_file)
    
    # Skip non-Excel files
    if not master_file_path.endswith((".xlsx", ".xls")):
        continue
    
    try:
        # Load master file
        master_df = pd.read_excel(master_file_path, engine='openpyxl')
        
        # Ensure necessary columns are present in the master file
        if "Mobile" not in master_df.columns or "Data Date" not in master_df.columns:
            continue
        
        # Normalize the Mobile column in the master file
        master_df["Mobile"] = master_df["Mobile"].astype(str).str.strip()
        
        # Find duplicates
        duplicates = df_to_check["Mobile"].isin(master_df["Mobile"])
        
        # Update 'Status' and 'Date' for duplicates
        df_to_check.loc[duplicates, "Status"] = "Duplicate"
        df_to_check.loc[duplicates, "Date"] = master_df.loc[
            master_df["Mobile"].isin(df_to_check.loc[duplicates, "Mobile"]), "Data Date"
        ].values
    
    except Exception as e:
        print(f"Skipping file {master_file} due to error: {e}")

# Save the updated file back to the Downloads folder
output_file_path = os.path.expanduser("~/Downloads/TEST3_001_WED 22-01-25 (1)_updated.xlsx")
df_to_check.to_excel(output_file_path, index=False)

print(f"File has been updated and saved to: {output_file_path}")

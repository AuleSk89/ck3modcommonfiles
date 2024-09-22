import os
import json
import pandas as pd
import logging
from datetime import datetime
import shutil
import openpyxl

# Set up directories
project_dir = r'C:\oned\OneDrive\python projects\240922 - ck3 mod incompatibility'
log_dir = os.path.join(project_dir, 'logs')
output_dir = os.path.join(project_dir, 'outputs')
code_dir = os.path.join(project_dir, 'codes')

# Create directories if they don't exist
os.makedirs(log_dir, exist_ok=True)
os.makedirs(output_dir, exist_ok=True)
os.makedirs(code_dir, exist_ok=True)

# Clean old files in directories
for folder in [log_dir, output_dir]:
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f'Failed to delete {file_path}. Reason: {e}')

# Set up logging
timestamp = datetime.now().strftime('%Y%m%d_%H%M')
log_filename = os.path.join(log_dir, f'log_{timestamp}.log')
logging.basicConfig(filename=log_filename, level=logging.DEBUG, format='%(asctime)s %(message)s')

# Log start of the script
logging.info('Script started.')

# Load JSON data
json_file_path = r'C:\oned\OneDrive\python projects\240922 - ck3 mod incompatibility\SK_Modded_ck3.json'
with open(json_file_path, 'r') as file:
    data = json.load(file)
logging.info(f'Loaded JSON file from {json_file_path}.')

mods = data['mods']
base_path = 'D:/SteamLibrary/steamapps/workshop/content/1158310'

# Extract mod directories
mod_directories = {mod['steamId']: os.path.join(base_path, mod['steamId']) for mod in mods}
logging.info(f'Extracted mod directories.')

# Define a list of subfolders to skip
skip_subfolders = ['.metadata', '.git']

# Dictionary to store file occurrences
file_occurrences = {}
second_sheet_data = []  # For documenting positions, mod names, and Steam IDs

# Iterate over each mod directory
for mod in mods:
    mod_dir = mod_directories[mod['steamId']]
    mod_name = mod['displayName']  # Assuming 'displayName' key exists in each mod's data
    for root, _, files in os.walk(mod_dir):
        rel_dir = os.path.relpath(root, mod_dir)
        # Skip files directly under root or in specified subfolders
        if rel_dir == "." or any(skip_folder in rel_dir.split(os.path.sep) for skip_folder in skip_subfolders):
            continue
        for file in files:
            file_path = os.path.join(rel_dir, file)
            if file_path not in file_occurrences:
                file_occurrences[file_path] = []
            file_occurrences[file_path].append(mod['steamId'])
            # Add to second sheet data
            second_sheet_data.append({"Position": rel_dir, "Mod Name": mod_name, "Steam ID": mod['steamId']})
logging.info('File occurrences recorded and second sheet data prepared.')

# Step 1: Identify Common Occurrences
common_files = {file_path: mods for file_path, mods in file_occurrences.items() if len(mods) > 1}

# Step 2: Filter Data to Include Only Mods in Common Occurrences
common_mods = set(mod for mods in common_files.values() for mod in mods)

# Prepare data for DataFrame
output_data = []
mod_columns_by_id = [mod['steamId'] for mod in mods if mod['steamId'] in common_mods]
mod_columns_by_name = [mod['displayName'] for mod in mods if mod['steamId'] in common_mods]
mod_columns_by_position = [str(mod['position']) for mod in mods if mod['steamId'] in common_mods]

for file_path, mod_ids in common_files.items():
    file_name = os.path.basename(file_path)
    file_subfolder = os.path.dirname(file_path)
    row_by_id = {
        "File Name": file_name,
        "File Subfolder": file_subfolder,
        "# of Mods": len(mod_ids),
    }
    row_by_name = row_by_id.copy()
    row_by_position = row_by_id.copy()
    
    for mod_id in mod_columns_by_id:
        row_by_id[mod_id] = 'X' if mod_id in mod_ids else ''
    
for file_path, mod_ids in common_files.items():
    file_name = os.path.basename(file_path)
    file_subfolder = os.path.dirname(file_path)
    row_by_id = {
        "File Name": file_name,
        "File Subfolder": file_subfolder,
        "# of Mods": len(mod_ids),
    }
    row_by_name = row_by_id.copy()
    row_by_position = row_by_id.copy()
    
    for mod_id in mod_columns_by_id:
        if mod_id in mod_ids:
            row_by_id[mod_id] = 'X'
            if mod_id in common_mods:  # Check if mod is in common_mods
                mod = next((m for m in mods if m['steamId'] == mod_id), None)
                if mod:
                    row_by_name[mod['displayName']] = 'X'
                    row_by_position[str(mod['position'])] = 'X'
        else:
            row_by_id[mod_id] = ''
            # No need to explicitly set '' for row_by_name and row_by_position here
    
    output_data.append((row_by_id, row_by_name, row_by_position))
logging.info('Data prepared for DataFrame.')

# Create DataFrames
df_by_id = pd.DataFrame([row[0] for row in output_data])
df_by_name = pd.DataFrame([row[1] for row in output_data])
df_by_position = pd.DataFrame([row[2] for row in output_data])
df_second_sheet = pd.DataFrame(second_sheet_data)

# Save DataFrames to Excel with multiple sheets
output_filename = os.path.join(output_dir, f'mod_file_comparison_{timestamp}.xlsx')
with pd.ExcelWriter(output_filename) as writer:
    df_by_id.to_excel(writer, sheet_name='File Occurrences by ID', index=False)
    df_by_name.to_excel(writer, sheet_name='File Occurrences by Name', index=False)
    df_by_position.to_excel(writer, sheet_name='File Occurrences by Position', index=False)
    df_second_sheet.to_excel(writer, sheet_name='Mod Details', index=False)
logging.info(f'DataFrame saved to {output_filename} with multiple sheets.')

# Copy the script to the codes subfolder
script_filename = __file__
shutil.copy(script_filename, os.path.join(code_dir, f'ck3_comp_check_{timestamp}.py'))
logging.info(f'Script copied to {os.path.join(code_dir, f"ck3_comp_check_{timestamp}.py")}.')

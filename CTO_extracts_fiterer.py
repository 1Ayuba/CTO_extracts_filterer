#!/usr/bin/env python
# coding: utf-8

# # CTO Extracts
# #### Author: Bayugo Ayuba Ahmed
# 
# The purpose of this script is to filter and extract records associated with a specified client name from multiple sheets within an Excel file, and save these records into a single workbook. The script handles scenarios where there are no records for the specified client in some sheets and informs the user accordingly. Additionally, it prompts the user to input the output path and desired file name for saving the processed Excel files.

# In[ ]:


import os
import pandas as pd

# Prompt the user to enter the Excel file name
file_name = input("Enter the Excel file name (including extension): ")

# Read all sheets into a dictionary of DataFrames
try:
    sheets_dict = pd.read_excel(file_name, sheet_name=None)
except FileNotFoundError:
    print(f"File '{file_name}' not found. Please make sure the file exists.")
    exit(1)
except pd.errors.EmptyDataError:
    print(f"File '{file_name}' is empty.")
    exit(1)
except pd.errors.ParserError:
    print(f"Error parsing the Excel file '{file_name}'. Please make sure it is a valid Excel file.")
    exit(1)

# Prompt the user to enter the client name
client_name = input("Enter the client name: ")

# Prompt the user to enter the output path
#output_path = input("Enter the output path for saving Excel files: ")

# Ensure the output path exists; create it if it doesn't
#if not os.path.exists(output_path):
#    os.makedirs(output_path)
    
# Prompt the user to enter the output path
output_path = input("Enter the output path for saving Excel files: ").strip()

# Ensure the output path exists; create it if it doesn't
if not os.path.exists(output_path):
    os.makedirs(output_path)


# Prompt the user to enter the file name for saving
output_file_name = input("Enter the file name for saving Excel workbook (without extension): ")

# Create an ExcelWriter to write multiple DataFrames to a single workbook
output_file_path = os.path.join(output_path, f"{output_file_name}_{client_name}_workbook.xlsx")
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    # Process each sheet and save filtered records to the same Excel writer
    for sheet_name, sheet_df in sheets_dict.items():
        # Check if 'Client' column exists in the sheet
        if 'Client' in sheet_df.columns:
            # Filter rows based on the entered client name
            filtered_df = sheet_df[sheet_df['Client'] == client_name]

            # Save the filtered DataFrame to a new worksheet with the original sheet name in the same workbook
            filtered_df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Check if any rows match the specified client name
            if not filtered_df.empty:
                print(f"Records associated with client '{client_name}' in sheet '{sheet_name}' saved to workbook.")
            else:
                print(f"No records found for client '{client_name}' in sheet '{sheet_name}'.")
        else:
            print(f"Sheet '{sheet_name}' does not have a 'Client' column.")

print(f"Workbook with records associated with client '{client_name}' saved to '{output_file_path}'.")
print("Processing complete.")


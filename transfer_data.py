import pandas as pd
import os

# Define file paths
excel1_path = 'C:\\Users\\Marye\\OneDrive\\Desktop\\C\\ExcelAutomation\\Excel1.xlsx'  # Path to Excel1
updated_excel2_path = 'C:\\Users\\Marye\\OneDrive\\Desktop\\C\\ExcelAutomation\\updated_excel2.xlsx'  # Path to updated Excel2

# Load Excel1
df1 = pd.read_excel(excel1_path)

# Load existing updated_excel2.xlsx if it exists, otherwise initialize an empty DataFrame
if os.path.exists(updated_excel2_path):
    df2 = pd.read_excel(updated_excel2_path)
    print("Existing updated_excel2.xlsx loaded successfully.")
else:
    df2 = pd.DataFrame()
    print("No existing updated_excel2.xlsx found. Creating a new file.")

# Define the required column order
required_columns = ['Name', 'TimeStamp', 'SourceTimeStamp', 'Area', 'Latitude', 'Longitude', 'Type', 'Unit', 'Value', 'Relationships']

# Map Excel1 columns to match the required structure
column_mapping = {
    'Well Id': 'Name',
    'Date of Sampling': 'TimeStamp',
    'Area': 'Area',
    'Parameter': 'Type',
    'Unit': 'Unit',
    'Result': 'Value'
}

# Rename columns in df1 to match the final structure
df1_selected = df1.rename(columns=column_mapping)

# Ensure 'TimeStamp' is in datetime format and convert to ISO format
df1_selected['TimeStamp'] = pd.to_datetime(df1_selected['TimeStamp'], errors='coerce').dt.strftime('%Y-%m-%dT%H:%M:%S')

# Create SourceTimeStamp identical to TimeStamp
df1_selected['SourceTimeStamp'] = df1_selected['TimeStamp']

# Create missing columns with NaN values
for col in required_columns:
    if col not in df1_selected.columns:
        df1_selected[col] = pd.NA

# Reorder columns to match the required structure
df1_selected = df1_selected[required_columns]

# If df2 is empty (first time running), use df1_selected as the base
if df2.empty:
    df2 = df1_selected
    print("Excel2 was empty, initializing with new data.")
else:
    # Ensure df2 has the same column structure as df1_selected
    df2 = df2.reindex(columns=required_columns, fill_value=pd.NA)

    # Remove rows from df1_selected that already exist in df2
    df_new_data = pd.concat([df2, df1_selected], ignore_index=True).drop_duplicates(keep=False)

    # Append only the new unique rows from df1_selected to df2
    df2 = pd.concat([df2, df_new_data], ignore_index=True)
    print("Only new data added to Excel2, no existing rows deleted.")

# Save the updated file
df2.to_excel(updated_excel2_path, index=False)
print("Data successfully updated and saved to", updated_excel2_path)








# import pandas as pd
# import os

# # Define file paths
# excel1_path = 'C:\\Users\\Marye\\OneDrive\\Desktop\\C\\ExcelAutomation\\Excel1.xlsx'  # Path to Excel1
# updated_excel2_path = 'C:\\Users\\Marye\\OneDrive\\Desktop\\C\\ExcelAutomation\\updated_excel2.xlsx'  # Path to updated Excel2

# # Load Excel1
# df1 = pd.read_excel(excel1_path)

# # Load existing updated_excel2.xlsx if it exists, otherwise initialize an empty DataFrame
# if os.path.exists(updated_excel2_path):
#     df2 = pd.read_excel(updated_excel2_path)
#     print("Existing updated_excel2.xlsx loaded successfully.")
# else:
#     df2 = pd.DataFrame()
#     print("No existing updated_excel2.xlsx found. Creating a new file.")

# # Define the required column order
# required_columns = ['Name', 'TimeStamp', 'SourceTimeStamp', 'Area', 'Latitude', 'Longitude', 'Type', 'Unit', 'Value', 'Relationships']

# # Map Excel1 columns to match the required structure
# column_mapping = {
#     'Well Id': 'Name',
#     'Date of Sampling': 'TimeStamp',
#     'Area': 'Area',
#     'Parameter': 'Type',
#     'Unit': 'Unit',
#     'Result': 'Value'
# }

# # Rename columns in df1 to match the final structure
# df1_selected = df1.rename(columns=column_mapping)

# # # Ensure 'TimeStamp' is in datetime format and convert to ISO format
# df1_selected['TimeStamp'] = pd.to_datetime(df1_selected['TimeStamp'], errors='coerce').dt.strftime('%Y-%m-%dT%H:%M:%S')

# # Create SourceTimeStamp identical to TimeStamp
# df1_selected['SourceTimeStamp'] = df1_selected['TimeStamp']

# # Create missing columns with NaN values
# for col in required_columns:
#     if col not in df1_selected.columns:
#         df1_selected[col] = pd.NA

# # Reorder columns to match the required structure
# df1_selected = df1_selected[required_columns]

# # If df2 is empty (first time running), use df1_selected as the base
# if df2.empty:
#     df2 = df1_selected
#     print("Excel2 was empty, initializing with new data.")
# else:
#     # Ensure df2 has the same column structure as df1_selected
#     df2 = df2.reindex(columns=required_columns, fill_value=pd.NA)


#     # Ensure 'TimeStamp' is in datetime format and convert to ISO format
#     df1_selected['TimeStamp'] = pd.to_datetime(df1_selected['TimeStamp'], errors='coerce').dt.strftime('%Y-%m-%dT%H:%M:%S')

#     # Create SourceTimeStamp identical to TimeStamp
#     df1_selected['SourceTimeStamp'] = df1_selected['TimeStamp']

#     # Combine old and new data while removing duplicates
#     df_combined = pd.concat([df2, df1_selected], ignore_index=True)
#     df_combined = df_combined.drop_duplicates(keep='first')

#     # Update df2 with the new unique data
#     df2 = df_combined
#     print("Checked for duplicates and added only new data.")

# # Save the updated file
# df2.to_excel(updated_excel2_path, index=False)
# print("Data successfully updated and saved to", updated_excel2_path)







# # Description: Transfer data from Excel1 to Excel2 based on required column structure
# # Excel1 is the source file and Excel2 is the destination file
# # Excel1 contains data that needs to be transferred to Excel2
# # Excel2 is updated with the new data from Excel1
# # The data is transferred based on the required column structure
# # The required column structure is defined in the 'required_columns' list
# # The column names in Excel1 are mapped to match the required structure
# # The 'TimeStamp' column is converted to datetime format and then to ISO format
# # The 'SourceTimeStamp' column is created identical to 'TimeStamp'
# # Missing columns are created with NaN values
# # The columns are reordered to match the required structure
# # The updated data is saved to updated_excel2.xlsx
# # If updated_excel2.xlsx does not exist, it is created
# # If updated_excel2.xlsx exists, the new data is added without duplicates


# import pandas as pd
# import os

# # Define file paths
# excel1_path = 'C:\\Users\\Marye\\OneDrive\\Desktop\\C\\ExcelAutomation\\Excel1.xlsx'  # Path to Excel1
# updated_excel2_path = 'C:\\Users\\Marye\\OneDrive\\Desktop\\C\\ExcelAutomation\\updated_excel2.xlsx'  # Path to updated Excel2

# # Load Excel1
# df1 = pd.read_excel(excel1_path)

# # Load existing updated_excel2.xlsx if it exists, otherwise initialize an empty DataFrame
# if os.path.exists(updated_excel2_path):
#     df2 = pd.read_excel(updated_excel2_path)
#     print("Existing updated_excel2.xlsx loaded successfully.")
# else:
#     df2 = pd.DataFrame()
#     print("No existing updated_excel2.xlsx found. Creating a new file.")

# # Define the required column order
# required_columns = ['Name', 'TimeStamp', 'SourceTimeStamp', 'Area', 'Latitude', 'Longitude', 'Type', 'Unit', 'Value', 'Relationships']

# # Map Excel1 columns to match the required structure
# column_mapping = {
#     'Well Id': 'Name',
#     'Date of Sampling': 'TimeStamp',
#     'Area': 'Area',
#     'Parameter': 'Type',
#     'Unit': 'Unit',
#     'Result': 'Value'
# }

# # Rename columns in df1 to match the final structure
# df1_selected = df1.rename(columns=column_mapping)

# # Ensure 'TimeStamp' is in datetime format and convert to ISO format
# df1_selected['TimeStamp'] = pd.to_datetime(df1_selected['TimeStamp'], errors='coerce').dt.strftime('%Y-%m-%dT%H:%M:%S')

# # Create SourceTimeStamp identical to TimeStamp
# df1_selected['SourceTimeStamp'] = df1_selected['TimeStamp']

# # Create missing columns with NaN values
# for col in required_columns:
#     if col not in df1_selected.columns:
#         df1_selected[col] = pd.NA

# # Reorder columns to match the required structure
# df1_selected = df1_selected[required_columns]

# # If df2 is empty (first time running), use df1_selected as the base
# if df2.empty:
#     df2 = df1_selected
#     print("Excel2 was empty, initializing with new data.")
# else:
#     # Ensure df2 has the same column structure as df1_selected
#     df2 = df2.reindex(columns=required_columns, fill_value=pd.NA)

#     # Combine old and new data while removing duplicates
#     df_combined = pd.concat([df2, df1_selected], ignore_index=True)
#     df_combined = df_combined.drop_duplicates(keep='first')

#     # Update df2 with the new unique data
#     df2 = df_combined
#     print("Checked for duplicates and added only new data.")

# # Save the updated file
# df2.to_excel(updated_excel2_path, index=False)
# print("Data successfully updated and saved to", updated_excel2_path)


import pandas as pd
import xlsxwriter

# Load the Excel documents into pandas DataFrames
df1 = pd.read_excel('new_df.xlsx', sheet_name='test sheet') # Adjust the sheet name as needed
df2 = pd.read_excel('CMETest.xlsx', sheet_name='Bucketted Skills') # Adjust the sheet name as needed

# Create a new Excel file and add a worksheet
workbook = xlsxwriter.Workbook('CMETest2.xlsx', {'nan_inf_to_errors': True})
worksheet = workbook.add_worksheet()


# Example: Write data from df1 to the new Excel file
# This part of the code is simplified and needs to be adjusted based on your specific requirements
for index, row in df1.iterrows():
    # Assuming 'Resource Enterprise ID' is in column 'B' of df1 and 'EID' is in column 'D' of df2
    # Assuming 'Skill' is in column 'C' of df1 and the skills are in the first row of df2
    # Assuming 'Level' is in column 'D' of df1
    eid = row['Resource Enterprise ID']
    skill = row['Skill']
    level = row['Level']
    
    # Find the matching row in df2
    matching_row = df2.loc[df2['EID'] == eid]
    
    # If a match is found, write the level to the corresponding cell in the new Excel file
    if not matching_row.empty:
        # Example: Write the level to a specific cell (adjust the row and column indices as needed)
        worksheet.write(index, 0, level) # Adjust the row and column indices as needed

# Example: Apply data validation to a cell
# This is a simplified example. You'll need to adjust the cell range and validation options based on your requirements.
data_validation = workbook.add_data_validation()
data_validation.add('A1:A10', {'validate': 'list', 'source': ['Option 1', 'Option 2', 'Option 3']})

# Close the workbook
workbook.close()

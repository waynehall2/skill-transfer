import pandas as pd

# Load the Excel documents into pandas DataFrames
df1 = pd.read_excel('updatedDF1.xlsx', sheet_name='test sheet') # Adjust the file name and sheet name as needed
df2 = pd.read_excel('CMETest.xlsx', sheet_name='Bucketted Skills') # Adjust the file name and sheet name as needed

# Assuming 'EID' is in column 'D' of df1 and df2, and 'Skill' is in column 'C' of df1
# Assuming 'Level' is in column 'D' of df1

# Create a mapping of skills to their row index in df2
skill_to_row = {skill: row for row, skill in enumerate(df2.iloc[0])}

# Loop through each row in df1
for index1, row1 in df1.iterrows():
    eid = row1['EID']
    skill = row1['Skill']
    level = row1['Level']
    
    # Find the matching row in df2
    matching_row_index = skill_to_row.get(skill)
    
    # If a match is found, update the corresponding cell in df2
    if matching_row_index is not None:
        # Assuming 'EID' is in column 'D' of df1 and df2
        # Find the row in df2 where 'EID' matches
        matching_eid_row = df2[df2['EID'] == eid].index[0]
        # Update the cell in df2 with the 'Level' value from df1
        df2.at[matching_eid_row, matching_row_index] = level

# Save the updated DataFrame back to the Excel workbook, maintaining the original format of df2
df2.to_excel('UpdatedExcelDoc2.xlsx', index=False, engine='openpyxl')

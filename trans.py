import pandas as pd

# Read the first Excel file (with values)
df1 = pd.read_excel("new_dataframe.xlsx", engine="openpyxl")

# Read the second Excel file (where you want to insert the values)
df2 = pd.read_excel("Book1.xlsx", engine="openpyxl")

# Create a mapping of B-C pairs to D values from df1
b_c_to_d = df1.set_index(['B', 'C'])['D'].to_dict()

# Function to find the D value for a given B-C pair
def find_d_value(b_value, c_value):

# Insert the D value into the intersecting cell of df2
for b_value in df2.columns:
    for c_value in df2.index:
        d_value = find_d_value(b_value, c_value)
        if d_value is not None:
            df2.at[c_value, b_value] = d_value

# Write the updated DataFrame back to the second Excel file
df2.to_excel("Book1.xlsx", index=False)

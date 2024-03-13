import pandas as pd

# Read the Excel file
df = pd.read_excel('CME Skills.xlsx', engine='openpyxl')

df['Resource Skills'] = df['Resource Skills'].str.split('|')

print(df)
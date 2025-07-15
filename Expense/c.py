import pandas as pd

# Load the Excel files
file1_path = 'Expense_coa.xlsx'
file2_path = 'Expense_gl.xlsx'

# Read the sheets (assumes data is in the first sheet of each file)
df1 = pd.read_excel(file1_path)
df2 = pd.read_excel(file2_path)

# Ensure 'Code' columns are treated as strings to preserve formatting
df1['Code'] = df1['Code'].astype(str)
df2['Code'] = df2['Code'].astype(str)

# Find codes in df1 but not in df2
missing_in_gl = df1[~df1['Code'].isin(df2['Code'])]

# Find codes in df2 but not in df1
missing_in_coa = df2[~df2['Code'].isin(df1['Code'])]

# Save the results to Excel files
missing_in_gl.to_excel('missing_in_gl.xlsx', index=False)
missing_in_coa.to_excel('missing_in_coa.xlsx', index=False)

# Optional: print results
print("Codes in COA but missing in GL:")
print(missing_in_gl)

print("\nCodes in GL but missing in COA:")
print(missing_in_coa)

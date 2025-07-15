import pandas as pd
import re

# Load the Excel file
df = pd.read_excel("y.xlsx", header=None)

# Extract code and name from the first column using regex
df[['code', 'name']] = df[0].str.extract(r'^(\d{10})\s+(.*)')

# Drop original column 0
df = df[['code', 'name']]

# Save to new file
df.to_excel("cleaned_output.xlsx", index=False)

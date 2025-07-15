import pandas as pd
import re

# Load the Excel file (replace with your file path)
file_path = "4.xlsx"
df = pd.read_excel(file_path, header=None)

# Flatten to a single-column Series and drop NaNs
rows = df[0].dropna().astype(str)

# Extract only rows with a 10-digit account code followed by a description
pattern = r"^(?P<code>\d{10})\s+(?P<description>.+)$"
structured_df = rows.str.extract(pattern)

# Drop rows where code is not matched (invalid format)
structured_df = structured_df.dropna(subset=["code"])

# Optionally: Export to a new Excel file
structured_df.to_excel("Structured_COA.xlsx", index=False)

print("✅ Process completed. File saved as 'Structured_COA.xlsx'")

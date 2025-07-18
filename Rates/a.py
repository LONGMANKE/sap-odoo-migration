import pandas as pd
from datetime import datetime

# === Settings ===
INPUT_FILE  = "sap_rates.xlsx"          # <-- change to your source file
OUTPUT_FILE = "sap_rates_normalized.xlsx"   # <-- output with fixed headers

# === Step 1: Load the Excel file ===
df = pd.read_excel(INPUT_FILE)

# === Step 2: Normalize all headers ===
new_cols = []

for col in df.columns:
    # Case 1: already a datetime (from Excel formatting)
    if isinstance(col, datetime):
        new_cols.append(col.strftime("%d.%m.%Y"))

    # Case 2: string that looks like a date (e.g. "9/9/2020")
    else:
        try:
            parsed = pd.to_datetime(str(col).strip(), dayfirst=True, errors='raise')
            new_cols.append(parsed.strftime("%d.%m.%Y"))
        except:
            new_cols.append(str(col).strip())  # keep as is (e.g. "Country")

# Apply the cleaned header
df.columns = new_cols

# === Step 3: Save to a new file ===
df.to_excel(OUTPUT_FILE, index=False)
print(f"✅ Saved normalized file: {OUTPUT_FILE}")

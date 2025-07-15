import pandas as pd

# Load Excel file
df = pd.read_excel("3.xlsx")  # <-- Replace with your actual filename

# Column name to filter
col = "code_mapping_ids/company_id"

# Define valid codes to keep
valid_codes = {
    "TBG", "ET10", "GB10", "IN10", "IN20", "KE10", "KE20", "KE30",
    "MU10", "MW10", "MZ10", "NA10", "NG10", "RW10", "SZ10",
    "TZ10", "UG10", "US10", "ZA10", "ZM10"
}

# Filter rows to keep only those with valid codes
df_cleaned = df[df[col].isin(valid_codes)].reset_index(drop=True)

# Save cleaned output
df_cleaned.to_excel("cleaned_codes_output.xlsx", index=False)

print("✅ File cleaned. Only valid codes retained in 'cleaned_codes_output.xlsx'")

import pandas as pd

# Load the latest and old COA Excel files
latest_df = pd.read_excel("full_Liability_COA.xlsx")
old_df = pd.read_excel("liabilities_COA.xlsx")

# --- Clean and extract 10-digit codes from latest COA ---
latest_df = latest_df.astype(str)
latest_df["code"] = latest_df.iloc[:, 0].str.extract(r"(\b\d{10}\b)")

# Drop rows without valid code
latest_df = latest_df.dropna(subset=["code"])

# --- Clean codes from old COA ---
old_df = old_df.dropna(subset=["code"])
old_df["code"] = old_df["code"].astype(float).astype(int).astype(str).str.zfill(10)

# --- Compare and find old codes that are NOT in the latest file ---
latest_codes = set(latest_df["code"])
old_codes = set(old_df["code"])
missing_in_latest = sorted(old_codes - latest_codes)

# --- Filter the missing entries from old_df ---
missing_df = old_df[old_df["code"].isin(missing_in_latest)]

# --- Export the result ---
missing_df.to_excel("old_codes_missing_in_latest.xlsx", index=False)

print("✅ Done! Old codes missing in the latest COA exported to 'old_codes_missing_in_latest.xlsx'")

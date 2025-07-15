import pandas as pd
import re

# Load the latest and old COA Excel files
latest_df = pd.read_excel("Assets_COA.xlsx")
old_df = pd.read_excel("Assets_GL.xlsx")

# --- Clean and extract 10-digit codes from latest COA ---
# Convert all to string and extract 10-digit codes
latest_df = latest_df.astype(str)
latest_df["Code"] = latest_df.iloc[:, 0].str.extract(r"(\b\d{10}\b)")
latest_df["description"] = latest_df.iloc[:, 0].str.extract(r"\b\d{10}\b\s+(.*)")

# Drop rows without valid code
latest_df = latest_df.dropna(subset=["Code"])

# --- Clean codes from old COA ---
old_df = old_df.dropna(subset=["Code"])
old_df["Code"] = old_df["Code"].astype(float).astype(int).astype(str).str.zfill(10)

# --- Compare and find missing codes ---
latest_codes = set(latest_df["Code"])
old_codes = set(old_df["Code"])
missing_codes = sorted(latest_codes - old_codes)

# --- Filter the missing entries from latest_df ---
missing_df = latest_df[latest_df["Code"].isin(missing_codes)]

# --- Export the result ---
missing_df.to_excel("missing_liability_accounts.xlsx", index=False)

print("✅ Done! Missing codes exported to 'missing_liability_accounts.xlsx'")

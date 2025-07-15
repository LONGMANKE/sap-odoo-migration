import pandas as pd

# Load the Excel file
df = pd.read_excel("2.xlsx")  # Replace with your actual file path

# Define the mapping of country names to codes
replacements = {
    "KENYA": "KE10,KE20,KE30",
    "INDIA": "IN10,IN20",
    "ETHIOPIA": "ET10",
    "MALAWI": "MW10",
    "MAURITIUS": "MU10",
    "MOZAMBIQUE": "MZ10",
    "NAMIBIA": "NA10",
    "NIGERIA": "NG10",
    "RWANDA": "RW10",
    "SOUTH AFRICA": "ZA10",
    "SWAZILAND": "SZ10",
    "TANZANIA": "TZ10",
    "UGANDA": "UG10",
    "UNITED KINGDOM": "GB10",
    "USA": "US10",
    "ZAMBIA": "ZM10",
    "TECHNO BRAIN GROUP": "TBG"
}

# Replace in 'company_ids' column only
def update_company_ids(value):
    if pd.isna(value):
        return value
    for country, code in replacements.items():
        value = value.replace(country, code)
    return value

df["company_ids"] = df["company_ids"].apply(update_company_ids)

# Save to a new Excel file
df.to_excel("3.xlsx", index=False)

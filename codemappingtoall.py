import pandas as pd

# Load your Excel file
df = pd.read_excel("banks.xlsx")

# Step 1: Extract all non-empty codes in order
codes_in_order = df["code_mapping_ids/code"].dropna().astype(str).tolist()

# Step 2: Repeat each code 35 times
expanded_codes = []
for code in codes_in_order:
    expanded_codes.extend([code] * 35)

# Step 3: Apply expanded codes back into the dataframe
df["code_mapping_ids/code"] = expanded_codes[:len(df)]

# Step 4: Define the constant company list
company_list = (
    "KE10,KE20,KE30,KE40,KE50,"
    "MU10,MU20,MU30,MU40,MU50,"
    "MW10,MZ10,NA10,NG10,RW10,SZ10,TZ10,"
    "UG10,UG20,US10,ZA10,ZA20,ZM10,"
    "TBG,AE00,AE10,AE20,BR10,CA10,ET10,GB10,GH10,IN10,IN20,IN30"
)

# Step 5: Initialize empty columns
df["company_ids"] = ""
df["code"] = ""

# Step 6: Fill only rows where company_id == "TBG"
block_size = 35
for start in range(0, len(df), block_size):
    end = start + block_size
    block_code = df["code_mapping_ids/code"].iloc[start]  # same repeated code for the block

    mask_tbg = (df.index >= start) & (df.index < end) & (df["code_mapping_ids/company_id"] == "TBG")

    # Fill company_ids and code only for TBG row
    df.loc[mask_tbg, "company_ids"] = company_list
    df.loc[mask_tbg, "code"] = block_code

# Step 7: Save the result
output_file = "banks_final_tbg_only_codes_and_companies.xlsx"
df.to_excel(output_file, index=False)

print("✅ Done! Code and company_ids added only where company_id == 'TBG'.")

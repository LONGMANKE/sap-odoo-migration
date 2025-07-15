import pandas as pd

# Load your input Excel file
input_file = "5.xlsx"
df = pd.read_excel(input_file)

# Step 1: Fill down the code column
df["code_mapping_ids/code"] = df["code_mapping_ids/code"].replace("", pd.NA).fillna(method="ffill")

# Step 2: Group by the filled code to get corresponding company_ids
company_group = df.groupby("code_mapping_ids/code")["code_mapping_ids/company_id"].apply(lambda x: ",".join(x))

# Step 3: Create new column 'company_ids' only for TBG rows
df["company_ids"] = ""
for code, company_list in company_group.items():
    mask = (df["code_mapping_ids/code"] == code) & (df["code_mapping_ids/company_id"] == "TBG")
    df.loc[mask, "company_ids"] = company_list

# Step 4: Save the result
df.to_excel("code_mapping_with_company_ids.xlsx", index=False)

print("✅ Done. File saved as 'code_mapping_with_company_ids.xlsx'")

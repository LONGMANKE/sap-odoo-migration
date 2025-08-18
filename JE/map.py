import pandas as pd
import json

# -------------------------------
# Load data
# -------------------------------
gl_df = pd.read_excel("sap_gl_extract.xlsx")
map_df = pd.read_excel("analytic_accounts_mapping.xlsx")

# Create mapping dict
analytic_map = dict(zip(map_df['Analytic Account'].astype(str), map_df['ID'].astype(str)))

# Normalize fields
gl_df['Document Number'] = gl_df['Document Number'].astype(str)
gl_df["Amount in doc. curr."] = pd.to_numeric(gl_df["Amount in doc. curr."], errors='coerce').fillna(0)
gl_df["Amount in local currency"] = pd.to_numeric(gl_df["Amount in local currency"], errors='coerce').fillna(0)

# -------------------------------
# Group and transform
# -------------------------------
output_rows = []
grouped = gl_df.groupby("Document Number")

for doc_number, group in grouped:
    first = group.iloc[0]
    reference = doc_number
    company = first.get("Company Code", "")
    date = pd.to_datetime(first.get("Document Date")).strftime('%Y-%m-%d')
    journal = "Miscellaneous Operations"
    number = ""
    partner = ""
    status = "Draft"
    currency = first.get("Document currency", "UGX")
    signed_total = group["Amount in doc. curr."].sum()

    total_debit = group[group["Amount in local currency"] > 0]["Amount in local currency"].sum()
    total_credit = -group[group["Amount in local currency"] < 0]["Amount in local currency"].sum()

    for _, row in group.iterrows():
        account = row.get("Account")
        label = row.get("Text") or row.get("Name of offsetting account", "")
        amount_doc = row["Amount in doc. curr."]
        local_amount = row["Amount in local currency"]
        debit = local_amount if local_amount > 0 else 0.0
        credit = -local_amount if local_amount < 0 else 0.0

        # Use Vendor Name directly
        vendor_name = row.get("Vendor Name", "")
        vendor_name = str(vendor_name).strip() if pd.notna(vendor_name) else ""

        # Analytic mapping
        analytic_dict = {}
        for col in ["Profit Center", "Cost Center", "WBS element"]:
            val = row.get(col, "")
            if pd.notna(val) and val.strip() != "":
                for entry in str(val).split(","):
                    entry = entry.strip()
                    if entry in analytic_map:
                        analytic_dict[analytic_map[entry]] = 100.0
        analytic_json = json.dumps(analytic_dict) if analytic_dict else ""

        output_rows.append({
            "Reference": reference,
            "Company": company,
            "Date": date,
            "Journal": journal,
            "Number": number,
            "Partner": partner,
            "Status": status,
            "Total Signed": signed_total,
            "Journal Items/Account": account,
            "Journal Items/Label": label,
            "Journal Items/Amount in Currency": amount_doc,
            "Journal Items/Partner": vendor_name,
            "Journal Items/Currency": row.get("Document currency", "UGX"),
            "Journal Items/Debit": debit,
            "Journal Items/Credit": credit,
            "Journal Items/Analytic Distribution": analytic_json,
        })

        # Reset metadata after the first line
        reference = company = date = journal = number = partner = status = signed_total = ""

# -------------------------------
# Export
# -------------------------------
final_df = pd.DataFrame(output_rows)
final_df.to_excel("odoo_journal_import_with_vendor_names.xlsx", index=False)
print("✅ Done: Exported using Vendor Name directly.")

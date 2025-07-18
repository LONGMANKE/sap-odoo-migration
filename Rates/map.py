import pandas as pd
from datetime import datetime

# ─────── 1. Input/output ─────────────────────────────────────
INPUT_FILE = "sap_rates.xlsx"
CURRENCY_META = {
    "AED": ("base.AED", "AED"),
    "BIF": ("base.BIF", "FBu"),
    "CAD": ("base.CAD", "$"),
    "ETB": ("base.ETB", "Br"),
    "GBP": ("base.GBP", "£"),
    "GHS": ("base.GHS", "GH¢"),
    "INR": ("base.INR", "₹"),
    "KES": ("base.KES", "KSh"),
    "MUR": ("base.MUR", "Rs"),
    "MWK": ("base.MWK", "MK"),
    "MZN": ("base.MZN", "MT"),
    "NAD": ("base.NAD", "$"),
    "NGN": ("base.NGN", "₦"),
    "RWF": ("base.RWF", "RF"),
    "SZL": ("base.SZL", "E"),
    "TZS": ("base.TZS", "TSh"),
    "UGX": ("base.UGX", "USh"),
    "USD": ("base.USD", "$"),
    "ZAR": ("base.ZAR", "R"),
    "ZMW": ("base.ZMW", "ZK"),
    "EUR": ("base.EUR", "€"),
}

OUTPUT_FILE = "odoo_currency_rates1.xlsx"

# ─────── 2. Read and prep ────────────────────────────────────
df = pd.read_excel(INPUT_FILE)
df.columns = [str(c).strip() for c in df.columns]

# detect date columns
date_cols = [
    col for col in df.columns
    if isinstance(col, (str, datetime)) and (
        (isinstance(col, str) and col.count('.') == 2 and len(col) == 10)
        or isinstance(col, datetime)
    )
]

# convert to ISO format
iso_map = {
    col: datetime.strptime(col, "%d.%m.%Y").strftime("%Y-%m-%d")
    if isinstance(col, str) else col.strftime("%Y-%m-%d")
    for col in date_cols
}
date_cols_desc = date_cols[::-1]

# ─────── 3. Build output ─────────────────────────────────────
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    for country, group in df.groupby("Country"):
        rows = []
        for _, row in group.iterrows():
            currency = row["To currency"]
            if currency not in CURRENCY_META:
                continue  # skip unknown currencies
            curr_id, symbol = CURRENCY_META[currency]
            for i, old_date in enumerate(date_cols_desc):
                rate = row[old_date]
                if pd.isna(rate):
                    continue
                iso_date = iso_map[old_date]
                rows.append({
                    "id": curr_id if i == 0 else "",
                    "Symbol": symbol if i == 0 else "",
                    "Active": "TRUE" if i == 0 else "",
                    "Rates/Date": iso_date,
                    "Rates/Inverse Company Rate": round(rate, 2),
                    "Rates/Currency": currency
                })
        if rows:
            out = pd.DataFrame(rows)
            out.to_excel(writer, index=False, sheet_name=country[:31])

print(f"✅ Finished. Saved '{OUTPUT_FILE}' with full Odoo format per country.")

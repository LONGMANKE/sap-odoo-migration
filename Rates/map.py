import pandas as pd
from datetime import datetime

# ─────── 1. Input/output ─────────────────────────────────────
INPUT_FILE = "1_normalized.xlsx"
OUTPUT_FILE = "odoo_currency_rates.xlsx"

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
    "SLL": ("base.SLL", "Le"),
}

# ─────── 2. Read and clean ───────────────────────────────────
df = pd.read_excel(INPUT_FILE)
df.columns = [str(c).strip() for c in df.columns]

def clean_cell(val):
    if isinstance(val, str):
        return val.replace('\xa0', ' ').strip()
    return val

df["Country"] = df["Country"].apply(clean_cell)
df["To currency"] = df["To currency"].apply(clean_cell)

# Detect and format date columns
date_cols = [
    col for col in df.columns
    if isinstance(col, (str, datetime)) and (
        (isinstance(col, str) and col.count('.') == 2 and len(col) == 10)
        or isinstance(col, datetime)
    )
]

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
        seen = set()  # track (currency) per country only

        for _, row in group.iterrows():
            currency = str(row["To currency"]).strip().upper()
            if not currency or currency not in CURRENCY_META:
                continue

            curr_id, symbol = CURRENCY_META[currency]
            currency_key = currency  # track only within current country

            is_first_for_currency = currency_key not in seen

            for old_date in date_cols_desc:
                rate = row.get(old_date)
                if pd.isna(rate):
                    continue
                iso_date = iso_map[old_date]
                rows.append({
                    "id": curr_id if is_first_for_currency else "",
                    "Symbol": symbol if is_first_for_currency else "",
                    "Active": "TRUE" if is_first_for_currency else "",
                    "Rates/Date": iso_date,
                    "Rates/Inverse Company Rate": round(rate, 2),
                    "Rates/Currency": currency
                })
                if is_first_for_currency:
                    seen.add(currency_key)
                    is_first_for_currency = False  # Only once

        if rows:
            out = pd.DataFrame(rows)
            out.to_excel(writer, index=False, sheet_name=country[:31])

print(f"✅ Done. Saved '{OUTPUT_FILE}' with one-time currency tagging per country.")

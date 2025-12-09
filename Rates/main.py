import pandas as pd
from datetime import datetime
import re

INPUT_FILE = "sap_rates_today.xlsx"   # normalized file
OUTPUT_FILE = "odoo_currency_rates_all_one_sheet.xlsx"

# Date range (DD.MM.YYYY)
DATE_FROM = "25.11.2025"
DATE_TO   = "08.12.2025"
date_from_dt = datetime.strptime(DATE_FROM, "%d.%m.%Y")
date_to_dt   = datetime.strptime(DATE_TO, "%d.%m.%Y")

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

# SAP Country -> company codes
COUNTRY_TO_COMPANIES = {
    "India": ["IN10", "IN20", "IN30"],
    "Tanzania": ["TZ10"],
    "Kenya": ["KE10", "KE20", "KE30", "KE40", "KE50"],
    "Uganda": ["UG10", "UG20"],
    "Ethiopia": ["ET10"],
    "Rwanda": ["RW10"],
    "UK": ["GB10"],
    "Malawi": ["MW10"],
    "Dubai": ["AE00", "AE10", "AE20", "TBG"],
    "Zambia": ["ZM10"],
    "Mauritius": ["MU10", "MU20", "MU30", "MU40", "MU50"],
    "South Africa": ["ZA10", "ZA20"],
    "USA": ["US10"],
    "Ghana": ["GH10"],
    "Canada": ["CA10"],
    "Burundi": ["BR10"],
    "Nigeria": ["NG10"],
    "Mozambique": ["MZ10"],
    "Namibia": ["NA10"],
    "Swaziland": ["SZ10"],
    "Sierra Leone": ["SL10"],
    # "Saudi Arabia": ["SA10"],
}

def clean_cell(val):
    if isinstance(val, str):
        return val.replace("\xa0", " ").strip()
    return val

# ─────── Read normalized file ─────────────────────────────────
df = pd.read_excel(INPUT_FILE)
df.columns = [str(c).strip() for c in df.columns]

# Clean object columns safely
for i in range(df.shape[1]):
    if df.iloc[:, i].dtype == object:
        df.iloc[:, i] = df.iloc[:, i].apply(clean_cell)

country_col = "Country"
from_col = "From Currency"
to_col = "To currency"

for req in (country_col, from_col, to_col):
    if req not in df.columns:
        raise ValueError(f"Missing column '{req}'. Found: {list(df.columns)[:30]} ...")

# Date columns are DD.MM.YYYY strings
date_cols_all = [c for c in df.columns if re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", str(c))]

# Filter by date range
date_cols = []
for c in date_cols_all:
    dt = datetime.strptime(c, "%d.%m.%Y")
    if date_from_dt <= dt <= date_to_dt:
        date_cols.append(c)

if not date_cols:
    raise ValueError(
        f"No date columns matched {DATE_FROM} to {DATE_TO}. "
        f"Found date columns like: {date_cols_all[:5]}"
    )

date_cols = sorted(date_cols, key=lambda x: datetime.strptime(x, "%d.%m.%Y"))
date_cols_desc = list(reversed(date_cols))
iso_map = {c: datetime.strptime(c, "%d.%m.%Y").strftime("%Y-%m-%d") for c in date_cols}

# ─────── Build rows (no header logic here) ────────────────────
raw_rows = []

for _, r in df.iterrows():
    sap_country = str(r.get(country_col, "")).strip()
    to_cur = str(r.get(to_col, "")).strip().upper()

    if not sap_country or not to_cur:
        continue
    if to_cur not in CURRENCY_META:
        continue

    companies = COUNTRY_TO_COMPANIES.get(sap_country)
    if not companies:
        continue

    for dcol in date_cols_desc:
        rate = r.get(dcol)
        if pd.isna(rate):
            continue

        iso_date = iso_map[dcol]

        for company in companies:
            raw_rows.append({
                "Rates/Date": iso_date,
                "Rates/Inverse Company Rate": float(rate),
                "Rates/Currency": to_cur,
                "Rates/Company": company,
            })

out = pd.DataFrame(raw_rows)

# Remove duplicates BEFORE grouping (this fixes Odoo unique constraint issues)
out = out.drop_duplicates(
    subset=["Rates/Date", "Rates/Currency", "Rates/Company"],
    keep="last"
)

# ─────── Force grouping by currency (THIS fixes your header issue) ───────
# Sort so currencies are grouped together, and within each currency newest date first
out = out.sort_values(
    by=["Rates/Currency", "Rates/Date", "Rates/Company"],
    ascending=[True, False, True]
).reset_index(drop=True)

# Now add the header row once per currency group (top of each block)
out["External ID"] = ""
out["Symbol"] = ""
out["Active"] = ""

for cur, (ext_id, sym) in CURRENCY_META.items():
    # first row index of that currency block
    idx = out.index[out["Rates/Currency"] == cur]
    if len(idx) == 0:
        continue
    first_idx = idx[0]
    out.at[first_idx, "External ID"] = ext_id
    out.at[first_idx, "Symbol"] = sym
    out.at[first_idx, "Active"] = "TRUE"

# Final column order
out = out[[
    "External ID", "Symbol", "Active",
    "Rates/Date", "Rates/Inverse Company Rate",
    "Rates/Currency", "Rates/Company"
]]

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    out.to_excel(writer, index=False, sheet_name="Rates")

print(f"✅ Done. Saved '{OUTPUT_FILE}'. Grouped by currency with header at top of each currency block.")

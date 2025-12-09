import pandas as pd
from datetime import datetime
import re

INPUT_FILE = "sap_rates_today.xlsx"

OUT_COMPANY = "odoo_rates_per_company.xlsx"          # your current one (optional)
OUT_USD_CONS = "odoo_rates_consolidation_USD.xlsx"   # NEW
OUT_AED_CONS = "odoo_rates_consolidation_AED.xlsx"   # NEW

DATE_FROM = "04.01.2016"
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
    "Saudi Arabia": ["SA10"],
}

def clean_cell(val):
    if isinstance(val, str):
        return val.replace("\xa0", " ").strip()
    return val

def add_currency_headers_grouped(df_out, currency_col="Rates/Currency"):
    """Sort by currency then date desc; add id/symbol/active only on first row of each currency block."""
    df_out = df_out.sort_values(
        by=[currency_col, "Rates/Date", "Rates/Currency"],
        ascending=[True, False, True]
    ).reset_index(drop=True)

    df_out["id"] = ""
    df_out["Symbol"] = ""
    df_out["Active"] = ""

    for cur, (ext_id, sym) in CURRENCY_META.items():
        idx = df_out.index[df_out[currency_col] == cur]
        if len(idx) == 0:
            continue
        first_idx = idx[0]
        df_out.at[first_idx, "id"] = ext_id
        df_out.at[first_idx, "Symbol"] = sym
        df_out.at[first_idx, "Active"] = "TRUE"

    return df_out

# ─────── Read normalized file ────────────────────────────────
df = pd.read_excel(INPUT_FILE)
df.columns = [str(c).strip() for c in df.columns]

for i in range(df.shape[1]):
    if df.iloc[:, i].dtype == object:
        df.iloc[:, i] = df.iloc[:, i].apply(clean_cell)

country_col = "Country"
from_col = "From Currency"
to_col = "To currency"

for req in (country_col, from_col, to_col):
    if req not in df.columns:
        raise ValueError(f"Missing column '{req}'. Found: {list(df.columns)[:30]} ...")

date_cols_all = [c for c in df.columns if re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", str(c))]

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

# ─────── 1) Build PER-COMPANY output (what you already have) ──
company_rows = []
for _, r in df.iterrows():
    sap_country = str(r.get(country_col, "")).strip()
    to_cur = str(r.get(to_col, "")).strip().upper()
    if not sap_country or not to_cur or to_cur not in CURRENCY_META:
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
            company_rows.append({
                "Rates/Date": iso_date,
                "Rates/Inverse Company Rate": round(float(rate), 2),
                "Rates/Currency": to_cur,
                "Rates/Company": company,
            })

out_company = pd.DataFrame(company_rows).drop_duplicates(
    subset=["Rates/Date", "Rates/Currency", "Rates/Company"],
    keep="last"
)

# add headers by currency block
out_company = out_company.sort_values(
    by=["Rates/Currency", "Rates/Date", "Rates/Company"],
    ascending=[True, False, True]
).reset_index(drop=True)
out_company["External ID"] = ""
out_company["Symbol"] = ""
out_company["Active"] = ""

for cur, (ext_id, sym) in CURRENCY_META.items():
    idx = out_company.index[out_company["Rates/Currency"] == cur]
    if len(idx) == 0:
        continue
    first_idx = idx[0]
    out_company.at[first_idx, "External ID"] = ext_id
    out_company.at[first_idx, "Symbol"] = sym
    out_company.at[first_idx, "Active"] = "TRUE"

out_company = out_company[[
    "External ID", "Symbol", "Active",
    "Rates/Date", "Rates/Inverse Company Rate",
    "Rates/Currency", "Rates/Company"
]]

# ─────── 2) Build CONSOLIDATION outputs (USD + AED) ──────────
def build_consolidation_file(target_to_currency: str, out_path: str):
    """
    Consolidation file: only rows where To currency == target_to_currency.
    Columns: id, Symbol, Active, Rates/Date, Rates/Company Rate, Rates/Currency
    """
    target_to_currency = target_to_currency.upper()

    cons_rows = []
    for _, r in df.iterrows():
        to_cur = str(r.get(to_col, "")).strip().upper()
        from_cur = str(r.get(from_col, "")).strip().upper()

        # We only want "all vs USD" or "all vs AED"
        if to_cur != target_to_currency:
            continue

        # From currency becomes the "Rates/Currency" we are defining rates for
        if not from_cur or from_cur not in CURRENCY_META:
            continue

        for dcol in date_cols_desc:
            rate = r.get(dcol)
            if pd.isna(rate):
                continue

            cons_rows.append({
                "Rates/Date": iso_map[dcol],
                "Rates/Company Rate": round(float(rate), 2),
                "Rates/Currency": from_cur,
            })

    out_cons = pd.DataFrame(cons_rows)

    # remove duplicates per day+currency (no company column here)
    out_cons = out_cons.drop_duplicates(
        subset=["Rates/Date", "Rates/Currency"],
        keep="last"
    )

    # group by currency + add headers at top of each currency block
    out_cons = add_currency_headers_grouped(out_cons, currency_col="Rates/Currency")

    out_cons = out_cons[[
        "id", "Symbol", "Active",
        "Rates/Date", "Rates/Company Rate",
        "Rates/Currency"
    ]]

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        out_cons.to_excel(writer, index=False, sheet_name="Rates")

    print(f"✅ Consolidation {target_to_currency} saved: {out_path}")

# Write files
with pd.ExcelWriter(OUT_COMPANY, engine="openpyxl") as writer:
    out_company.to_excel(writer, index=False, sheet_name="Rates")
print(f"✅ Per-company file saved: {OUT_COMPANY}")

build_consolidation_file("USD", OUT_USD_CONS)
build_consolidation_file("AED", OUT_AED_CONS)

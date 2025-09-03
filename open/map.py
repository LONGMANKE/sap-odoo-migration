import pandas as pd
from pathlib import Path

# ========================
# Settings
# ========================
INPUT_FILE = "input.xlsx"                 # your vendor file
OUTPUT_FILE = "Odoo_Vendor_Bills.xlsx"    # output file
SHEET = 0                                  # sheet index or name
JOURNAL_NAME = "Open Vendor Invoices"
DEFAULT_GL = "2060900000 Suspense account"   # <-- fixed account for all lines
LABEL_TEMPLATE = "Payable to Vendor as on {date}-{docnum}"  # used if no Text column


def normalize_cols(df: pd.DataFrame) -> dict:
    def norm(s: str) -> str:
        return s.strip().lower().replace("  ", " ")
    return {norm(c): c for c in df.columns}


def pick(cols: dict, *aliases, required=False):
    for a in aliases:
        k = a.strip().lower()
        if k in cols:
            return cols[k]
    if required:
        raise KeyError(f"Missing required column. Tried: {aliases}")
    return None


def main():
    xl = pd.ExcelFile(INPUT_FILE)
    src = xl.parse(SHEET)

    # --- map columns (robust to slight header differences) ---
    cols = normalize_cols(src)
    col_vendor = pick(cols, "vendor name", "invoice partner display name", required=True)
    col_docnum = pick(cols, "document number", "document no", "document number", required=True)
    col_date = pick(cols, "posting date", "invoice/bill date", required=True)
    col_curr = pick(cols, "general ledger currency", "currency", required=True)
    col_text = pick(cols, "text", "invoice lines/label")  # optional
    col_amount = pick(
        cols,
        "amount in doc. curr.",
        "amount in doc curr",
        "invoice lines/unit price",
        required=True,
    )

    # --- build base frame in Odoo column order ---
    out = pd.DataFrame()
    out["Invoice Partner Display Name"] = src[col_vendor]
    out["Reference"] = src[col_docnum].apply(lambda x: str(int(x)) if pd.notna(x) else "")
    out["Invoice/Bill Date"] = pd.to_datetime(src[col_date], errors="coerce").dt.strftime("%Y-%m-%d")
    out["Journal"] = JOURNAL_NAME
    out["Currency"] = src[col_curr]

    if col_text:
        out["Invoice lines/Label"] = src[col_text].astype(str)
    else:
        out["Invoice lines/Label"] = out.apply(
            lambda r: LABEL_TEMPLATE.format(date=r["Invoice/Bill Date"], docnum=int(r["Reference"]))
            if pd.notna(r["Reference"]) else "",
            axis=1,
        )

    # >>> Fixed account for ALL rows
    out["Invoice lines/Account"] = DEFAULT_GL

    # Flip sign for Unit Price (negative -> positive, positive -> negative)
    out["Invoice lines/Unit Price"] = pd.to_numeric(src[col_amount], errors="coerce") * -1

    # drop rows without essential values
    out = out.dropna(subset=["Reference", "Invoice/Bill Date"])

    # sort so groups are together
    out = out.sort_values(
        ["Reference", "Invoice Partner Display Name", "Invoice/Bill Date"]
    ).reset_index(drop=True)

    # --- blank repeated group columns (show once per group) ---
    group_cols = [
        "Invoice Partner Display Name",
        "Reference",
        "Invoice/Bill Date",
        "Journal",
        "Currency",
    ]

    # Ensure group columns can safely accept strings (avoid dtype warnings)
    out[group_cols] = out[group_cols].astype("string")

    # Blank repeated values for group display
    for ref, idx in out.groupby("Reference").groups.items():
        idx = list(idx)
        if len(idx) > 1:
            out.loc[idx[1:], group_cols] = ""

    # save
    ext = Path(OUTPUT_FILE).suffix.lower()
    if ext in (".xlsx", ".xls"):
        out.to_excel(OUTPUT_FILE, index=False)
    else:
        out.to_csv(Path(OUTPUT_FILE).with_suffix(".csv"), index=False)

    print(f"✅ Wrote grouped file: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()

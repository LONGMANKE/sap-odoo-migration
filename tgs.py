import os, re, sys
import pandas as pd

# ---- SETTINGS --------------------------------------------------------------
ONLY_FILL_EMPTY = True  # set to False to overwrite existing tag_ids
OUT_FILE = "coa_mapped.xlsx"
# ---------------------------------------------------------------------------

def find_file(*candidates):
    for c in candidates:
        if os.path.exists(c):
            return c
    return None

def load_any(path_xlsx, path_csv):
    path = find_file(path_xlsx, path_csv)
    if not path:
        raise FileNotFoundError(f"Neither '{path_xlsx}' nor '{path_csv}' found.")
    if path.endswith(".xlsx"):
        return pd.read_excel(path), path
    return pd.read_csv(path), path

def normalize_columns(df):
    # trim and lower for matching; keep original names
    mapping = {c: re.sub(r"\s+", " ", c.strip()).lower() for c in df.columns}
    return mapping

def pick_column(df, *candidates):
    norm = normalize_columns(df)
    for want in candidates:
        for orig, low in norm.items():
            if low == want:
                return orig
    # Try contains for flexible matches (e.g., 'tag name' variations)
    for want in candidates:
        for orig, low in norm.items():
            if want in low:
                return orig
    return None

def first_three_digits(val):
    if pd.isna(val):
        return None
    s = str(val).strip()
    # remove .0 if it came from Excel numeric formatting
    if re.match(r"^\d+\.0$", s):
        s = s[:-2]
    # accept first three digits from start; fallback: first any 3 digits
    m = re.match(r"^\s*(\d{3})", s)
    if m:
        return m.group(1)
    m2 = re.search(r"(\d{3})", s)
    return m2.group(1) if m2 else None

def main():
    # Load files
    coa_df, coa_path = load_any("coa.xlsx", "coa.csv")
    tags_df, tags_path = load_any("tags.xlsx", "tags.csv")

    # Identify columns robustly
    id_col    = pick_column(coa_df, "id")
    name_col  = pick_column(coa_df, "name")
    code_col  = pick_column(coa_df, "code")
    tagids_col= pick_column(coa_df, "tag_ids", "tag ids", "tags", "account tag")
    if not code_col:
        sys.exit("❌ Couldn't find a 'code' column in COA (case/space-insensitive).")
    if not tagids_col:
        # create one if missing
        tagids_col = "tag_ids"
        if "tag_ids" not in coa_df.columns:
            coa_df[tagids_col] = None

    tagname_col = pick_column(tags_df, "tag name", "tag_name", "name")
    if not tagname_col:
        sys.exit("❌ Couldn't find a 'Tag Name' column in TAGS.")

    # Build map from first-3 digits -> full Tag Name
    tags_df = tags_df.copy()
    tags_df[tagname_col] = tags_df[tagname_col].astype(str)
    tags_df["__prefix__"] = tags_df[tagname_col].map(first_three_digits)
    tag_map = (
        tags_df.dropna(subset=["__prefix__"])
               .drop_duplicates("__prefix__", keep="first")
               .set_index("__prefix__")[tagname_col]
               .to_dict()
    )
    if not tag_map:
        sys.exit("❌ No valid 3-digit prefixes found in tags file.")

    # Prepare COA
    coa_df = coa_df.copy()
    coa_df[code_col] = coa_df[code_col].astype(str)
    coa_df["__prefix__"] = coa_df[code_col].map(first_three_digits)

    before_nonempty = (coa_df[tagids_col].astype(str).str.strip() != "").sum()

    def compute_tag(row):
        existing = str(row[tagids_col]).strip()
        mapped = tag_map.get(row["__prefix__"])
        if ONLY_FILL_EMPTY:
            return existing if existing not in ("", "nan", "None") else mapped
        return mapped if mapped else (existing if existing not in ("", "nan", "None") else None)

    coa_df[tagids_col] = coa_df.apply(compute_tag, axis=1)

    # Stats
    after_nonempty = (coa_df[tagids_col].astype(str).str.strip() != "").sum()
    filled = after_nonempty - before_nonempty

    # Save
    coa_df.drop(columns=["__prefix__"], errors="ignore").to_excel(OUT_FILE, index=False)

    # Report a few sample matches
    sample = coa_df.loc[
        coa_df[tagids_col].astype(str).str.strip() != ""
    ][[code_col, tagids_col]].head(10)

    print("✅ Mapping complete.")
    print(f"   COA file:  {coa_path}")
    print(f"   TAGS file: {tags_path}")
    print(f"   Output:    {OUT_FILE}")
    print(f"   Filled/updated tag_ids for {filled} row(s) out of {len(coa_df)}.")
    if not sample.empty:
        print("\n   Sample mapped rows:")
        for _, r in sample.iterrows():
            print(f"     code={r[code_col]}  ->  tag_ids='{r[tagids_col]}'")
    else:
        print("   (No rows ended up with non-empty tag_ids. Check that your tag names start with 3 digits like '101 ...').")

if __name__ == "__main__":
    main()

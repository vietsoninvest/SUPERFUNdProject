#!/usr/bin/env python3
import argparse, re, unicodedata
import pandas as pd

def normalize_name(name: str) -> str:
    if not isinstance(name, str):
        return ""
    s = str(name).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.upper().replace("&", " AND ")
    s = re.sub(r"[^A-Z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def blocking_key(norm: str, k: int = 6) -> str:
    comp = re.sub(r"[^A-Z0-9]", "", norm)
    return comp[:k]

def canonical_lookup_map(columns, requested_names):
    """Map user-requested extra columns to actual df columns (case/space-insensitive)."""
    canon = {c.strip().lower(): c for c in columns}
    out = []
    for req in requested_names:
        key = req.strip().lower()
        if key in canon:
            out.append(canon[key])
    return out

def main():
    ap = argparse.ArgumentParser(description="Compare company names across two CSVs using 'contains' logic (case/punct-insensitive).")
    ap.add_argument("--file1", required=True)
    ap.add_argument("--file2", required=True)
    ap.add_argument("--col1", default="Name/Kind of Investment Item")
    ap.add_argument("--col2", default="Name/Kind of Investment Item")
    ap.add_argument("--include-cols-file1", default="Listed Country", help="Comma-separated extra columns from file1 to include")
    ap.add_argument("--output", required=True)
    ap.add_argument("--unmatched2")
    ap.add_argument("--unmatched1")
    ap.add_argument("--encoding1", default=None)
    ap.add_argument("--encoding2", default=None)
    ap.add_argument("--block-len", type=int, default=6)
    args = ap.parse_args()

    f1 = pd.read_csv(args.file1, encoding=args.encoding1)
    f2 = pd.read_csv(args.file2, encoding=args.encoding2)

    if args.col1 not in f1.columns: raise SystemExit(f"Column '{args.col1}' not found in file1")
    if args.col2 not in f2.columns: raise SystemExit(f"Column '{args.col2}' not found in file2")

    # Keep originals for output, but use renamed for processing
    f1 = f1.rename(columns={args.col1: "name_f1"}).copy()
    f2 = f2.rename(columns={args.col2: "name_f2"}).copy()

    # Normalize names
    f1["_norm"] = f1["name_f1"].apply(normalize_name)
    f2["_norm"] = f2["name_f2"].apply(normalize_name)
    f1 = f1[f1["_norm"] != ""].copy()
    f2 = f2[f2["_norm"] != ""].copy()

    # Build blocking keys
    f1["_block"] = f1["_norm"].apply(lambda s: blocking_key(s, args.block_len))
    f2["_block"] = f2["_norm"].apply(lambda s: blocking_key(s, args.block_len))

    # Resolve extra columns robustly (case/space-insensitive)
    requested = [c for c in (args.include_cols_file1 or "").split(",") if c.strip()]
    extra_cols = canonical_lookup_map(f1.columns, requested)

    # Pre-sort f1 to favor non-null extra cols (esp. Listed Country) when duplicates exist
    sort_cols = []
    for c in extra_cols:
        sort_cols.append(f"{c}__nn")
        f1[f"{c}__nn"] = ~f1[c].isna()  # True first
    if sort_cols:
        f1 = f1.sort_values(by=sort_cols, ascending=False).copy()
        f1.drop(columns=sort_cols, inplace=True)

    # Candidate pairs via block join
    cand = f1[["_norm","_block","name_f1"] + extra_cols].merge(
        f2[["_norm","_block","name_f2"]],
        on="_block", how="inner", suffixes=("_f1","_f2")
    )

    def is_match(a: str, b: str) -> bool:
        return a in b or b in a

    matches = cand[cand.apply(lambda r: is_match(r["_norm_f1"], r["_norm_f2"]), axis=1)]

    # Build output with original columns
    out = matches[["name_f1"] + extra_cols + ["name_f2"]].drop_duplicates()
    out = out.rename(columns={"name_f1": args.col1, "name_f2": args.col2})
    out = out.sort_values(by=[args.col1, args.col2]).reset_index(drop=True)

    out.to_csv(args.output, index=False)
    print(f"✔ Wrote {len(out)} matched rows to: {args.output}")

    # Unmatched from file2
    if args.unmatched2:
        matched_f2 = set(out[args.col2].dropna().astype(str))
        um2 = f2[~f2["name_f2"].astype(str).isin(matched_f2)].copy()
        um2 = um2.rename(columns={"name_f2": args.col2})
        um2.to_csv(args.unmatched2, index=False)
        print(f"✔ Wrote {len(um2)} unmatched rows from file2 to: {args.unmatched2}")

    # Unmatched from file1
    if args.unmatched1:
        matched_f1 = set(out[args.col1].dropna().astype(str))
        um1 = f1[~f1["name_f1"].astype(str).isin(matched_f1)].copy()
        um1 = um1.rename(columns={"name_f1": args.col1})
        um1.to_csv(args.unmatched1, index=False)
        print(f"✔ Wrote {len(um1)} unmatched rows from file1 to: {args.unmatched1}")

if __name__ == "__main__":
    main()

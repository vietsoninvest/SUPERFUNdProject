#!/usr/bin/env python3
import os, time, json, re, argparse
from typing import Optional, List, Dict, Tuple
import pandas as pd
import requests

OPENFIGI_URL = "https://api.openfigi.com/v3/mapping"
ISIN_RE = re.compile(r"\b[A-Za-z]{2}[A-Za-z0-9]{9}\d\b")

# ---------- helpers ----------
def _is_nonempty_stockid(x) -> bool:
    """Treat empty/placeholder Stock IDs as empty (NOT invalid)."""
    if pd.isna(x):
        return False
    s = str(x).strip()
    return s != "" and s.lower() not in {"n/a", "na", "none", "null"}

# ---------- ISIN validator ----------
def _luhn_check_isin(isin: str) -> bool:
    if len(isin) != 12:
        return False
    def expand(ch: str) -> str:
        if ch.isdigit():
            return ch
        return str(ord(ch.upper()) - 55)
    expanded = "".join(expand(c) for c in isin[:-1]) + isin[-1]
    total, dbl = 0, True
    for d in expanded[-2::-1]:
        n = int(d)
        n = n*2 if dbl else n
        total += (n // 10) + (n % 10)
        dbl = not dbl
    return (total + int(expanded[-1])) % 10 == 0

def extract_isin(val: str) -> Optional[str]:
    if not isinstance(val, str) or not val:
        return None
    m = ISIN_RE.search(val)
    if not m:
        return None
    cand = m.group(0).upper()
    return cand if _luhn_check_isin(cand) else None

# ---------- OpenFIGI lookup ----------
def openfigi_map_isins(isins: List[str], api_key: Optional[str], batch_size: int, sleep_sec: float) -> pd.DataFrame:
    headers = {"Content-Type": "application/json"}
    if api_key:
        headers["X-OPENFIGI-APIKEY"] = api_key
    rows: List[Dict] = []
    for i in range(0, len(isins), batch_size):
        batch = isins[i:i+batch_size]
        payload = [{"idType":"ID_ISIN","idValue":isin} for isin in batch]
        attempt = 0
        while True:
            attempt += 1
            try:
                resp = requests.post(OPENFIGI_URL, headers=headers, data=json.dumps(payload), timeout=60)
                if resp.status_code in (429,500,502,503,504) and attempt <= 6:
                    ra = resp.headers.get("Retry-After")
                    delay = float(ra) if ra else min(8.0, (2**(attempt-1))*sleep_sec)
                    time.sleep(delay)
                    continue
                resp.raise_for_status()
            except requests.RequestException:
                if attempt <= 6:
                    delay = min(8.0, (2**(attempt-1))*sleep_sec)
                    time.sleep(delay)
                    continue
                else:
                    raise
            break
        results = resp.json()
        for req_isin, res in zip(batch, results):
            data = res.get("data") or []
            if not data:
                rows.append({
                    "ISIN": req_isin,
                    "ticker": None, "exchCode": None, "mic": None,
                    "name": None, "marketSector": None, "compositeFIGI": None
                })
            else:
                for d in data:
                    rows.append({
                        "ISIN": req_isin,
                        "ticker": d.get("ticker"),
                        "exchCode": d.get("exchCode"),
                        "mic": d.get("mic"),
                        "name": d.get("name"),
                        "marketSector": d.get("marketSector"),
                        "compositeFIGI": d.get("compositeFIGI"),
                    })
        time.sleep(sleep_sec)
    return pd.DataFrame(rows)

# ---------- Load Bloomberg exchange code map ----------
def load_exchcode_map(path: Optional[str]) -> pd.DataFrame:
    """
    Load your 'bloomberg-exchange-codes-full.csv' and return:
      exchCode, listed_country, priority
    Handles BOM/encoding issues and variant header names.
    """
    if not path:
        return pd.DataFrame(columns=["exchCode","listed_country","priority"])

    df = None
    for enc in ("utf-8-sig", "cp1252", "utf-8"):
        try:
            df = pd.read_csv(path, encoding=enc)
            break
        except Exception:
            df = None
    if df is None:
        raise ValueError(f"Could not read mapping file: {path}")

    # Normalize headers
    def norm(col: str) -> str:
        c = col.replace("\ufeff", "").replace("ï»¿", "").strip().lower()
        for ch in " -_/()[]{}":
            c = c.replace(ch, "")
        return c
    df.rename(columns={c: norm(c) for c in df.columns}, inplace=True)

    cand_code = [c for c in df.columns if c in {"bbgcode","bbg_code","code"}]
    cand_iso  = [c for c in df.columns if c in {"iso_alpha2","isoalpha2","iso2","iso"}]

    if not cand_code or not cand_iso:
        raise ValueError(f"Could not find 'BBG_Code' and 'ISO_Alpha2' (or equivalents). Columns present: {list(df.columns)}")

    col_code = cand_code[0]
    col_iso  = cand_iso[0]

    out = pd.DataFrame({
        "exchCode": df[col_code].astype(str).str.strip().str.upper(),
        "listed_country": df[col_iso].astype(str).str.strip().str.upper()
    })
    out["priority"] = 1  # default: all venues same priority (refine later if you want)
    return out.drop_duplicates()

# ---------- Primary selection helpers ----------
def _ticker_plainness_score(ticker: Optional[str]) -> Tuple[int, int]:
    """Lower tuple sorts better: (nonalpha_count, length)."""
    if not isinstance(ticker, str) or not ticker:
        return (999, 999)
    nonalpha = sum(1 for c in ticker if not c.isalpha())
    return (nonalpha, len(ticker))

def mark_primary_rows(out: pd.DataFrame, exch_map: pd.DataFrame) -> pd.DataFrame:
    """
    Adds columns:
      - HomeCountry (from ISIN prefix)
      - mapped_country (from exch_map)
      - mapped_priority (int; lower is better)
      - Primary (bool)
      - PrimaryListingCountry (string) on primary row else NaN
      - PrimaryTicker (string) repeated across group
    """
    out = out.copy()
    out["HomeCountry"] = out["ISIN"].astype(str).str[:2].str.upper()

    # Normalize exchCode and merge mapping
    if "exchCode" not in out.columns:
        out["exchCode"] = None
    out["exchCode_norm"] = out["exchCode"].astype("string").str.strip().str.upper()

    if exch_map is not None and not exch_map.empty:
        m = exch_map.rename(columns={"priority": "mapped_priority"})
        out = out.merge(m, how="left", left_on="exchCode_norm", right_on="exchCode", suffixes=("", "_map"))
        out.drop(columns=["exchCode_map"], inplace=True, errors="ignore")
        out.rename(columns={"listed_country": "mapped_country"}, inplace=True)
    else:
        out["mapped_country"] = pd.NA
        out["mapped_priority"] = pd.NA

    out["Primary"] = False
    out["PrimaryListingCountry"] = pd.NA

    def choose_primary(group: pd.DataFrame) -> Optional[int]:
        g = group.copy()
        # Prefer rows mapped to the HomeCountry
        cand = g[g["mapped_country"].fillna("") == g["HomeCountry"].fillna("")]
        if not cand.empty:
            cand = cand.assign(_plain=cand["ticker"].apply(_ticker_plainness_score)) \
                       .sort_values(by=["mapped_priority", "_plain"], ascending=[True, True], na_position="last")
            return int(cand.index[0])

        # Otherwise, lowest priority overall; tie-break by plainness
        cand = g.assign(_plain=g["ticker"].apply(_ticker_plainness_score)) \
                .sort_values(by=["mapped_priority", "_plain"], ascending=[True, True], na_position="last")
        if cand["mapped_priority"].isna().all():
            cand = g.assign(_plain=g["ticker"].apply(_ticker_plainness_score)) \
                    .sort_values(by=["_plain"], ascending=True, na_position="last")
        return int(cand.index[0])

    # Pick one primary per ISIN
    primary_rows: Dict[str, int] = {}
    for isin, grp in out.groupby("ISIN", dropna=False):
        if not isinstance(isin, str) or not isin:
            continue
        idx = choose_primary(grp)
        if idx is not None:
            primary_rows[isin] = idx

    if primary_rows:
        prim_df = out.loc[list(primary_rows.values())].copy()
        out.loc[prim_df.index, "Primary"] = True
        out.loc[prim_df.index, "PrimaryListingCountry"] = prim_df["mapped_country"].where(
            prim_df["mapped_country"].notna(),
            prim_df["HomeCountry"]  # fallback if mapping missing
        )
        # Broadcast PrimaryTicker to all rows of the ISIN group
        primary_ticker_by_isin = prim_df.set_index("ISIN")["ticker"].to_dict()
        out["PrimaryTicker"] = out["ISIN"].map(primary_ticker_by_isin)
    else:
        out["PrimaryTicker"] = pd.NA

    out.drop(columns=["exchCode_norm"], inplace=True, errors="ignore")
    return out

# ---------- Main ----------
def main():
    ap = argparse.ArgumentParser(description="Extract ISINs, fetch OpenFIGI (all listings), and mark primary listings using Bloomberg exchange codes.")
    ap.add_argument("--input", required=True, help="Source CSV (must contain 'Stock ID')")
    ap.add_argument("--output", required=True, help="Output CSV (all listings + primary columns)")
    ap.add_argument("--invalid-output", help="Optional CSV path for rows with non-empty Stock ID but invalid ISIN")
    ap.add_argument("--primary-output", help="Optional CSV path for rows marked Primary=True")
    ap.add_argument("--stockid-col", default="Stock ID")
    ap.add_argument("--name-col", default="Name/Kind of Investment Item")
    ap.add_argument("--encoding", default=None)
    ap.add_argument("--api-key", default=os.getenv("OPENFIGI_API_KEY"))
    ap.add_argument("--batch-size", type=int, default=50)
    ap.add_argument("--sleep-sec", type=float, default=0.4)
    ap.add_argument("--exchcode-map", required=True, help="Path to 'bloomberg-exchange-codes-full.csv'")
    args = ap.parse_args()

    # Load source
    df_src = pd.read_csv(args.input, encoding=args.encoding, dtype={args.stockid_col:"string"}, engine="python")

    # Keep only non-empty Stock IDs (placeholders treated as empty)
    df_src = df_src[df_src[args.stockid_col].apply(_is_nonempty_stockid)].copy()

    # Extract ISINs
    df_src["ISIN"] = df_src[args.stockid_col].apply(extract_isin)
    valid_mask = df_src["ISIN"].notna()
    invalid_mask = ~valid_mask

    # Write invalid ISINs: ONLY rows where Stock ID is non-empty but ISIN invalid
    if args.invalid_output:
        invalid_rows = df_src.loc[invalid_mask & df_src[args.stockid_col].apply(_is_nonempty_stockid),
                                  [args.name_col, args.stockid_col]]
        if not invalid_rows.empty:
            invalid_rows.to_csv(args.invalid_output, index=False)
            print(f"✔ Wrote {len(invalid_rows)} invalid-ISIN rows to: {args.invalid_output}")
        else:
            print("✔ No invalid ISIN rows → no invalid file produced")

    rows_valid = int(valid_mask.sum())
    rows_invalid = int(invalid_mask.sum())
    total_nonempty = len(df_src)

    unique_isins = sorted(set(df_src.loc[valid_mask, "ISIN"]))
    df_map = openfigi_map_isins(unique_isins, args.api_key, args.batch_size, args.sleep_sec) if unique_isins else pd.DataFrame()

    # Join all listings
    out = df_src.loc[valid_mask, [args.name_col, args.stockid_col, "ISIN"]].merge(df_map, on="ISIN", how="left")

    # Ensure columns in desired order
    want_cols = [args.name_col, args.stockid_col, "ISIN", "ticker", "exchCode", "mic", "name", "marketSector", "compositeFIGI"]
    for col in want_cols:
        if col not in out.columns:
            out[col] = None
    out = out[want_cols]

    # Load Bloomberg exchCode map and mark primary
    exch_map = load_exchcode_map(args.exchcode_map)
    out = mark_primary_rows(out, exch_map)

    # Write all-listings output
    out.to_csv(args.output, index=False)
    print(f"✔ Wrote {len(out)} rows to: {args.output}")

    # Write primary-only output if requested
    if args.primary_output:
        out_primary = out[out.get("Primary") == True].copy()
        out_primary.to_csv(args.primary_output, index=False)
        print(f"✔ Wrote {len(out_primary)} primary rows to: {args.primary_output}")

    # Summary
    num_primary = int(out["Primary"].sum()) if "Primary" in out.columns else 0
    print("\nSummary:")
    print(f"  Total rows with non-empty Stock ID : {total_nonempty}")
    print(f"  Rows with valid ISIN               : {rows_valid}")
    print(f"  Rows with invalid ISIN             : {rows_invalid}")
    print(f"  Unique valid ISINs queried         : {len(unique_isins)}")
    print(f"  Primary rows marked                : {num_primary}")

if __name__ == "__main__":
    main()

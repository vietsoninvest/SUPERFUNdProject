import pandas as pd
import numpy as np
import os
import requests
import yfinance as yf
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm


def sector_industry_match(input_df):
    # retrieve all stock ids
    stock_ids_mask = ~input_df['Stock ID'].isna()
    stock_ids = input_df['Stock ID'][stock_ids_mask]

    # map each symbol to *all* table indices where it appears (dedupe network calls)
    sym_to_indices = defaultdict(list)
    for idx, sym in stock_ids.items():   # idx = original DataFrame index
        sym_to_indices[sym].append(idx)

    unique_symbols = list(sym_to_indices.keys())
    max_workers = 8

    # call yahoo finance api and retrieve sector and industry for each stock id
    sector_arr, industry_arr, error_id_list = parallel_run(len(input_df), max_workers, unique_symbols, sym_to_indices)

    # run ids in error_id_list through yahoo_resolve function then rerun the fetch sector industry function
    sector_arr, industry_arr, final_err_list = parallel_resolve(max_workers, sector_arr, industry_arr, error_id_list, sym_to_indices)

    # insert new cols Sector & Industry right after Listed Country
    cols = list(input_df.columns)
    insert_at = cols.index("Listed Country") + 1

    if 'Sector' in input_df.columns:
        input_df['Sector'] = sector_arr   # overwrite
    else:
        input_df.insert(insert_at, 'Sector', sector_arr)

    if 'Industry' in input_df.columns:
        input_df['Industry'] = industry_arr   # overwrite
    else:
        input_df.insert(insert_at + 1, 'Industry', industry_arr)

    return input_df, final_err_list

def parallel_resolve(max_workers, sector_arr, industry_arr, error_id_list, sym_to_indices):
    final_err_list = []
    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        revised_syms = {ex.submit(yahoo_resolve_symbol, sym): sym for sym, err in error_id_list}
        futures = {ex.submit(fetch_sector_industry, key): key for key,val in revised_syms.items()}
        for fut in tqdm(as_completed(futures), total=len(futures), desc="Resolving Yahoo IDs -> Rerun fetching"):
            sym, sector, industry, err = fut.result()
            if (sector is None) or (industry is None):
                final_err_list.append((sym, err))
            # write the same result to all rows that have this symbol
            for idx in sym_to_indices[sym]:
                sector_arr[idx] = sector
                industry_arr[idx] = industry
    
    return sector_arr, industry_arr, final_err_list

# 5) run in parallel and fill arrays by original table index
def parallel_run(arr_len, max_workers, unique_symbols, sym_to_indices):
    sector_arr = np.full(arr_len, None, dtype=object)
    industry_arr = np.full(arr_len, None, dtype=object)
    error_id_list = []

    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futures = {ex.submit(fetch_sector_industry, sym): sym for sym in unique_symbols}
        for fut in tqdm(as_completed(futures), total=len(futures), desc="Fetching sectors/industries"):
            sym, sector, industry, err = fut.result()
            if (sector is None) or (industry is None):
                error_id_list.append((sym, err))
            # write the same result to all rows that have this symbol
            for idx in sym_to_indices[sym]:
                sector_arr[idx] = sector
                industry_arr[idx] = industry
                
    return sector_arr, industry_arr, error_id_list

# 4) worker that mimics original logic (uses .info)
def fetch_sector_industry(symbol):
    try:
        t = yf.Ticker(symbol)
        info = t.info
        # Explicitly raise if sector/industry missing
        sector = info.get('sector')
        industry = info.get('industry')
        if sector is None and industry is None:
            raise ValueError("No sector/industry returned")
        return symbol, sector, industry, None
    except Exception as e:
        return symbol, None, None, str(e)

def yahoo_resolve_symbol(query, lang="en-US", region="AU"):
    """
    Resolve a query (SEDOL/ISIN/name/ticker) to a Yahoo symbol using Yahoo's search API.
    Returns the best match symbol (str) or None.
    """
    url = "https://query2.finance.yahoo.com/v1/finance/search"
    params = {"q": query, "lang": lang, "region": region}
    r = requests.get(url, params=params, headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
    r.raise_for_status()
    data = r.json()
    quotes = data.get("quotes") or []
    if not quotes:
        return None

    # pick the first equity-like hit; adjust ranking logic as you like
    for q in quotes:
        sym = q.get("symbol")
        if sym:
            return sym
    return None


if __name__ == "__main__":
    
    df_dict = {
        # "equip": "Data/equip_cleaned.csv",
        # "hesta": "Data/hesta_cleaned.csv",
        # "hostplus": "Data/hostplus_cleaned.csv",
        "vision": "Data/vision_cleaned.csv"
    }

    for fund_name, filepath in df_dict.items():
        input_df = pd.read_csv(filepath)
        # update input df with 2 more cols Sector & Industry
        input_df, final_err_list = sector_industry_match(input_df)
        # rewrite to csv file
        input_df.to_csv(filepath, index=False)  # write back to the same file

        ### still have error ids remained to address 
        print(f"Number of error IDs for {fund_name}: {len(final_err_list)}")



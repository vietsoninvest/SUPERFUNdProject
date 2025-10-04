import pandas as pd
import numpy as np
import os
import re
import requests
import time, random, threading
from collections import deque
import yfinance as yf
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm

from stock_id_filter import file_directory

# helping functions
class RateLimiter:
    """Allow up to max_calls per 'period' seconds (token-bucket-ish)."""
    def __init__(self, max_calls: int, period: float):
        self.max_calls = max_calls
        self.period = period
        self._lock = threading.Lock()
        self._calls = deque()  # timestamps

    def acquire(self):
        while True:
            with self._lock:
                now = time.monotonic()
                # drop old calls
                while self._calls and (now - self._calls[0]) > self.period:
                    self._calls.popleft()
                if len(self._calls) < self.max_calls:
                    self._calls.append(now)
                    return
                # need to sleep until the oldest call expires
                sleep_for = self.period - (now - self._calls[0])
            if sleep_for > 0:
                time.sleep(sleep_for)

# Tune these to your comfort; start conservative
YAHOO_SEARCH_RL = RateLimiter(max_calls=8, period=1.0)   # ~4 req/sec
YF_INFO_RL      = RateLimiter(max_calls=8, period=1.0)   # ~4 req/sec


def _with_retries(fn, attempts=5, base=0.75, cap=8):
    last_exc = None
    for i in range(attempts):
        try:
            return fn()
        except requests.HTTPError as e:
            status = getattr(e.response, "status_code", None)
            # back off on rate limit / server errors
            if status in (429, 500, 502, 503, 504):
                delay = min(cap, base * (2 ** i)) + random.random() * 0.25
                time.sleep(delay)
                last_exc = e
                continue
            raise
        except Exception as e:
            # brief backoff for other transient errors
            delay = min(cap, base * (2 ** i)) + random.random() * 0.25
            time.sleep(delay)
            last_exc = e
            continue
    # give up
    if last_exc:
        raise last_exc

# side function to sequence into chunk
def _chunks(seq, size):
    for i in range(0, len(seq), size):
        yield seq[i:i+size]


######################### End helping functions #########################

def update_final_file():
    pre_path, final_path, pre_final_names = file_directory("pre_final_data", "final_data")
    pre_final_names.remove('vision_final.csv')
    for file_name in pre_final_names:
        super_name = file_name.split("_",1)[0]
        final_name = f"{super_name}_final.csv"

        # read pre cleaned file
        input_df = pd.read_csv(f"{pre_path}/{file_name}")

        # update input df with ID type col
        updated_df, err_list = new_cols_insert(input_df)

        # rewrite to csv file, overwrite if already exists
        updated_df.to_csv(f"{final_path}/{final_name}", mode='w', index=False)  # write to csv

        print(f"Number of errors in fund {super_name}: {len(err_list)}")



def new_cols_insert(input_df):
    
    sector_arr, industry_arr, err_list = sector_industry_arr(input_df)

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

    return input_df, err_list


def sector_industry_arr(input_df):

    # create a dict to link symbols with indices of them in original df
    sym_to_indices = prep_sym_ind_dict(input_df)

    # create sector and industry arr ready for sub series run and batch fill
    sector_arr = np.full(len(input_df), None, dtype=object)
    industry_arr = np.full(len(input_df), None, dtype=object)
    max_workers = 8

    possible_ids = ["ISIN", "BB Ticker", "SEDOL", "Other"]
    
    id_dict = {}
    err_list_total = []
    for id_type in possible_ids:
        id_dict[id_type] = input_df[input_df["ID Type"] == id_type]
        
        if id_dict[id_type].empty:  #  if df empty pass
            continue
        
        sub_unique_id_list = id_dict[id_type]["Stock ID"].unique().tolist()

        if id_type == "ISIN":
            sector_arr, industry_arr, err_list = parallel_fetch(sector_arr, industry_arr, max_workers, 
                                                                sub_unique_id_list, sym_to_indices)
        elif id_type == "BB Ticker":
            sector_arr, industry_arr, err_list = parallel_fetch(sector_arr, industry_arr, max_workers, 
                                                                sub_unique_id_list, sym_to_indices)
            error_ids = [i for i,e in err_list]
            # run ids in error_id_list through yahoo_resolve function then rerun the fetch sector industry function
            sector_arr, industry_arr, err_list = parallel_resolve_fetch(max_workers, sector_arr, 
                                                                industry_arr, error_ids, sym_to_indices)
        elif id_type == "SEDOL" or id_type == "Other":
            sector_arr, industry_arr, err_list = parallel_resolve_fetch(max_workers, sector_arr, 
                                                                industry_arr, sub_unique_id_list, sym_to_indices)
        else:
            continue

        err_list_total.append(err_list)
            
    return sector_arr, industry_arr, err_list_total


# create a dict to link symbols with indices of them in original df
def prep_sym_ind_dict(input_df):
    stock_ids_mask = ~input_df['Stock ID'].isna()
    stock_ids = input_df['Stock ID'][stock_ids_mask]
    stock_ids = stock_ids.astype(str)   # turn all ids into strings

    # Pad only numeric values shorter than 7 digits & id type is indeed SEDOL
    cond_1 = stock_ids.str.isdigit() & (stock_ids.str.len() < 7)
    cond_2 = input_df['ID Type'][stock_ids_mask] == "SEDOL"
    stock_ids.loc[cond_1 * cond_2] = stock_ids.loc[cond_1 * cond_2].str.zfill(7)

    # map each symbol to *all* table indices where it appears (dedupe network calls)
    sym_to_indices = defaultdict(list)
    for idx, sym in stock_ids.items():   # idx = original DataFrame index
        sym_to_indices[sym].append(idx)

    return sym_to_indices


# run in parallel and fill arrays by original table index
def parallel_fetch(sector_arr, industry_arr, max_workers, unique_symbols, sym_to_indices,
                    batch_size=100, batch_pause=1.5):
    error_id_list = []

    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        for batch in _chunks(unique_symbols, batch_size):
            futures = {ex.submit(fetch_sector_industry, sym): sym for sym in batch}
            for fut in tqdm(as_completed(futures), total=len(futures), desc="Fetching sectors/industries"):
                sym, sector, industry, err = fut.result()
                if (sector is None) or (industry is None):
                    error_id_list.append((sym, err))
                # write the same result to all rows that have this symbol
                for idx in sym_to_indices[sym]:
                    sector_arr[idx] = sector
                    industry_arr[idx] = industry

            # brief pause between batches to let any upstream throttles recover
            time.sleep(batch_pause)

    return sector_arr, industry_arr, error_id_list

# run resolve & fetch in parallel
def parallel_resolve_fetch(max_workers, sector_arr, industry_arr, unique_ids, sym_to_indices,
                            batch_size=100, batch_pause=1.5):
    error_id_list = []

    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        for batch in _chunks(unique_ids, batch_size):
            futures = {ex.submit(resolve_fetch_combine, sym): sym for sym in batch}
            for fut in tqdm(as_completed(futures), total=len(futures), desc="Re-Fetching sectors/industries"):
                try:
                    orig_id, resolved_sym, sector, industry, err = fut.result()
                except Exception as e:
                    error_id_list.append((orig_id, f"task_error: {e}"))
                    continue

                if (resolved_sym is None) or (sector is None and industry is None):
                    error_id_list.append((orig_id, err or "no_sector_industry"))
                    continue

                # write back using the ORIGINAL id’s row indices
                for idx in sym_to_indices[orig_id]:
                    if sector is not None:
                        sector_arr[idx] = sector
                    if industry is not None:
                        industry_arr[idx] = industry

            # brief pause between batches to let any upstream throttles recover
            time.sleep(batch_pause)
                
    return sector_arr, industry_arr, error_id_list

# combine both resolve and fetch functions
def resolve_fetch_combine(original_id):
    try:
        resolved_id = yahoo_resolve_symbol(original_id)
        if not resolved_id:
            return original_id, None, None, None, "resolve_failed"

        resolved_id, sector, industry, err = fetch_sector_industry(resolved_id)
        return original_id, resolved_id, sector, industry, err

    except Exception as e:
        return original_id, None, None, None, str(e)

# resolve non-ticker stock IDs through yahoo UI
def yahoo_resolve_symbol(query, lang="en-US", region="AU"):
    """
    Resolve a query (SEDOL/ISIN/name/ticker) to a Yahoo symbol using Yahoo's search API.
    Returns the best match symbol (str) or None.
    """
    def _call():
        YAHOO_SEARCH_RL.acquire() 
        url = "https://query2.finance.yahoo.com/v1/finance/search"
        params = {"q": query, "lang": lang, "region": region}
        r = requests.get(url, params=params, headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
        r.raise_for_status()
        return r.json()

    data = _with_retries(_call)
    quotes = data.get("quotes") or []
    if not quotes:
        return None

    # pick the first equity-like hit; adjust ranking logic as you like
    for q in quotes:
        sym = q.get("symbol")
        if sym:
            return sym

def _yahoo_json(url, params, attempts=4, base=0.5, cap=3.0):
    """Tiny helper for retries with backoff + jitter."""
    last = None
    for i in range(attempts):
        try:
            r = requests.get(url, params=params, headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
            r.raise_for_status()
            return r.json()
        except Exception as e:
            last = e
            time.sleep(min(cap, base * (2**i)) + random.random()*0.25)
    raise last

def yahoo_resolve_symbol_2(query, lang="en-US", region="AU"):
    """
    Resolve a query (SEDOL/ISIN/name/ticker) to a Yahoo symbol.
    Tries the finance/search endpoint with richer params in several regions,
    then falls back to the autocomplete endpoint.
    """
    q = (query or "").strip().upper()
    if not q:
        return None

    # Prefer broader regions Yahoo often uses in the UI
    regions_to_try = [region] + [r for r in ("US", "GB", "AU", "HK") if r != region]

    # 1) Primary: v1/finance/search with richer params (closer to UI behavior)
    search_url = "https://query2.finance.yahoo.com/v1/finance/search"
    base_params = {
        "q": q,
        "lang": lang,
        "region": regions_to_try[0],
        "quotesCount": 10,                # increase from default
        "newsCount": 0,
        "enableFuzzyQuery": True,
        "quotesQueryId": "tss_match_phrase_query",   # UI-style hint
        "multiQuoteQueryId": "multi_quote_single_token_query",
        "enableEnhancedTrivialQuery": True,
    }

    # Try multiple regions
    for reg in regions_to_try:
        params = dict(base_params, region=reg)
        try:
            data = _yahoo_json(search_url, params)
        except Exception:
            continue
        quotes = (data or {}).get("quotes") or []
        # If you supplied a SEDOL, prefer exact id match in the result payload fields
        if quotes:
            # First pass: return first symbol (as your original code)
            for qte in quotes:
                sym = qte.get("symbol")
                if sym:
                    return sym

    # 2) Fallback: autocomplete endpoint (often better for identifiers like ISIN/SEDOL)
    #    The AU site’s search box frequently hits this too.
    ac_url = "https://autoc.finance.yahoo.com/autoc"
    for reg in regions_to_try:
        ac_params = {"query": q, "region": reg, "lang": lang}
        try:
            data = _yahoo_json(ac_url, ac_params)
        except Exception:
            continue
        # The shape is usually {"ResultSet": {"Query": "...", "Result": [...]}}
        results = ((data or {}).get("ResultSet") or {}).get("Result") or []
        # Find any obvious equity ticker; prefer exact query hits for ids
        for r in results:
            sym = r.get("symbol")
            if sym:
                return sym

    # Nothing found
    return None



# fetch sector and industry function
def fetch_sector_industry(symbol):
    """
    Fetch sector/industry for a *ticker* symbol via yfinance.
    Adds rate limiting and retries around the .info call.
    """

    def _call():
        YF_INFO_RL.acquire()
        t = yf.Ticker(symbol)
        return t.info

    try:
        info = _with_retries(_call)
        # Explicitly raise if sector/industry missing
        sector = info.get('sector')
        industry = info.get('industry')
        if sector is None and industry is None:
            raise ValueError("No sector/industry returned")
        return symbol, sector, industry, None
    except Exception as e:
        return symbol, None, None, str(e)


if __name__ == "__main__":

    update_final_file()

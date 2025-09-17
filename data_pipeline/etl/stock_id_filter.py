import re
import os
import pandas as pd
import numpy as np

def update_pre_file():
    pre_path, final_path, pre_file_names = file_directory("pre_data", "pre_final_data")

    for file_name in pre_file_names:
        super_name = file_name.split("_",1)[0]
        final_name = f"{super_name}_final.csv"

        # read pre cleaned file
        input_df = pd.read_csv(f"{pre_path}/{file_name}")

        # update input df with ID type col
        updated_df = id_type_create(input_df)

        # rewrite to csv file, overwrite if already exists
        updated_df.to_csv(f"{final_path}/{final_name}", mode='w', index=False)  # write to csv

        possible_ids = ["SEDOL", "ISIN", "BB Ticker", "Other"]

        for id_type in possible_ids:
            row_count = len(updated_df[updated_df["ID Type"] == id_type])
            print(f"Number of {id_type} rows for fund - {super_name}: {row_count}")


def file_directory(pre_data_folder, final_data_folder):
    # Get the current working directory
    current_dir = os.getcwd()

    # pre data folder path
    pre_data_path = os.path.join(current_dir, pre_data_folder)

    # List all entries in the folder
    all_files = os.listdir(pre_data_path)

    # Keep only files (exclude subfolders)
    file_names = [f for f in all_files if os.path.isfile(os.path.join(pre_data_folder, f))]

    # Construct the full path for the final data folder
    final_data_path = os.path.join(current_dir, final_data_folder)

    # Create the new folder (and any necessary parent directories)
    os.makedirs(final_data_path, exist_ok=True)

    return pre_data_path, final_data_path, file_names


# Function to create ID Type column in df
def id_type_create(df):
    id_type_series = df["Stock ID"].apply(detect_id_type)

    # insert new col ID Type right after Stock ID
    cols = list(df.columns)
    insert_at = cols.index("Stock ID") + 1

    if 'ID Type' in df.columns:
        df['ID Type'] = id_type_series   # overwrite
    else:
        df.insert(insert_at, 'ID Type', id_type_series)

    return df

# Function to detect stock ID type
def detect_id_type(orig_id):
    """
    Detect type of stock identifier.
    Returns one of: "ISIN", "SEDOL", "BB Ticker", "Other".
    """
    # check for nan values
    if pd.isna(orig_id) or str(orig_id).strip() == "":
        return "N/A"

    if not isinstance(orig_id, str):
        orig_id = str(orig_id)
    s = orig_id.strip().upper()

    # --- ISIN ---
    if check_isin(s):
        return "ISIN"

    # --- SEDOL ---
    # Either: 7 digits OR 7 chars (digits + consonants)
    ### However, if its full of number, and less than 7 digits, may want to add 0 to the front and check
    if check_sedol(s):
        return "SEDOL"

    # --- Bloomberg Ticker ---
    # e.g., "AAPL US", "BHP AU", "BRK/B US", "RDS-A LN"
    # Case 1: Full Bloomberg ticker (symbol + exchange code)
    bb_cond1 = re.fullmatch(r'^[A-Z0-9./-]+ [A-Z]{2,3}$', s)
    # Case 2: Symbol only (no exchange code)
    bb_cond2 = re.fullmatch(r'^[A-Z0-9./-]+$', s)
    if bb_cond1 or bb_cond2:
        return "BB Ticker"


    # --- Other ---
    return "Other"


def check_isin(s: str) -> bool:

    if not re.fullmatch(r'^[A-Z]{2}[A-Z0-9]{9}[0-9]$', s):
        return False

    # Convert ISIN string to a list of digits (ints) per ISO 6166
    digits = []
    for ch in s:
        val = int(ch, 36)  # works for 0–9 and A–Z
        # Expand numbers >= 10 into their decimal digits
        for d in str(val):
            digits.append(int(d))
            
    total = 0
    for i, d in enumerate(reversed(digits)):
        if i % 2 == 1:  # double every 2nd digit
            d *= 2
            if d > 9:
                d -= 9
        total += d

    return total % 10 == 0


def check_sedol(s: str) -> bool:
    """
    Validate a SEDOL (Stock Exchange Daily Official List) identifier.
    Must be 7 chars, digits + consonants (no vowels), last char is check digit.
    """
    # If it's numeric and too short, pad to 7
    if s.isdigit() and len(s) < 7:
        s = s.zfill(7)

    if len(s) != 7:
        return False

    weight = [1, 3, 1, 7, 3, 9, 1]

    # First 6 characters must match allowed pattern
    if not re.fullmatch(r'^[0-9BCDFGHJKLMNPQRSTVWXYZ]{6}[0-9]$', s):
        return False

    body, check_digit = s[:6], s[6]

    # Compute check digit
    total = 0
    for i, ch in enumerate(body):
        val = int(ch, 36)  # base-36 conversion (0-9=0-9, A=10, B=11, etc.)
        total += weight[i] * val
    calc_check = (10 - (total % 10)) % 10

    return check_digit == str(calc_check)


if __name__ == "__main__":

    update_pre_file()

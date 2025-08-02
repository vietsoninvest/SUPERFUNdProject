import pandas as pd
import os
import re # Import re for regular expressions

def tokenize_string(s):
    """
    Cleans a string by removing non-alphanumeric characters (except spaces),
    converts it to lowercase, and splits it into a set of words (tokens).
    Handles NaN values by returning an empty set.
    """
    if pd.isna(s) or not isinstance(s, str):
        return set()
    cleaned_s = re.sub(r'[^a-z0-9\s]', '', str(s).lower()).strip()
    return set(cleaned_s.split())

def infer_country_from_text(text, country_keywords, australian_states):
    """
    Infers a country name from a given text string based on a list of keywords.
    Prioritizes a match with an Australian state/territory abbreviation.
    If no Australian state is found, it finds the longest country keyword match.
    
    Args:
        text (str): The string to search within (e.g., combined Address and Investment Name).
        country_keywords (dict): A dictionary where keys are lowercase country names
                                 and values are the original proper-cased country names.
        australian_states (list): A list of Australian state/territory abbreviations.
                                 
    Returns:
        str: The proper-cased country name if a match is found, otherwise None.
    """
    if pd.isna(text) or not isinstance(text, str):
        return None

    # First, check for Australian state abbreviations as a high-confidence signal
    for state in australian_states:
        if re.search(r'\b' + re.escape(state) + r'\b', text, re.IGNORECASE):
            return "Australia"
            
    text_tokens = tokenize_string(text)
    
    # Early exit if the text is empty after tokenization
    if not text_tokens:
        return None

    best_match = None
    max_common_tokens = 0

    # New logic: check for multi-word matches by comparing token sets
    for keyword_lower, original_name in country_keywords.items():
        keyword_tokens = set(keyword_lower.split()) # e.g., 'united states' -> {'united', 'states'}
        
        # Check if all tokens of the country keyword are a subset of the text's tokens
        if keyword_tokens.issubset(text_tokens):
            num_tokens_matched = len(keyword_tokens)
            
            # Prioritize the match with the most tokens (the longest country name)
            if num_tokens_matched > max_common_tokens:
                max_common_tokens = num_tokens_matched
                best_match = original_name
    
    return best_match

def process_assets_data(input_filename="D:\\LinhDao\\Programming\\SUPERFUNdProject\\Linh-ausiesuper-CleanedData_Assets.csv",
                        currency_codes_path="D:\\LinhDao\\Programming\\SUPERFUNdProject\\CleanedCurrencyCodes.csv",
                        output_filename="D:\\LinhDao\\Programming\\SUPERFUNdProject\\Linh-Ausiesuper-CleanedData.csv"):
    """
    Reads the Linh-ausiesuper-CleanedData_Assets.csv file,
    updates the 'Listed Country' column for Australian states/territories to 'Australia',
    infers countries from 'Address' and 'Name/Kind of Investment Item' if 'Listed Country' is empty
    (using the data's existing countries as a reference),
    updates the 'Currency' column based on a keyword-based lookup,
    and saves the modified data to a new CSV file.

    Args:
        input_filename (str): The path to the input Assets CSV file.
        currency_codes_path (str): The path to the CleanedCurrencyCodes.csv file.
        output_filename (str): The path for the output CSV file with updated data.
    """
    print(f"[INFO] Attempting to load assets data from: {input_filename}")
    try:
        df_assets = pd.read_csv(input_filename, encoding='utf-8')
        print(f"[SUCCESS] Successfully loaded '{input_filename}'. Shape: {df_assets.shape}")
    except FileNotFoundError:
        print(f"[ERROR] The file '{input_filename}' was not found. Please double-check the path.")
        return
    except pd.errors.EmptyDataError:
        print(f"[ERROR] The file '{input_filename}' is empty. No data to process.")
        return
    except Exception as e:
        print(f"[ERROR] An unexpected error occurred while loading assets file: {e}")
        return

    # A list of Australian state/territory abbreviations for direct lookup
    australian_states = ["NSW", "NT", "QLD", "VIC", "SA", "TAS", "WA", "ACT"]

    # --- Step 1: Update Australian states/territories to 'Australia' (Direct Replacement) ---
    print("[INFO] Starting to update 'Listed Country' values for Australian states/territories.")
    
    if 'Listed Country' in df_assets.columns:
        rows_to_update = df_assets['Listed Country'].isin(australian_states)
        updated_count = rows_to_update.sum()
        
        if updated_count > 0:
            df_assets.loc[rows_to_update, 'Listed Country'] = 'Australia'
            print(f"[SUCCESS] Updated 'Listed Country' for {updated_count} rows from Australian states/territories to 'Australia'.")
        else:
            print("[INFO] No rows found with Australian states/territories in 'Listed Country'.")
    else:
        print("[WARNING] 'Listed Country' column not found in the assets data. Skipping this step.")

    # --- Step 2: Prepare for country inference using existing data as reference ---
    country_keywords_dict = {}
    if 'Listed Country' in df_assets.columns:
        print("[INFO] Building country reference list from unique values in 'Listed Country' column.")
        unique_countries = df_assets['Listed Country'].dropna().unique()
        for country in unique_countries:
            country_keywords_dict[country.lower()] = country
        print(f"[INFO] Found {len(country_keywords_dict)} unique countries to use for inference.")
    else:
        print("[WARNING] 'Listed Country' column not found. Country inference will be skipped.")
    
    # --- Step 3: Infer countries for empty 'Listed Country' rows ---
    if country_keywords_dict:
        print("[INFO] Starting country inference for rows with empty 'Listed Country'...")
        inferred_count = 0
        
        for index, row in df_assets.iterrows():
            listed_country = row.get('Listed Country')
            address = row.get('Address')
            investment_name = row.get('Name/Kind of Investment Item')
            
            if pd.isna(listed_country) and pd.notna(address):
                combined_text = f"{address} {investment_name}"
                inferred_country = infer_country_from_text(combined_text, country_keywords_dict, australian_states)
                
                if inferred_country:
                    df_assets.loc[index, 'Listed Country'] = inferred_country
                    inferred_count += 1
        
        print(f"[SUCCESS] Completed country inference. Updated 'Listed Country' for {inferred_count} rows.")
    
    # --- Step 4: Perform the currency lookup (still requires the currency codes file) ---
    print(f"[INFO] Attempting to load currency codes data from: {currency_codes_path}")
    df_currency_codes = None
    try:
        df_currency_codes = pd.read_csv(currency_codes_path, encoding='utf-8')
        print(f"[SUCCESS] Successfully loaded '{currency_codes_path}'. Shape: {df_currency_codes.shape}")
        
        if 'Country' not in df_currency_codes.columns or 'Code' not in df_currency_codes.columns:
            print("[ERROR] 'CleanedCurrencyCodes.csv' must contain 'Country' and 'Code' columns. Currency lookup will be skipped.")
            df_currency_codes = None
        else:
            df_currency_codes['Extracted_Currency_Name'] = df_currency_codes['Country'].apply(
                lambda x: re.search(r'\((.*?)\)', str(x)).group(1).strip() if re.search(r'\((.*?)\)', str(x)) else str(x).strip()
            )
            df_currency_codes['Country_Tokens'] = df_currency_codes['Country'].apply(tokenize_string)
            df_currency_codes['Code_Tokens'] = df_currency_codes['Code'].apply(tokenize_string)
            df_currency_codes['Extracted_Currency_Name_Tokens'] = df_currency_codes['Extracted_Currency_Name'].apply(tokenize_string)

    except FileNotFoundError:
        print(f"[ERROR] The file '{currency_codes_path}' was not found. Currency lookup will be skipped.")
        df_currency_codes = None
    except pd.errors.EmptyDataError:
        print(f"[ERROR] The file '{currency_codes_path}' is empty. Currency lookup will be skipped.")
        df_currency_codes = None
    except Exception as e:
        print(f"[ERROR] An unexpected error occurred while loading currency codes file: {e}. Currency lookup will be skipped.")
        df_currency_codes = None
    
    if df_currency_codes is not None:
        print("[INFO] Starting currency lookup and update...")
        updated_currency_count = 0

        for index, row in df_assets.iterrows():
            current_currency_value = row['Currency']
            if pd.notna(current_currency_value) and str(current_currency_value).strip() != '':
                found_match = False
                current_currency_tokens = tokenize_string(current_currency_value)
                
                if not current_currency_tokens:
                    continue

                best_match_code = None
                max_common_tokens = 0

                for _, currency_lookup_row in df_currency_codes.iterrows():
                    lookup_country_tokens = currency_lookup_row['Country_Tokens']
                    lookup_currency_name_tokens = currency_lookup_row['Extracted_Currency_Name_Tokens']
                    lookup_code_tokens = currency_lookup_row['Code_Tokens']
                    actual_lookup_code = currency_lookup_row['Code']

                    all_lookup_tokens = lookup_country_tokens.union(lookup_currency_name_tokens).union(lookup_code_tokens)
                    common_tokens = len(current_currency_tokens.intersection(all_lookup_tokens))

                    if common_tokens > max_common_tokens:
                        max_common_tokens = common_tokens
                        best_match_code = actual_lookup_code
                        found_match = True

                if found_match and best_match_code is not None:
                    df_assets.loc[index, 'Currency'] = best_match_code
                    updated_currency_count += 1
        
        print(f"[INFO] Completed currency lookup. Updated {updated_currency_count} currency values.")
    else:
        print("[WARNING] Currency code lookup was skipped due to file loading issues.")

    # Save the modified DataFrame
    try:
        df_assets.to_csv(output_filename, index=False, encoding='utf-8-sig')
        print(f"[SUCCESS] Modified assets data saved to '{output_filename}'.")
    except Exception as e:
        print(f"[ERROR] Error saving the output CSV file: {e}")

if __name__ == "__main__":
    assets_csv_file_path = "D:\\LinhDao\\Programming\\SUPERFUNdProject\\Linh-ausiesuper-CleanedData_Assets.csv"
    currency_codes_csv_path = "D:\\LinhDao\\Programming\\SUPERFUNdProject\\CleanedCurrencyCodes.csv"
    output_assets_csv_path = "D:\\LinhDao\\Programming\\SUPERFUNdProject\\Linh-Ausiesuper-CleanedData.csv"
    
    process_assets_data(
        input_filename=assets_csv_file_path,
        currency_codes_path=currency_codes_csv_path,
        output_filename=output_assets_csv_path
    )
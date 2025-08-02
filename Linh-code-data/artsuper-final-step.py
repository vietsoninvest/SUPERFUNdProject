import pandas as pd
import re
import os # Import os for path manipulation
import datetime # Import datetime for date handling

def finalize_artsuper_data_csv(
    source_file_path,
    output_file_name="Linh-artsuper-CleanedData.csv",
    country_codes_file="D:\\LinhDao\\Programming\\SUPERFUNdProject\\InternationalCountryCodes.csv" # Added for country lookup
):
    """
    Performs final data cleaning and column reordering for a file (reads Excel, outputs CSV):
    - Ensures all columns are present and reorders them to a specific desired sequence.
    - Converts "Effective Date" to date-only format.
    - Maps "Int/Ext" values to 0 or 1.
    - Fills empty cells in 'Name/Kind of Investment Item' with 'Sub Total'.
    - Extracts last 2 characters from 'Stock ID' for 'Listed Country' lookup.
    - Looks up country names using an external CSV; retains code if not found.
    - Removes all other 'n/a' (case-insensitive) values from the entire DataFrame (leaving cells empty).

    Args:
        source_file_path (str): Path to the input Excel file (e.g., ARTsuper_final.xlsx).
        output_file_name (str): The desired name for the final output CSV file.
        country_codes_file (str): Path to the CSV file containing country codes and names.
    """
    print(f"\n[INFO] Starting finalization steps on '{source_file_path}'...")
    
    try:
        # Load the Excel file into a pandas DataFrame.
        # Assuming the Excel file has headers in the first row of its active sheet.
        df_source = pd.read_excel(source_file_path)
        print(f"[DEBUG] Successfully loaded '{source_file_path}'. Shape: {df_source.shape}")
    except FileNotFoundError:
        print(f"[ERROR] Source file '{source_file_path}' not found. Please check the path.")
        return
    except Exception as e:
        print(f"[ERROR] Error loading source Excel file: {e}")
        return

    # --- Load International Country Codes CSV for lookup ---
    country_code_map = {}
    try:
        df_country_codes = pd.read_csv(country_codes_file, encoding='utf-8')
        if 'Code' in df_country_codes.columns and 'Country' in df_country_codes.columns:
            # Create a dictionary mapping Code to Country
            country_code_map = pd.Series(df_country_codes['Country'].values, index=df_country_codes['Code']).to_dict()
            print(f"[DEBUG] Successfully loaded country codes from '{country_codes_file}'.")
        else:
            print(f"[WARNING] Country codes file '{country_codes_file}' does not contain 'Code' and 'Country' columns. Listed Country will not be fully mapped.")
    except FileNotFoundError:
        print(f"[ERROR] Country codes file '{country_codes_file}' not found. Listed Country will not be mapped via lookup.")
    except Exception as e:
        print(f"[ERROR] Error loading country codes CSV: {e}. Listed Country will not be mapped via lookup.")

    # --- Define the FINAL desired column order ---
    FINAL_HEADERS_ORDER = [
        "Effective Date",
        "Fund Name",
        "Option Name",
        "Asset Class Name",
        "Int/Ext",
        "Name/Kind of Investment Item",
        "Currency",
        "Stock ID", 
        "Listed Country", 
        "Units Held",
        "% Ownership",
        "Address",
        "Value (AUD)",
        "Weighting"
    ]

    # --- Pre-process source DataFrame column names for robust matching ---
    # Normalize source column names to match the desired case/format
    # This helps if the input Excel has "Stock id" or "stock ID" etc.
    df_source.columns = [col.replace(' ', '_').replace('-', '_').replace('(', '').replace(')', '').replace('%', '').strip().lower() for col in df_source.columns]
    
    # Create a mapping from normalized source column names to FINAL_HEADERS_ORDER
    # This is crucial for correctly identifying columns like 'Stock ID' if they have variations
    # in the source file.
    source_col_mapping = {}
    for final_header in FINAL_HEADERS_ORDER:
        normalized_final_header = final_header.replace(' ', '_').replace('-', '_').replace('(', '').replace(')', '').replace('%', '').strip().lower()
        if normalized_final_header == 'stock_id': # Special handling for Stock ID
            # Check for common variations of 'stock id' in the source columns
            if 'stock_id' in df_source.columns:
                source_col_mapping['stock_id'] = final_header
            elif 'stock_id' in df_source.columns: # This line seems redundant, keeping for now
                source_col_mapping['stock_id'] = final_header
            elif 'stock_id' in df_source.columns: # This line seems redundant, keeping for now
                source_col_mapping['stock_id'] = final_header
        elif normalized_final_header in df_source.columns:
            source_col_mapping[normalized_final_header] = final_header
        
    # Rename columns in df_source based on the mapping
    df_source.rename(columns={k: v for k, v in source_col_mapping.items() if k in df_source.columns}, inplace=True)
    print(f"[DEBUG] Source DataFrame columns after normalization and initial rename: {df_source.columns.tolist()}")


    # --- 1. Ensure all final desired columns are present in the DataFrame ---
    # Add any missing columns from FINAL_HEADERS_ORDER to df_source, filling with None (NaN in pandas)
    missing_cols = [col for col in FINAL_HEADERS_ORDER if col not in df_source.columns]
    for col in missing_cols:
        df_source[col] = None
        print(f"[DEBUG] Added missing column: '{col}' to DataFrame.")

    # --- 2. Reorder columns to the desired sequence ---
    # This will also drop any columns from df_source that are not in FINAL_HEADERS_ORDER
    df_final = df_source[FINAL_HEADERS_ORDER].copy()
    print(f"[INFO] Data reordered into desired column sequence. Final columns: {df_final.columns.tolist()}")

    # --- Apply requested transformations ---

    # 1. "Effective Date" column to be Date only, no time
    date_col = "Effective Date"
    if date_col in df_final.columns:
        # Convert to datetime, coerce errors to NaT (Not a Time), then extract date part
        df_final[date_col] = pd.to_datetime(df_final[date_col], errors='coerce').dt.date
        print(f"[INFO] '{date_col}' column converted to date-only format.")
    else:
        print(f"[WARNING] '{date_col}' column not found for date formatting.")

    # 2. "Int/Ext" column mapping: "Externally Managed" -> 1, "Internally Managed" -> 0
    int_ext_col = "Int/Ext"
    if int_ext_col in df_final.columns:
        # Create a mapping dictionary for direct replacement
        int_ext_mapping = {
            "Externally Managed": 1,
            "Internally Managed": 0
        }
        # Apply the mapping. Values not in the map (including NaN) will remain as is.
        df_final[int_ext_col] = df_final[int_ext_col].replace(int_ext_mapping)
        print(f"[INFO] '{int_ext_col}' column values mapped to 0 or 1.")
    else:
        print(f"[WARNING] '{int_ext_col}' column not found for mapping.")

    # 3. Fill empty cells in "Name/Kind of Investment Item" with "Sub Total"
    name_kind_col = "Name/Kind of Investment Item"
    if name_kind_col in df_final.columns:
        # Convert column to string type to handle empty strings consistently
        # Replace empty strings (after stripping) with None, then fill None/NaN with "Sub Total"
        initial_empty_count = df_final[name_kind_col].apply(lambda x: str(x).strip() == '' or pd.isna(x)).sum()
        df_final[name_kind_col] = df_final[name_kind_col].apply(lambda x: None if str(x).strip() == '' else x)
        df_final[name_kind_col] = df_final[name_kind_col].fillna("Sub Total")
        print(f"[INFO] Filled {initial_empty_count} empty cells in '{name_kind_col}' with 'Sub Total'.")
    else:
        print(f"[WARNING] Column '{name_kind_col}' not found in final DataFrame. Skipping filling empty cells.")

    # 4. "Listed Country" extraction and lookup (last 2 characters of Stock ID)
    stock_id_col = "Stock ID" 
    listed_country_col = "Listed Country"

    # Helper function to apply the country lookup logic
    def get_listed_country(stock_id_value, country_map):
        if pd.isna(stock_id_value) or str(stock_id_value).strip() == '':
            return None # If Stock ID is empty or NaN, Listed Country should be empty
        
        stock_id_str = str(stock_id_value).strip()
        if len(stock_id_str) >= 2:
            country_code = stock_id_str[-2:].upper() # Get LAST two chars and convert to uppercase
            mapped_country = country_map.get(country_code)
            if mapped_country:
                return mapped_country
            else:
                return country_code # If no country mapping found, retain the 2-character code
        return None # If Stock ID is not long enough, return None

    if stock_id_col in df_final.columns and listed_country_col in df_final.columns:
        # Apply the helper function to the 'Stock ID' column
        df_final[listed_country_col] = df_final[stock_id_col].apply(lambda x: get_listed_country(x, country_code_map))
        print(f"[INFO] '{listed_country_col}' column populated based on last 2 characters of '{stock_id_col}' and country lookup.")
    else:
        print(f"[WARNING] '{stock_id_col}' or '{listed_country_col}' column not found for country lookup.")

    # --- 5. Remove all other "n/a" values from the entire data table (leaving cells empty) ---
    # This should be the last cleaning step for 'n/a'
    initial_na_count_global = df_final.astype(str).apply(lambda x: x.str.lower() == 'n/a').sum().sum()
    df_final.replace(to_replace=r'(?i)n/a', value=None, regex=True, inplace=True)
    print(f"[INFO] Removed {initial_na_count_global} instances of 'n/a' from the entire table (set to empty).")

    # --- Final Save to CSV ---
    try:
        df_final.to_csv(output_file_name, index=False, encoding='utf-8')
        print(f"[SUCCESS] Final cleaned file saved as '{output_file_name}'.")
    except Exception as e:
        print(f"[ERROR] Error saving the final CSV file: {e}")

# =========================================================================
# Main execution block
# =========================================================================
if __name__ == "__main__":
    # The input for this script is the output from the previous step, now expected to be an Excel file
    input_file_from_previous_step = r"D:\LinhDao\Programming\SUPERFUNdProject\ARTsuper_final.xlsx" 
    # The final desired output file name (CSV)
    final_output_csv_name = r"D:\LinhDao\Programming\SUPERFUNdProject\Linh-artsuper-CleanedData.csv" 
    # Path to the country codes CSV file
    country_codes_path = r"D:\LinhDao\Programming\SUPERFUNdProject\InternationalCountryCodes.csv" 

    finalize_artsuper_data_csv(
        source_file_path=input_file_from_previous_step,
        output_file_name=final_output_csv_name,
        country_codes_file=country_codes_path 
    )

    print("\n[IMPORTANT NOTE ON WORKFLOW]:")
    print("This script is designed to take an Excel file as input (e.g., 'ARTsuper_final.xlsx')")
    print("and perform the final cleaning and reordering, saving it as a new CSV file.")
    print("Ensure 'ARTsuper_final.xlsx' and 'InternationalCountryCodes.csv' exist and contain the data you expect.")

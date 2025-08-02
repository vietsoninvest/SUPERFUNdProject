import pandas as pd
import re
import os # Import os for path manipulation

def parse_value_string_to_number(s):
    """
    Converts strings like '$2m', '1.5Bn', '500k' to numerical values.
    Handles 'm' for million, 'bn'/'b' for billion, and 'k' for thousand.
    """
    if pd.isna(s) or not isinstance(s, str):
        return None
    
    # Remove currency symbols, commas, and convert to lowercase for consistent parsing
    s = s.strip().replace('$', '').replace(',', '').lower()
    
    multiplier = 1.0
    if 'm' in s:
        multiplier = 1_000_000
        s = s.replace('m', '')
    elif 'bn' in s or 'b' in s:
        multiplier = 1_000_000_000
        s = s.replace('bn', '').replace('b', '')
    elif 'k' in s:
        multiplier = 1_000
        s = s.replace('k', '')

    try:
        return float(s) * multiplier
    except ValueError:
        return None

def extract_and_average_value_range(value_range_str):
    """
    Extracts numerical value(s) from a value range string and returns an average.
    Examples:
    "$2m to $10m" -> (2m + 10m) / 2 = 6,000,000
    ">1.5Bn" -> 1,500,000,000
    "$500k" -> 500,000
    """
    if pd.isna(value_range_str) or not isinstance(value_range_str, str):
        return None

    value_range_str = value_range_str.strip()

    # Handle ranges like "$2m to $10m"
    if 'to' in value_range_str.lower():
        parts = value_range_str.lower().split('to')
        if len(parts) == 2:
            val1 = parse_value_string_to_number(parts[0])
            val2 = parse_value_string_to_number(parts[1])
            if val1 is not None and val2 is not None:
                return (val1 + val2) / 2
    # Handle "greater than" values like ">1.5Bn"
    elif '>' in value_range_str:
        num_str = value_range_str.replace('>', '').strip()
        return parse_value_string_to_number(num_str)
    # Handle single values like "$5m" or "100000"
    else:
        return parse_value_string_to_number(value_range_str)
    
    return None # Return None if no valid pattern matched or parsing failed

def process_excel_to_csv_file(input_filename="D:\\LinhDao\\Programming\\SUPERFUNdProject\\Linh-ausiesuper.xlsx", output_base_name="D:\\LinhDao\\Programming\\SUPERFUNdProject\\Linh-ausiesuper-CleanedData"):
    """
    Reads an Excel file, updates the 'Asset Class' column based on the
    'Filter' column, separates 'Derivatives' rows into a new CSV file,
    processes 'Assets' rows and saves them to another new CSV file.
    Also adds and updates an 'Int/Ext' column based on 'Filter' and 'Sub-Filter' columns.
    Additionally, it extracts values from 'Value Range' to fill missing 'Value(AUD)' for 'Assets' data,
    and then reorders and renames columns for the 'Assets' data.

    Args:
        input_filename (str): The path to the input Excel file.
        output_base_name (str): The base path for the output CSV files (e.g., "path/to/output_file" will create "path/to/output_file_Assets.csv" and "path/to/output_file_Derivatives.csv").
    """
    try:
        # Read the Excel file into a pandas DataFrame
        # Assuming the Excel file has headers in the first row.
        df = pd.read_excel(input_filename) # Changed to pd.read_excel
        print(f"Successfully loaded '{input_filename}'.")

        # --- Step 0.1: Add new columns with default values to the main DataFrame ---
        # Add 'Effective Date' as the first column
        df.insert(0, "Effective Date", "31 Dec 2024")
        # Add 'Fund Name' as the second column
        df.insert(1, "Fund Name", "AusieSuper")
        print("Initialized 'Effective Date' and 'Fund Name' columns.")

        # --- Step 0.2: Add 'Int/Ext' column with default value ---
        # Initialize the new 'Int/Ext' column with 'Externally Managed' as the default
        # This column will be positioned later during reordering
        df["Int/Ext"] = "Externally Managed"
        print("Initialized 'Int/Ext' column with 'Externally Managed' as default.")

        # --- Step 1: Apply the "Listed/Unlisted" logic and update 'Int/Ext' column ---
        for index, row in df.iterrows():
            filter_value = str(row["Filter"]).strip()
            sub_filter_value = str(row["Sub-Filter"]).strip()

            if filter_value == "Listed":
                df.loc[index, "Asset Class"] = "Listed " + str(row["Asset Class"])
            elif filter_value == "Unlisted":
                df.loc[index, "Asset Class"] = "Unlisted " + str(row["Asset Class"])

            # Update 'Int/Ext' based on 'Filter' and 'Sub-Filter' (case-insensitive)
            if "Internally" in filter_value or "Internally" in sub_filter_value:
                df.loc[index, "Int/Ext"] = "Internally Managed"
            elif "Externally" in filter_value or "Externally" in sub_filter_value:
                df.loc[index, "Int/Ext"] = "Externally Managed"

        print("Completed 'Listed/Unlisted' asset class updates and 'Int/Ext' column updates.")

        # --- NEW: Convert 'Int/Ext' text values to 0 or 1 ---
        # This should be done after the text values are finalized from Filter/Sub-Filter
        df["Int/Ext"] = df["Int/Ext"].map({"Externally Managed": 1, "Internally Managed": 0}).fillna(df["Int/Ext"])
        print("Converted 'Int/Ext' column values to 0 (Internally Managed) or 1 (Externally Managed).")

        # --- Step 2: Separate data into 'Derivatives' and 'Assets' DataFrames ---
        derivatives_df = df[df["Asset Class"] == "Derivatives"].copy()
        assets_df = df[df["Asset Class"] != "Derivatives"].copy()

        print(f"Separated {len(derivatives_df)} rows into 'Derivatives' data and {len(assets_df)} rows into 'Assets' data.")

        # --- Step 3: Process 'Assets' DataFrame: Extract from 'Value Range' for missing '$ Value' ---
        # Ensure '$ Value' is numeric, coercing errors to NaN (using original column name)
        assets_df["$ Value"] = pd.to_numeric(assets_df["$ Value"], errors='coerce')

        # Fill missing '$ Value' using 'Value Range'
        # Apply the extraction function only to rows where '$ Value' is NaN
        # and 'Value Range' is not NaN
        missing_value_mask = assets_df["$ Value"].isna()
        value_range_available_mask = assets_df["Value Range"].notna()

        assets_df.loc[missing_value_mask & value_range_available_mask, "$ Value"] = \
            assets_df.loc[missing_value_mask & value_range_available_mask, "Value Range"].apply(extract_and_average_value_range)

        print("Filled missing '$ Value' values from 'Value Range' in 'Assets' DataFrame.")

        # Fill any remaining NaN values in '$ Value' with 0
        assets_df["$ Value"] = assets_df["$ Value"].fillna(0)
        print("Filled any remaining NaN values in '$ Value' with 0.")

        # Columns to be deleted from the 'Assets' data
        columns_to_delete = [
            "Option Code", "Filter", "Sub-Filter", "Name Type", "Issuer Type",
            "Actual Currency Exposure (%)", "Actual Asset Allocation (%)",
            "Effect of Derivatives Exposure (%)", "Classification", "Sort Order",
            "Value Range", "Geo Latitude", "Geo Longitude"
        ]
        # Drop specified columns from assets_df
        assets_df = assets_df.drop(columns=columns_to_delete, errors='ignore')
        print("Deleted specified columns from 'Assets' DataFrame.")

        # Columns to be renamed in the 'Assets' data
        columns_to_rename_map = {
            "$ Value": "Value (AUD)",
            "Weighting (%)": "Weighting",
            "Asset Class": "Asset Class Name",
            "Name": "Name/Kind of Investment Item",
            "Security Identifier": "Stock ID", # Ensure this is 'Stock ID'
            "Location": "Listed Country"
        }
        # Rename columns in assets_df
        assets_df = assets_df.rename(columns=columns_to_rename_map)
        print("Renamed specified columns in 'Assets' DataFrame.")
        
        # Fill empty 'Name/Kind of Investment Item' values with "Total"
        # This will fill NaN/None and empty strings after stripping
        assets_df["Name/Kind of Investment Item"] = assets_df["Name/Kind of Investment Item"].fillna("").apply(lambda x: "Total" if x.strip() == "" else x)
        print("Filled empty 'Name/Kind of Investment Item' values with 'Total' in 'Assets' DataFrame.")

        # --- NEW: Change "Total" to "Sub Total" in "Name/Kind of Investment Item" ---
        name_kind_col = "Name/Kind of Investment Item"
        if name_kind_col in assets_df.columns:
            # Ensure the column is string type for comparison
            # Use .loc for setting values to avoid SettingWithCopyWarning
            # Check for exact match "Total" (case-insensitive)
            initial_total_count = assets_df[name_kind_col].astype(str).str.strip().str.lower().eq("total").sum()
            assets_df.loc[assets_df[name_kind_col].astype(str).str.strip().str.lower() == "total", name_kind_col] = "Sub Total"
            print(f"Changed {initial_total_count} instances of 'Total' to 'Sub Total' in '{name_kind_col}'.")
        else:
            print(f"[WARNING] Column '{name_kind_col}' not found for 'Total' to 'Sub Total' conversion.")


        # Define the desired order of columns for the 'Assets' data
        desired_assets_column_order = [
            "Effective Date",
            "Fund Name",
            "Option Name",
            "Asset Class Name",
            "Int/Ext",
            "Name/Kind of Investment Item",
            "Currency",
            "Stock ID", # Ensure this is 'Stock ID'
            "Listed Country",
            "Units Held",
            "% Ownership",
            "Address",
            "Value (AUD)",
            "Weighting"
        ]
        # Reorder columns in assets_df. This will also implicitly drop any columns
        # that are not in this list and were not explicitly dropped before.
        assets_df = assets_df[desired_assets_column_order]
        print("Reordered columns in 'Assets' DataFrame.")

        # --- Step 4: Save the modified DataFrames to separate CSV files ---
        assets_output_path = f"{output_base_name}_Assets.csv"
        derivatives_output_path = f"{output_base_name}_Derivatives.csv"

        assets_df.to_csv(assets_output_path, index=False, encoding='utf-8')
        print(f"Successfully wrote 'Assets' data to '{assets_output_path}'.")

        derivatives_df.to_csv(derivatives_output_path, index=False, encoding='utf-8')
        print(f"Successfully wrote 'Derivatives' data to '{derivatives_output_path}'.")

        print(f"Processing complete. Modified data saved to '{assets_output_path}' and '{derivatives_output_path}'.")

    except FileNotFoundError:
        print(f"Error: The file '{input_filename}' was not found. Please double-check the path.")
    except KeyError as e:
        print(f"Error: Missing expected column. Please ensure your Excel file has all necessary columns. Details: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

# Call the function to process your Excel file with the specified path
if __name__ == "__main__":
    # Input Excel file
    input_excel_file = "D:\\LinhDao\\Programming\\SUPERFUNdProject\\Linh-ausiesuper.xlsx" # Changed to .xlsx
    # Base name for output CSV files (e.g., will create Linh-ausiesuper-CleanedData_Assets.csv and Linh-ausiesuper-CleanedData_Derivatives.csv)
    output_csv_base_name = "D:\\LinhDao\\Programming\\SUPERFUNdProject\\Linh-ausiesuper-CleanedData" 

    process_excel_to_csv_file( # Changed function name
        input_filename=input_excel_file,
        output_base_name=output_csv_base_name
    )

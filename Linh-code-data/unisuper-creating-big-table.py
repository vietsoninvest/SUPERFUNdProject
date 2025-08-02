import pandas as pd
import re
import os
import datetime
import csv # Import the csv module for basic CSV operations

# =========================================================================
# Function 1: create_csv_file_with_headers
# =========================================================================
def create_csv_file_with_headers(file_name="CleanedData.csv"):
    """
    Creates a CSV file with the specified headers.

    Args:
        file_name (str): The path and name for the new CSV file.
    """
    # Define column headers for the CleanedData table
    column_headers = [
        "Effective Date", "Fund Name", "Option Name", "Asset Class Name",
        "Int/Ext", "Name/Kind of Investment Item", "Currency", "Stock Id",
        "Listed Country", "Units Held", "% Ownership", "Address",
        "Value (AUD)", "Weighting" # Changed from "Weighting (%)" to "Weighting"
    ]

    try:
        # Open the CSV file in write mode, 'newline=""' is crucial for CSV
        with open(file_name, 'w', newline='', encoding='utf-8') as csvfile:
            csv_writer = csv.writer(csvfile)
            # Write the header row
            csv_writer.writerow(column_headers)
        print(f"CSV file '{file_name}' with headers created successfully.")
    except Exception as e:
        print(f"Error creating the CSV file: {e}")

# =========================================================================
# Function 2: process_complicated_source_csv
# =========================================================================
def process_complicated_source_csv(
    source_file_path,
    destination_file="CleanedData.csv",
    country_codes_file="D:\\LinhDao\\Programming\\SUPERFUNdProject\\InternationalCountryCodes.csv"
):
    """
    Processes a complicated source CSV file with multiple tables,
    extracting data and populating a destination CSV file.
    
    Looks for Asset Class Name and Int/Ext in separate rows
    in Column A above the table header.
    Also maps Stock ID prefixes to Listed Country using an external CSV.
    """
    print(f"\n[DEBUG] Starting to process complicated source file: '{source_file_path}'")
    
    # --- Load Source CSV into pandas DataFrame ---
    try:
        df_source = pd.read_csv(source_file_path, header=None, skipinitialspace=True)
        print(f"[DEBUG] Successfully loaded source CSV. Total rows: {len(df_source)}")
        if len(df_source) < 2:
            print("[ERROR] Source CSV appears almost empty (less than 2 rows). Exiting.")
            return
    except FileNotFoundError:
        print(f"[ERROR] Source file '{source_file_path}' not found. Please check the path.")
        return
    except Exception as e:
        print(f"[ERROR] Error loading source CSV: {e}")
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
            print(f"[WARNING] Country codes file '{country_codes_file}' does not contain 'Code' and 'Country' columns. Skipping country lookup.")
    except FileNotFoundError:
        print(f"[ERROR] Country codes file '{country_codes_file}' not found. Listed Country will not be mapped.")
    except Exception as e:
        print(f"[ERROR] Error loading country codes CSV: {e}. Listed Country will not be mapped.")

    # --- Load Destination CSV (or create initial DataFrame if empty/new) ---
    destination_columns_ordered = [
        "Effective Date", "Fund Name", "Option Name", "Asset Class Name",
        "Int/Ext", "Name/Kind of Investment Item", "Currency", "Stock Id",
        "Listed Country", "Units Held", "% Ownership", "Address",
        "Value (AUD)", "Weighting" # Changed from "Weighting (%)" to "Weighting"
    ]
    
    try:
        df_dest = pd.read_csv(destination_file, encoding='utf-8')
        print(f"[DEBUG] Successfully loaded existing destination CSV: '{destination_file}'.")
        df_dest = df_dest[destination_columns_ordered]
    except FileNotFoundError:
        print(f"[WARNING] Destination file '{destination_file}' not found or empty. Initializing new DataFrame.")
        df_dest = pd.DataFrame(columns=destination_columns_ordered)
    except Exception as e:
        print(f"[ERROR] Error loading destination CSV '{destination_file}': {e}. Initializing new DataFrame.")
        df_dest = pd.DataFrame(columns=destination_columns_ordered)

    # --- Configuration for Column Mapping and Data Extraction ---
    HEADER_KEYWORDS = ["name", "value", "weighting"]
    TOTAL_KEYWORD = "total" 

    def normalize_text(text):
        if pd.isna(text):
            return ""
        return str(text).strip().lower()

    source_to_dest_header_map = {
        'security identifier': 'Stock Id', 
        'stock id': 'Stock Id', 
        'id': 'Stock Id', 
        
        'units held': 'Units Held', 
        'units': 'Units Held', 
        
        'value': 'Value (AUD)',
        'value (aud)': 'Value (AUD)',
        'investment value': 'Value (AUD)',
        'investment value (aud)': 'Value (AUD)',
        
        'weighting %': 'Weighting', # Mapped to new column name "Weighting"
        'weighting (%)': 'Weighting', # Mapped to new column name "Weighting"
        'weight': 'Weighting',
        'proportional weight': 'Weighting',
        
        'currency': 'Currency',
        'currency code': 'Currency',

        '% ownership': '% Ownership',
        '% of property held': '% Ownership', 
        
        'listed country': 'Listed Country',
        'country': 'Listed Country',
        
        'address': 'Address',
        'full address details': 'Address',
    }

    print("\n[DEBUG] Starting data extraction from source sheet...")
    rows_processed_count = 0
    tables_found_count = 0
    data_rows_transferred = 0

    current_asset_class_name = None
    current_int_ext = None
    current_effective_date = datetime.date(2024, 12, 31) 

    current_state = 0 # 0: Searching for header, 2: Processing table data
    current_table_column_indices = {} # Maps destination header to source column index
    
    found_int_ext_for_current_table = False
    found_asset_class_for_current_table = False
    
    ASSET_CLASS_KEYWORDS = [
        "listed property", "unlisted property",
        "listed infrastructure", "unlisted infrastructure",
        "listed equity", "unlisted equity",
        "cash", "fixed income", "listed alternatives",
        "unlisted alternatives", "property", "infrastructure", "equities", "alternatives" # General terms last
    ]
    INT_EXT_KEYWORDS = ["internally managed", "externally managed"]

    all_extracted_rows_for_append = []

    for row_idx in range(len(df_source)):
        rows_processed_count += 1
        row_cells_values = df_source.iloc[row_idx].tolist()
        current_row_values = [normalize_text(cell_val) for cell_val in row_cells_values]
        
        col_a_value = current_row_values[0] if current_row_values else ""

        if rows_processed_count <= 10 or rows_processed_count % 100 == 0: 
            print(f"\n[DEBUG] Row {row_idx + 1} (Current State: {current_state}):")
            print(f"  Raw values: {row_cells_values}") 
            print(f"  Normalized values: {current_row_values}")
            print(f"  Column A value: '{col_a_value}'")

        # --- PRE-SCANNING PHASE (Runs until a header is found) ---
        if current_state == 0:
            if not found_asset_class_for_current_table:
                for ac_keyword in ASSET_CLASS_KEYWORDS:
                    if ac_keyword in col_a_value:
                        current_asset_class_name = ac_keyword.replace(" ", " ").title()
                        found_asset_class_for_current_table = True
                        print(f"[DEBUG]   Found Asset Class Name: '{current_asset_class_name}' at row {row_idx + 1}, Col A.")
                        break 

            if not found_int_ext_for_current_table:
                for ie_keyword in INT_EXT_KEYWORDS:
                    if ie_keyword in col_a_value:
                        current_int_ext = ie_keyword.replace(" ", " ").title()
                        found_int_ext_for_current_table = True
                        print(f"[DEBUG]   Found Int/Ext: '{current_int_ext}' at row {row_idx + 1}, Col A.")
                        break 
            
            is_header_candidate = all(
                any(kw in cell_val for cell_val in current_row_values) for kw in HEADER_KEYWORDS
            )

            if is_header_candidate:
                print(f"[DEBUG]   Header candidate found at row {row_idx + 1}.")
                
                if current_asset_class_name is None:
                    current_asset_class_name = "Unknown Asset Class"
                    print(f"  [WARNING] Asset Class Name not found before header. Defaulting to '{current_asset_class_name}'.")
                
                if current_int_ext is None:
                    current_int_ext = "Externally Managed"
                    print(f"  [WARNING] Int/Ext not found before header. Defaulting to '{current_int_ext}'.")
                                
                print(f"  [INFO] Entering Table Processing State. Context: Asset Class: '{current_asset_class_name}', Int/Ext: '{current_int_ext}'")
                tables_found_count += 1
                current_state = 2 # Transition to processing table data

                current_table_column_indices = {}
                found_source_col_indices = set()

                name_col_found = False
                for col_idx, cell_value in enumerate(current_row_values):
                    if cell_value.startswith("name") and col_idx not in found_source_col_indices:
                        current_table_column_indices['Name/Kind of Investment Item'] = col_idx
                        found_source_col_indices.add(col_idx)
                        name_col_found = True
                        break 
                
                for dest_col in destination_columns_ordered:
                    if dest_col in ["Effective Date", "Fund Name", "Option Name", "Asset Class Name", "Int/Ext", "Name/Kind of Investment Item"]:
                        continue

                    found_mapping_for_dest_col = False
                    for src_header_norm, mapped_dest_header in source_to_dest_header_map.items():
                        if mapped_dest_header == dest_col:
                            for col_idx, cell_value in enumerate(current_row_values):
                                if src_header_norm == cell_value and col_idx not in found_source_col_indices:
                                    current_table_column_indices[dest_col] = col_idx
                                    found_source_col_indices.add(col_idx)
                                    found_mapping_for_dest_col = True
                                    break 
                            if found_mapping_for_dest_col:
                                break 
                        
                print(f"  [DEBUG] Current Table Mapped Headers: {current_table_column_indices}")
                if not current_table_column_indices:
                    print(f"    [WARNING] No mappable headers found in table at row {row_idx + 1}. Resetting state to search for next table.")
                    current_state = 0 
                    current_asset_class_name = None 
                    current_int_ext = None
                    found_int_ext_for_current_table = False
                    found_asset_class_for_current_table = False
                    continue

            continue

        # --- TABLE PROCESSING PHASE ---
        elif current_state == 2:
            is_total_row = any(cell_val == TOTAL_KEYWORD for cell_val in current_row_values)

            name_col_index = current_table_column_indices.get('Name/Kind of Investment Item')
            name_cell_value = row_cells_values[name_col_index] if name_col_index is not None and name_col_index < len(row_cells_values) else None
            
            if is_total_row or (name_cell_value is not None and str(name_cell_value).strip() != ""):
                data_to_append = [None] * len(destination_columns_ordered)

                data_to_append[destination_columns_ordered.index("Effective Date")] = current_effective_date
                data_to_append[destination_columns_ordered.index("Fund Name")] = "UniSuper"
                data_to_append[destination_columns_ordered.index("Option Name")] = "Balanced"
                data_to_append[destination_columns_ordered.index("Asset Class Name")] = current_asset_class_name
                
                # Apply Int/Ext mapping
                if current_int_ext == "Externally Managed":
                    data_to_append[destination_columns_ordered.index("Int/Ext")] = 1
                elif current_int_ext == "Internally Managed":
                    data_to_append[destination_columns_ordered.index("Int/Ext")] = 0
                else:
                    data_to_append[destination_columns_ordered.index("Int/Ext")] = None
                
                row_has_meaningful_data_beyond_context = False 

                # Store Stock Id value temporarily for Listed Country lookup
                stock_id_value = None

                for dest_col_name, source_col_index in current_table_column_indices.items():
                    try:
                        cell_value = row_cells_values[source_col_index]
                        processed_value = cell_value 
                        
                        # Handle "nan" values
                        if pd.isna(processed_value):
                            processed_value = None
                        
                        # Type-specific processing
                        if dest_col_name in ["Units Held", "Value (AUD)", "Weighting", "% Ownership"]: # Updated "Weighting (%)" to "Weighting"
                            if isinstance(processed_value, (int, float)):
                                pass
                            elif isinstance(processed_value, str):
                                try:
                                    cleaned_str = processed_value.strip().replace(",", "").replace("$", "").replace("%", "")
                                    if cleaned_str: 
                                        processed_value = float(cleaned_str)
                                        if dest_col_name == "% Ownership" and processed_value > 1 and processed_value <= 100:
                                            processed_value /= 100.0 
                                    else: 
                                        processed_value = None 
                                except ValueError:
                                    processed_value = None 
                            else:
                                processed_value = None 

                        elif dest_col_name == "Effective Date":
                            if isinstance(processed_value, (datetime.date, datetime.datetime)):
                                processed_value = processed_value.date()
                            elif isinstance(processed_value, str):
                                try: 
                                    processed_value = datetime.datetime.strptime(processed_value, '%Y-%m-%d').date()
                                except ValueError:
                                    try:
                                        processed_value = datetime.datetime.strptime(processed_value, '%m/%d/%Y').date()
                                    except ValueError:
                                        processed_value = None
                            else:
                                processed_value = None
                        
                        # Store Stock Id value for later lookup
                        elif dest_col_name == "Stock Id":
                            stock_id_value = str(processed_value).strip() if processed_value is not None else None
                            # Ensure Stock Id itself is added to data_to_append
                            processed_value = stock_id_value
                        
                        # For other string-based columns, ensure it's a cleaned string and handle "Total"
                        elif isinstance(processed_value, (str, int, float, datetime.date, datetime.datetime)):
                            processed_value = str(processed_value).strip() if processed_value is not None else None
                            # Change "Total" to "Sub Total"
                            if processed_value is not None and processed_value.lower() == TOTAL_KEYWORD:
                                processed_value = "Sub Total"
                        else:
                            processed_value = None 

                        dest_col_index = destination_columns_ordered.index(dest_col_name)
                        data_to_append[dest_col_index] = processed_value

                        if dest_col_name not in ["Effective Date", "Fund Name", "Option Name", "Asset Class Name", "Int/Ext", "Listed Country"]: # Exclude Listed Country here
                            if processed_value is not None and str(processed_value).strip() != "":
                                row_has_meaningful_data_beyond_context = True

                    except IndexError:
                        print(f"    [WARNING] Column index {source_col_index} out of bounds for row {row_idx + 1}. Skipping value for {dest_col_name}.")
                        data_to_append[destination_columns_ordered.index(dest_col_name)] = None 
                    except Exception as e:
                        print(f"    [ERROR] Error processing value for column '{dest_col_name}' in row {row_idx + 1}: '{cell_value}' - {e}")
                        data_to_append[destination_columns_ordered.index(dest_col_name)] = None
                
                # --- LOGIC for "Listed Country" ---
                listed_country_index = destination_columns_ordered.index("Listed Country")
                if stock_id_value and len(stock_id_value) >= 2:
                    country_code = stock_id_value[:2].upper() # Get first two chars and convert to uppercase
                    mapped_country = country_code_map.get(country_code)
                    if mapped_country:
                        data_to_append[listed_country_index] = mapped_country
                        print(f"    [DEBUG] Mapped Stock Id '{stock_id_value}' to Listed Country: '{mapped_country}'")
                    else:
                        # If no country mapping found, retain the 2-character code
                        data_to_append[listed_country_index] = country_code
                        print(f"    [DEBUG] No country mapping found for code '{country_code}' from Stock Id '{stock_id_value}'. Retaining code.")
                else:
                    data_to_append[listed_country_index] = None # No Stock Id or too short
                    print(f"    [DEBUG] No valid Stock Id for country lookup for row {row_idx + 1}.")
                # --- END LOGIC for "Listed Country" ---


                if row_has_meaningful_data_beyond_context or is_total_row:
                    all_extracted_rows_for_append.append(data_to_append)
                    data_rows_transferred += 1
                    if is_total_row:
                        print(f"    [DEBUG] Appended Total row from row {row_idx + 1}.")
                else:
                    print(f"    [DEBUG] Skipping row {row_idx + 1} within table due to no meaningful data detected (after context columns).")

            else:
                print(f"    [DEBUG] Skipping row {row_idx + 1} (no name/total keyword) within table boundaries.")

            if is_total_row:
                print(f"  [INFO] End of Table at Row {row_idx + 1} (Keyword '{TOTAL_KEYWORD}' found and processed). Resetting state for next table.")
                current_state = 0
                current_asset_class_name = None 
                current_int_ext = None
                current_table_column_indices = {} 
                found_int_ext_for_current_table = False
                found_asset_class_for_current_table = False
            
    print(f"\n[INFO] Processing complete.")
    print(f"[INFO] Total rows processed in source sheet: {rows_processed_count}")
    print(f"[INFO] Total tables identified: {tables_found_count}")
    print(f"[INFO] Total data rows extracted from source: {data_rows_transferred}")

    # --- Append extracted data to destination DataFrame and save ---
    if all_extracted_rows_for_append:
        df_new_data = pd.DataFrame(all_extracted_rows_for_append, columns=destination_columns_ordered)
        df_combined = pd.concat([df_dest, df_new_data], ignore_index=True)
        
        try:
            df_combined.to_csv(destination_file, index=False, encoding='utf-8')
            print(f"[INFO] Successfully saved '{destination_file}' with extracted and appended data.")
        except Exception as e:
            print(f"[ERROR] Error saving the destination CSV file after data transfer: {e}")
    else:
        print("[INFO] No new data rows were extracted from the source file to append.")

# =========================================================================
# Main execution block
# =========================================================================
if __name__ == "__main__":
    cleaned_data_filename = "D:\\LinhDao\\Programming\\SUPERFUNdProject\\Linh-Unisuper-CleanedData.csv"
    source_file_path = r"D:\LinhDao\Programming\SUPERFUNdProject\Unisuper_First_Table_Extracted.csv"
    country_codes_path = r"D:\LinhDao\Programming\SUPERFUNdProject\InternationalCountryCodes.csv" # New path for country codes

    print(f"[INFO] Attempting to create or overwrite '{cleaned_data_filename}'...")
    create_csv_file_with_headers(cleaned_data_filename)

    print(f"[INFO] Attempting to process source file '{source_file_path}' and populate '{cleaned_data_filename}'...")
    process_complicated_source_csv(
        source_file_path=source_file_path,
        destination_file=cleaned_data_filename,
        country_codes_file=country_codes_path # Pass the new argument
    )

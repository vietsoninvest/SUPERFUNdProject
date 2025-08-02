import pandas as pd
import numpy as np
import re
import os

def clean_sheet_name(name):
    """Cleans a string to be a valid Excel sheet name (max 31 chars, no invalid chars)."""
    # Convert to string, remove leading/trailing whitespace
    name = str(name).strip()
    # Remove invalid characters: / \ ? * [ ] :
    invalid_chars = r'[\\/?*\[\]:]'
    cleaned_name = re.sub(invalid_chars, '', name)
    # Replace single quotes with nothing (often problematic)
    cleaned_name = cleaned_name.replace("'", "")
    # Replace double quotes with nothing
    cleaned_name = cleaned_name.replace('"', "")
    # Truncate to 31 characters
    return cleaned_name[:31]


def split_excel_tables_by_marker_only(
    input_excel_path,
    output_excel_path,
    table_start_marker_keyword="Table",
    header_confirmation_keywords=None
):
    """
    Splits multiple tables from one Excel sheet into separate sheets in a new Excel file.
    Tables are defined solely by a 'table_start_marker_keyword' row, followed by an empty row,
    then the header. A new marker signifies the end of the previous table.

    Args:
        input_excel_path (str): Path to the input Excel file.
        output_excel_path (str): Path to the output Excel file where tables will be saved.
        table_start_marker_keyword (str): The keyword marking the start of a new table block.
                                          The content of this row is used as the sheet name.
        header_confirmation_keywords (list, optional): A list of strings (e.g., ['Value', 'Weighting'])
                                                      that must all be present in a row for it to be
                                                      confirmed as the header. Case-insensitive.
                                                      If None or empty, the row 2 positions after the marker is always used.

    Returns:
        list: A list of sheet names successfully created in the output Excel file. Returns an empty
              list if no tables are found or an error occurs.
    """
    print(f"--- Starting Stage 1: Initial Table Split (Marker-Only Logic) ---")
    print(f"  Input file: '{input_excel_path}'")
    print(f"  Output file: '{output_excel_path}'")
    print(f"  Table start marker: '{table_start_marker_keyword}' (will define both start and end of tables)")
    if header_confirmation_keywords:
        print(f"  Header confirmation keywords: {header_confirmation_keywords}")
    else:
        print(f"  No header confirmation keywords provided. The row 2 positions after the marker will be used as header.")

    generated_sheet_names = [] # To store names of sheets successfully created

    try:
        df = pd.read_excel(input_excel_path, sheet_name=0, header=None)
        print(f"  Successfully read {len(df)} rows from the Excel file.")
        if df.empty:
            print("  Warning: The DataFrame read from the Excel file is empty. No data to process.")
            print(f"--- Stage 1 Halted ---")
            return [] # Return empty list on empty DataFrame
    except FileNotFoundError:
        print(f"  ERROR: Input file not found at '{input_excel_path}'. Please ensure the path is correct and the file exists.")
        print(f"--- Stage 1 Halted ---")
        return [] # Return empty list on error
    except Exception as e:
        print(f"  ERROR: An unexpected error occurred while reading the Excel file: {e}")
        print(f"--- Stage 1 Halted ---")
        return [] # Return empty list on error

    # Prepare pattern for table start marker (case-insensitive, whole word match)
    table_marker_pattern = re.compile(rf'\b{re.escape(table_start_marker_keyword)}\b', re.IGNORECASE)

    # Helper to check if a row contains a pattern
    def _contains_pattern(row, pattern):
        return any(pattern.search(str(cell)) for cell in row if pd.notna(cell))

    # Identify table marker rows and blank rows (for skipping)
    is_table_marker_row = df.apply(lambda row: _contains_pattern(row, table_marker_pattern), axis=1)
    is_blank_row = df.apply(lambda row: row.astype(str).str.strip().eq('').all(), axis=1)

    identified_table_info = [] # Stores (sheet_name_content, header_row_idx, last_data_row_idx)
    
    current_table_header_idx_candidate = None
    current_sheet_name_raw = None

    print("\n  Analyzing rows to identify primary table sections...")
    for i in range(len(df)):
        # Check if the current row is a table start marker
        if is_table_marker_row.iloc[i]:
            # If we were tracking a previous table, it ends just before this new marker
            if current_table_header_idx_candidate is not None:
                # The previous table data ends at 'i - 1'
                if i - 1 >= current_table_header_idx_candidate: # Ensure there's data between header and end
                    identified_table_info.append((current_sheet_name_raw, current_table_header_idx_candidate, i - 1))
                    print(f"    --> Ended previous table. Header: {current_table_header_idx_candidate+1}, Data Ends: {i} (original Excel row numbers).")
                else:
                    print(f"    Warning: Previous table (header {current_table_header_idx_candidate+1}) had no data before new marker at row {i+1}. Skipping.")
            
            # Start tracking a new table
            # Marker at 'i', empty row at 'i+1', header at 'i+2'
            if i + 2 < len(df): # Check if the header row actually exists within bounds
                current_sheet_name_raw_list = df.iloc[i].astype(str).str.strip().tolist()
                # Take the first non-empty string in the marker row as the base for sheet name
                current_sheet_name_raw = next((s for s in current_sheet_name_raw_list if s), f"Table_{len(identified_table_info) + 1}")
                
                current_table_header_idx_candidate = i + 2
                print(f"    Found Table Marker at row {i+1}. New table header expected at row {current_table_header_idx_candidate+1}. Sheet name base: '{current_sheet_name_raw}'.")
                
                # Check for the expected empty row at i+1
                if not is_blank_row.iloc[i+1]:
                    print(f"    Warning: Expected an empty row at {i+2} after marker, but it's not empty for table '{current_sheet_name_raw}'. Proceeding with header at {i+3}.")
                
            else:
                print(f"    Warning: Table Marker found at row {i+1}, but no subsequent header row (needs i+2). Skipping this marker.")
                current_table_header_idx_candidate = None # Reset if table definition is incomplete

    # Handle the very last table if it hasn't been closed by another marker
    if current_table_header_idx_candidate is not None:
        identified_table_info.append((current_sheet_name_raw, current_table_header_idx_candidate, len(df) - 1))
        print(f"    --> Ended final table (reached end of file). Header: {current_table_header_idx_candidate+1}, Data Ends: {len(df)}.")

    if not identified_table_info:
        print("  INFO: No primary tables found based on the 'Table' marker criteria.")
        print(f"--- Stage 1 Finished (No Tables Found) ---")
        return [] # Return empty list if no tables found

    print(f"\n  Identified {len(identified_table_info)} primary tables for extraction.")
    for idx, (name, header_idx, end_idx) in enumerate(identified_table_info):
        print(f"    Table {idx+1}: Sheet Name: '{name}', Header at original Excel row {header_idx+1}, Data ends at original Excel row {end_idx+1}.")

    # Ensure the output directory exists
    output_dir = os.path.dirname(output_excel_path)
    if output_dir and not os.path.exists(output_dir):
        print(f"  Output directory '{output_dir}' does not exist. Attempting to create it...")
        try:
            os.makedirs(output_dir)
            print(f"  Successfully created output directory: {output_dir}")
        except OSError as e:
            print(f"  ERROR: Could not create output directory '{output_dir}'. Please check permissions or specify an existing directory.")
            print(f"  Error details: {e}")
            print(f"--- Stage 1 Halted ---")
            return [] # Return empty list on error
    else:
        print(f"  Output directory '{output_dir}' already exists or is current directory.")

    # Use ExcelWriter to write to multiple sheets
    try:
        # Using openpyxl to support appending in the next stage
        with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
            for i, (sheet_name_raw, header_row_idx, last_data_row_idx) in enumerate(identified_table_info):
                # Clean sheet name to be Excel-compatible
                sheet_name = clean_sheet_name(sheet_name_raw if sheet_name_raw else f"Table_{i+1}")
                generated_sheet_names.append(sheet_name) # Add to list of generated names
                print(f"\n  Processing primary Table {i+1} ('{sheet_name}'). Header: {header_row_idx+1}, Data Ends: {last_data_row_idx+1}.")

                # Extract the raw block from header row to last data row
                table_df_raw_block = df.iloc[header_row_idx : last_data_row_idx + 1].copy()
                
                # --- Header Extraction & Data Assignment ---
                header = [str(col).strip() if pd.notna(col) else '' for col in table_df_raw_block.iloc[0].values]
                data_rows_df = table_df_raw_block.iloc[1:].copy()
                data_rows_df.columns = header

                # Optional: Header confirmation using header_confirmation_keywords
                if header_confirmation_keywords:
                    lower_confirm_keywords = [k.lower() for k in header_confirmation_keywords]
                    header_cells_lower = [str(cell).lower().strip() for cell in header if pd.notna(cell)]
                    if not all(any(keyword_part in cell_value for cell_value in header_cells_lower) for keyword_part in lower_confirm_keywords):
                        print(f"    Warning: Header for '{sheet_name}' (original Excel row {header_row_idx+1}) does not contain all confirmation keywords: {header_confirmation_keywords}. It will still be used as header.")

                if data_rows_df.empty:
                    print(f"    Warning: Table '{sheet_name}' has no data rows after header. Skipping further processing for this table.")
                    continue

                # --- CLEANING STEPS ---
                print(f"    Initial data shape for '{sheet_name}': {data_rows_df.shape}")

                # 1. Remove rows that are entirely empty or contain only whitespace characters
                rows_before_row_drop = len(data_rows_df)
                rows_to_keep_mask = data_rows_df.apply(lambda row: not row.astype(str).str.strip().eq('').all(), axis=1)
                data_rows_df = data_rows_df[rows_to_keep_mask].copy()
                rows_after_row_drop = len(data_rows_df)
                if rows_before_row_drop > rows_after_row_drop:
                    print(f"    Removed {rows_before_row_drop - rows_after_row_drop} rows that were entirely empty or contained only whitespace from '{sheet_name}'.")
                else:
                    print(f"    No completely empty or whitespace-only rows removed from '{sheet_name}'.")

                # 2. Remove entirely empty columns
                cols_before_col_drop = len(data_rows_df.columns)
                data_rows_df.replace('', np.nan, inplace=True) # Convert empty strings to NaN for column dropping
                data_rows_df.dropna(axis=1, how='all', inplace=True) # Drop columns where ALL cells are NaN
                cols_after_col_drop = len(data_rows_df.columns)
                if cols_before_col_drop > cols_after_col_drop:
                    print(f"    Removed {cols_before_col_drop - cols_after_col_drop} entirely empty columns from '{sheet_name}'.")
                else:
                    print(f"    No entirely empty columns removed from '{sheet_name}'.")

                # Reset index after dropping rows, for cleaner output
                data_rows_df = data_rows_df.reset_index(drop=True)
                print(f"    Final data shape for '{sheet_name}': {data_rows_df.shape}")
                
                if data_rows_df.empty:
                    print(f"    Warning: Table '{sheet_name}' became empty after cleaning. Skipping saving this table.")
                    continue

                data_rows_df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"  Primary table successfully processed and saved to sheet '{sheet_name}'.")

        print(f"\nSUCCESS: Stage 1 tables successfully split, cleaned, and saved to '{output_excel_path}'")

    except Exception as e:
        print(f"  ERROR: An error occurred while writing to Excel file '{output_excel_path}': {e}")
        print(f"--- Stage 1 Halted ---")
        return [] # Return empty list upon error

    print(f"--- Stage 1 Finished ---")
    return generated_sheet_names # Return the list of generated sheet names


# def split_single_sheet_into_sub_tables(
#     excel_file_path,
#     parent_sheet_name,
#     sub_table_end_keyword="Total",
#     sub_header_confirmation_keywords=None
# ):
#     """
#     Performs a second level of splitting on a single specified sheet within an Excel file.
#     Sub-tables are defined by a 'name row' (not blank, not total), immediately followed by
#     a header row (containing specific keywords), and ending strictly with a 'Total' row.

#     Args:
#         excel_file_path (str): Path to the Excel file containing the parent sheet.
#         parent_sheet_name (str): The exact name of the sheet to split into sub-tables.
#         sub_table_end_keyword (str): The keyword marking the strict end of a sub-table.
#         sub_header_confirmation_keywords (list, optional): List of strings expected in sub-table headers.
#     """
#     print(f"\n--- Starting Stage 2: Sub-Table Split for sheet '{parent_sheet_name}' ---")
#     print(f"  Target Excel file: '{excel_file_path}'")
#     print(f"  Sub-table end keyword: '{sub_table_end_keyword}' (strict)")
#     if sub_header_confirmation_keywords:
#         print(f"  Sub-header confirmation keywords: {sub_header_confirmation_keywords}")
#     else:
#         print(f"  No sub-header confirmation keywords provided. Will assume row after name is header.")
    
#     try:
#         # Read the parent sheet (e.g., Table 1)
#         df_parent = pd.read_excel(excel_file_path, sheet_name=parent_sheet_name, header=None)
#         print(f"  Successfully read {len(df_parent)} rows from sheet '{parent_sheet_name}'.")
#         if df_parent.empty:
#             print(f"  Warning: Parent sheet '{parent_sheet_name}' is empty. No sub-tables to process.")
#             print(f"--- Stage 2 Finished (No Sub-Tables Found) ---")
#             return
#     except Exception as e:
#         print(f"  ERROR: Could not read parent sheet '{parent_sheet_name}' from '{excel_file_path}': {e}")
#         print(f"--- Stage 2 Halted ---")
#         return

#     # Prepare patterns
#     total_pattern = re.compile(rf'\b{re.escape(sub_table_end_keyword)}\b', re.IGNORECASE)

#     # Helper to check if a row contains a pattern
#     def _contains_pattern(row, pattern):
#         return any(pattern.search(str(cell)) for cell in row if pd.notna(cell))

#     # Identify row types
#     is_total_row = df_parent.apply(lambda row: _contains_pattern(row, total_pattern), axis=1)
#     is_blank_row = df_parent.apply(lambda row: row.astype(str).str.strip().eq('').all(), axis=1)

#     identified_sub_table_info = [] # Stores (sub_sheet_name_content, sub_header_row_idx, sub_last_data_row_idx)
    
#     current_sub_table_name_row_idx = None
#     current_sub_table_header_idx_candidate = None
#     current_sub_sheet_name_raw = None

#     print("\n  Analyzing rows in parent sheet for sub-tables...")
#     for i in range(len(df_parent)):
#         # If we are NOT currently tracking a sub-table, look for a new sub-table start (name row)
#         if current_sub_table_name_row_idx is None:
#             # A potential sub-table name row is NOT blank, NOT a total row
#             if not is_blank_row.iloc[i] and not is_total_row.iloc[i]:
#                 # And the next row (i+1) should be the header row
#                 if i + 1 < len(df_parent):
#                     candidate_header_row = df_parent.iloc[i+1]
#                     lower_candidate_header_cells = [str(cell).lower().strip() for cell in candidate_header_row if pd.notna(cell)]
                    
#                     is_valid_header_candidate = False
#                     if sub_header_confirmation_keywords:
#                         lower_confirm_keywords = [k.lower() for k in sub_header_confirmation_keywords]
#                         # Check if all confirmation keywords are present in the *next* row
#                         if all(any(keyword_part in cell_value for cell_value in lower_candidate_header_cells) for keyword_part in lower_confirm_keywords):
#                             is_valid_header_candidate = True
#                     else: # No header confirmation keywords provided, assume i+1 is header if i is name
#                         is_valid_header_candidate = True # Any non-blank row after a potential name is a header

#                     if is_valid_header_candidate:
#                         current_sub_table_name_row_idx = i
#                         # Extract sub-sheet name from the current row (the name row)
#                         # Filter out NaN/None, convert to string, strip whitespace
#                         clean_cells = [str(cell).strip() for cell in df_parent.iloc[i] if pd.notna(cell)]
#                         current_sub_sheet_name_raw = next((s for s in clean_cells if s), f"SubTable_{len(identified_sub_table_info) + 1}")

#                         current_sub_table_header_idx_candidate = i + 1 # Header is immediately after name row (Option B)
#                         print(f"    Found potential sub-table name at row {i+1}. Header expected at {i+2}. Identified name: '{current_sub_sheet_name_raw}'")
#                 # else: current row could be a name, but no header follows, so it's not a valid sub-table start

#         # If currently tracking a sub-table, check for its end (Total row)
#         elif current_sub_table_name_row_idx is not None:
#             if is_total_row.iloc[i]:
#                 # Sub-table ends at this Total row
#                 identified_sub_table_info.append((current_sub_sheet_name_raw, current_sub_table_header_idx_candidate, i))
#                 print(f"    --> Ended sub-table '{current_sub_sheet_name_raw}' with Total row at {i+1}.")
#                 # Reset for next sub-table
#                 current_sub_table_name_row_idx = None
#                 current_sub_table_header_idx_candidate = None
#                 current_sub_sheet_name_raw = None
#             # IMPORTANT: If a blank row appears, it's NOT a strict end according to the new rule,
#             # so we just continue looking for Total or next sub-table start.
    
#     # After the loop, if there's an active sub-table that didn't end with a 'Total'
#     # (because the sheet ended before finding a 'Total' for the last one)
#     if current_sub_table_name_row_idx is not None:
#         # Per user request, tables end "strictly with word Total".
#         # So, if a table is being tracked but no final Total is found, it's considered incomplete and not added.
#         print(f"    Warning: Last sub-table ('{current_sub_sheet_name_raw}') did not end with a 'Total' keyword. It will NOT be included as per 'strict end' rule.")
        
#     if not identified_sub_table_info:
#         print(f"  INFO: No sub-tables found in sheet '{parent_sheet_name}' based on the criteria.")
#         print(f"--- Stage 2 Finished (No Sub-Tables Found) ---")
#         return

#     print(f"\n  Identified {len(identified_sub_table_info)} sub-tables in '{parent_sheet_name}'.")

#     # Write sub-tables to the SAME Excel file in append mode
#     try:
#         # Open in append mode 'a'. This modifies the existing file.
#         # if_sheet_exists='replace' will overwrite sheets if they have the same name,
#         # which is good for re-running the script.
#         with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
#             for i, (sub_sheet_name_raw, sub_header_row_idx, sub_last_data_row_idx) in enumerate(identified_sub_table_info):
#                 # Construct combined sheet name: "Parent Name - Sub Table Name"
#                 full_sheet_name = clean_sheet_name(f"{parent_sheet_name} - {sub_sheet_name_raw}")
#                 print(f"\n  Processing Sub-Table {i+1} ('{full_sheet_name}')...")

#                 # Extract the raw block for the sub-table, starting from its name row
#                 # up to and including the total row.
#                 # `sub_header_row_idx` is already relative to df_parent's 0-index.
#                 # The sub-table's data block starts from the sub_table_name_row_idx
#                 # and goes up to the sub_last_data_row_idx (which is the Total row).
#                 # To include the Total row, we use +1 in slicing.
                
#                 # The header is at index 1 within this `sub_table_df_raw_block` (index 0 is the name row).
#                 # Data starts at index 2 within this `sub_table_df_raw_block`.
#                 sub_table_df_raw_block = df_parent.iloc[sub_header_row_idx - 1 : sub_last_data_row_idx + 1].copy()

#                 sub_header = [str(col).strip() if pd.notna(col) else '' for col in sub_table_df_raw_block.iloc[1].values] # Header is at index 1 of this slice
#                 sub_data_df = sub_table_df_raw_block.iloc[2:].copy() # Data starts after name row (index 0) + header row (index 1)
#                 sub_data_df.columns = sub_header

#                 if sub_data_df.empty:
#                     print(f"    Warning: Sub-Table '{full_sheet_name}' has no data rows after header. Skipping saving.")
#                     continue

#                 # --- CLEANING STEPS (re-apply for sub-tables) ---
#                 print(f"    Initial data shape for '{full_sheet_name}': {sub_data_df.shape}")

#                 # 1. Remove rows that are entirely empty or contain only whitespace
#                 rows_before_row_drop = len(sub_data_df)
#                 rows_to_keep_mask = sub_data_df.apply(lambda row: not row.astype(str).str.strip().eq('').all(), axis=1)
#                 sub_data_df = sub_data_df[rows_to_keep_mask].copy()
#                 rows_after_row_drop = len(sub_data_df)
#                 if rows_before_row_drop > rows_after_row_drop:
#                     print(f"    Removed {rows_before_row_drop - rows_after_row_drop} empty/whitespace rows from '{full_sheet_name}'.")
#                 else:
#                     print(f"    No completely empty or whitespace-only rows removed from '{full_sheet_name}'.")

#                 # 2. Remove entirely empty columns
#                 cols_before_col_drop = len(sub_data_df.columns)
#                 sub_data_df.replace('', np.nan, inplace=True)
#                 sub_data_df.dropna(axis=1, how='all', inplace=True)
#                 cols_after_col_drop = len(sub_data_df.columns)
#                 if cols_before_col_drop > cols_after_col_drop:
#                     print(f"    Removed {cols_before_col_drop - cols_after_col_drop} empty columns from '{full_sheet_name}'.")
#                 else:
#                     print(f"    No entirely empty columns removed from '{full_sheet_name}'.")

#                 sub_data_df = sub_data_df.reset_index(drop=True)
#                 print(f"    Final data shape for '{full_sheet_name}': {sub_data_df.shape}")
                
#                 if sub_data_df.empty:
#                     print(f"    Warning: Sub-Table '{full_sheet_name}' became empty after cleaning. Skipping saving.")
#                     continue

#                 sub_data_df.to_excel(writer, sheet_name=full_sheet_name, index=False)
#                 print(f"  Sub-table successfully saved to sheet '{full_sheet_name}'.")
#         print(f"\nSUCCESS: Sub-tables from '{parent_sheet_name}' successfully processed and added to '{excel_file_path}'")

#     except Exception as e:
#         print(f"  ERROR: An error occurred while writing sub-tables to Excel file '{excel_file_path}': {e}")
#         print(f"--- Stage 2 Halted ---")

#     print(f"--- Stage 2 Finished ---")


# --- Main execution block ---
if __name__ == "__main__":
    # Define your input and output file paths
    # IMPORTANT: Use raw string (r"...") for Windows paths to avoid issues with backslashes
    # Example: your_input_file = r"C:\Users\YourUser\Documents\my_excel_file.xlsx"
    your_input_file = r"D:\LinhDao\Programming\SUPERFUNdProject\Linh-caresuper.xlsx"
    
    # Suggesting an output file name and path in the same directory as the input
    output_dir = os.path.dirname(your_input_file)
    input_filename_without_ext = os.path.splitext(os.path.basename(your_input_file))[0]
    # This will be the output file for the first stage and input for the second stage
    main_output_file = os.path.join(output_dir, f"{input_filename_without_ext}_splitted_tables.xlsx")

    # --- Stage 1 Configuration Parameters ---
    # Keyword marking the START of each new primary table (e.g., "Table 1", "Table A", etc.)
    # The content of this row will be used as the base for the primary sheet name.
    primary_table_marker = "Table"
    # A list of words expected to be present in the primary table header row.
    # Used for confirmation; the header is always assumed to be 2 rows after the marker.
    primary_header_confirmation_keywords = ['Value', 'Weighting'] 

    # --- Stage 2 Configuration Parameters ---
    # Keyword for ending sub-tables (e.g., "Total", "Grand Total"). This is a STRICT end.
    sub_table_end_keyword = "Total"

    # Header confirmation keywords for sub-tables.
    # Now includes 'Name', 'Value', and 'Weighting'.
    # Sub-table headers are assumed to be immediately after their name row.
    sub_header_confirmation_keywords = ['Name', 'Value', 'Weighting'] # <--- UPDATED THIS LINE

    # --- Execute Stage 1: Initial Table Splitting ---
    print("\nStarting the Excel splitting process...")
    
    # Call Stage 1 and capture the list of sheet names it generates
    generated_primary_sheet_names = split_excel_tables_by_marker_only(
        input_excel_path=your_input_file,
        output_excel_path=main_output_file,
        table_start_marker_keyword=primary_table_marker,
        header_confirmation_keywords=primary_header_confirmation_keywords
    )

    # --- Execute Stage 2: Sub-Table Splitting ---
    # Check if Stage 1 generated any sheets and if the output file exists
    # if generated_primary_sheet_names and os.path.exists(main_output_file):
    #     # We will target the first sheet generated by Stage 1 for the sub-split
    #     sheet_to_further_split = generated_primary_sheet_names[0] 
        
    #     print(f"\n--- Executing Stage 2: Sub-Table Splitting for automatically determined sheet '{sheet_to_further_split}' ---")
    #     split_single_sheet_into_sub_tables(
    #         excel_file_path=main_output_file,
    #         parent_sheet_name=sheet_to_further_split, # Automatically determined from Stage 1 output
    #         sub_table_end_keyword=sub_table_end_keyword,
    #         sub_header_confirmation_keywords=sub_header_confirmation_keywords
    #     )
    # elif not generated_primary_sheet_names:
    #     print(f"\nStage 2 skipped: Stage 1 did not generate any primary tables to further split.")
    # else: # os.path.exists(main_output_file) is False
    #     print(f"\nStage 2 skipped: Main output file '{main_output_file}' not found. Stage 1 might have failed.")

    print("\n--- All Excel splitting processes finished ---")
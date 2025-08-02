import pandas as pd
import re
import os # Import os module for path manipulation

def extract_first_table_by_marker_to_csv(input_filename="D:\\LinhDao\\Programming\\SUPERFUNdProject\\Unisuper_Balanced_Investment_Option.csv", output_filename="D:\\LinhDao\\Programming\\SUPERFUNdProject\\Unisuper_First_Table_Extracted.csv"):
    """
    Reads a CSV file, identifies the first table strictly starting with "TABLE X",
    extracts that table into a DataFrame, and saves it to a single CSV file.
    The extraction stops when the row "TOTAL INVESTMENT ITEMS" is encountered,
    excluding that row.
    Ensures all columns are retained, assigning generic names to empty headers.

    Args:
        input_filename (str): The path to the input CSV file.
        output_filename (str): The path for the single output CSV file.
    """
    try:
        # Read the CSV file into a pandas DataFrame.
        # header=None is crucial because the "headers" for sub-tables are within the data,
        # not at the very top of the file. This reads all rows as data.
        # skipinitialspace=True helps handle CSVs where fields might have leading spaces after commas.
        df = pd.read_csv(input_filename, header=None, skipinitialspace=True)
        print(f"Successfully loaded '{input_filename}'.")

        # Step 1: Find the start of the FIRST table marked by "TABLE X"
        first_table_marker_idx = -1
        table_name_for_first_table = "Unnamed_First_Table" # Default name

        # Iterate through the DataFrame to find the first "TABLE X" marker
        for i, cell_value in enumerate(df.iloc[:, 0].astype(str)):
            # Use regex to strictly match the pattern "TABLE " followed by digits.
            if re.match(r"TABLE\s+\d+", cell_value.upper()):
                first_table_marker_idx = i
                print(f"Found first 'TABLE X' marker at row {i}: '{cell_value}'.")
                
                # Attempt to get the descriptive name of this table (usually the row after "TABLE X")
                try:
                    table_name_candidate = df.iloc[i + 1, 0]
                    if pd.isna(table_name_candidate) or str(table_name_candidate).strip() == "":
                        table_name_for_first_table = "Unnamed_First_Table"
                    else:
                        table_name_for_first_table = str(table_name_candidate).strip()
                    print(f"Identified name for the first table: '{table_name_for_first_table}'")
                except IndexError:
                    print(f"Warning: 'TABLE' marker at row {i} but no subsequent row for table name. Using 'Unnamed_First_Table'.")
                    table_name_for_first_table = "Unnamed_First_Table"
                
                break # Found the first table marker, stop searching

        if first_table_marker_idx == -1:
            print("Error: Could not find any 'TABLE X' markers in the file. No table extracted.")
            return

        # Step 2: Determine the end of the first table
        # The actual data for the table (including its own header row) starts
        # 2 rows after the "TABLE X" marker (1 for "TABLE X", 1 for descriptive name).
        data_start_row_for_extraction = first_table_marker_idx + 2
        
        # Initialize end_idx to the end of the DataFrame as a fallback.
        # This will be updated if the specific stopping condition ("TOTAL INVESTMENT ITEMS") is met.
        end_idx = len(df)

        # Iterate from the row where the data *should* start to find the stopping condition.
        # This loop defines where the extraction should stop for the first table.
        for i in range(data_start_row_for_extraction, len(df)):
            current_cell_value = str(df.iloc[i, 0]).strip()
            
            # NEW STOPPING CONDITION: Stop when "TOTAL INVESTMENT ITEMS" is found
            # Exclude this row by setting end_idx to 'i'
            if current_cell_value == "TOTAL INVESTMENT ITEMS":
                end_idx = i # Set end_idx to the row *before* "TOTAL INVESTMENT ITEMS"
                print(f"Stopping extraction at row {i} (excluding 'TOTAL INVESTMENT ITEMS').")
                break # Stop scanning for end condition
        
        # Check if there's actual data to extract for this table.
        # This checks if the determined end_idx is valid relative to the start.
        if data_start_row_for_extraction >= end_idx:
            print(f"Warning: No data found for the first table after its header at row {first_table_marker_idx}. Skipping extraction.")
            return

        # Step 3: Extract the data for the first table.
        # The slice starts from the row *after* the descriptive name (which is data_start_row_for_extraction)
        # and goes up to (but not including) end_idx.
        # .copy() is used to ensure we're working on a separate copy of the DataFrame slice,
        # preventing potential SettingWithCopyWarning issues.
        table_df = df.iloc[data_start_row_for_extraction:end_idx].copy()
        
        # Step 4: Process column headers.
        # The first row of the extracted table_df contains the actual column headers for this sub-table.
        raw_columns = table_df.iloc[0].astype(str)
        
        new_columns = []
        unnamed_counter = 1
        for col_name in raw_columns:
            cleaned_col_name = str(col_name).strip() # Ensure it's a string and remove leading/trailing whitespace
            if not cleaned_col_name: # If the header is empty after stripping
                new_columns.append(f"Unnamed_Column_{unnamed_counter}") # Assign a generic name
                unnamed_counter += 1
            else:
                new_columns.append(cleaned_col_name) # Use the cleaned header
        
        # Assign the newly generated column names to the DataFrame.
        table_df.columns = new_columns
        
        # Remove the row that was used as the header from the DataFrame's data.
        # This is the row at index 0 of the *newly created* table_df.
        # .reset_index(drop=True) resets the DataFrame index after slicing.
        table_df = table_df[1:].reset_index(drop=True) 

        # Further clean column names:
        # 1. Remove any non-alphanumeric characters (except spaces) using regex.
        # 2. Replace one or more spaces with a single underscore.
        # This makes column names more consistent and easier to use programmatically.
        table_df.columns = table_df.columns.str.replace(r'[^\w\s]', '', regex=True).str.replace(r'\s+', '_', regex=True)
        
        # Step 5: Write the extracted table DataFrame to the single CSV file.
        if not table_df.empty:
            # The output filename is set to "First_Table_Extracted.csv" by default
            table_df.to_csv(output_filename, index=False, encoding='utf-8')
            print(f"Successfully extracted the first table ('{table_name_for_first_table}') and saved it to '{output_filename}'. Rows: {len(table_df)}")
        else:
            print(f"The first table ('{table_name_for_first_table}') had no data after extraction. No CSV file was created.")

    except FileNotFoundError:
        print(f"Error: The file '{input_filename}' was not found. Please make sure it's in the correct directory.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

# Call the function to process your CSV file
if __name__ == "__main__":
    extract_first_table_by_marker_to_csv()

import openpyxl
import csv

# Path to your Excel file, updated to your specified location
excel_file_path = "D:\\LinhDao\\Programming\\SUPERFUNdProject\\UniSuperOptionHoldings.xlsx"

# The name of the output CSV file
csv_output_file = "Unisuper_Balanced_Investment_Option.csv"

# List to hold the data of the target table
table_data = []

try:
    # Load the workbook
    workbook = openpyxl.load_workbook(excel_file_path)
    
    # Select the active sheet
    sheet = workbook.active
    
    # Flag to indicate when we should start capturing rows
    capturing_data = False
    
    print(f"Searching for the table with 'Investment Option Name' and 'Balanced'...")
    
    # Iterate through all rows in the sheet
    for row in sheet.iter_rows():
        # Get the value from the first cell of the current row
        cell_1_value = row[0].value
        
        # Check if this is a header row (start of a table)
        is_header_row = str(cell_1_value).strip() == "INVESTMENT OPTION NAME"
        
        # If we are currently capturing data, and we see a new header, stop.
        if capturing_data and is_header_row:
            capturing_data = False
            break  # Stop processing after the table is captured

        # Check if this is the start of our *target* table
        if is_header_row and str(row[1].value).strip() == "Balanced":
            capturing_data = True
            print("Target table found. Starting data capture.")
        
        # If we are currently capturing data, append the row's values
        if capturing_data:
            # Append the row values to our data list
            table_data.append([cell.value for cell in row])

    # Check if any data was captured
    if table_data:
        # Write the captured data to a new CSV file
        with open(csv_output_file, 'w', newline='', encoding='utf-8') as csvfile:
            csv_writer = csv.writer(csvfile)
            csv_writer.writerows(table_data)
            
        print(f"Successfully extracted the table and saved it to '{csv_output_file}'")
    else:
        print("No table matching the criteria was found.")

except FileNotFoundError:
    print(f"Error: The file '{excel_file_path}' was not found.")
except Exception as e:
    print(f"An error occurred: {e}")
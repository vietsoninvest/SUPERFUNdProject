import pandas as pd
import os

def extract_country_codes_to_csv(url, output_filename="country_codes.csv"):
    """
    Extracts the 'Country' and 'Alpha-2 Code' columns from a table on a given URL,
    renames 'Alpha-2 Code' to 'Code', and saves the data to a CSV file.
    """
    print(f"[INFO] Attempting to read tables from {url}...")
    try:
        # Use pandas to read all tables on the page, explicitly setting the encoding.
        tables = pd.read_html(url, encoding='utf-8')
    except ValueError as e:
        print(f"[ERROR] No tables found on the page or an error occurred: {e}")
        return
    except Exception as e:
        print(f"[ERROR] An unexpected error occurred: {e}")
        return

    if tables:
        df = tables[0]
        print(f"[INFO] Found a table with columns: {list(df.columns)}")

        # Clean and rename columns to a consistent format.
        df.columns = [col.strip() for col in df.columns]
        
        # Correct the column name to match the website's table
        required_columns = ['Country', 'Alpha-2 code'] 
        
        # The column name on the website is actually 'Alpha-2 code',
        # with a lowercase 'c'
        # The website might have changed, or the initial check failed.
        # Let's check for both possibilities and adjust the code
        if 'Alpha-2 code' in df.columns:
            required_columns = ['Country', 'Alpha-2 code']
        elif 'Alpha-2 Code' in df.columns:
            required_columns = ['Country', 'Alpha-2 Code']
        else:
            print("[ERROR] The required columns 'Country' and 'Alpha-2 code' were not found in the table.")
            print(f"[INFO] Available columns are: {list(df.columns)}")
            return

        if all(col in df.columns for col in required_columns):
            # Select the required columns
            country_codes_df = df[required_columns]

            # Rename the 'Alpha-2 code' column to 'Code' as requested
            country_codes_df.rename(columns={required_columns[1]: 'Code'}, inplace=True)

            # Save the DataFrame to a CSV file, explicitly setting the encoding to UTF-8.
            country_codes_df.to_csv(output_filename, index=False, encoding='utf-8-sig')
            print(f"[SUCCESS] Data successfully saved to '{output_filename}'")
            print(f"[INFO] {len(country_codes_df)} rows were written.")
        else:
            print("[ERROR] The required columns 'Country' and 'Alpha-2 code' were not found in the table.")
            print(f"[INFO] Available columns are: {list(df.columns)}")
    else:
        print("[ERROR] No tables were found on the specified webpage.")

# --- Main execution block ---
if __name__ == "__main__":
    target_url = "https://www.iban.com/country-codes"
    output_csv_file = "InternationalCountryCodes.csv"

    extract_country_codes_to_csv(target_url, output_csv_file)
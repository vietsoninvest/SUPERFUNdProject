import pandas as pd
import requests
import time
import urllib.parse
import os

# =========================================================================
# Function to get country using a more robust Nominatim query
# =========================================================================
def get_country_from_location_nominatim(location_query):
    """
    Uses the OpenStreetMap Nominatim API with a free-form query.
    Returns the country name or None if not found.
    """
    if not location_query or not isinstance(location_query, str) or len(location_query.strip()) == 0:
        return None

    encoded_query = urllib.parse.quote(location_query)
    url = f"https://nominatim.openstreetmap.org/search?q={encoded_query}&format=json&addressdetails=1&limit=1"

    headers = {'User-Agent': 'SUPERFUNdProject_LinhDao'}

    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        data = response.json()
        
        if data and len(data) > 0 and 'address' in data[0] and 'country' in data[0]['address']:
            return data[0]['address']['country']
        else:
            return None
    except requests.exceptions.RequestException as e:
        print(f"[ERROR] API request failed for query: '{location_query}'. Error: {e}")
        return None
    except Exception as e:
        print(f"[ERROR] An unexpected error occurred during API call: {e}")
        return None
    return None

# =========================================================================
# Main function to process the CSV file
# =========================================================================
def lookup_countries_in_csv(file_path):
    """
    Opens a CSV file, finds rows where 'Listed Country' is empty,
    and performs a multi-step lookup to populate the column.
    """
    print(f"[INFO] Starting country lookup process for file: '{file_path}'")
    
    if not os.path.exists(file_path):
        print(f"[ERROR] The file '{file_path}' was not found. Please check the path.")
        return

    try:
        df = pd.read_csv(file_path, encoding='utf-8-sig')
    except Exception as e:
        print(f"[ERROR] An error occurred while opening the CSV file: {e}")
        return

    required_columns = ['Address', 'Name/Kind of Investment Item', 'Listed Country']
    if not all(col in df.columns for col in required_columns):
        print(f"[ERROR] Required columns '{required_columns}' not found in the CSV file. Available columns are: {list(df.columns)}")
        return

    rows_updated = 0
    total_rows_to_process = len(df[(df['Address'].notna()) & (df['Listed Country'].isna())])
    print(f"[INFO] {total_rows_to_process} rows require a country lookup.")

    for index, row in df.iterrows():
        address = row['Address']
        investment_name = row['Name/Kind of Investment Item']
        listed_country = row['Listed Country']

        if pd.notna(address) and pd.isna(listed_country):
            address_str = str(address).strip()
            
            country = None
            
            # --- Method 1: Combined Query (Name, Address) ---
            if pd.notna(investment_name):
                combined_query = f"{str(investment_name).strip()}, {address_str}"
                print(f"[INFO] Processing row {index + 1} with combined query: '{combined_query}'")
                country = get_country_from_location_nominatim(combined_query)

            # --- Method 2: Fallback to just the Address ---
            if country is None:
                print(f"[INFO] Combined query failed. Falling back to simple address query: '{address_str}'")
                country = get_country_from_location_nominatim(address_str)
            
            if country:
                df.loc[index, 'Listed Country'] = country
                rows_updated += 1
                print(f"[SUCCESS] Updated row {index + 1} with country: '{country}'")
            else:
                print(f"[WARNING] Could not determine country for row {index + 1} after all attempts.")
            
            # Add a delay to be respectful of the API's usage policy
            time.sleep(1.5)

    print(f"\n[INFO] Processing complete.")
    print(f"[INFO] {rows_updated} rows were updated with country information.")

    try:
        df.to_csv(file_path, index=False, encoding='utf-8-sig')
        print(f"[SUCCESS] File '{file_path}' has been successfully saved.")
    except Exception as e:
        print(f"[ERROR] An error occurred while saving the file: {e}")

# =========================================================================
# Main execution block
# =========================================================================
if __name__ == "__main__":
    csv_file_path = "Linh-caresuper-CleanedData.csv"
    lookup_countries_in_csv(csv_file_path)
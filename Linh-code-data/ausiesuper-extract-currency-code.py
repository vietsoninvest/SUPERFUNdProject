import pandas as pd
import os
import re # Import re for regular expressions
import requests # Import requests for fetching web content

def clean_country_currency_merge(country, currency):
    """
    Merges country and currency names, removing redundant *exact whole words*
    of the country name from the beginning of the currency part.
    Example: "United States" and "United States Dollar" becomes "United States (Dollar)".
    "Land" and "Landmark Dollar" remains "Land (Landmark Dollar)".
    """
    country_str = str(country).strip() if pd.notna(country) else ""
    currency_str = str(currency).strip() if pd.notna(currency) else ""

    if not currency_str:
        return country_str
    
    country_lower = country_str.lower()
    currency_lower = currency_str.lower()

    # Pattern to match the exact country name at the beginning of the currency string,
    # followed by a word boundary (\b) or end of string ($).
    # re.escape() is used to handle special characters in the country name (e.g., "Côte d'Ivoire")
    # re.IGNORECASE makes the match case-insensitive.
    pattern = r"^{}\b".format(re.escape(country_lower))
    
    # Check if the country name is a redundant prefix (e.g., "United States Dollar" and "United States")
    # and if it matches as a whole word at the beginning of the currency string.
    if country_lower and re.match(pattern, currency_lower, flags=re.IGNORECASE):
        # Replace the matched country name with an empty string, then strip leading/trailing spaces.
        # Use count=1 to replace only the first occurrence.
        remaining_currency_part = re.sub(pattern, '', currency_str, 1, flags=re.IGNORECASE).strip()
        
        # If there's a meaningful remaining part, append it in parentheses.
        # This handles cases like "United States" + "United States Dollar" -> "United States (Dollar)"
        if remaining_currency_part:
            return f"{country_str} ({remaining_currency_part})"
        else:
            # If after removing the country, nothing meaningful is left (e.g., "France" + "France"),
            # or if the currency was just the country name, return just the country.
            return country_str
    else:
        # If no whole-word redundancy at the beginning, just combine normally.
        return f"{country_str} ({currency_str})"


def extract_currency_data(url="https://www.iban.com/currency-codes", output_filename="D:\\LinhDao\\Programming\\SUPERFUNdProject\\CleanedCurrencyCodes.csv"):
    """
    Extracts currency code data from a given URL, merges 'Country' and 'Currency' columns
    (removing duplicated exact words), extracts 'Code' column, cleans special characters,
    and saves the result to a CSV file.

    Args:
        url (str): The URL of the webpage containing the currency codes table.
        output_filename (str): The path and name for the output CSV file.
    """
    print(f"[INFO] Attempting to extract data from: {url}")
    try:
        # Use requests to fetch the HTML content with explicit UTF-8 encoding
        response = requests.get(url)
        response.raise_for_status() # Raise an HTTPError for bad responses (4xx or 5xx)
        
        # Explicitly set encoding to UTF-8 for the content, as this is a common web encoding
        response.encoding = 'utf-8' 
        html_content = response.text # Get the content as text using the specified encoding

        # pandas.read_html returns a list of DataFrames found in the HTML.
        # Pass the correctly decoded HTML content.
        tables = pd.read_html(html_content)
        if not tables:
            print("[ERROR] No tables found on the webpage.")
            return

        df = tables[0] # Get the first table
        print(f"[DEBUG] Successfully extracted a table from the URL. Initial columns: {df.columns.tolist()}")

        # Normalize column names for easier access (strip whitespace)
        df.columns = df.columns.str.strip()

        # Check for required columns
        required_cols = ["Country", "Currency", "Code", "Number"]
        if not all(col in df.columns for col in required_cols):
            print(f"[ERROR] Expected columns not found in the extracted table. Found: {df.columns.tolist()}")
            print(f"Expected to find: {required_cols}")
            return

        # 1. Merge "Country" and "Currency" columns into a single "Country" column
        # Use the new helper function to handle duplicate words
        df['Country'] = df.apply(
            lambda row: clean_country_currency_merge(row['Country'], row['Currency']),
            axis=1
        )
        print("[INFO] 'Country' and 'Currency' columns merged into 'Country', with duplicate exact words removed.")

        # 2. Select "Country" and "Code" columns
        df_cleaned = df[["Country", "Code"]].copy()
        print(f"[INFO] Selected 'Country' and 'Code' columns. Final columns: {df_cleaned.columns.tolist()}")

        # 3. Fix errors for some special characters in the 'Country' column
        # This regex now explicitly includes the apostrophe (')
        # It keeps alphanumeric characters, spaces, parentheses, hyphens, and apostrophes.
        # The re.UNICODE flag (or re.U) makes \w match Unicode word characters (like Ô).
        df_cleaned['Country'] = df_cleaned['Country'].apply(
            lambda x: re.sub(r'[^\w\s\(\)\'-]', '', str(x), flags=re.UNICODE).strip() if pd.notna(x) else x
        )
        print("[INFO] Special characters cleaned in 'Country' column (apostrophes now allowed).")

        # Save the cleaned DataFrame to a CSV file with 'utf-8-sig' encoding
        df_cleaned.to_csv(output_filename, index=False, encoding='utf-8-sig') # Changed encoding here
        print(f"[SUCCESS] Cleaned currency codes saved to '{output_filename}'.")

    except requests.exceptions.HTTPError as e:
        print(f"[ERROR] HTTP Error: {e}. Could not retrieve data from the URL.")
    except requests.exceptions.ConnectionError as e:
        print(f"[ERROR] Connection Error: {e}. Please check your internet connection or the URL.")
    except requests.exceptions.Timeout as e:
        print(f"[ERROR] Timeout Error: {e}. The request timed out.")
    except requests.exceptions.RequestException as e:
        print(f"[ERROR] An unexpected requests error occurred: {e}")
    except Exception as e:
        print(f"[ERROR] An error occurred during data extraction or processing: {e}")

# Main execution block
if __name__ == "__main__":
    output_csv_file = "D:\\LinhDao\\Programming\\SUPERFUNdProject\\CleanedCurrencyCodes.csv"
    extract_currency_data(output_filename=output_csv_file)

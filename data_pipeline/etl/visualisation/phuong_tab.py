import streamlit as st
import pandas as pd
import sys
import os

# Ensure db_connection is importable
ETL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ETL_DIR not in sys.path:
    sys.path.append(ETL_DIR)

from db_connection.db_connect import get_connection

def render():
    st.header("Phuong")
    st.info("Container for Phuong's charts and tables.")
    
    # Create a container for the content
    with st.container():
        placeholder = st.empty()  # will be used to display dataframe or messages
        try:
            # Get list of unique fund_name values for the dropdown
            with get_connection() as conn:
                with conn.cursor() as cur:
                    cur.execute("SELECT DISTINCT fund_name FROM final_data ORDER BY fund_name;")
                    rows = cur.fetchall()
            funds = [r[0] for r in rows]

            if not funds:
                placeholder.warning("No funds found in final_data table.")
                return {"placeholder": placeholder}

            selected_fund = st.selectbox("Choose fund_name", funds, index=0)

            # Query rows for the selected fund and render as a DataFrame
            with get_connection() as conn:
                with conn.cursor() as cur:
                    cur.execute(
                        "SELECT * FROM final_data WHERE fund_name = %s;",
                        (selected_fund,)
                    )
                    rows = cur.fetchall()
                    cols = [desc[0] for desc in cur.description]

            df = pd.DataFrame(rows, columns=cols)
            if df.empty:
                placeholder.info(f"No rows for fund '{selected_fund}'.")
            else:
                placeholder.dataframe(df)

        except Exception as e:
            placeholder.error(f"Error accessing database: {e}")

    return {"placeholder": placeholder}
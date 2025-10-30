# ...existing code...
import os
import sys
import importlib.util
import streamlit as st
import pandas as pd

# Ensure sibling package `db_connection` is importable (etl directory)
ETL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ETL_DIR not in sys.path:
    sys.path.append(ETL_DIR)

try:
    from db_connection.db_connect import get_connection
except Exception:
    get_connection = None

st.set_page_config(page_title="SUPERFUNd Project Dashboard", layout="wide")
st.title("SUPERFUNd Project Dashboard")

# Sidebar controls
with st.sidebar:
    st.header("Controls")
    test_conn = st.button("Test DB connection")
    refresh = st.button("Refresh placeholders")

# Helper to load a tab module from a file
def load_tab_module(path, module_name):
    if not os.path.exists(path):
        return None
    spec = importlib.util.spec_from_file_location(module_name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod

# Paths for tab modules (files in same directory as this dashboard)
base_dir = os.path.dirname(os.path.abspath(__file__))
tab_files = [
    ("Linh", os.path.join(base_dir, "linh_tab.py")),
    ("Son", os.path.join(base_dir, "son_tab.py")),
    ("Phuong", os.path.join(base_dir, "phuong_tab.py")),
]

# Create tabs and call each module.render() inside its tab
tabs = st.tabs([t[0] for t in tab_files])
placeholders = []

for idx, (tab_name, path) in enumerate(tab_files):
    mod = load_tab_module(path, f"visualisation.{tab_name.lower()}_tab")
    with tabs[idx]:
        if mod is None or not hasattr(mod, "render"):
            st.warning(f"Tab module for {tab_name} not found or has no render() — expected file: {os.path.basename(path)}")
            placeholders.append(None)
        else:
            try:
                returned = mod.render()  # expected to return dict with 'placeholder'
            except Exception as e:
                st.error(f"Error rendering {tab_name}: {e}")
                returned = None
            placeholders.append(returned)

# Handle sidebar actions (keep minimal)
if test_conn:
    if get_connection is None:
        st.sidebar.error("get_connection not available (import failed).")
    else:
        try:
            conn = get_connection()
            conn.close()
            st.sidebar.success("DB connection OK")
        except Exception as e:
            st.sidebar.error(f"DB connection failed: {e}")

if refresh:
    # Update each tab placeholder if available
    for p in placeholders:
        if p and isinstance(p, dict) and "placeholder" in p and p["placeholder"] is not None:
            try:
                p["placeholder"].write("Refreshed placeholder")
            except Exception:
                pass
# ...existing code...
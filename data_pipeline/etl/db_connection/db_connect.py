import psycopg2
import pandas as pd

# --- Database connection string---
SUPABASE_HOST = "aws-0-ap-southeast-1.pooler.supabase.com"
SUPABASE_PORT = 6543
SUPABASE_DB = "postgres"
SUPABASE_USER = "postgres.laremattjpwkgpmwitzv"
SUPABASE_PASSWORD = "Wombat2025@@!!"

def get_connection():
    return psycopg2.connect(
        host=SUPABASE_HOST,
        port=SUPABASE_PORT,
        dbname=SUPABASE_DB,
        user=SUPABASE_USER,
        password=SUPABASE_PASSWORD,
        sslmode="require"
    )

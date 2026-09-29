"""Tiny script run by GitHub Actions: touches the database so Supabase (free plan) never pauses."""
import os
import sys

import psycopg2

url = os.environ.get("DATABASE_URL", "")
if not url:
    sys.exit("DATABASE_URL secret is missing")
conn = psycopg2.connect(url, sslmode="require", connect_timeout=20)
cur = conn.cursor()
cur.execute("SELECT count(*) FROM transactions")
print("database alive, bills:", cur.fetchone()[0])
conn.close()

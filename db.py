"""
db.py - all database work lives here.

* On the cloud   -> PostgreSQL (Supabase). Set DATABASE_URL in Streamlit "Secrets".
* On your laptop -> if DATABASE_URL is not set, a local SQLite file (local_data.db)
                    is used automatically, so you can test without any setup.

The app keeps the same DataFrame model as before; this module only loads
them from / saves them to the database (inside ONE transaction, so a failed
save never leaves half-written data).
"""

from __future__ import annotations

import os

import pandas as pd
import streamlit as st
from sqlalchemy import create_engine, text
from sqlalchemy.pool import NullPool

import logic as lg

LOCAL_SQLITE = "sqlite:///local_data.db"

# app column name  ->  database column name
TX_DB = {
    "Txn ID": "txn_id", "Date": "bill_date", "Bill No": "bill_no", "Party / Company": "party",
    "Product / Item": "product_item", "Qty": "qty", "Subtotal": "subtotal", "Discount": "discount",
    "Tax": "tax", "Bill Amount": "bill_amount", "Paid Amount": "paid_amount",
    "Return Amount": "return_amount", "Balance": "balance", "Payment Method": "payment_method",
    "Cheque No": "cheque_no", "Cheque Date": "cheque_date", "Bank Name": "bank_name",
    "Cheque Status": "cheque_status", "Sale Base": "sale_base", "Status": "status", "Remarks": "remarks",
}
ITEM_DB = {"Txn ID": "txn_id", "Product": "product", "Qty": "qty", "Rate": "rate", "Amount": "amount"}
PROD_DB = {"Product": "product", "Unit": "unit", "Rate": "rate"}

TABLES = ("transactions", "items", "products", "parties")

DDL = [
    """CREATE TABLE IF NOT EXISTS transactions (
        txn_id TEXT PRIMARY KEY, bill_date DATE, bill_no TEXT, party TEXT, product_item TEXT,
        qty NUMERIC(14,3), subtotal NUMERIC(14,2), discount NUMERIC(14,2), tax NUMERIC(14,2),
        bill_amount NUMERIC(14,2), paid_amount NUMERIC(14,2), return_amount NUMERIC(14,2),
        balance NUMERIC(14,2), payment_method TEXT, cheque_no TEXT, cheque_date DATE, bank_name TEXT,
        cheque_status TEXT, sale_base TEXT, status TEXT, remarks TEXT)""",
    """CREATE TABLE IF NOT EXISTS items (
        txn_id TEXT, product TEXT, qty NUMERIC(14,3), rate NUMERIC(14,2), amount NUMERIC(14,2))""",
    "CREATE TABLE IF NOT EXISTS products (product TEXT PRIMARY KEY, unit TEXT, rate NUMERIC(14,2))",
    "CREATE TABLE IF NOT EXISTS parties (party TEXT PRIMARY KEY)",
    "CREATE INDEX IF NOT EXISTS idx_items_txn ON items (txn_id)",
    "CREATE INDEX IF NOT EXISTS idx_tx_party ON transactions (party)",
]


# ---------------------------------------------------------------- connection
def _secret(name: str) -> str:
    """Read a value from Streamlit secrets, then from environment variables."""
    try:
        if name in st.secrets:
            return str(st.secrets[name]).strip()
    except Exception:  # no secrets file (normal on your laptop)
        pass
    return os.environ.get(name, "").strip()


def get_secret(name: str) -> str:
    return _secret(name)


def database_url() -> str:
    url = _secret("DATABASE_URL")
    if not url:
        return LOCAL_SQLITE
    if url.startswith("postgres://"):
        url = "postgresql://" + url[len("postgres://"):]
    if url.startswith("postgresql://"):
        url = "postgresql+psycopg2://" + url[len("postgresql://"):]
    return url


def is_cloud() -> bool:
    """True when a real (PostgreSQL) database is configured."""
    return not database_url().startswith("sqlite")


@st.cache_resource(show_spinner=False)
def get_engine():
    url = database_url()
    if url.startswith("sqlite"):
        return create_engine(url)
    # NullPool = open a fresh connection per operation and close it. This is the
    # safest choice with Supabase's connection pooler (no stale connections).
    return create_engine(url, poolclass=NullPool, connect_args={"sslmode": "require", "connect_timeout": 15})


@st.cache_resource(show_spinner=False)
def init_db() -> bool:
    """Create the tables on first run (safe to run every time)."""
    eng = get_engine()
    with eng.begin() as conn:
        for stmt in DDL:
            conn.execute(text(stmt))
        if eng.dialect.name == "postgresql":
            # Row Level Security ON with no policies: Supabase's public web API can NOT
            # read your data. The app connects directly, so it is not affected.
            for t in TABLES:
                conn.execute(text(f"ALTER TABLE {t} ENABLE ROW LEVEL SECURITY"))
    return True


def ping() -> None:
    with get_engine().connect() as conn:
        conn.execute(text("SELECT 1"))


# ---------------------------------------------------------------- load / save
def _empty_frames():
    return (lg.recompute(lg._empty(lg.COLUMNS)),
            lg._empty(lg.ITEM_COLS).astype({"Qty": float, "Rate": float, "Amount": float}),
            lg._empty(lg.PRODUCT_COLS).astype({"Rate": float}), [])


@st.cache_data(ttl=300, show_spinner=False)
def _load_cached():
    """Read everything from the database. Cached; the cache is cleared on every save."""
    init_db()
    with get_engine().connect() as conn:
        tx = pd.read_sql(text("SELECT * FROM transactions"), conn)
        items = pd.read_sql(text("SELECT * FROM items"), conn)
        prods = pd.read_sql(text("SELECT * FROM products"), conn)
        parties = pd.read_sql(text("SELECT party FROM parties"), conn)

    tx = lg.recompute(tx.rename(columns={v: k for k, v in TX_DB.items()}))
    items = items.rename(columns={v: k for k, v in ITEM_DB.items()}).reindex(columns=lg.ITEM_COLS)
    prods = prods.rename(columns={v: k for k, v in PROD_DB.items()}).reindex(columns=lg.PRODUCT_COLS)
    for d, cols in ((items, ["Qty", "Rate", "Amount"]), (prods, ["Rate"])):
        for c in cols:
            d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0.0)
    items["Txn ID"] = items["Txn ID"].astype(str)
    items["Product"] = items["Product"].astype(str)
    prods["Product"] = prods["Product"].astype(str)
    prods["Unit"] = prods["Unit"].fillna("Pcs").astype(str)
    prods = prods.sort_values("Product", key=lambda s: s.str.lower()).reset_index(drop=True)
    return tx, items.reset_index(drop=True), prods, sorted(parties["party"].dropna().astype(str).tolist())


def load_all():
    """Returns fresh copies: (transactions, items, products, extra_parties)."""
    return _load_cached()


def _replace_table(conn, table: str, df: pd.DataFrame) -> None:
    conn.execute(text(f"DELETE FROM {table}"))
    if len(df):
        df.to_sql(table, conn, if_exists="append", index=False, method="multi", chunksize=200)


def save_all(tx, items, products, parties, only: tuple[str, ...] = TABLES) -> None:
    """Replace the chosen tables inside ONE transaction (all-or-nothing)."""
    init_db()
    with get_engine().begin() as conn:
        if "transactions" in only:
            _replace_table(conn, "transactions", tx[lg.COLUMNS].rename(columns=TX_DB))
        if "items" in only:
            _replace_table(conn, "items", items[lg.ITEM_COLS].rename(columns=ITEM_DB))
        if "products" in only:
            _replace_table(conn, "products", products[lg.PRODUCT_COLS].rename(columns=PROD_DB))
        if "parties" in only:
            p = pd.DataFrame({"party": sorted(set(x for x in parties if x))})
            _replace_table(conn, "parties", p)
    _load_cached.clear()  # next load sees the new data (for every user)

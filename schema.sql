-- OPTIONAL. The app creates these tables by itself on first start.
-- Run this in Supabase -> SQL Editor only if you want to create them manually first.

CREATE TABLE IF NOT EXISTS transactions (
    txn_id TEXT PRIMARY KEY, bill_date DATE, bill_no TEXT, party TEXT, product_item TEXT,
    qty NUMERIC(14,3), subtotal NUMERIC(14,2), discount NUMERIC(14,2), tax NUMERIC(14,2),
    bill_amount NUMERIC(14,2), paid_amount NUMERIC(14,2), return_amount NUMERIC(14,2),
    balance NUMERIC(14,2), payment_method TEXT, cheque_no TEXT, cheque_date DATE, bank_name TEXT,
    cheque_status TEXT, sale_base TEXT, status TEXT, remarks TEXT);
CREATE TABLE IF NOT EXISTS items (
    txn_id TEXT, product TEXT, qty NUMERIC(14,3), rate NUMERIC(14,2), amount NUMERIC(14,2));
CREATE TABLE IF NOT EXISTS products (product TEXT PRIMARY KEY, unit TEXT, rate NUMERIC(14,2));
CREATE TABLE IF NOT EXISTS parties (party TEXT PRIMARY KEY);
CREATE INDEX IF NOT EXISTS idx_items_txn ON items (txn_id);
CREATE INDEX IF NOT EXISTS idx_tx_party ON transactions (party);

-- Security: block Supabase's public web API from reading your business data.
-- (The app connects directly with the database password, so it keeps working.)
ALTER TABLE transactions ENABLE ROW LEVEL SECURITY;
ALTER TABLE items        ENABLE ROW LEVEL SECURITY;
ALTER TABLE products     ENABLE ROW LEVEL SECURITY;
ALTER TABLE parties      ENABLE ROW LEVEL SECURITY;

"""Run with:  python -m pytest tests/   (tests the calculations, Excel/PDF export and import engine)"""
import io
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import logic as lg  # noqa: E402


def make_tx():
    raw = pd.DataFrame({
        "Txn ID": ["TXN-000001", "TXN-000002", "TXN-000003", "TXN-000004"],
        "Date": ["2026-01-05", "2026-01-06", None, "2026-02-01"],
        "Bill No": ["B-0001", "B-0002", "B-0003", "B-0004"],
        "Party / Company": ["ALI TRADERS", "ALI TRADERS", "KHAN & SONS / CO", "ZED (PVT)"],
        "Bill Amount": [1000, 500, 200, 300],
        "Paid Amount": [1000, 200, 0, 300],
        "Return Amount": [0, 0, 0, 300],
        "Cheque Status": [None, None, "Pending", None],
    })
    return lg.recompute(raw)


def test_balance_and_status():
    tx = make_tx()
    assert tx["Balance"].tolist() == [0, 300, 200, -300]
    assert tx["Status"].tolist() == ["PAID", "PARTIALLY PAID", "CHEQUE PENDING", "RETURNED"]


def test_recompute_handles_empty():
    assert lg.recompute(lg._empty(lg.COLUMNS)).empty


def test_excel_export_with_illegal_sheet_name_and_control_chars():
    tx = make_tx()
    tx.loc[0, "Remarks"] = lg._txt(pd.Series(["bad\x01char"]))[0]
    data = lg.to_excel_bytes({"Ledger - KHAN & SONS / CO: [x]": tx})
    xl = pd.ExcelFile(io.BytesIO(data))
    assert len(xl.sheet_names[0]) <= 31 and "/" not in xl.sheet_names[0]


def test_pdf_and_invoice():
    tx = make_tx()
    assert lg.to_pdf_bytes("Test <report> & more", tx).startswith(b"%PDF")
    items = pd.DataFrame({"Txn ID": ["TXN-000002"], "Product": ["Rice"], "Qty": [5.0], "Rate": [100.0], "Amount": [500.0]})
    assert lg.invoice_pdf(tx.iloc[1], items).startswith(b"%PDF")
    assert lg.invoice_pdf(tx.iloc[0], items.iloc[0:0]).startswith(b"%PDF")


def test_import_flow():
    tx = make_tx()
    raw = pd.DataFrame({"Customer": ["new party", "ali traders", ""], "Invoice": ["", "B-0001", "X"],
                        "Date": ["01/03/2026", "2026-01-05", "01/03/2026"], "Item": ["Sugar", "Tea", "Tea"],
                        "Quantity": [2, 1, 1], "Price": ["1,500", "1000", "10"], "Paid": [0, 1000, 0]})
    d, used, ignored = lg.prepare_import(raw)
    assert len(d) == 2 and d["Bill Amount"].tolist() == [3000, 1000]
    new, dups = lg.split_duplicates(d, tx)
    assert dups == 1 and len(new) == 1
    items = lg._empty(lg.ITEM_COLS)
    prods = lg._empty(lg.PRODUCT_COLS)
    tx2, items2, prods2, n = lg.commit_import(new, True, tx, items, prods)
    assert n == 1 and len(tx2) == 5 and tx2["Txn ID"].is_unique and len(items2) == 1 and len(prods2) == 1


def test_import_products_and_ids():
    prods = lg._empty(lg.PRODUCT_COLS)
    p, n = lg.import_products(pd.DataFrame({"Name": ["Rice", "rice ", "Oil"], "Price": [10, 12, 5]}), prods)
    assert n == 3 and set(p["Product"]) >= {"Oil"}
    tx = make_tx()
    tx.loc[1, "Txn ID"] = "TXN-000001"  # duplicate id
    fixed = lg.ensure_txn_ids(tx)
    assert fixed["Txn ID"].is_unique
    assert lg.next_bill_no(fixed) == "B-0005"


def test_aging_and_party_statement():
    tx = make_tx()
    aging = lg.compute_aging(tx)
    assert "Current (0-30d)" in aging.columns
    assert "Total Outstanding" in aging.columns
    stmt = lg.compute_party_statement(tx, "ALI TRADERS")
    assert len(stmt) == 2
    assert "Running Balance" in stmt.columns
    # Net impact: 1st bill was 1000 - 1000 = 0. 2nd bill was 500 - 200 = 300. Running balance at end should be 300.
    assert stmt["Running Balance"].iloc[-1] == 300.0


def test_whatsapp_and_thermal_receipt():
    tx = make_tx()
    row = tx.iloc[1]
    wa = lg.whatsapp_invoice_text(row)
    assert "ALI TRADERS" in wa
    assert "B-0002" in wa
    assert "300" in wa  # balance
    url = lg.whatsapp_url("03001234567", wa)
    assert url.startswith("https://wa.me/03001234567?text=")

    items = pd.DataFrame([{"Txn ID": "TXN-000002", "Product": "Super Basmati", "Qty": 2, "Rate": 250, "Amount": 500}])
    receipt = lg.thermal_receipt_html(row, items)
    assert "Super Basmati" in receipt
    assert "SALE RECEIPT" in receipt


"""
Business Management System (v3 - cloud ready)
---------------------------------------------
app.py   = the screens (this file)
logic.py = calculations, Excel/PDF export, import engine
db.py    = PostgreSQL / Supabase storage

Run locally:   streamlit run app.py
"""

from __future__ import annotations

import hmac
import inspect
import re
import time
from datetime import date, timedelta

import pandas as pd
import streamlit as st

import db
import logic as lg
from logic import (
    BUSINESS_NAME, COLUMNS, CHEQUE_STATUSES, CUR, DATE_COLS, GROUPS, ITEM_COLS, MONEY, NEW_PARTY,
    NO_PRODUCT, PAYMENT_METHODS, PRODUCT_COLS, STATUS_OPTIONS, XLSX,
    _empty, _txt, commit_import, import_products, invoice_pdf, prepare_import, read_upload,
    recompute, split_duplicates,
)

# Works on both older and newer Streamlit versions
_STRETCH = ({"width": "stretch"} if "width" in inspect.signature(st.dataframe).parameters
            else {"use_container_width": True})


def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    return lg.to_excel_bytes(sheets)


def to_pdf_bytes(title: str, df: pd.DataFrame) -> bytes:
    return lg.to_pdf_bytes(title, df)


st.set_page_config(page_title=BUSINESS_NAME, page_icon="\U0001F4BC", layout="wide", initial_sidebar_state="expanded")
S = st.session_state
for _k, _v in {"page": "Dashboard", "ntx_fid": 0, "ntx_n": 1, "ver": 0, "line_items": pd.DataFrame(columns=ITEM_COLS)}.items():
    S.setdefault(_k, _v)


def workbook_sheets() -> dict[str, pd.DataFrame]:
    return {"Transactions": S.tx, "Items": S.line_items, "Products": S.products,
            "Parties": pd.DataFrame({"Party": sorted(set(S.extra_parties))})}


def save_all(*only: str) -> None:
    """Save to the database. If saving fails we STOP and say so - never pretend it worked."""
    try:
        db.save_all(S.tx, S.line_items, S.products, S.extra_parties, only or db.TABLES)
    except Exception as exc:
        st.error("Could not save to the database, so this change was NOT stored. "
                 "Please check your internet / database and try again.")
        st.caption(f"Technical detail: {type(exc).__name__}: {str(exc)[:300]}")
        st.stop()


def require_login() -> None:
    pwd = db.get_secret("APP_PASSWORD")
    if not pwd:
        if db.is_cloud():
            st.error("APP_PASSWORD is missing in Streamlit Secrets. Add it (see README) - "
                     "the app refuses to open without a password when it is online.")
            st.stop()
        return  # local testing: no password needed
    if S.get("auth"):
        return
    _, mid, _ = st.columns([1, 1.2, 1])
    with mid:
        st.markdown(f'<div class="page-title">{BUSINESS_NAME}</div>'
                    '<div class="page-sub">Please log in to continue.</div>', unsafe_allow_html=True)
        with st.form("login"):
            p = st.text_input("Password", type="password")
            if st.form_submit_button("Log in", type="primary"):
                if hmac.compare_digest(p.encode(), pwd.encode()):
                    S.auth, S.fails = True, 0
                    st.rerun()
                S.fails = S.get("fails", 0) + 1
                time.sleep(min(S.fails, 5))  # slows down password guessing
                st.error("Wrong password.")
    st.stop()


def all_parties() -> list[str]:
    return sorted(set(p for p in S.tx["Party / Company"] if p) | set(S.extra_parties))


def add_party_if_new(name: str) -> None:
    if name and name not in all_parties():
        S.extra_parties.append(name)


def next_txn_ids(n: int) -> list[str]:
    return lg.next_txn_ids(S.tx, n)


def next_bill_no() -> str:
    return lg.next_bill_no(S.tx)


def party_picker(label: str, key: str, allow_all: bool = False) -> str:
    opts = (["ALL PARTIES"] if allow_all else []) + all_parties() + [NEW_PARTY]
    choice = st.selectbox(label, opts, key=f"{key}_c")
    if choice == NEW_PARTY:
        return st.text_input("New party / company name", key=f"{key}_n", placeholder="Type the full name").strip().upper()
    return choice


def update_bill(tid: str, remark: str | None = None, **changes) -> None:
    m = S.tx["Txn ID"] == tid
    for k, v in changes.items():
        S.tx.loc[m, k] = v
    if remark:
        cur = S.tx.loc[m, "Remarks"].iloc[0]
        S.tx.loc[m, "Remarks"] = f"{cur} | {remark}".strip(" |")
    S.tx = recompute(S.tx)
    save_all("transactions")


def money(v: float) -> str:
    return f"{CUR} {v:,.0f}"


def badge(status: str) -> str:
    cls = {"PAID": "ok", "UNPAID": "bad", "PARTIALLY PAID": "warn", "CHEQUE PENDING": "orange", "RETURNED": "info"}
    return f'<span class="badge b-{cls.get(status, "warn")}">{status}</span>'



# =====================================================================
# THEME
# =====================================================================
st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"], .stApp { font-family: 'Inter', 'Segoe UI', Arial, sans-serif; }
.stApp { background: #F4F6FA; }
.block-container { padding-top: 2rem; max-width: 1500px; }
header[data-testid="stHeader"] { background: transparent; }

.page-title { font-size: 26px; font-weight: 800; color: #0F1F4D; letter-spacing: -.4px; }
.page-sub { font-size: 14px; color: #64748B; margin: 2px 0 22px 0; }
.section { font-size: 13px; font-weight: 700; color: #1E3A8A; text-transform: uppercase; letter-spacing: .8px;
  margin: 22px 0 10px 0; padding-bottom: 6px; border-bottom: 2px solid #DBE4F5; }

/* KPI cards */
div[data-testid="stMetric"] { background: #fff; border: 1px solid #E2E8F0; border-radius: 12px;
  padding: 16px 18px; box-shadow: 0 1px 2px rgba(15,31,77,.05); transition: all .18s ease; }
div[data-testid="stMetric"]:hover { transform: translateY(-2px); box-shadow: 0 8px 20px rgba(30,58,138,.12); border-color: #B6C7EE; }
div[data-testid="stMetricLabel"] p { color: #64748B; font-weight: 600; font-size: 13px; }
div[data-testid="stMetricValue"] { color: #0F1F4D; font-weight: 700; }

/* sidebar */
section[data-testid="stSidebar"] { background: linear-gradient(180deg, #0B1840 0%, #1E3A8A 100%); }
section[data-testid="stSidebar"] .brand { color: #fff; font-size: 18px; font-weight: 800; letter-spacing: -.2px; padding: 6px 4px 0; }
section[data-testid="stSidebar"] .tag { color: #A9BCEB; font-size: 12px; padding: 0 4px 6px; }
section[data-testid="stSidebar"] .grp { color: #7F97D6; font-size: 11px; font-weight: 700; letter-spacing: 1.2px;
  text-transform: uppercase; padding: 14px 6px 4px; }
section[data-testid="stSidebar"] .foot { color: #A9BCEB; font-size: 12px; padding: 4px; }
section[data-testid="stSidebar"] button { justify-content: flex-start; border: none; border-radius: 8px;
  padding: 8px 12px; font-weight: 500; transition: all .15s ease; box-shadow: none; }
section[data-testid="stSidebar"] button[data-testid="stBaseButton-secondary"] { background: transparent; color: #D6E0FA; }
section[data-testid="stSidebar"] button[data-testid="stBaseButton-secondary"]:hover { background: rgba(255,255,255,.10); color: #fff; padding-left: 16px; }
section[data-testid="stSidebar"] button[data-testid="stBaseButton-primary"] { background: rgba(255,255,255,.16); color: #fff;
  font-weight: 700; border-left: 3px solid #60A5FA; }
section[data-testid="stSidebar"] button p { text-align: left; width: 100%; }

/* buttons */
.stButton > button, .stDownloadButton > button, .stFormSubmitButton > button { width: 100%; border-radius: 8px;
  font-weight: 600; transition: all .15s ease; }
.stApp [data-testid="stBaseButton-primary"], .stApp [data-testid="stBaseButton-primaryFormSubmit"] { background: #1E3A8A; border: 1px solid #1E3A8A; color: #fff; }
.stApp [data-testid="stBaseButton-primary"]:hover, .stApp [data-testid="stBaseButton-primaryFormSubmit"]:hover {
  background: #2563EB; border-color: #2563EB; box-shadow: 0 6px 14px rgba(37,99,235,.28); transform: translateY(-1px); }
.stApp [data-testid="stBaseButton-secondary"], .stApp [data-testid="stBaseButton-secondaryFormSubmit"] { background: #fff; border: 1px solid #C7D2EA; color: #1E3A8A; }
.stApp [data-testid="stBaseButton-secondary"]:hover, .stApp [data-testid="stBaseButton-secondaryFormSubmit"]:hover {
  background: #EEF3FF; border-color: #1E3A8A; color: #1E3A8A; }

/* inputs + tables */
div[data-baseweb="input"], div[data-baseweb="select"] > div, div[data-baseweb="textarea"] { border-radius: 8px !important; transition: box-shadow .15s ease; }
div[data-baseweb="input"]:hover, div[data-baseweb="select"] > div:hover { box-shadow: 0 0 0 2px #DBE4F5; }
div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] { border: 1px solid #E2E8F0; border-radius: 10px; overflow: hidden; background: #fff; }
div[data-testid="stExpander"] { background: #fff; border: 1px solid #E2E8F0; border-radius: 10px; }
button[data-baseweb="tab"] { font-weight: 600; }

/* new-transaction pieces */
.cols-head { display: flex; font-size: 12px; font-weight: 700; color: #64748B; text-transform: uppercase; letter-spacing: .6px; }
.amt { background: #EEF3FF; border-radius: 8px; padding: 8px 12px; text-align: right; font-weight: 700; color: #0F1F4D; min-height: 40px; }
.sum-card { background: #fff; border: 1px solid #DBE4F5; border-radius: 14px; padding: 18px 20px; box-shadow: 0 4px 16px rgba(30,58,138,.08); }
.sum-row { display: flex; justify-content: space-between; padding: 6px 0; color: #475569; font-size: 14px; }
.sum-row b { color: #0F1F4D; }
.sum-total { display: flex; justify-content: space-between; padding: 12px 0; margin-top: 4px; border-top: 2px solid #1E3A8A;
  font-size: 18px; font-weight: 800; color: #0F1F4D; }
.sum-bal { display: flex; justify-content: space-between; padding: 10px 14px; margin-top: 8px; border-radius: 10px;
  background: #1E3A8A; color: #fff; font-size: 16px; font-weight: 700; }

.badge { display: inline-block; padding: 4px 12px; border-radius: 999px; font-size: 12px; font-weight: 700; letter-spacing: .3px; }
.b-ok { background: #DCFCE7; color: #166534; } .b-bad { background: #FEE2E2; color: #991B1B; }
.b-warn { background: #FEF9C3; color: #854D0E; } .b-orange { background: #FFEDD5; color: #9A3412; }
.b-info { background: #E0E7FF; color: #3730A3; }
hr { border-color: #E2E8F0; }
</style>
""",
    unsafe_allow_html=True,
)


def header(title: str, sub: str) -> None:
    st.markdown(f'<div class="page-title">{title}</div><div class="page-sub">{sub}</div>', unsafe_allow_html=True)
    if S.get("flash"):
        st.success(S.pop("flash"))


def section(title: str) -> None:
    st.markdown(f'<div class="section">{title}</div>', unsafe_allow_html=True)


def show(df: pd.DataFrame, cols: list[str] | None = None, height: int | None = None) -> None:
    d = df[cols] if cols else df
    cfg = {}
    for c in d.columns:
        if c in MONEY:
            cfg[c] = st.column_config.NumberColumn(c, format="localized")
        elif c in DATE_COLS:
            cfg[c] = st.column_config.DateColumn(c, format="DD MMM YYYY")
    kwargs = dict(_STRETCH)
    if height is not None:
        kwargs["height"] = height
    st.dataframe(d, hide_index=True, column_config=cfg, **kwargs)


def export_row(title: str, df: pd.DataFrame, fname: str, key: str) -> None:
    """Excel + PDF download buttons side by side."""
    safe_fname = re.sub(r'[^A-Za-z0-9_-]+', '_', str(fname or 'report')).strip('_') or 'report'
    c1, c2, _ = st.columns([1.2, 1.2, 3])
    c1.download_button("Download Excel", to_excel_bytes({title[:31]: df}), f"{safe_fname}.xlsx", XLSX, key=f"{key}_x")
    c2.download_button("Download PDF", to_pdf_bytes(title, df), f"{safe_fname}.pdf", "application/pdf", key=f"{key}_p")


def sidebar() -> None:
    sb = st.sidebar
    sb.markdown('<div class="brand">Business Manager</div><div class="tag">Enter once. Find in seconds.</div>',
                unsafe_allow_html=True)
    for grp, pages in GROUPS:
        sb.markdown(f'<div class="grp">{grp}</div>', unsafe_allow_html=True)
        for p in pages:
            if sb.button(p, key=f"nav_{p}", type="primary" if S.page == p else "secondary"):
                S.page = p
                st.rerun()
    sb.markdown(f'<div class="grp">Summary</div><div class="foot">{len(S.tx)} bills &nbsp;|&nbsp; '
                f'{len(all_parties())} parties &nbsp;|&nbsp; {len(S.products)} products<br/>'
                f'Storage: {"Cloud database" if db.is_cloud() else "Local test database"}</div>', unsafe_allow_html=True)
    if S.get("auth") and sb.button("Log out", key="logout"):
        S.auth = False
        st.rerun()


# =====================================================================
# PAGES
# =====================================================================
def page_dashboard() -> None:
    df = S.tx
    header("Dashboard", "A live picture of your business. Everything here calculates automatically.")
    c = st.columns(4)
    c[0].metric("Total Billed", money(df["Bill Amount"].sum()))
    c[1].metric("Total Received", money(df["Paid Amount"].sum()))
    c[2].metric("Outstanding", money(df["Balance"].sum()))
    c[3].metric("Returns", money(df["Return Amount"].sum()))
    c = st.columns(4)
    c[0].metric("Unpaid Bills", int((df["Status"] == "UNPAID").sum()))
    c[1].metric("Partially Paid", int((df["Status"] == "PARTIALLY PAID").sum()))
    c[2].metric("Fully Paid", int((df["Status"] == "PAID").sum()))
    c[3].metric("Cheques Pending", int((df["Status"] == "CHEQUE PENDING").sum()))
    if df.empty:
        st.info("No records yet. Use New Transaction to add a bill, or Import Data to load your existing records.")
        return
    a, b = st.columns(2)
    with a:
        section("Top 5 Outstanding Parties")
        top = df.groupby("Party / Company")["Balance"].sum().sort_values(ascending=False).head(5)
        try:
            st.bar_chart(top, color="#1E3A8A", horizontal=True)
        except TypeError:
            st.bar_chart(top, color="#1E3A8A")
    with b:
        section("Monthly Billed vs Received")
        m = df.dropna(subset=["Date"]).assign(Month=lambda x: x["Date"].dt.strftime("%Y-%m"))
        m = m.groupby("Month")[["Bill Amount", "Paid Amount"]].sum()
        st.bar_chart(m, color=["#1E3A8A", "#60A5FA"])
    section("Recent Transactions")
    show(df.sort_values("Date", ascending=False).head(10),
         ["Txn ID", "Date", "Bill No", "Party / Company", "Bill Amount", "Paid Amount", "Balance", "Status"])


def _add_line() -> None:
    S.ntx_n += 1


def _drop_line() -> None:
    S.ntx_n = max(1, S.ntx_n - 1)


def page_new_transaction() -> None:
    header("New Transaction", "Build the bill line by line. Every total updates instantly as you type.")
    fid, prods = S.ntx_fid, S.products
    names = prods["Product"].astype(str).tolist()
    summary = None
    left, right = st.columns([3, 2], gap="large")

    with left:
        section("Bill Details")
        c1, c2, c3 = st.columns([2, 1, 1])
        with c1:
            party = party_picker("Party / Company", f"ntx_party_{fid}")
        bill_date = c2.date_input("Date", value=date.today(), key=f"ntx_date_{fid}")
        bill_no = c3.text_input("Bill No", value=next_bill_no(), key=f"ntx_bill_{fid}")

        section("Items")
        if not names:
            st.warning("Your product list is empty. Add products below or open the Products page.")
        st.markdown('<div class="cols-head"><div style="flex:4">Product</div><div style="flex:1.4">Qty</div>'
                    '<div style="flex:1.6">Rate</div><div style="flex:1.6;text-align:right">Amount</div></div>',
                    unsafe_allow_html=True)
        lines, subtotal = [], 0.0
        for i in range(S.ntx_n):
            c = st.columns([4, 1.4, 1.6, 1.6])
            prod = c[0].selectbox("Product", [NO_PRODUCT] + names, key=f"p_{fid}_{i}", label_visibility="collapsed")
            info = prods[prods["Product"].astype(str) == prod]
            unit = str(info["Unit"].iloc[0]) if not info.empty else ""
            d_rate = float(info["Rate"].iloc[0]) if not info.empty and pd.notna(info["Rate"].iloc[0]) else 0.0
            qty = c[1].number_input(f"Qty {unit}", min_value=0.0, value=0.0, step=1.0, key=f"q_{fid}_{i}",
                                    label_visibility="collapsed", help=f"Quantity ({unit})" if unit else None)
            rate = c[2].number_input("Rate", min_value=0.0, value=d_rate, step=1.0, key=f"r_{fid}_{i}_{prod}",
                                     label_visibility="collapsed")
            amt = round(qty * rate, 2)
            c[3].markdown(f'<div class="amt">{amt:,.2f}</div>', unsafe_allow_html=True)
            if prod != NO_PRODUCT and qty > 0:
                lines.append((prod, qty, rate, amt))
                subtotal += amt
        b1, b2, _ = st.columns([1, 1, 3])
        b1.button("Add Item", on_click=_add_line, key=f"add_{fid}")
        b2.button("Remove Last", on_click=_drop_line, key=f"rem_{fid}")

        with st.expander("Quick add a new product"):
            with st.form(f"qp_{fid}", clear_on_submit=True):
                q1, q2, q3 = st.columns([3, 1, 1])
                pn = q1.text_input("Product name")
                pu = q2.text_input("Unit", value="Pcs")
                pr = q3.number_input("Default rate", min_value=0.0, step=1.0)
                if st.form_submit_button("Add Product") and pn.strip():
                    S.products = pd.concat([S.products, pd.DataFrame([[pn.strip(), pu.strip() or "Pcs", pr]], columns=PRODUCT_COLS)]) \
                        .drop_duplicates("Product", keep="last").reset_index(drop=True)
                    save_all("products")
                    st.rerun()
        remarks = st.text_area("Remarks / Notes", key=f"ntx_rem_{fid}", height=80)

    with right:
        summary = st.container()
        section("Discount, Tax and Payment")
        d1, d2 = st.columns([1, 1])
        dtype = d1.radio("Discount type", ["Percent (%)", "Amount"], horizontal=True, key=f"dt_{fid}")
        dval = d2.number_input("Discount", min_value=0.0, value=0.0, key=f"dv_{fid}")
        discount = round(subtotal * dval / 100 if dtype.startswith("Percent") else dval, 2)
        discount = min(discount, subtotal)
        tax_pct = st.number_input("Tax / GST (%)", min_value=0.0, max_value=100.0, value=0.0, key=f"tx_{fid}")
        tax = round((subtotal - discount) * tax_pct / 100, 2)
        total = round(subtotal - discount + tax, 2)
        paid = st.number_input("Amount Paid Now", min_value=0.0, value=0.0, key=f"pd_{fid}")
        method = st.selectbox("Payment Method", PAYMENT_METHODS, key=f"pm_{fid}")
        chq_no = chq_bank = chq_status = sale_base = ""
        chq_date = None
        if method == "Cheque":
            k1, k2 = st.columns(2)
            chq_no = k1.text_input("Cheque No", key=f"cn_{fid}")
            chq_bank = k2.text_input("Bank Name", key=f"cb_{fid}")
            k3, k4 = st.columns(2)
            chq_date = k3.date_input("Cheque Date", value=None, key=f"cd_{fid}")
            chq_status = k4.selectbox("Cheque Status", CHEQUE_STATUSES[1:], key=f"cs_{fid}")
        sale_base = st.text_input("Sale Base (optional)", key=f"sb_{fid}", placeholder="e.g. 5/8 or S/B")
        balance = round(total - paid, 2)

        tax_lbl = f"Tax ({tax_pct:g}%)" if tax_pct else "Tax"
        disc_lbl = f"Discount ({dval:g}%)" if dtype.startswith("Percent") and dval else "Discount"
        with summary:
            section("Bill Summary")
            st.markdown(
                f'<div class="sum-card">'
                f'<div class="sum-row"><span>Subtotal ({len(lines)} item{"s" if len(lines) != 1 else ""})</span><b>{CUR} {subtotal:,.2f}</b></div>'
                f'<div class="sum-row"><span>{disc_lbl}</span><b>- {CUR} {discount:,.2f}</b></div>'
                f'<div class="sum-row"><span>{tax_lbl}</span><b>+ {CUR} {tax:,.2f}</b></div>'
                f'<div class="sum-total"><span>Grand Total</span><span>{CUR} {total:,.2f}</span></div>'
                f'<div class="sum-row"><span>Paid Now</span><b>{CUR} {paid:,.2f}</b></div>'
                f'<div class="sum-bal"><span>Balance Due</span><span>{CUR} {balance:,.2f}</span></div></div>',
                unsafe_allow_html=True)

        st.write("")
        if st.button("Save Transaction", type="primary", key=f"save_{fid}"):
            dup = S.tx[(S.tx["Party / Company"] == party) & (S.tx["Bill No"] == bill_no)]
            if not party:
                st.error("Select a party, or choose '+ Add a new party...' and type a name.")
            elif not lines:
                st.error("Add at least one item with a product and a quantity greater than zero.")
            elif paid > total:
                st.error("Amount paid cannot be more than the grand total.")
            elif not bill_no.strip():
                st.error("Bill No cannot be empty.")
            elif not dup.empty:
                st.error(f"Bill No {bill_no} already exists for {party}. Please use a different number.")
            else:
                tid = next_txn_ids(1)[0]
                add_party_if_new(party)
                row = {"Txn ID": tid, "Date": pd.to_datetime(bill_date), "Bill No": bill_no.strip(),
                       "Party / Company": party, "Product / Item": ", ".join(f"{p} x{q:g}" for p, q, _, _ in lines),
                       "Qty": sum(q for _, q, _, _ in lines), "Subtotal": subtotal, "Discount": discount, "Tax": tax,
                       "Bill Amount": total, "Paid Amount": paid, "Return Amount": 0.0, "Payment Method": method,
                       "Cheque No": chq_no, "Cheque Date": pd.to_datetime(chq_date) if chq_date else pd.NaT,
                       "Bank Name": chq_bank, "Cheque Status": chq_status, "Sale Base": sale_base, "Remarks": remarks}
                S.tx = recompute(pd.concat([S.tx, pd.DataFrame([row])], ignore_index=True))
                S.line_items = pd.concat([S.line_items, pd.DataFrame(
                    [(tid, p, q, r, a) for p, q, r, a in lines], columns=ITEM_COLS)], ignore_index=True)
                save_all("transactions", "items", "parties")
                S.flash = f"Saved {bill_no} for {party}: total {CUR} {total:,.2f}, balance {CUR} {balance:,.2f}."
                S.ntx_fid += 1
                S.ntx_n = 1
                st.rerun()


def page_bills() -> None:
    header("Bills & Payments", "Open a bill to receive a payment, record a return, review items or print the invoice.")
    df = S.tx
    if df.empty:
        st.info("No bills yet. Create one from New Transaction.")
        return
    q = st.text_input("Find a bill", placeholder="Type a party name or bill number to narrow the list")
    view = df
    if q:
        view = df[df["Party / Company"].str.contains(q, case=False, regex=False) | df["Bill No"].str.contains(q, case=False, regex=False)]
    if view.empty:
        st.warning("No bills match your search.")
        return
    lab = {r["Txn ID"]: f"{r['Bill No']}  |  {r['Party / Company']}  |  Total {r['Bill Amount']:,.0f}  |  Balance {r['Balance']:,.0f}"
           for _, r in view.iterrows()}
    ids = view.sort_values("Date", ascending=False)["Txn ID"].tolist()
    tid = st.selectbox("Select bill", ids, format_func=lambda t: lab[t])
    row = df[df["Txn ID"] == tid].iloc[0]

    st.markdown(badge(row["Status"]), unsafe_allow_html=True)
    c = st.columns(5)
    c[0].metric("Bill Total", money(row["Bill Amount"]))
    c[1].metric("Paid", money(row["Paid Amount"]))
    c[2].metric("Returned", money(row["Return Amount"]))
    c[3].metric("Balance", money(row["Balance"]))
    c[4].metric("Discount / Tax", f"{row['Discount']:,.0f} / {row['Tax']:,.0f}")

    a, b = st.columns(2, gap="large")
    with a:
        section("Receive Payment")
        with st.form(f"pay_{tid}", clear_on_submit=True):
            amt = st.number_input("Amount received", min_value=0.0, value=0.0)
            pm = st.selectbox("Method", PAYMENT_METHODS[1:])
            note = st.text_input("Note (optional)")
            if st.form_submit_button("Record Payment", type="primary"):
                if amt <= 0:
                    st.error("Enter an amount greater than zero.")
                elif amt > max(row["Balance"], 0):
                    st.error(f"Amount exceeds the balance of {money(row['Balance'])}.")
                else:
                    ch = {"Paid Amount": row["Paid Amount"] + amt}
                    if row["Payment Method"] == "Not specified":
                        ch["Payment Method"] = pm
                    update_bill(tid, f"Payment {CUR} {amt:,.0f} ({pm}) on {date.today():%d %b %Y}" + (f" - {note}" if note else ""), **ch)
                    S.flash = f"Payment of {money(amt)} recorded."
                    st.rerun()
    with b:
        section("Record Return")
        limit = max(row["Bill Amount"] - row["Return Amount"], 0.0)
        with st.form(f"ret_{tid}", clear_on_submit=True):
            ramt = st.number_input(f"Return amount (max {limit:,.0f})", min_value=0.0, value=0.0)
            why = st.text_input("Reason (optional)")
            if st.form_submit_button("Record Return", type="primary"):
                if ramt <= 0:
                    st.error("Enter an amount greater than zero.")
                elif ramt > limit:
                    st.error("Return cannot exceed the bill amount.")
                else:
                    update_bill(tid, f"Return {CUR} {ramt:,.0f} on {date.today():%d %b %Y}" + (f" - {why}" if why else ""),
                                **{"Return Amount": row["Return Amount"] + ramt})
                    S.flash = f"Return of {money(ramt)} recorded."
                    st.rerun()

    section("Bill Items")
    it = S.line_items[S.line_items["Txn ID"] == tid]
    if it.empty:
        st.caption(f"No item breakdown stored for this bill. Summary: {row['Product / Item'] or 'not specified'}")
    else:
        show(it, ["Product", "Qty", "Rate", "Amount"])
    st.download_button("Download Invoice (PDF)", invoice_pdf(row, it), f"Invoice_{row['Bill No']}.pdf",
                       "application/pdf", key=f"inv_{tid}")
    if row["Remarks"]:
        st.caption(f"History: {row['Remarks']}")


def page_ledger() -> None:
    header("Party Ledger", "Pick a party to see their complete account.")
    sel = party_picker("Party / Company", "ledger", allow_all=True)
    d = S.tx if sel == "ALL PARTIES" else S.tx[S.tx["Party / Company"] == sel]
    c = st.columns(4)
    c[0].metric("Bills", len(d))
    c[1].metric("Total Billed", money(d["Bill Amount"].sum()))
    c[2].metric("Total Paid", money(d["Paid Amount"].sum()))
    c[3].metric("Outstanding", money(d["Balance"].sum()))
    section("Transaction History")
    if d.empty:
        st.info("No transactions for this party yet.")
        return
    cols = ["Txn ID", "Date", "Bill No", "Product / Item", "Bill Amount", "Paid Amount", "Return Amount",
            "Balance", "Payment Method", "Status"]
    if sel == "ALL PARTIES":
        cols.insert(3, "Party / Company")
    d = d.sort_values("Date", ascending=False)[cols]
    show(d)
    safe_sel = re.sub(r'[^A-Za-z0-9_-]+', '_', str(sel or 'all')).strip('_') or 'all'
    export_row(f"Ledger - {sel}", d, f"Ledger_{safe_sel}", "led")


def page_parties() -> None:
    header("Parties Directory", "Manage your customer & supplier accounts, view total business and outstanding balances.")
    parties = all_parties()
    df = S.tx
    
    records = []
    for p in parties:
        sub = df[df["Party / Company"] == p]
        tot_billed = sub["Bill Amount"].sum() if not sub.empty else 0.0
        tot_paid = sub["Paid Amount"].sum() if not sub.empty else 0.0
        tot_returns = sub["Return Amount"].sum() if not sub.empty else 0.0
        balance = tot_billed - tot_paid - tot_returns
        bills_count = len(sub)
        last_date = sub["Date"].max() if not sub.empty and pd.notna(sub["Date"].max()) else pd.NaT
        records.append({
            "Party / Company": p,
            "Total Invoiced": tot_billed,
            "Total Paid": tot_paid,
            "Current Balance": balance,
            "Total Bills": bills_count,
            "Last Transaction": last_date
        })
    pdf = pd.DataFrame(records)
    if not pdf.empty:
        pdf = pdf.sort_values("Current Balance", ascending=False).reset_index(drop=True)
    
    c = st.columns(4)
    c[0].metric("Total Parties", len(parties))
    c[1].metric("Parties with Balance", int((pdf["Current Balance"] > 0).sum()) if not pdf.empty else 0)
    c[2].metric("Total Receivables", money(pdf["Current Balance"].sum()) if not pdf.empty else money(0))
    c[3].metric("Total Business Done", money(pdf["Total Invoiced"].sum()) if not pdf.empty else money(0))
    
    section("Register New Party")
    with st.expander("+ Add New Customer / Supplier Party", expanded=False):
        with st.form("new_party_form", clear_on_submit=True):
            np_name = st.text_input("Party / Company Name", placeholder="e.g. TARIQ TRADERS").strip().upper()
            if st.form_submit_button("Add Party", type="primary") and np_name:
                if np_name in parties:
                    st.warning(f"Party '{np_name}' already exists.")
                else:
                    add_party_if_new(np_name)
                    save_all("parties")
                    S.flash = f"Added party: {np_name}"
                    st.rerun()

    section("All Parties Directory")
    if pdf.empty:
        st.info("No parties registered yet.")
    else:
        q = st.text_input("Search party in directory", placeholder="Type party name to filter...", key="pdir_q")
        view_pdf = pdf
        if q:
            view_pdf = view_pdf[view_pdf["Party / Company"].str.contains(q, case=False, regex=False)]
        
        cfg = {
            "Total Invoiced": st.column_config.NumberColumn("Total Invoiced", format="localized"),
            "Total Paid": st.column_config.NumberColumn("Total Paid", format="localized"),
            "Current Balance": st.column_config.NumberColumn("Current Balance", format="localized"),
            "Last Transaction": st.column_config.DateColumn("Last Transaction", format="DD MMM YYYY"),
        }
        st.dataframe(view_pdf, hide_index=True, column_config=cfg, **_STRETCH)
        export_row("Parties Directory", view_pdf, "Parties_Directory", "pdir")


def page_search() -> None:
    header("Search Center", "Combine any filters. Leave a box empty or on All to ignore it.")
    df = S.tx
    c = st.columns(4)
    party = c[0].text_input("Party contains")
    bill = c[1].text_input("Bill No contains")
    prod = c[2].text_input("Product contains")
    status = c[3].selectbox("Status", ["All"] + STATUS_OPTIONS)
    c = st.columns(4)
    d1 = c[0].date_input("Date from", value=None, key="sc_f")
    d2 = c[1].date_input("Date to", value=None, key="sc_t")
    pay = c[2].selectbox("Payment method", ["All"] + PAYMENT_METHODS)
    chq = c[3].selectbox("Cheque status", ["All"] + CHEQUE_STATUSES[1:])
    r = df
    if party: r = r[r["Party / Company"].str.contains(party, case=False, regex=False)]
    if bill: r = r[r["Bill No"].str.contains(bill, case=False, regex=False)]
    if prod: r = r[r["Product / Item"].str.contains(prod, case=False, regex=False)]
    if d1: r = r[r["Date"] >= pd.to_datetime(d1)]
    if d2: r = r[r["Date"] <= pd.to_datetime(d2)]
    if status != "All": r = r[r["Status"] == status]
    if pay != "All": r = r[r["Payment Method"] == pay]
    if chq != "All": r = r[r["Cheque Status"] == chq]
    st.markdown(f"**{len(r)}** matching record(s) &nbsp;|&nbsp; Outstanding in results: **{money(r['Balance'].sum())}**")
    if not r.empty:
        cols = ["Txn ID", "Date", "Bill No", "Party / Company", "Product / Item", "Bill Amount", "Paid Amount",
                "Return Amount", "Balance", "Payment Method", "Status", "Remarks"]
        r = r.sort_values("Date", ascending=False)[cols]
        show(r)
        export_row("Search Results", r, "Search_Results", "srch")


def page_cheques() -> None:
    header("Cheque Tracker", "Every cheque payment in one place: pending, cleared and coming due.")
    df = S.tx
    ch = df[df["Payment Method"] == "Cheque"].copy()
    today = pd.Timestamp(date.today())
    pend = ch[ch["Cheque Status"] == "Pending"]
    soon = pend[pend["Cheque Date"].between(today, today + timedelta(days=7))]
    c = st.columns(4)
    c[0].metric("Pending Cheques", len(pend))
    c[1].metric("Pending Amount", money(pend["Paid Amount"].sum()))
    c[2].metric("Cleared Amount", money(ch.loc[ch["Cheque Status"] == "Cleared", "Paid Amount"].sum()))
    c[3].metric("Due in 7 Days", len(soon))
    flt = st.selectbox("Show", ["All", "Pending", "Cleared", "Bounced"])
    if flt != "All":
        ch = ch[ch["Cheque Status"] == flt]

    def due(r):
        if r["Cheque Status"] != "Pending" or pd.isna(r["Cheque Date"]):
            return ""
        n = (r["Cheque Date"] - today).days
        return "OVERDUE" if n < 0 else "DUE TODAY" if n == 0 else f"in {n} day(s)"

    section("Cheque Transactions")
    if ch.empty:
        st.info("No cheques match this filter.")
    else:
        ch["Due"] = ch.apply(due, axis=1)
        ch = ch.sort_values("Cheque Date", na_position="last")[
            ["Txn ID", "Party / Company", "Bill No", "Cheque No", "Cheque Date", "Bank Name", "Paid Amount",
             "Balance", "Cheque Status", "Due"]]
        show(ch)
        export_row("Cheque Tracker", ch, "Cheque_Tracker", "chq")
    if not pend.empty:
        section("Update Cheque Status")
        u = st.columns([4, 1.5, 1], vertical_alignment="bottom")
        lab = {r["Txn ID"]: f"{r['Party / Company']}  |  {r['Bill No']}  |  Cheque {r['Cheque No'] or '-'}  |  {r['Paid Amount']:,.0f}"
               for _, r in pend.iterrows()}
        tid = u[0].selectbox("Pending cheque", list(lab), format_func=lambda t: lab[t])
        new = u[1].selectbox("New status", ["Cleared", "Bounced"])
        if u[2].button("Update", type="primary"):
            update_bill(tid, f"Cheque {new.lower()} on {date.today():%d %b %Y}", **{"Cheque Status": new})
            S.flash = f"Cheque marked {new}."
            st.rerun()
        st.caption("If a cheque bounces, correct the Paid Amount in Master Database so the balance reopens.")


def page_products() -> None:
    header("Products", "Your product list feeds the dropdown in New Transaction, so bills take seconds to enter.")
    st.caption("Add a row at the bottom of the table to create a product. To delete, tick the box at the left of a row "
               "and press the Delete key or the trash icon. Click Save Products when finished.")
    ed = st.data_editor(
        S.products, num_rows="dynamic", hide_index=True, key=f"prod_ed_{S.ver}", **_STRETCH,
        column_config={
            "Product": st.column_config.TextColumn("Product", required=True),
            "Unit": st.column_config.TextColumn("Unit", default="Pcs"),
            "Rate": st.column_config.NumberColumn("Default Rate", min_value=0, default=0, format="localized"),
        })
    c1, _ = st.columns([1, 3])
    if c1.button("Save Products", type="primary"):
        d = ed.copy()
        d["Product"] = _txt(d["Product"])
        d["Unit"] = _txt(d["Unit"]).replace("", "Pcs")
        d["Rate"] = pd.to_numeric(d["Rate"], errors="coerce").fillna(0.0)
        d = d[d["Product"] != ""].drop_duplicates("Product", keep="last").reset_index(drop=True)
        S.products = d[PRODUCT_COLS]
        save_all("products")
        S.ver += 1
        S.flash = f"Product list saved ({len(d)} products)."
        st.rerun()
    if not S.products.empty:
        export_row("Product List", S.products, "Products", "prd")


def page_master() -> None:
    header("Master Database", "Edit any field directly. Balance and Status recalculate when you save.")
    q = st.text_input("Quick filter", placeholder="Type any text to filter rows (party, bill no, product, status...)")
    view = S.tx.copy()
    if q:
        view = view[view.astype(str).apply(lambda c: c.str.contains(q, case=False, na=False, regex=False)).any(axis=1)]
    view.insert(0, "Select", False)
    pm = sorted(set(PAYMENT_METHODS) | set(v for v in S.tx["Payment Method"] if v))
    cs = sorted(set(CHEQUE_STATUSES) | set(v for v in S.tx["Cheque Status"] if v))
    cfg = {c: st.column_config.NumberColumn(c, format="localized") for c in MONEY if c in view.columns}
    cfg.update({
        "Select": st.column_config.CheckboxColumn("Delete?", width="small"),
        "Date": st.column_config.DateColumn("Date", format="DD MMM YYYY"),
        "Cheque Date": st.column_config.DateColumn("Cheque Date", format="DD MMM YYYY"),
        "Payment Method": st.column_config.SelectboxColumn(options=pm),
        "Cheque Status": st.column_config.SelectboxColumn(options=cs),
    })
    ed = st.data_editor(view, hide_index=True, num_rows="fixed", key=f"mdb_{S.ver}_{q}", column_config=cfg,
                        disabled=["Txn ID", "Balance", "Status"], height=520, **_STRETCH)
    ids = ed.loc[ed["Select"], "Txn ID"].tolist()
    c1, c2, c3 = st.columns([1, 1, 2], vertical_alignment="bottom")
    if c1.button("Save Changes", type="primary"):
        upd = ed.drop(columns=["Select", "Balance", "Status"]).set_index("Txn ID")
        base = S.tx.set_index("Txn ID")
        base.loc[upd.index, upd.columns] = upd
        S.tx = recompute(base.reset_index())
        save_all("transactions")
        S.ver += 1
        S.flash = "Changes saved."
        st.rerun()
    ok = c3.checkbox(f"Confirm permanent deletion of {len(ids)} record(s)", disabled=not ids)
    if c2.button("Delete Selected", disabled=not (ids and ok)):
        S.tx = S.tx[~S.tx["Txn ID"].isin(ids)].reset_index(drop=True)
        S.line_items = S.line_items[~S.line_items["Txn ID"].isin(ids)].reset_index(drop=True)
        save_all("transactions", "items")
        S.ver += 1
        S.flash = f"Deleted {len(ids)} record(s)."
        st.rerun()
    st.caption("Tip: save your edits before deleting. Deleted records cannot be recovered unless you kept a backup.")
    export_row("Master Database", S.tx, "Master_Database", "mdb")


def page_import() -> None:
    header("Import Data", "Bring in your existing records from Excel, CSV or PDF. Columns are recognised automatically.")
    t1, t2, t3 = st.tabs(["Import Records", "Import Products", "Backup & Restore"])

    with t1:
        st.markdown("Upload a file with a header row. Recognised columns include **Party, Date, Bill No, Product, Qty, "
                    "Rate, Amount, Paid, Return, Payment Method, Cheque No, Cheque Date, Bank, Cheque Status, "
                    "Sale Base, Remarks**. Only **Party** is required.")
        tpl = _empty(["Date", "Bill No", "Party / Company", "Product / Item", "Qty", "Rate", "Bill Amount",
                      "Paid Amount", "Return Amount", "Payment Method", "Cheque No", "Cheque Date", "Bank Name",
                      "Cheque Status", "Sale Base", "Remarks"])
        st.download_button("Download blank Excel template", to_excel_bytes({"Transactions": tpl}),
                           "Import_Template.xlsx", XLSX, key="tpl")
        f = st.file_uploader("Choose file", type=["xlsx", "xls", "csv", "pdf"], key=f"imp_{S.ver}")
        if f:
            try:
                d, used, ignored = prepare_import(read_upload(f))
            except Exception as exc:
                st.error(str(exc))
                return
            new, dups = split_duplicates(d, S.tx)
            st.success(f"Recognised columns: {', '.join(used)}")
            if ignored:
                st.caption(f"Ignored columns: {', '.join(ignored)}")
            c = st.columns(3)
            c[0].metric("Rows in file", len(d))
            c[1].metric("New records", len(new))
            c[2].metric("Duplicates skipped", dups)
            show(new.head(50), [x for x in COLUMNS if x in new.columns and x not in ("Balance", "Status", "Txn ID")], height=300)
            addp = st.checkbox("Add new products found in this file to the Products list", value=True)
            if st.button(f"Import {len(new)} record(s)", type="primary", disabled=new.empty):
                S.tx, S.line_items, S.products, n = commit_import(new, addp, S.tx, S.line_items, S.products)
                save_all()
                S.ver += 1
                S.flash = f"Imported {n} record(s) successfully."
                st.rerun()

    with t2:
        st.markdown("Upload a list with **Product** (or Item), optional **Unit** and **Rate**. Existing products are updated.")
        f = st.file_uploader("Choose file", type=["xlsx", "xls", "csv", "pdf"], key=f"impp_{S.ver}")
        if f and st.button("Import Products", type="primary"):
            try:
                S.products, n = import_products(read_upload(f), S.products)
                save_all("products")
                S.ver += 1
                S.flash = f"{n} product(s) imported."
                st.rerun()
            except Exception as exc:
                st.error(str(exc))

    with t3:
        st.info("Moving from the old local version? Choose your old **data_store.xlsx** under "
                "'Restore from backup' below - all bills, items, products and parties go into the cloud database.")
        st.markdown("Download a full backup regularly. It contains bills, items, products and parties in one Excel file.")
        st.download_button("Download Full Backup (Excel)", to_excel_bytes(workbook_sheets()),
                           f"Backup_{date.today():%Y%m%d}.xlsx", XLSX, key="bk")
        section("Restore from backup")
        st.warning("Restoring replaces ALL current data with the contents of the backup file.")
        f = st.file_uploader("Backup file (.xlsx)", type=["xlsx"], key=f"rest_{S.ver}")
        ok = st.checkbox("I understand that current data will be replaced")
        if f and ok and st.button("Restore Backup", type="primary"):
            try:
                sh = pd.read_excel(f, sheet_name=None)
                S.tx = lg.ensure_txn_ids(recompute(sh["Transactions"]))
                S.line_items = sh.get("Items", _empty(ITEM_COLS)).reindex(columns=ITEM_COLS)
                S.products = sh.get("Products", _empty(PRODUCT_COLS)).reindex(columns=PRODUCT_COLS)
                S.products["Product"] = _txt(S.products["Product"])
                S.products["Unit"] = _txt(S.products["Unit"]).replace("", "Pcs")
                S.products["Rate"] = pd.to_numeric(S.products["Rate"], errors="coerce").fillna(0.0)
                S.products = S.products[S.products["Product"] != ""].drop_duplicates("Product").reset_index(drop=True)
                S.extra_parties = (sh["Parties"]["Party"].dropna().astype(str).tolist()
                                   if "Parties" in sh and "Party" in sh["Parties"] else [])
                save_all()
                S.ver += 1
                S.flash = "Backup restored."
                st.rerun()
            except Exception as exc:
                st.error(f"This does not look like a valid backup file: {exc}")


def page_reports() -> None:
    header("Reports & Analytics", "Interactive financial reports, aging analysis, and statements. Download any report as Excel or PDF.")
    df = S.tx
    t_stmt, t_age, t_out, t_sales, t_book = st.tabs([
        "Party Statement", "Aged Receivables", "Outstanding Summary", "Sales & Cheques", "Full Business Workbook"
    ])

    with t_stmt:
        section("Customer / Party Statement of Account")
        c_top = st.columns([2, 1, 1])
        sel = party_picker("Party / Company", "rep", allow_all=True)
        d1 = c_top[1].date_input("From Date", value=None, key="rp_f")
        d2 = c_top[2].date_input("To Date", value=None, key="rp_t")
        
        st_df = lg.compute_party_statement(df, sel)
        if d1: st_df = st_df[st_df["Date"] >= pd.to_datetime(d1)]
        if d2: st_df = st_df[st_df["Date"] <= pd.to_datetime(d2)]
        
        c = st.columns(4)
        c[0].metric("Total Invoiced", money(st_df["Bill Amount"].sum()))
        c[1].metric("Total Received", money(st_df["Paid Amount"].sum()))
        c[2].metric("Total Returns", money(st_df["Return Amount"].sum()))
        net_bal = (st_df["Bill Amount"].sum() - st_df["Paid Amount"].sum() - st_df["Return Amount"].sum())
        c[3].metric("Net Balance Due", money(net_bal))
        
        if st_df.empty:
            st.info("No records found for the selected party and date range.")
        else:
            show(st_df)
            safe_sel = re.sub(r'[^A-Za-z0-9_-]+', '_', str(sel or 'all')).strip('_') or 'all'
            export_row(f"Statement - {sel}", st_df, f"Statement_{safe_sel}", "stmt")

    with t_age:
        section("Aged Receivables (Aging Report)")
        st.caption("Breakdown of unpaid balances by age. Identifies overdue debts to prioritize collections.")
        aging_df = lg.compute_aging(df)
        if aging_df.empty:
            st.info("No outstanding balances at this time! All bills are fully paid.")
        else:
            ca = st.columns(4)
            ca[0].metric("Current (0-30 days)", money(aging_df["Current (0-30d)"].sum()))
            ca[1].metric("31-60 Days", money(aging_df["31-60 Days"].sum()))
            ca[2].metric("61-90 Days", money(aging_df["61-90 Days"].sum()))
            ca[3].metric("Over 90 Days", money(aging_df["90+ Days"].sum()))
            show(aging_df)
            export_row("Aged Receivables", aging_df, "Aged_Receivables_Report", "age")

    with t_out:
        section("Outstanding Balance Summary")
        out = df[df["Balance"] > 0].sort_values("Balance", ascending=False)
        if out.empty:
            st.success("No outstanding balances! All invoices are settled.")
        else:
            by_party = out.groupby("Party / Company")[["Bill Amount", "Paid Amount", "Balance"]].sum() \
                .sort_values("Balance", ascending=False).reset_index()
            c_out = st.columns(3)
            c_out[0].metric("Parties with Balance", len(by_party))
            c_out[1].metric("Total Unpaid Balance", money(by_party["Balance"].sum()))
            c_out[2].metric("Highest Single Balance", money(by_party["Balance"].max()))
            show(by_party)
            export_row("Outstanding by Party", by_party, "Outstanding_By_Party", "outp")

    with t_sales:
        section("Sales & Revenue Breakdown")
        if df.empty:
            st.info("No sales records available yet.")
        else:
            c1, c2 = st.columns(2)
            with c1:
                st.subheader("Monthly Billed vs Received")
                m = df.dropna(subset=["Date"]).assign(Month=lambda x: x["Date"].dt.strftime("%Y-%m"))
                if not m.empty:
                    m_grp = m.groupby("Month")[["Bill Amount", "Paid Amount"]].sum()
                    st.bar_chart(m_grp, color=["#1E3A8A", "#60A5FA"])
            with c2:
                st.subheader("Payment Methods Share")
                pm_grp = df.groupby("Payment Method")["Paid Amount"].sum()
                if not pm_grp.empty and pm_grp.sum() > 0:
                    st.bar_chart(pm_grp, color="#2563EB")
                else:
                    st.caption("No payments recorded yet.")

    with t_book:
        section("Complete Business Workbook")
        st.caption("One master Excel file containing all business data sheets: Transactions, Outstanding, Cheques, Returns, Items Breakdown, and Aging Analysis.")
        out_full = df[df["Balance"] > 0].sort_values("Balance", ascending=False)
        aging_full = lg.compute_aging(df)
        st.download_button("Download Complete Master Workbook (Excel)", to_excel_bytes({
            "Transactions": df,
            "Outstanding": out_full,
            "Aged Receivables": aging_full,
            "Cheques": df[df["Payment Method"] == "Cheque"],
            "Returns": df[df["Return Amount"] > 0],
            "Items": S.line_items
        }), f"Business_Master_Report_{date.today():%Y%m%d}.xlsx", XLSX, key="full_wb")


# =====================================================================
# ROUTING
# =====================================================================
PAGES = {
    "Dashboard": page_dashboard, "New Transaction": page_new_transaction, "Bills & Payments": page_bills,
    "Cheque Tracker": page_cheques, "Party Ledger": page_ledger, "Parties Directory": page_parties,
    "Search Center": page_search, "Products": page_products, "Master Database": page_master,
    "Import Data": page_import, "Reports": page_reports,
}
require_login()
try:
    S.tx, S.line_items, S.products, S.extra_parties = db.load_all()
except Exception as exc:
    st.error("Could not connect to the database.")
    st.markdown("Most common reasons: wrong `DATABASE_URL` in Secrets, wrong password, or the Supabase project is paused "
                "(open supabase.com and click *Restore*). See README, section *Troubleshooting*.")
    st.caption(f"Technical detail: {type(exc).__name__}: {str(exc)[:300]}")
    st.stop()
sidebar()
PAGES[S.page]()

"""
Business Management System (v4 - Enterprise Executive Edition)
-------------------------------------------------------------
app.py   = Screens & UI interactions
logic.py = Calculations, Excel/PDF/WhatsApp/Thermal exports, Import engine
db.py    = Multi-user authentication, PostgreSQL / SQLite storage, Audit logs

Run locally: streamlit run app.py
"""

from __future__ import annotations

import inspect
import re
import time
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st

import db
import logic as lg
from logic import (
    BUSINESS_NAME, CHEQUE_STATUSES, COLUMNS, CUR, DATE_COLS, GROUPS, ITEM_COLS, MONEY,
    NEW_PARTY, NO_PRODUCT, PAYMENT_METHODS, PRODUCT_COLS, STATUS_OPTIONS, XLSX,
    _empty, _txt, commit_import, import_products, invoice_pdf, prepare_import, read_upload,
    recompute, split_duplicates, thermal_receipt_html, to_excel_bytes, to_pdf_bytes,
    whatsapp_invoice_text, whatsapp_url,
)

# Responsive stretch configuration
_STRETCH = ({"width": "stretch"} if "width" in inspect.signature(st.dataframe).parameters
            else {"use_container_width": True})

st.set_page_config(
    page_title=BUSINESS_NAME,
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded",
)

S = st.session_state
for _k, _v in {
    "page": "Dashboard",
    "ntx_fid": 0,
    "ntx_items": [],
    "ver": 0,
    "user": None,
    "tx": pd.DataFrame(columns=COLUMNS),
    "line_items": pd.DataFrame(columns=ITEM_COLS),
    "products": pd.DataFrame(columns=PRODUCT_COLS),
    "extra_parties": [],
}.items():
    S.setdefault(_k, _v)


# =====================================================================
# STYLING & DESIGN SYSTEM (TOP 1% EXECUTIVE FINTECH LOOK)
# =====================================================================
st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

:root {
  --primary-navy: #0F172A;
  --primary-blue: #1E3A8A;
  --accent-blue: #2563EB;
  --accent-cyan: #0284C7;
  --bg-light: #F8FAFC;
  --card-bg: #FFFFFF;
  --border-light: #E2E8F0;
  --text-main: #0F172A;
  --text-muted: #64748B;
  --success: #10B981;
  --warning: #F59E0B;
  --danger: #EF4444;
}

html, body, [class*="css"], .stApp {
  font-family: 'Inter', system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
  background-color: var(--bg-light);
  color: var(--text-main);
}

.stApp {
  background: radial-gradient(circle at 15% 10%, rgba(37, 99, 235, 0.03) 0%, transparent 40%),
              radial-gradient(circle at 85% 90%, rgba(14, 165, 233, 0.03) 0%, transparent 40%),
              #F8FAFC;
}

.block-container {
  padding-top: 1.5rem;
  padding-bottom: 3rem;
  max-width: 1540px;
}

header[data-testid="stHeader"] {
  background: rgba(248, 250, 252, 0.85);
  backdrop-filter: blur(8px);
}

/* Page Headers */
.page-title {
  font-size: 28px;
  font-weight: 800;
  color: #0F172A;
  letter-spacing: -0.6px;
  margin-bottom: 4px;
  display: flex;
  align-items: center;
  gap: 10px;
}
.page-sub {
  font-size: 14px;
  color: #64748B;
  margin-bottom: 22px;
  font-weight: 400;
}
.section-title {
  font-size: 13px;
  font-weight: 700;
  color: #1E3A8A;
  text-transform: uppercase;
  letter-spacing: 0.8px;
  margin: 24px 0 12px 0;
  padding-bottom: 6px;
  border-bottom: 2px solid #E2E8F0;
  display: flex;
  align-items: center;
  gap: 8px;
}

/* Metric / KPI Cards */
div[data-testid="stMetric"] {
  background: #FFFFFF;
  border: 1px solid #E2E8F0;
  border-radius: 14px;
  padding: 16px 20px;
  box-shadow: 0 2px 8px -2px rgba(15, 23, 42, 0.05), 0 1px 3px rgba(15, 23, 42, 0.03);
  transition: all 0.2s cubic-bezier(0.16, 1, 0.3, 1);
  position: relative;
  overflow: hidden;
}
div[data-testid="stMetric"]:hover {
  transform: translateY(-2px);
  box-shadow: 0 12px 24px -4px rgba(30, 58, 138, 0.10);
  border-color: #BFDBFE;
}
div[data-testid="stMetricLabel"] p {
  color: #64748B !important;
  font-weight: 600 !important;
  font-size: 12px !important;
  text-transform: uppercase !important;
  letter-spacing: 0.6px !important;
}
div[data-testid="stMetricValue"] {
  color: #0F172A !important;
  font-weight: 800 !important;
  font-size: 24px !important;
  letter-spacing: -0.5px;
}

/* Sidebar styling */
section[data-testid="stSidebar"] {
  background: linear-gradient(180deg, #09122C 0%, #0F172A 45%, #1E293B 100%) !important;
  border-right: 1px solid rgba(255, 255, 255, 0.06);
}
section[data-testid="stSidebar"] .brand-title {
  color: #FFFFFF;
  font-size: 19px;
  font-weight: 900;
  letter-spacing: -0.3px;
  padding: 8px 4px 2px 4px;
}
section[data-testid="stSidebar"] .brand-badge {
  display: inline-block;
  background: rgba(37, 99, 235, 0.25);
  border: 1px solid rgba(96, 165, 250, 0.3);
  color: #93C5FD;
  font-size: 10px;
  font-weight: 700;
  padding: 2px 8px;
  border-radius: 999px;
  letter-spacing: 0.8px;
  text-transform: uppercase;
  margin-bottom: 8px;
}
section[data-testid="stSidebar"] .nav-grp {
  color: #94A3B8;
  font-size: 11px;
  font-weight: 700;
  letter-spacing: 1.2px;
  text-transform: uppercase;
  padding: 16px 8px 6px 8px;
}
section[data-testid="stSidebar"] button {
  justify-content: flex-start;
  border: none;
  border-radius: 10px;
  padding: 9px 14px;
  font-size: 13.5px;
  font-weight: 500;
  transition: all 0.15s ease;
  box-shadow: none;
  margin-bottom: 3px;
}
section[data-testid="stSidebar"] button[data-testid="stBaseButton-secondary"] {
  background: transparent;
  color: #CBD5E1;
}
section[data-testid="stSidebar"] button[data-testid="stBaseButton-secondary"]:hover {
  background: rgba(255, 255, 255, 0.08);
  color: #FFFFFF;
  padding-left: 18px;
}
section[data-testid="stSidebar"] button[data-testid="stBaseButton-primary"] {
  background: linear-gradient(90deg, rgba(37, 99, 235, 0.4) 0%, rgba(37, 99, 235, 0.15) 100%);
  color: #FFFFFF;
  font-weight: 700;
  border-left: 3px solid #60A5FA;
}
section[data-testid="stSidebar"] .user-card {
  background: rgba(255, 255, 255, 0.06);
  border: 1px solid rgba(255, 255, 255, 0.10);
  border-radius: 12px;
  padding: 12px 14px;
  margin: 10px 4px 14px 4px;
  color: #FFFFFF;
}
section[data-testid="stSidebar"] .user-name {
  font-weight: 700;
  font-size: 13px;
  color: #FFFFFF;
}
section[data-testid="stSidebar"] .user-role {
  font-size: 11px;
  color: #93C5FD;
}
section[data-testid="stSidebar"] .foot-info {
  color: #94A3B8;
  font-size: 11px;
  padding: 8px 4px;
  line-height: 1.5;
  border-top: 1px solid rgba(255, 255, 255, 0.08);
  margin-top: 16px;
}

/* Buttons */
.stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
  width: 100%;
  border-radius: 10px;
  font-weight: 600;
  font-size: 14px;
  padding: 8px 16px;
  transition: all 0.18s ease;
}
.stApp [data-testid="stBaseButton-primary"], .stApp [data-testid="stBaseButton-primaryFormSubmit"] {
  background: #1E3A8A !important;
  border: 1px solid #1E3A8A !important;
  color: #FFFFFF !important;
  box-shadow: 0 2px 6px rgba(30, 58, 138, 0.2);
}
.stApp [data-testid="stBaseButton-primary"]:hover, .stApp [data-testid="stBaseButton-primaryFormSubmit"]:hover {
  background: #2563EB !important;
  border-color: #2563EB !important;
  box-shadow: 0 6px 16px rgba(37, 99, 235, 0.35);
  transform: translateY(-1px);
}
.stApp [data-testid="stBaseButton-secondary"], .stApp [data-testid="stBaseButton-secondaryFormSubmit"] {
  background: #FFFFFF !important;
  border: 1px solid #CBD5E1 !important;
  color: #1E3A8A !important;
}
.stApp [data-testid="stBaseButton-secondary"]:hover, .stApp [data-testid="stBaseButton-secondaryFormSubmit"]:hover {
  background: #EFF6FF !important;
  border-color: #93C5FD !important;
  color: #1D4ED8 !important;
}

/* Danger / Delete button style */
.danger-box {
  background: #FEF2F2;
  border: 1px solid #FCA5A5;
  border-radius: 12px;
  padding: 16px;
  margin: 12px 0;
}

/* Badges */
.badge {
  display: inline-block;
  padding: 4px 12px;
  border-radius: 999px;
  font-size: 11.5px;
  font-weight: 700;
  letter-spacing: 0.4px;
}
.b-ok { background: #DCFCE7; color: #166534; border: 1px solid #BBF7D0; }
.b-bad { background: #FEE2E2; color: #991B1B; border: 1px solid #FECACA; }
.b-warn { background: #FEF9C3; color: #854D0E; border: 1px solid #FEF08A; }
.b-orange { background: #FFEDD5; color: #9A3412; border: 1px solid #FED7AA; }
.b-info { background: #E0E7FF; color: #3730A3; border: 1px solid #C7D2FE; }

/* Bill Summary Card */
.sum-card {
  background: #FFFFFF;
  border: 1px solid #DBE4F5;
  border-radius: 16px;
  padding: 20px 22px;
  box-shadow: 0 8px 24px -4px rgba(30, 58, 138, 0.08);
}
.sum-row {
  display: flex;
  justify-content: space-between;
  padding: 7px 0;
  color: #475569;
  font-size: 14px;
}
.sum-row b { color: #0F172A; }
.sum-total {
  display: flex;
  justify-content: space-between;
  padding: 12px 0;
  margin-top: 6px;
  border-top: 2px solid #E2E8F0;
  font-size: 19px;
  font-weight: 900;
  color: #0F172A;
}
.sum-bal {
  display: flex;
  justify-content: space-between;
  padding: 12px 16px;
  margin-top: 10px;
  border-radius: 12px;
  background: linear-gradient(135deg, #1E3A8A 0%, #1D4ED8 100%);
  color: #FFFFFF;
  font-size: 16px;
  font-weight: 800;
  box-shadow: 0 4px 12px rgba(30, 58, 138, 0.25);
}

/* Quick Action Tiles on Dashboard */
.action-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
  gap: 12px;
  margin-bottom: 24px;
}
.action-tile {
  background: #FFFFFF;
  border: 1px solid #E2E8F0;
  border-radius: 12px;
  padding: 14px 16px;
  display: flex;
  align-items: center;
  gap: 12px;
  text-decoration: none;
  cursor: pointer;
  transition: all 0.2s ease;
  box-shadow: 0 1px 3px rgba(0,0,0,0.03);
}
.action-tile:hover {
  transform: translateY(-2px);
  border-color: #93C5FD;
  box-shadow: 0 6px 16px rgba(37, 99, 235, 0.10);
}
.action-icon {
  width: 40px;
  height: 40px;
  border-radius: 10px;
  background: #EFF6FF;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 18px;
  color: #2563EB;
}
.action-title {
  font-weight: 700;
  font-size: 13.5px;
  color: #0F172A;
}
.action-desc {
  font-size: 11.5px;
  color: #64748B;
}

/* Dataframe & Tables styling */
div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] {
  border: 1px solid #E2E8F0 !important;
  border-radius: 12px !important;
  overflow: hidden !important;
  background: #FFFFFF !important;
  box-shadow: 0 1px 3px rgba(0,0,0,0.03);
}

/* Tabs */
button[data-baseweb="tab"] {
  font-weight: 600 !important;
  font-size: 14px !important;
}

/* Item cards in transaction form */
.item-card {
  background: #FFFFFF;
  border: 1px solid #E2E8F0;
  border-radius: 10px;
  padding: 12px 14px;
  margin-bottom: 10px;
  transition: border-color 0.15s ease;
}
.item-card:hover {
  border-color: #BFDBFE;
}

/* =====================================================================
   MOBILE FIRST RESPONSIVENESS OVERRIDES
   ===================================================================== */
@media (max-width: 768px) {
  .block-container {
    padding-top: 1rem !important;
    padding-left: 0.6rem !important;
    padding-right: 0.6rem !important;
  }
  .page-title {
    font-size: 22px !important;
  }
  .page-sub {
    font-size: 12.5px !important;
    margin-bottom: 16px !important;
  }
  div[data-testid="stMetric"] {
    padding: 12px 14px !important;
    border-radius: 10px !important;
  }
  div[data-testid="stMetricValue"] {
    font-size: 18px !important;
  }
  div[data-testid="stMetricLabel"] p {
    font-size: 10.5px !important;
  }
  .sum-card {
    padding: 14px 16px !important;
  }
  .sum-total {
    font-size: 17px !important;
  }
  .sum-bal {
    font-size: 15px !important;
    padding: 10px 14px !important;
  }
  /* Allow metric columns to wrap neatly into 2 columns on mobile */
  div[data-testid="stHorizontalBlock"]:has(div[data-testid="stMetric"]) {
    flex-wrap: wrap !important;
    gap: 8px !important;
  }
  div[data-testid="stHorizontalBlock"]:has(div[data-testid="stMetric"]) > div[data-testid="column"] {
    min-width: calc(50% - 6px) !important;
    flex: 1 1 calc(50% - 6px) !important;
  }
}
</style>
""",
    unsafe_allow_html=True,
)


# =====================================================================
# AUTHENTICATION & LOGIN SCREEN
# =====================================================================
def require_login() -> None:
    if S.get("user"):
        return

    # Executive Centered Login Card
    _, mid, _ = st.columns([1, 1.25, 1])
    with mid:
        st.markdown(
            f"""
            <div style="background:#FFFFFF; border:1px solid #E2E8F0; border-radius:20px; padding:32px 30px;
                        box-shadow:0 12px 36px -4px rgba(15,23,42,0.10); margin-top:40px;">
                <div style="text-align:center; margin-bottom:24px;">
                    <div style="width:54px; height:54px; border-radius:14px; background:linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%);
                                margin:0 auto 12px auto; display:flex; align-items:center; justify-content:center; color:#fff; font-size:26px;">
                        💼
                    </div>
                    <div style="font-size:22px; font-weight:800; color:#0F172A; letter-spacing:-0.4px;">{BUSINESS_NAME}</div>
                    <div style="font-size:12.5px; color:#64748B; margin-top:4px;">Sign in to access your secure business management portal</div>
                </div>
            """,
            unsafe_allow_html=True,
        )

        with st.form("auth_login_form"):
            uname = st.text_input("Username", placeholder="e.g. admin").strip()
            pwd = st.text_input("Password", type="password", placeholder="Enter your password")
            submit = st.form_submit_button("Sign In Securely", type="primary")

            if submit:
                if not uname or not pwd:
                    st.warning("Please enter both username and password.")
                else:
                    user = db.authenticate(uname, pwd)
                    if user:
                        S.user = user
                        S.auth = True
                        db.log_activity(user["username"], "LOGIN", f"User {user['username']} logged in successfully")
                        st.rerun()
                    else:
                        S.fails = S.get("fails", 0) + 1
                        time.sleep(min(S.fails * 0.5, 2))
                        st.error("Invalid username or password. Please verify credentials.")

        st.markdown(
            """
            <div style="text-align:center; margin-top:16px; font-size:11.5px; color:#94A3B8;">
                Default Administrator: <b>admin</b> &nbsp;|&nbsp; <b>admin123</b><br/>
                Password can be changed inside User Accounts.
            </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    st.stop()


# =====================================================================
# COMMON HELPERS
# =====================================================================
def current_user() -> dict:
    return S.get("user") or {"username": "admin", "full_name": "Administrator", "role": "Admin"}


def is_admin() -> bool:
    return current_user().get("role") == "Admin"


def header(title: str, sub: str) -> None:
    st.markdown(f'<div class="page-title">{title}</div><div class="page-sub">{sub}</div>', unsafe_allow_html=True)
    if S.get("flash"):
        st.success(S.pop("flash"))


def section(title: str) -> None:
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)


def money(v: float) -> str:
    return f"{CUR} {v:,.0f}"


def badge(status: str) -> str:
    cls = {"PAID": "ok", "UNPAID": "bad", "PARTIALLY PAID": "warn", "CHEQUE PENDING": "orange", "RETURNED": "info"}
    return f'<span class="badge b-{cls.get(status, "warn")}">{status}</span>'


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
    safe_fname = re.sub(r'[^A-Za-z0-9_-]+', '_', str(fname or 'report')).strip('_') or 'report'
    c1, c2, _ = st.columns([1.2, 1.2, 3])
    c1.download_button("Download Excel", to_excel_bytes({title[:31]: df}), f"{safe_fname}.xlsx", XLSX, key=f"{key}_x")
    c2.download_button("Download PDF", to_pdf_bytes(title, df), f"{safe_fname}.pdf", "application/pdf", key=f"{key}_p")


def workbook_sheets() -> dict[str, pd.DataFrame]:
    return {"Transactions": S.tx, "Items": S.line_items, "Products": S.products,
            "Parties": pd.DataFrame({"Party": sorted(set(S.extra_parties))})}


def save_all(*only: str) -> None:
    try:
        db.save_all(S.tx, S.line_items, S.products, S.extra_parties, only or db.TABLES)
    except Exception as exc:
        st.error("Could not save to the database, so this change was NOT stored. Please check your connection.")
        st.caption(f"Technical detail: {type(exc).__name__}: {str(exc)[:300]}")
        st.stop()


def update_bill(tid: str, remark: str | None = None, **changes) -> None:
    m = S.tx["Txn ID"] == tid
    for k, v in changes.items():
        S.tx.loc[m, k] = v
    if remark:
        cur = S.tx.loc[m, "Remarks"].iloc[0]
        S.tx.loc[m, "Remarks"] = f"{cur} | {remark}".strip(" |")
    S.tx = recompute(S.tx)
    save_all("transactions")


# =====================================================================
# SIDEBAR
# =====================================================================
def sidebar() -> None:
    sb = st.sidebar
    sb.markdown('<div class="brand-title">💼 Business Manager</div>'
                '<div class="brand-badge">Executive Edition</div>', unsafe_allow_html=True)

    # Current User Profile Box
    u = current_user()
    sb.markdown(
        f"""
        <div class="user-card">
            <div style="display:flex; align-items:center; gap:10px;">
                <div style="width:34px; height:34px; border-radius:50%; background:#2563EB; display:flex;
                            align-items:center; justify-content:center; font-weight:800; font-size:14px; color:#fff;">
                    {u['full_name'][:1].upper()}
                </div>
                <div>
                    <div class="user-name">{u['full_name']}</div>
                    <div class="user-role">Role: {u['role']}</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Navigation menu
    for grp, pages in GROUPS:
        # Hide admin pages for non-admin staff
        filtered_pages = [p for p in pages if (p != "User Accounts" or is_admin())]
        if not filtered_pages:
            continue
        sb.markdown(f'<div class="nav-grp">{grp}</div>', unsafe_allow_html=True)
        for p in filtered_pages:
            if sb.button(p, key=f"nav_{p}", type="primary" if S.page == p else "secondary"):
                S.page = p
                st.rerun()

    # System status
    sb.markdown(
        f'<div class="foot-info">'
        f'<b>Database Status:</b> {"Cloud (Supabase)" if db.is_cloud() else "Local SQLite"}<br/>'
        f'<b>Records:</b> {len(S.tx)} bills &bull; {len(all_parties())} parties &bull; {len(S.products)} products'
        f'</div>',
        unsafe_allow_html=True,
    )

    if sb.button("Sign Out", key="logout_btn"):
        db.log_activity(u["username"], "LOGOUT", "User signed out")
        S.user = None
        S.auth = False
        st.rerun()


# =====================================================================
# PAGE: DASHBOARD (TOP 1% FINTECH INTELLIGENCE)
# =====================================================================
def page_dashboard() -> None:
    df = S.tx
    header("Business Dashboard", "Live executive intelligence. Real-time calculations of revenue, collections and receivables.")

    # Quick action navigation pills
    col_a, col_b, col_c, col_d, col_e = st.columns(5)
    if col_a.button("➕ New Transaction", key="qa_ntx", type="primary"):
        S.page = "New Transaction"
        st.rerun()
    if col_b.button("💳 Bills & Payments", key="qa_bills"):
        S.page = "Bills & Payments"
        st.rerun()
    if col_c.button("📋 Party Ledger", key="qa_led"):
        S.page = "Party Ledger"
        st.rerun()
    if col_d.button("⏳ Aging Report", key="qa_age"):
        S.page = "Reports"
        st.rerun()
    if col_e.button("🔍 Search Center", key="qa_search"):
        S.page = "Search Center"
        st.rerun()

    st.write("")

    # Financial Core KPIs
    tot_billed = df["Bill Amount"].sum()
    tot_paid = df["Paid Amount"].sum()
    tot_bal = df["Balance"].sum()
    tot_ret = df["Return Amount"].sum()
    col_pct = (tot_paid / tot_billed * 100) if tot_billed > 0 else 0.0

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Total Invoiced", money(tot_billed), help="Cumulative gross amount billed across all time")
    kpi2.metric("Total Collected", money(tot_paid), f"{col_pct:.1f}% Recovery")
    kpi3.metric("Outstanding Balance", money(tot_bal), help="Net pending receivables due from all parties")
    kpi4.metric("Total Returns", money(tot_ret))

    # Bill status breakdown
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Unpaid Invoices", int((df["Status"] == "UNPAID").sum()))
    c2.metric("Partially Paid", int((df["Status"] == "PARTIALLY PAID").sum()))
    c3.metric("Fully Paid", int((df["Status"] == "PAID").sum()))
    c4.metric("Pending Cheques", int((df["Status"] == "CHEQUE PENDING").sum()))

    # Pending Cheques Alert Banner
    pend_chqs = df[df["Cheque Status"] == "Pending"]
    if not pend_chqs.empty:
        st.warning(f"⚠️ **Attention Required:** You have **{len(pend_chqs)} pending cheque(s)** totaling "
                   f"**{money(pend_chqs['Paid Amount'].sum())}** waiting for bank clearance. "
                   f"Check the **Cheque Tracker** to review maturity dates.")

    if df.empty:
        st.info("No records recorded yet. Start by creating your first bill from **New Transaction**.")
        return

    # Visual Insights
    g1, g2 = st.columns(2)
    with g1:
        section("Top 5 Debtors (Highest Balance Due)")
        top_debtors = df.groupby("Party / Company")["Balance"].sum().sort_values(ascending=False).head(5)
        top_debtors = top_debtors[top_debtors > 0]
        if not top_debtors.empty:
            try:
                st.bar_chart(top_debtors, color="#1E3A8A", horizontal=True)
            except TypeError:
                st.bar_chart(top_debtors, color="#1E3A8A")
        else:
            st.caption("All accounts are fully settled. No pending debtor balances!")

    with g2:
        section("Monthly Sales vs Collections")
        m_df = df.dropna(subset=["Date"]).assign(Month=lambda x: x["Date"].dt.strftime("%Y-%m"))
        if not m_df.empty:
            m_summary = m_df.groupby("Month")[["Bill Amount", "Paid Amount"]].sum()
            st.bar_chart(m_summary, color=["#1E3A8A", "#60A5FA"])
        else:
            st.caption("No dated records available for monthly chart.")

    section("Recent 10 Transactions")
    recent = df.sort_values("Date", ascending=False).head(10)
    show(recent, ["Txn ID", "Date", "Bill No", "Party / Company", "Bill Amount", "Paid Amount", "Balance", "Status"])


# =====================================================================
# PAGE: NEW TRANSACTION (RESPONSIVE ITEM BUILDER)
# =====================================================================
def page_new_transaction() -> None:
    header("New Transaction", "Create sales invoices with itemized line items, auto-rates, tax & discounts.")
    fid = S.ntx_fid
    prods = S.products
    names = prods["Product"].astype(str).tolist()

    left, right = st.columns([3, 2], gap="large")

    with left:
        section("Invoice Details")
        c1, c2, c3 = st.columns([2, 1, 1])
        with c1:
            party = party_picker("Party / Company *", f"ntx_party_{fid}")
        bill_date = c2.date_input("Date", value=date.today(), key=f"ntx_date_{fid}")
        bill_no = c3.text_input("Bill No *", value=next_bill_no(), key=f"ntx_bill_{fid}")

        section("Line Items")
        if not names:
            st.warning("Product catalog is empty. Add products in the expander below or under 'Products' page.")

        # Line items builder
        if not S.ntx_items:
            S.ntx_items = [{"prod": NO_PRODUCT, "qty": 0.0, "rate": 0.0, "amt": 0.0}]

        lines = []
        subtotal = 0.0

        for i, it in enumerate(list(S.ntx_items)):
            with st.container():
                st.markdown(f'<div class="item-card"><b>Item #{i+1}</b>', unsafe_allow_html=True)
                ic1, ic2, ic3, ic4 = st.columns([3.5, 1.3, 1.5, 1.7])
                sel_prod = ic1.selectbox("Product", [NO_PRODUCT] + names,
                                         index=([NO_PRODUCT] + names).index(it["prod"]) if it["prod"] in ([NO_PRODUCT] + names) else 0,
                                         key=f"p_{fid}_{i}")
                
                # Fetch default product rate if product changed
                p_match = prods[prods["Product"].astype(str) == sel_prod]
                default_rate = float(p_match["Rate"].iloc[0]) if not p_match.empty and pd.notna(p_match["Rate"].iloc[0]) else 0.0
                unit_label = str(p_match["Unit"].iloc[0]) if not p_match.empty else "Pcs"
                
                init_rate = default_rate if sel_prod != it["prod"] and default_rate > 0 else it["rate"]
                qty = ic2.number_input(f"Qty ({unit_label})", min_value=0.0, value=float(it["qty"]), step=1.0, key=f"q_{fid}_{i}")
                rate = ic3.number_input("Rate", min_value=0.0, value=float(init_rate), step=1.0, key=f"r_{fid}_{i}")
                amt = round(qty * rate, 2)
                ic4.metric("Amount", f"{amt:,.2f}")

                # Update current item in state
                S.ntx_items[i] = {"prod": sel_prod, "qty": qty, "rate": rate, "amt": amt}
                if sel_prod != NO_PRODUCT and qty > 0:
                    lines.append((sel_prod, qty, rate, amt))
                    subtotal += amt
                st.markdown('</div>', unsafe_allow_html=True)

        ib1, ib2, ib3 = st.columns([1, 1, 1])
        if ib1.button("➕ Add Item", key=f"btn_add_item_{fid}"):
            S.ntx_items.append({"prod": NO_PRODUCT, "qty": 0.0, "rate": 0.0, "amt": 0.0})
            st.rerun()
        if ib2.button("➖ Remove Last", key=f"btn_rem_item_{fid}", disabled=len(S.ntx_items) <= 1):
            S.ntx_items.pop()
            st.rerun()
        if ib3.button("🔄 Clear Items", key=f"btn_clr_item_{fid}"):
            S.ntx_items = [{"prod": NO_PRODUCT, "qty": 0.0, "rate": 0.0, "amt": 0.0}]
            st.rerun()

        with st.expander("+ Quick Register New Product"):
            with st.form(f"quick_prod_{fid}", clear_on_submit=True):
                qp1, qp2, qp3 = st.columns([3, 1, 1])
                pn = qp1.text_input("Product Name")
                pu = qp2.text_input("Unit", value="Pcs")
                pr = qp3.number_input("Default Rate", min_value=0.0, step=1.0)
                if st.form_submit_button("Add Product") and pn.strip():
                    S.products = pd.concat([S.products, pd.DataFrame([[pn.strip(), pu.strip() or "Pcs", pr]], columns=PRODUCT_COLS)]) \
                        .drop_duplicates("Product", keep="last").reset_index(drop=True)
                    save_all("products")
                    st.success(f"Added product: {pn.strip()}")
                    st.rerun()

        remarks = st.text_area("Remarks / Notes", key=f"ntx_rem_{fid}", height=80,
                               placeholder="Add any delivery notes, order reference, or special terms...")

    with right:
        summary_container = st.container()
        section("Financials & Settlement")
        d1, d2 = st.columns(2)
        dtype = d1.radio("Discount Type", ["Percentage (%)", "Fixed Amount"], horizontal=True, key=f"dt_{fid}")
        dval = d2.number_input("Discount Value", min_value=0.0, value=0.0, key=f"dv_{fid}")
        discount = round(subtotal * dval / 100 if dtype.startswith("Percentage") else dval, 2)
        discount = min(discount, subtotal)

        tax_pct = st.number_input("Tax / GST (%)", min_value=0.0, max_value=100.0, value=0.0, key=f"tx_{fid}")
        tax = round((subtotal - discount) * tax_pct / 100, 2)
        total = round(subtotal - discount + tax, 2)
        paid = st.number_input("Amount Received Now", min_value=0.0, value=0.0, key=f"pd_{fid}")
        balance = round(total - paid, 2)

        method = st.selectbox("Payment Method", PAYMENT_METHODS, key=f"pm_{fid}")
        chq_no = chq_bank = chq_status = sale_base = ""
        chq_date = None
        if method == "Cheque":
            k1, k2 = st.columns(2)
            chq_no = k1.text_input("Cheque No", key=f"cn_{fid}")
            chq_bank = k2.text_input("Bank Name", key=f"cb_{fid}")
            k3, k4 = st.columns(2)
            chq_date = k3.date_input("Cheque Maturity Date", value=None, key=f"cd_{fid}")
            chq_status = k4.selectbox("Cheque Status", CHEQUE_STATUSES[1:], key=f"cs_{fid}")

        sale_base = st.text_input("Sale Base (Optional)", key=f"sb_{fid}", placeholder="e.g. 5/8 or S/B")

        with summary_container:
            section("Bill Financial Summary")
            st.markdown(
                f"""
                <div class="sum-card">
                    <div class="sum-row">
                        <span>Items Subtotal ({len(lines)} item{'s' if len(lines) != 1 else ''})</span>
                        <b>{CUR} {subtotal:,.2f}</b>
                    </div>
                    <div class="sum-row">
                        <span>Discount {f'({dval:g}%)' if dtype.startswith('Percentage') and dval else ''}</span>
                        <b style="color:#DC2626;">- {CUR} {discount:,.2f}</b>
                    </div>
                    <div class="sum-row">
                        <span>Tax / GST ({tax_pct:g}%)</span>
                        <b>+ {CUR} {tax:,.2f}</b>
                    </div>
                    <div class="sum-total">
                        <span>Grand Total</span>
                        <span>{CUR} {total:,.2f}</span>
                    </div>
                    <div class="sum-row">
                        <span>Paid Now</span>
                        <b style="color:#16A34A;">{CUR} {paid:,.2f}</b>
                    </div>
                    <div class="sum-bal">
                        <span>Balance Due</span>
                        <span>{CUR} {balance:,.2f}</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.write("")
        if st.button("💾 Confirm & Save Transaction", type="primary", key=f"save_btn_{fid}"):
            dup = S.tx[(S.tx["Party / Company"] == party) & (S.tx["Bill No"] == bill_no.strip())]
            if not party or party == NEW_PARTY:
                st.error("Please select a valid customer / party name.")
            elif not lines:
                st.error("Add at least one line item with quantity greater than zero.")
            elif paid > total:
                st.error("Paid amount cannot exceed the grand total bill amount.")
            elif not bill_no.strip():
                st.error("Bill number cannot be blank.")
            elif not dup.empty:
                st.error(f"Bill No '{bill_no}' already exists for {party}. Please assign a unique bill number.")
            else:
                tid = next_txn_ids(1)[0]
                add_party_if_new(party)
                prod_summary = ", ".join(f"{p} x{q:g}" for p, q, _, _ in lines)
                tot_qty = sum(q for _, q, _, _ in lines)

                row = {
                    "Txn ID": tid, "Date": pd.to_datetime(bill_date), "Bill No": bill_no.strip(),
                    "Party / Company": party, "Product / Item": prod_summary,
                    "Qty": tot_qty, "Subtotal": subtotal, "Discount": discount, "Tax": tax,
                    "Bill Amount": total, "Paid Amount": paid, "Return Amount": 0.0,
                    "Payment Method": method, "Cheque No": chq_no,
                    "Cheque Date": pd.to_datetime(chq_date) if chq_date else pd.NaT,
                    "Bank Name": chq_bank, "Cheque Status": chq_status,
                    "Sale Base": sale_base, "Remarks": remarks,
                }
                S.tx = recompute(pd.concat([S.tx, pd.DataFrame([row])], ignore_index=True))
                new_items_df = pd.DataFrame([(tid, p, q, r, a) for p, q, r, a in lines], columns=ITEM_COLS)
                S.line_items = pd.concat([S.line_items, new_items_df], ignore_index=True)
                save_all("transactions", "items", "parties")

                db.log_activity(current_user()["username"], "CREATE_BILL",
                                f"Created Bill #{bill_no.strip()} for {party} (Total: {CUR} {total:,.2f})")

                S.flash = f"✅ Bill #{bill_no.strip()} saved for {party}: Total {CUR} {total:,.2f} (Balance: {CUR} {balance:,.2f})."
                S.ntx_fid += 1
                S.ntx_items = []
                st.rerun()


# =====================================================================
# PAGE: BILLS & PAYMENTS (WITH FULL EDITING & DELETING FOR WRONG ENTRIES)
# =====================================================================
def page_bills() -> None:
    header("Bills & Payments Manager", "Inspect invoices, record collections, print slips, and edit or remove wrong entries.")
    df = S.tx
    if df.empty:
        st.info("No bills stored yet. Create one from **New Transaction**.")
        return

    q = st.text_input("🔍 Quick Search Bill", placeholder="Type customer name or bill number to find instantly...")
    view = df
    if q:
        view = df[df["Party / Company"].str.contains(q, case=False, regex=False) |
                  df["Bill No"].str.contains(q, case=False, regex=False)]
    if view.empty:
        st.warning("No bills found matching your search term.")
        return

    labels = {
        r["Txn ID"]: f"#{r['Bill No']}  |  {r['Party / Company']}  |  Total: {money(r['Bill Amount'])}  |  Due: {money(r['Balance'])}  [{r['Status']}]"
        for _, r in view.iterrows()
    }
    sorted_ids = view.sort_values("Date", ascending=False)["Txn ID"].tolist()
    tid = st.selectbox("Select Bill to Manage", sorted_ids, format_func=lambda t: labels[t])
    row = df[df["Txn ID"] == tid].iloc[0]

    # Tabs for modern multi-action interface
    tab_view, tab_edit, tab_delete = st.tabs(["⚡ Overview & Actions", "✏️ Edit Bill (Correct Entry)", "🗑️ Delete Transaction"])

    # ------------------ TAB 1: OVERVIEW & ACTIONS ------------------
    with tab_view:
        st.markdown(badge(row["Status"]), unsafe_allow_html=True)
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Invoice Total", money(row["Bill Amount"]))
        m2.metric("Paid to Date", money(row["Paid Amount"]))
        m3.metric("Returned", money(row["Return Amount"]))
        m4.metric("Net Balance Due", money(row["Balance"]))
        m5.metric("Discount / Tax", f"{row['Discount']:,.0f} / {row['Tax']:,.0f}")

        # Quick Export / Share Bar
        st.write("")
        sh1, sh2, sh3 = st.columns([1.5, 1.5, 1.5])
        with sh1:
            it_for_inv = S.line_items[S.line_items["Txn ID"] == tid]
            st.download_button("📄 Download Invoice (PDF)", invoice_pdf(row, it_for_inv),
                               f"Invoice_{row['Bill No']}.pdf", "application/pdf", key=f"inv_pdf_{tid}")
        with sh2:
            # WhatsApp share
            wa_text = whatsapp_invoice_text(row)
            wa_link = whatsapp_url("", wa_text)
            st.link_button("📲 Share via WhatsApp", wa_link)
        with sh3:
            with st.popover("🧾 View Thermal POS Slip"):
                st.markdown(thermal_receipt_html(row, it_for_inv), unsafe_allow_html=True)
                st.caption("Tip: Press Ctrl+P (or Cmd+P) to print directly to a thermal or standard printer.")

        act_col1, act_col2 = st.columns(2, gap="large")
        with act_col1:
            section("Record Payment (Deposit)")
            with st.form(f"pay_form_{tid}", clear_on_submit=True):
                p_amt = st.number_input("Amount Received", min_value=0.0, value=0.0)
                p_meth = st.selectbox("Payment Method", PAYMENT_METHODS[1:])
                p_note = st.text_input("Payment Note / Reference")
                if st.form_submit_button("Record Payment", type="primary"):
                    if p_amt <= 0:
                        st.error("Please enter an amount greater than zero.")
                    elif p_amt > max(row["Balance"], 0):
                        st.error(f"Payment of {money(p_amt)} exceeds current balance of {money(row['Balance'])}.")
                    else:
                        changes = {"Paid Amount": row["Paid Amount"] + p_amt}
                        if row["Payment Method"] == "Not specified":
                            changes["Payment Method"] = p_meth
                        hist_note = f"Payment {CUR} {p_amt:,.0f} ({p_meth}) on {date.today():%d %b %Y}" + (f": {p_note}" if p_note else "")
                        update_bill(tid, hist_note, **changes)
                        db.log_activity(current_user()["username"], "PAYMENT",
                                        f"Received {money(p_amt)} for Bill #{row['Bill No']} ({row['Party / Company']})")
                        S.flash = f"Payment of {money(p_amt)} recorded successfully."
                        st.rerun()

        with act_col2:
            section("Record Goods Return")
            ret_limit = max(row["Bill Amount"] - row["Return Amount"], 0.0)
            with st.form(f"ret_form_{tid}", clear_on_submit=True):
                r_amt = st.number_input(f"Return Value (Max {ret_limit:,.0f})", min_value=0.0, value=0.0)
                r_reason = st.text_input("Return Reason")
                if st.form_submit_button("Record Return", type="primary"):
                    if r_amt <= 0:
                        st.error("Please enter an amount greater than zero.")
                    elif r_amt > ret_limit:
                        st.error("Return value cannot exceed bill total.")
                    else:
                        hist_ret = f"Return {CUR} {r_amt:,.0f} on {date.today():%d %b %Y}" + (f": {r_reason}" if r_reason else "")
                        update_bill(tid, hist_ret, **{"Return Amount": row["Return Amount"] + r_amt})
                        db.log_activity(current_user()["username"], "RETURN",
                                        f"Recorded return of {money(r_amt)} for Bill #{row['Bill No']}")
                        S.flash = f"Return of {money(r_amt)} recorded."
                        st.rerun()

        section("Line Items Breakdown")
        it = S.line_items[S.line_items["Txn ID"] == tid]
        if it.empty:
            st.caption(f"Summary goods: {row['Product / Item'] or 'Goods & Services'}")
        else:
            show(it, ["Product", "Qty", "Rate", "Amount"])

        if row["Remarks"]:
            st.caption(f"**Audit History & Remarks:** {row['Remarks']}")

    # ------------------ TAB 2: EDIT TRANSACTION (CORRECT ENTRY) ------------------
    with tab_edit:
        section("Correct or Edit Transaction Data")
        st.info("You can correct any field in this bill. The balance and status will recalculate automatically.")

        with st.form(f"edit_txn_form_{tid}"):
            e1, e2, e3 = st.columns([2, 1, 1])
            e_party = e1.selectbox("Party / Company", all_parties(),
                                   index=all_parties().index(row["Party / Company"]) if row["Party / Company"] in all_parties() else 0)
            
            cur_date = row["Date"].date() if pd.notna(row["Date"]) and hasattr(row["Date"], "date") else date.today()
            e_date = e2.date_input("Date", value=cur_date)
            e_bill_no = e3.text_input("Bill No", value=str(row["Bill No"]))

            e_prod = st.text_input("Product Summary / Items", value=str(row.get("Product / Item", "")))

            f1, f2, f3, f4 = st.columns(4)
            e_subtotal = f1.number_input("Subtotal", min_value=0.0, value=float(row.get("Subtotal", row["Bill Amount"])))
            e_disc = f2.number_input("Discount", min_value=0.0, value=float(row.get("Discount", 0)))
            e_tax = f3.number_input("Tax / GST", min_value=0.0, value=float(row.get("Tax", 0)))
            e_paid = f4.number_input("Paid Amount", min_value=0.0, value=float(row.get("Paid Amount", 0)))

            p1, p2, p3 = st.columns(3)
            e_pm = p1.selectbox("Payment Method", PAYMENT_METHODS,
                                index=PAYMENT_METHODS.index(row["Payment Method"]) if row["Payment Method"] in PAYMENT_METHODS else 0)
            e_chq_no = p2.text_input("Cheque No", value=str(row.get("Cheque No") or ""))
            e_chq_bank = p3.text_input("Bank Name", value=str(row.get("Bank Name") or ""))

            e_remarks = st.text_area("Remarks / Notes", value=str(row.get("Remarks") or ""), height=70)

            if st.form_submit_button("💾 Save Corrected Bill", type="primary"):
                # Duplicate check against other bills
                dup = S.tx[(S.tx["Party / Company"] == e_party) & (S.tx["Bill No"] == e_bill_no.strip()) & (S.tx["Txn ID"] != tid)]
                if not dup.empty:
                    st.error(f"Bill No '{e_bill_no}' is already used by another record for {e_party}.")
                elif not e_bill_no.strip():
                    st.error("Bill No cannot be blank.")
                else:
                    new_bill_amt = round(e_subtotal - e_disc + e_tax, 2)
                    updates = {
                        "Party / Company": e_party,
                        "Date": pd.to_datetime(e_date),
                        "Bill No": e_bill_no.strip(),
                        "Product / Item": e_prod,
                        "Subtotal": e_subtotal,
                        "Discount": e_disc,
                        "Tax": e_tax,
                        "Bill Amount": new_bill_amt,
                        "Paid Amount": e_paid,
                        "Payment Method": e_pm,
                        "Cheque No": e_chq_no,
                        "Bank Name": e_chq_bank,
                        "Remarks": e_remarks,
                    }
                    update_bill(tid, f"Edited on {date.today():%d %b %Y} by {current_user()['username']}", **updates)
                    db.log_activity(current_user()["username"], "EDIT_BILL",
                                    f"Updated Bill #{e_bill_no.strip()} ({e_party})")
                    S.flash = f"Bill #{e_bill_no.strip()} updated successfully."
                    st.rerun()

    # ------------------ TAB 3: DELETE TRANSACTION ------------------
    with tab_delete:
        section("Danger Zone: Remove Wrong Record")
        st.markdown(
            f"""
            <div class="danger-box">
                <div style="font-weight:700; color:#991B1B; font-size:15px; margin-bottom:6px;">⚠️ Permanently Remove Record</div>
                <div style="color:#7F1D1D; font-size:13px; line-height:1.5;">
                    Are you sure you want to delete <b>Bill #{row['Bill No']}</b> for <b>{row['Party / Company']}</b>?<br/>
                    This will permanently remove the invoice, all associated line items, and adjust the party ledger balance immediately.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        confirm_del = st.checkbox(f"I confirm that I want to permanently delete Bill #{row['Bill No']} (Amount: {money(row['Bill Amount'])})", key=f"conf_del_{tid}")
        if st.button("🗑️ Delete This Transaction Now", type="secondary", disabled=not confirm_del, key=f"del_btn_{tid}"):
            # Atomic deletion from DB
            db.delete_transaction_atomic(tid)
            # Remove from local session frames
            S.tx = S.tx[S.tx["Txn ID"] != tid].reset_index(drop=True)
            S.line_items = S.line_items[S.line_items["Txn ID"] != tid].reset_index(drop=True)
            db.log_activity(current_user()["username"], "DELETE_BILL",
                            f"Deleted Bill #{row['Bill No']} of {row['Party / Company']} (Amount: {money(row['Bill Amount'])})")
            S.flash = f"Bill #{row['Bill No']} for {row['Party / Company']} was permanently deleted."
            st.rerun()


# =====================================================================
# PAGE: CHEQUE TRACKER
# =====================================================================
def page_cheques() -> None:
    header("Cheque Management", "Track deposited cheques, monitor maturity dates, and mark clearances or bounces.")
    df = S.tx
    ch = df[df["Payment Method"] == "Cheque"].copy()
    today = pd.Timestamp(date.today())
    pend = ch[ch["Cheque Status"] == "Pending"]
    soon = pend[pend["Cheque Date"].between(today, today + timedelta(days=7))]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Pending Cheques", len(pend))
    c2.metric("Pending Value", money(pend["Paid Amount"].sum()))
    c3.metric("Cleared Value", money(ch.loc[ch["Cheque Status"] == "Cleared", "Paid Amount"].sum()))
    c4.metric("Maturing in 7 Days", len(soon))

    flt = st.selectbox("Filter Cheques by Status", ["All", "Pending", "Cleared", "Bounced"])
    if flt != "All":
        ch = ch[ch["Cheque Status"] == flt]

    def calc_due(r):
        if r["Cheque Status"] != "Pending" or pd.isna(r["Cheque Date"]):
            return ""
        n = (r["Cheque Date"] - today).days
        return "⚠️ OVERDUE" if n < 0 else "🔔 DUE TODAY" if n == 0 else f"in {n} day(s)"

    section("Cheque Records")
    if ch.empty:
        st.info("No cheques found matching the selected filter.")
    else:
        ch["Due Status"] = ch.apply(calc_due, axis=1)
        ch_view = ch.sort_values("Cheque Date", na_position="last")[
            ["Txn ID", "Party / Company", "Bill No", "Cheque No", "Cheque Date", "Bank Name", "Paid Amount", "Cheque Status", "Due Status"]
        ]
        show(ch_view)
        export_row("Cheque Tracker", ch_view, "Cheque_Tracker", "chq")

    if not pend.empty:
        section("Update Cheque Clearance")
        u = st.columns([3, 1.5, 1], vertical_alignment="bottom")
        labels = {r["Txn ID"]: f"{r['Party / Company']}  |  Bill #{r['Bill No']}  |  Cheque #{r['Cheque No'] or '-'}  |  {money(r['Paid Amount'])}"
                  for _, r in pend.iterrows()}
        tid = u[0].selectbox("Select Pending Cheque", list(labels), format_func=lambda t: labels[t])
        new_st = u[1].selectbox("New Status", ["Cleared", "Bounced"])
        if u[2].button("Update Status", type="primary"):
            update_bill(tid, f"Cheque marked {new_st.lower()} on {date.today():%d %b %Y}", **{"Cheque Status": new_st})
            db.log_activity(current_user()["username"], "CHEQUE_UPDATE", f"Cheque for {labels[tid]} marked {new_st}")
            S.flash = f"Cheque marked as {new_st}."
            st.rerun()


# =====================================================================
# PAGE: PARTY LEDGER
# =====================================================================
def page_ledger() -> None:
    header("Customer & Supplier Ledger", "Generate customer account statements with running balances, exports and WhatsApp sharing.")
    sel = party_picker("Select Party / Company", "ledger", allow_all=True)
    d = S.tx if sel == "ALL PARTIES" else S.tx[S.tx["Party / Company"] == sel]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Bills", len(d))
    c2.metric("Total Invoiced", money(d["Bill Amount"].sum()))
    c3.metric("Total Paid", money(d["Paid Amount"].sum()))
    c4.metric("Net Outstanding Balance", money(d["Balance"].sum()))

    if sel != "ALL PARTIES" and not d.empty:
        # Quick WhatsApp balance reminder
        wa_ledger_msg = (
            f"Assalam-o-Alaikum / Dear {sel},\n\n"
            f"Account statement update from {BUSINESS_NAME}:\n"
            f"Total Invoices: {len(d)}\n"
            f"Total Invoiced: {money(d['Bill Amount'].sum())}\n"
            f"Total Paid: {money(d['Paid Amount'].sum())}\n"
            f"Current Balance Due: {money(d['Balance'].sum())}\n\n"
            f"Thank you for your business!"
        )
        st.link_button(f"📲 Send Balance Summary to {sel} via WhatsApp", whatsapp_url("", wa_ledger_msg))

    section("Ledger Transaction History")
    if d.empty:
        st.info("No recorded transactions for this party yet.")
        return

    stmt = lg.compute_party_statement(S.tx, sel)
    show(stmt)
    safe_sel = re.sub(r'[^A-Za-z0-9_-]+', '_', str(sel or 'all')).strip('_') or 'all'
    export_row(f"Ledger - {sel}", stmt, f"Ledger_{safe_sel}", "led")


# =====================================================================
# PAGE: PARTIES DIRECTORY
# =====================================================================
def page_parties() -> None:
    header("Parties Directory", "Manage customer and supplier master accounts, monitor total receivables and activity.")
    parties = all_parties()
    df = S.tx

    records = []
    for p in parties:
        sub = df[df["Party / Company"] == p]
        tot_billed = sub["Bill Amount"].sum() if not sub.empty else 0.0
        tot_paid = sub["Paid Amount"].sum() if not sub.empty else 0.0
        tot_ret = sub["Return Amount"].sum() if not sub.empty else 0.0
        bal = tot_billed - tot_paid - tot_ret
        cnt = len(sub)
        last_dt = sub["Date"].max() if not sub.empty and pd.notna(sub["Date"].max()) else pd.NaT
        records.append({
            "Party / Company": p,
            "Total Invoiced": tot_billed,
            "Total Paid": tot_paid,
            "Current Balance": bal,
            "Bills Count": cnt,
            "Last Transaction": last_dt,
        })
    pdf = pd.DataFrame(records)
    if not pdf.empty:
        pdf = pdf.sort_values("Current Balance", ascending=False).reset_index(drop=True)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Parties", len(parties))
    c2.metric("Parties with Balance", int((pdf["Current Balance"] > 0).sum()) if not pdf.empty else 0)
    c3.metric("Total Receivables", money(pdf["Current Balance"].sum()) if not pdf.empty else money(0))
    c4.metric("Cumulative Volume", money(pdf["Total Invoiced"].sum()) if not pdf.empty else money(0))

    section("Register New Party")
    with st.expander("+ Add New Customer or Supplier Party", expanded=False):
        with st.form("new_party_form", clear_on_submit=True):
            np_name = st.text_input("Party / Company Name", placeholder="e.g. TARIQ TRADERS").strip().upper()
            if st.form_submit_button("Add Party", type="primary") and np_name:
                if np_name in parties:
                    st.warning(f"Party '{np_name}' already exists.")
                else:
                    add_party_if_new(np_name)
                    save_all("parties")
                    db.log_activity(current_user()["username"], "ADD_PARTY", f"Added party {np_name}")
                    S.flash = f"Added party: {np_name}"
                    st.rerun()

    section("Directory Records")
    if pdf.empty:
        st.info("No parties registered.")
    else:
        q = st.text_input("Filter Directory", placeholder="Search by party name...", key="pdir_q")
        view_pdf = pdf
        if q:
            view_pdf = view_pdf[view_pdf["Party / Company"].str.contains(q, case=False, regex=False)]
        show(view_pdf)
        export_row("Parties Directory", view_pdf, "Parties_Directory", "pdir")


# =====================================================================
# PAGE: SEARCH CENTER
# =====================================================================
def page_search() -> None:
    header("Search Center", "Multi-parameter filter across dates, parties, invoices, statuses and remarks.")
    df = S.tx
    c = st.columns(4)
    party = c[0].text_input("Party contains")
    bill = c[1].text_input("Bill No contains")
    prod = c[2].text_input("Product contains")
    status = c[3].selectbox("Status", ["All"] + STATUS_OPTIONS)

    c = st.columns(4)
    d1 = c[0].date_input("Date From", value=None, key="sc_f")
    d2 = c[1].date_input("Date To", value=None, key="sc_t")
    pay = c[2].selectbox("Payment Method", ["All"] + PAYMENT_METHODS)
    chq = c[3].selectbox("Cheque Status", ["All"] + CHEQUE_STATUSES[1:])

    r = df
    if party: r = r[r["Party / Company"].str.contains(party, case=False, regex=False)]
    if bill: r = r[r["Bill No"].str.contains(bill, case=False, regex=False)]
    if prod: r = r[r["Product / Item"].str.contains(prod, case=False, regex=False)]
    if d1: r = r[r["Date"] >= pd.to_datetime(d1)]
    if d2: r = r[r["Date"] <= pd.to_datetime(d2)]
    if status != "All": r = r[r["Status"] == status]
    if pay != "All": r = r[r["Payment Method"] == pay]
    if chq != "All": r = r[r["Cheque Status"] == chq]

    st.markdown(f"**{len(r)}** record(s) found &nbsp;|&nbsp; Outstanding in filter: **{money(r['Balance'].sum())}**")
    if not r.empty:
        cols = ["Txn ID", "Date", "Bill No", "Party / Company", "Product / Item", "Bill Amount", "Paid Amount",
                "Return Amount", "Balance", "Payment Method", "Status", "Remarks"]
        r = r.sort_values("Date", ascending=False)[cols]
        show(r)
        export_row("Search Results", r, "Search_Results", "srch")


# =====================================================================
# PAGE: PRODUCTS
# =====================================================================
def page_products() -> None:
    header("Product Catalog", "Manage items, units and default pricing used to speed up bill generation.")
    st.caption("Edit values directly in the table or add a new row at the bottom. Click Save Products when done.")

    ed = st.data_editor(
        S.products, num_rows="dynamic", hide_index=True, key=f"prod_ed_{S.ver}", **_STRETCH,
        column_config={
            "Product": st.column_config.TextColumn("Product Name", required=True),
            "Unit": st.column_config.TextColumn("Unit", default="Pcs"),
            "Rate": st.column_config.NumberColumn("Default Rate", min_value=0, default=0, format="localized"),
        },
    )

    c1, _ = st.columns([1, 3])
    if c1.button("Save Products", type="primary"):
        d = ed.copy()
        d["Product"] = _txt(d["Product"])
        d["Unit"] = _txt(d["Unit"]).replace("", "Pcs")
        d["Rate"] = pd.to_numeric(d["Rate"], errors="coerce").fillna(0.0)
        d = d[d["Product"] != ""].drop_duplicates("Product", keep="last").reset_index(drop=True)
        S.products = d[PRODUCT_COLS]
        save_all("products")
        db.log_activity(current_user()["username"], "SAVE_PRODUCTS", f"Updated product list ({len(d)} items)")
        S.ver += 1
        S.flash = f"Product catalog updated ({len(d)} items)."
        st.rerun()

    if not S.products.empty:
        export_row("Product Catalog", S.products, "Product_Catalog", "prd")


# =====================================================================
# PAGE: MASTER DATABASE (WITH MASS EDIT & SINGLE ROW DELETE)
# =====================================================================
def page_master() -> None:
    header("Master Database", "Direct tabular spreadsheet view of all raw records. Balance & Status recalculate on save.")
    
    # Quick single row manager
    with st.expander("🔍 Quick Row Selector (Inspect, Edit or Remove Single Record)", expanded=False):
        if not S.tx.empty:
            sel_tid = st.selectbox("Select Record", S.tx["Txn ID"].tolist(),
                                   format_func=lambda t: f"{t} - {S.tx.loc[S.tx['Txn ID']==t, 'Bill No'].iloc[0]} ({S.tx.loc[S.tx['Txn ID']==t, 'Party / Company'].iloc[0]})")
            m_row = S.tx[S.tx["Txn ID"] == sel_tid].iloc[0]
            st.write(f"**Customer:** {m_row['Party / Company']} &bull; **Bill Amount:** {money(m_row['Bill Amount'])} &bull; **Balance:** {money(m_row['Balance'])}")
            if st.button("🗑️ Delete Selected Record", key="btn_del_single_master"):
                db.delete_transaction_atomic(sel_tid)
                S.tx = S.tx[S.tx["Txn ID"] != sel_tid].reset_index(drop=True)
                S.line_items = S.line_items[S.line_items["Txn ID"] != sel_tid].reset_index(drop=True)
                db.log_activity(current_user()["username"], "DELETE_BILL", f"Deleted {sel_tid} from Master Database")
                S.flash = f"Deleted record {sel_tid}."
                st.rerun()

    q = st.text_input("Quick Filter", placeholder="Type party, bill no, product or status to narrow rows...")
    view = S.tx.copy()
    if q:
        view = view[view.astype(str).apply(lambda c: c.str.contains(q, case=False, na=False, regex=False)).any(axis=1)]

    view.insert(0, "Select", False)
    pm = sorted(set(PAYMENT_METHODS) | set(v for v in S.tx["Payment Method"] if v))
    cs = sorted(set(CHEQUE_STATUSES) | set(v for v in S.tx["Cheque Status"] if v))
    cfg = {c: st.column_config.NumberColumn(c, format="localized") for c in MONEY if c in view.columns}
    cfg.update({
        "Select": st.column_config.CheckboxColumn("Select", width="small"),
        "Date": st.column_config.DateColumn("Date", format="DD MMM YYYY"),
        "Cheque Date": st.column_config.DateColumn("Cheque Date", format="DD MMM YYYY"),
        "Payment Method": st.column_config.SelectboxColumn(options=pm),
        "Cheque Status": st.column_config.SelectboxColumn(options=cs),
    })

    ed = st.data_editor(view, hide_index=True, num_rows="fixed", key=f"mdb_{S.ver}_{q}", column_config=cfg,
                        disabled=["Txn ID", "Balance", "Status"], height=520, **_STRETCH)
    ids = ed.loc[ed["Select"], "Txn ID"].tolist()

    c1, c2, c3 = st.columns([1, 1, 2], vertical_alignment="bottom")
    if c1.button("Save Grid Edits", type="primary"):
        upd = ed.drop(columns=["Select", "Balance", "Status"]).set_index("Txn ID")
        base = S.tx.set_index("Txn ID")
        base.loc[upd.index, upd.columns] = upd
        S.tx = recompute(base.reset_index())
        save_all("transactions")
        db.log_activity(current_user()["username"], "EDIT_MASTER", "Saved modifications in Master Database grid")
        S.ver += 1
        S.flash = "Modifications saved successfully."
        st.rerun()

    ok = c3.checkbox(f"Confirm permanent deletion of {len(ids)} selected record(s)", disabled=not ids)
    if c2.button("Delete Selected", disabled=not (ids and ok)):
        for tid in ids:
            db.delete_transaction_atomic(tid)
        S.tx = S.tx[~S.tx["Txn ID"].isin(ids)].reset_index(drop=True)
        S.line_items = S.line_items[~S.line_items["Txn ID"].isin(ids)].reset_index(drop=True)
        db.log_activity(current_user()["username"], "DELETE_MASS", f"Deleted {len(ids)} transactions from Master Database")
        S.ver += 1
        S.flash = f"Permanently deleted {len(ids)} record(s)."
        st.rerun()

    export_row("Master Database", S.tx, "Master_Database", "mdb")


# =====================================================================
# PAGE: ACTIVITY LOG (AUDIT TRAIL)
# =====================================================================
def page_activity_log() -> None:
    header("Audit Trail & Activity Log", "Complete transparent log of all user activities, bill creations, edits and deletions.")
    logs_df = db.get_activity_logs(limit=250)

    if logs_df.empty:
        st.info("No activity logs recorded yet.")
        return

    q = st.text_input("Filter Activity Logs", placeholder="Search by user, action or keyword...")
    if q:
        logs_df = logs_df[logs_df.astype(str).apply(lambda c: c.str.contains(q, case=False, na=False, regex=False)).any(axis=1)]

    st.dataframe(logs_df, hide_index=True, **_STRETCH)
    export_row("Audit Trail", logs_df, "Audit_Activity_Log", "audit_log")


# =====================================================================
# PAGE: USER ACCOUNTS (ADMIN MANAGEMENT)
# =====================================================================
def page_users() -> None:
    header("User Accounts & Permissions", "Manage system users, assign administrative or staff roles, and update passwords.")
    if not is_admin():
        st.error("Access Restricted. Only administrators can view and manage user accounts.")
        return

    users_df = db.get_all_users()
    st.dataframe(users_df, hide_index=True, **_STRETCH)

    section("Create New User")
    with st.form("create_user_form", clear_on_submit=True):
        u1, u2, u3, u4 = st.columns(4)
        new_uname = u1.text_input("Username *", placeholder="e.g. staff1").strip().lower()
        new_fn = u2.text_input("Full Name", placeholder="e.g. Tariq Mehmood")
        new_pwd = u3.text_input("Password *", type="password")
        new_role = u4.selectbox("Role", ["Staff", "Admin"])

        if st.form_submit_button("Create User Account", type="primary"):
            success, msg = db.add_user(new_uname, new_pwd, new_fn, new_role)
            if success:
                db.log_activity(current_user()["username"], "CREATE_USER", f"Created user {new_uname} ({new_role})")
                S.flash = msg
                st.rerun()
            else:
                st.error(msg)

    section("Change User Password")
    with st.form("chg_pwd_form", clear_on_submit=True):
        cp1, cp2 = st.columns(2)
        target_u = cp1.selectbox("Select User", users_df["username"].tolist())
        target_p = cp2.text_input("New Password", type="password")
        if st.form_submit_button("Update Password"):
            success, msg = db.update_user_password(target_u, target_p)
            if success:
                db.log_activity(current_user()["username"], "UPDATE_PASSWORD", f"Changed password for user {target_u}")
                S.flash = f"Password updated for {target_u}."
                st.rerun()
            else:
                st.error(msg)


# =====================================================================
# PAGE: IMPORT DATA & BACKUP
# =====================================================================
def page_import() -> None:
    header("Import & Backup", "Import records from Excel/CSV/PDF or export full system backups.")
    t1, t2, t3 = st.tabs(["Import Records", "Import Products", "Backup & Restore"])

    with t1:
        st.markdown("Upload files containing headers: **Party, Date, Bill No, Product, Qty, Rate, Bill Amount, Paid Amount, Return Amount, Payment Method, Remarks**.")
        tpl = _empty(["Date", "Bill No", "Party / Company", "Product / Item", "Qty", "Rate", "Bill Amount",
                      "Paid Amount", "Return Amount", "Payment Method", "Cheque No", "Cheque Date", "Bank Name",
                      "Cheque Status", "Sale Base", "Remarks"])
        st.download_button("Download Blank Excel Template", to_excel_bytes({"Transactions": tpl}), "Import_Template.xlsx", XLSX, key="tpl_btn")

        f = st.file_uploader("Upload Transaction File", type=["xlsx", "xls", "csv", "pdf"], key=f"imp_{S.ver}")
        if f:
            try:
                d, used, ignored = prepare_import(read_upload(f))
            except Exception as exc:
                st.error(str(exc))
                return
            new, dups = split_duplicates(d, S.tx)
            st.success(f"Recognized columns: {', '.join(used)}")
            if ignored:
                st.caption(f"Ignored columns: {', '.join(ignored)}")
            c1, c2, c3 = st.columns(3)
            c1.metric("Rows Found", len(d))
            c2.metric("New Records", len(new))
            c3.metric("Duplicates Skipped", dups)
            show(new.head(50), [x for x in COLUMNS if x in new.columns and x not in ("Balance", "Status", "Txn ID")], height=280)
            addp = st.checkbox("Add new products found in file to catalog", value=True)
            if st.button(f"Import {len(new)} Record(s)", type="primary", disabled=new.empty):
                S.tx, S.line_items, S.products, n = commit_import(new, addp, S.tx, S.line_items, S.products)
                save_all()
                db.log_activity(current_user()["username"], "IMPORT", f"Imported {n} transaction records")
                S.ver += 1
                S.flash = f"Imported {n} record(s) successfully."
                st.rerun()

    with t2:
        st.markdown("Upload a list with **Product** (or Item), optional **Unit** and **Rate**.")
        f = st.file_uploader("Upload Product File", type=["xlsx", "xls", "csv", "pdf"], key=f"impp_{S.ver}")
        if f and st.button("Import Products", type="primary"):
            try:
                S.products, n = import_products(read_upload(f), S.products)
                save_all("products")
                db.log_activity(current_user()["username"], "IMPORT_PRODUCTS", f"Imported {n} products")
                S.ver += 1
                S.flash = f"{n} product(s) imported."
                st.rerun()
            except Exception as exc:
                st.error(str(exc))

    with t3:
        st.markdown("Download a full backup regularly. It contains all bills, items, products and parties in one Excel workbook.")
        st.download_button("📥 Download Master Backup (Excel)", to_excel_bytes(workbook_sheets()),
                           f"Business_Backup_{date.today():%Y%m%d}.xlsx", XLSX, key="master_bk_btn")

        if is_admin():
            section("Restore from Backup File")
            st.warning("Restoring will replace all existing transactions, items, products and parties with the contents of the backup file.")
            rf = st.file_uploader("Upload Backup File (.xlsx)", type=["xlsx"], key=f"rest_{S.ver}")
            ok = st.checkbox("I understand and confirm that all current data will be overwritten")
            if rf and ok and st.button("Restore Backup Now", type="primary"):
                try:
                    sh = pd.read_excel(rf, sheet_name=None)
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
                    db.log_activity(current_user()["username"], "RESTORE_BACKUP", "Restored master system backup")
                    S.ver += 1
                    S.flash = "System backup restored successfully."
                    st.rerun()
                except Exception as exc:
                    st.error(f"Invalid backup file: {exc}")


# =====================================================================
# PAGE: REPORTS & ANALYTICS
# =====================================================================
def page_reports() -> None:
    header("Financial Reports & Business Analytics", "Comprehensive financial statements, aging analysis, collections and master workbooks.")
    df = S.tx
    t_stmt, t_age, t_out, t_sales, t_book = st.tabs([
        "Party Statement", "Aged Receivables", "Outstanding Summary", "Sales Breakdown", "Complete Business Workbook"
    ])

    with t_stmt:
        section("Customer Statement of Account")
        c_top = st.columns([2, 1, 1])
        sel = party_picker("Party / Company", "rep", allow_all=True)
        d1 = c_top[1].date_input("From Date", value=None, key="rp_f")
        d2 = c_top[2].date_input("To Date", value=None, key="rp_t")

        st_df = lg.compute_party_statement(df, sel)
        if d1: st_df = st_df[st_df["Date"] >= pd.to_datetime(d1)]
        if d2: st_df = st_df[st_df["Date"] <= pd.to_datetime(d2)]

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Invoiced", money(st_df["Bill Amount"].sum()))
        c2.metric("Total Received", money(st_df["Paid Amount"].sum()))
        c3.metric("Total Returns", money(st_df["Return Amount"].sum()))
        net_bal = (st_df["Bill Amount"].sum() - st_df["Paid Amount"].sum() - st_df["Return Amount"].sum())
        c4.metric("Net Balance Due", money(net_bal))

        if st_df.empty:
            st.info("No records found for the selected party and date range.")
        else:
            show(st_df)
            safe_sel = re.sub(r'[^A-Za-z0-9_-]+', '_', str(sel or 'all')).strip('_') or 'all'
            export_row(f"Statement - {sel}", st_df, f"Statement_{safe_sel}", "stmt")

    with t_age:
        section("Aged Receivables (Aging Report)")
        aging_df = lg.compute_aging(df)
        if aging_df.empty:
            st.success("No outstanding balances at this time! All invoices are settled.")
        else:
            ca = st.columns(4)
            ca[0].metric("Current (0-30 days)", money(aging_df["Current (0-30d)"].sum()))
            ca[1].metric("31-60 Days", money(aging_df["31-60 Days"].sum()))
            ca[2].metric("61-90 Days", money(aging_df["61-90 Days"].sum()))
            ca[3].metric("Over 90 Days", money(aging_df["90+ Days"].sum()))
            show(aging_df)
            export_row("Aged Receivables", aging_df, "Aged_Receivables_Report", "age")

    with t_out:
        section("Outstanding Debtors Summary")
        out = df[df["Balance"] > 0].sort_values("Balance", ascending=False)
        if out.empty:
            st.success("No outstanding balances! All invoices are settled.")
        else:
            by_party = out.groupby("Party / Company")[["Bill Amount", "Paid Amount", "Balance"]].sum() \
                .sort_values("Balance", ascending=False).reset_index()
            c_out = st.columns(3)
            c_out[0].metric("Parties with Balance", len(by_party))
            c_out[1].metric("Total Pending Balance", money(by_party["Balance"].sum()))
            c_out[2].metric("Highest Individual Balance", money(by_party["Balance"].max()))
            show(by_party)
            export_row("Outstanding by Party", by_party, "Outstanding_By_Party", "outp")

    with t_sales:
        section("Revenue & Payment Methods")
        if df.empty:
            st.info("No sales records available.")
        else:
            sc1, sc2 = st.columns(2)
            with sc1:
                st.subheader("Monthly Billed vs Received")
                m = df.dropna(subset=["Date"]).assign(Month=lambda x: x["Date"].dt.strftime("%Y-%m"))
                if not m.empty:
                    m_grp = m.groupby("Month")[["Bill Amount", "Paid Amount"]].sum()
                    st.bar_chart(m_grp, color=["#1E3A8A", "#60A5FA"])
            with sc2:
                st.subheader("Collections by Payment Method")
                pm_grp = df.groupby("Payment Method")["Paid Amount"].sum()
                if not pm_grp.empty and pm_grp.sum() > 0:
                    st.bar_chart(pm_grp, color="#2563EB")

    with t_book:
        section("Complete Business Master Workbook")
        st.caption("One master Excel file containing all data sheets: Transactions, Outstanding, Aged Receivables, Cheques, Returns and Items.")
        out_full = df[df["Balance"] > 0].sort_values("Balance", ascending=False)
        aging_full = lg.compute_aging(df)
        st.download_button("Download Complete Master Workbook (Excel)", to_excel_bytes({
            "Transactions": df,
            "Outstanding": out_full,
            "Aged Receivables": aging_full,
            "Cheques": df[df["Payment Method"] == "Cheque"],
            "Returns": df[df["Return Amount"] > 0],
            "Items": S.line_items,
        }), f"Business_Master_Report_{date.today():%Y%m%d}.xlsx", XLSX, key="full_wb_btn")


# =====================================================================
# ROUTING & ENTRYPOINT
# =====================================================================
PAGES = {
    "Dashboard": page_dashboard,
    "New Transaction": page_new_transaction,
    "Bills & Payments": page_bills,
    "Cheque Tracker": page_cheques,
    "Party Ledger": page_ledger,
    "Parties Directory": page_parties,
    "Search Center": page_search,
    "Products": page_products,
    "Master Database": page_master,
    "Reports": page_reports,
    "Activity Log": page_activity_log,
    "User Accounts": page_users,
    "Import Data": page_import,
}

# 1. Require user login
require_login()

# 2. Load live data from database
try:
    S.tx, S.line_items, S.products, S.extra_parties = db.load_all()
except Exception as exc:
    st.error("Could not connect to the database.")
    st.caption(f"Technical detail: {type(exc).__name__}: {str(exc)[:300]}")
    st.stop()

# 3. Render sidebar navigation
sidebar()

# 4. Render selected page
page_func = PAGES.get(S.page, page_dashboard)
page_func()

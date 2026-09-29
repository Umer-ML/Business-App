# Enterprise Business & Accounts Management System (Cloud Ready)

Professional Streamlit Billing, Invoicing & Accounts Ledger app with PostgreSQL / SQLite storage.
Data is automatically synced with cloud database (Supabase / Neon / PostgreSQL) or local SQLite, ensuring zero data loss across redeployments or server restarts.

---

## 🌟 Key Features & Capabilities

1. **Executive Dashboard**:
   - Live KPI cards: Total Billed, Total Received, Outstanding Receivables, Total Returns.
   - Live status counts: Unpaid, Partially Paid, Fully Paid, Cheques Pending.
   - Interactive charts: Top 5 Outstanding Debtors & Monthly Billed vs Received trend.
   - Recent Transactions table.

2. **New Transaction (Invoicing)**:
   - Dynamic line-item billing with instant auto-calculation.
   - Product auto-completion with default unit and unit price.
   - Quick inline product creation.
   - Real-time Discount (percentage or fixed amount), Tax/GST (%), Grand Total, and Balance Due summary card.
   - Payment method selection with dedicated Cheque Tracker fields (Cheque No, Bank, Cheque Date, Cheque Status).

3. **Bills & Payments**:
   - Search & open any bill to view line items.
   - One-click payment receipt logging with payment method and remarks.
   - Goods return recording with automatic balance adjustments.
   - Instant professional Invoice PDF download with item breakdown and payment status.

4. **Cheque Tracker**:
   - Status indicators: Pending, Cleared, Bounced.
   - Automated due date alerts: Overdue, Due Today, Due within 7 Days.
   - Quick one-click status update (`Cleared` or `Bounced`).

5. **Party Ledger (Customer Accounts)**:
   - Account overview for any customer or supplier.
   - Full transaction history with debit, credit, and net balance.
   - Side-by-side one-click download of Excel and PDF statements.

6. **Parties Directory (NEW)**:
   - Complete directory of all customers and suppliers.
   - Shows Total Invoiced, Total Paid, Current Outstanding Balance, Total Bills count, and Last Transaction date.
   - Quick registration of new parties.
   - Excel & PDF export.

7. **Search Center (Multi-filter Hub)**:
   - Filter by Party, Bill No, Product, Payment Status, Payment Method, Cheque Status, and Date Range.
   - Summary statistics of filtered results.
   - Download filtered results as Excel or PDF.

8. **Products Catalog**:
   - Direct editable product list with Unit and Default Rates.
   - Instant Excel export.

9. **Master Database**:
   - Central database viewer and editor.
   - Recalculates balance and payment status automatically upon saving.
   - Safe bulk deletion with checkbox confirmation.

10. **Import Data & Backups**:
    - Automatic recognition of column headers from Excel, CSV, or PDF.
    - Duplicate detection and prevention.
    - Downloadable Excel import template.
    - One-click complete workbook backup and restore.

11. **Reports & Analytics (NEW & ENHANCED)**:
    - **Tab 1: Party Statement of Account**: Filter by date range and party; displays running balance and account summary.
    - **Tab 2: Aged Receivables (Aging Report)**: Analyzes outstanding debts by age: Current (0-30 days), 31-60 days, 61-90 days, and 90+ days overdue.
    - **Tab 3: Outstanding Summary**: Grouped by debtor with highest single balance metric.
    - **Tab 4: Sales & Cheques Breakdown**: Monthly revenue comparison and payment mode share.
    - **Tab 5: Master Business Workbook**: Single-click download of all business sheets in one Excel file (Transactions, Outstanding, Aged Receivables, Cheques, Returns, Items).

---

## 📁 Repository Structure

| File / Folder | Purpose |
|---|---|
| `app.py` | Complete Streamlit web interface and 11 screens |
| `logic.py` | Accounting calculations, aging analysis, running balance, Excel & PDF generators |
| `db.py` | Database engine (SQLAlchemy) connecting to PostgreSQL / Supabase or SQLite |
| `requirements.txt` | Python packages required |
| `schema.sql` | PostgreSQL schema (auto-created on first run) |
| `.streamlit/config.toml` | UI Theme & production server configuration |
| `.streamlit/secrets.toml.example` | Template for environment secrets |
| `keepalive.py` | Automated ping script to keep free Supabase databases active |
| `.github/workflows/keepalive.yml` | GitHub Actions cron workflow to prevent Supabase auto-pause |
| `render.yaml` | 1-click deployment configuration for Render.com |
| `Procfile` | Startup command for Render, Railway, and Heroku |
| `Dockerfile` | Production Docker image definition |
| `tests/test_logic.py` | 7 automated unit tests covering calculations and exports |

---

## 🚀 Deployment Guide

### Option 1: Streamlit Community Cloud (Recommended - 100% Free & Fast)

1. Push your repository to GitHub (Private repository recommended).
2. Go to [share.streamlit.io](https://share.streamlit.io/) and log in with your GitHub account.
3. Click **New app** > Select your repository, Branch `main`, Main file path `app.py`.
4. Click **Advanced settings**:
   - Python version: **3.10** or **3.12**
   - In the **Secrets** text box, paste:
     ```toml
     APP_PASSWORD = "YourSecurePasswordHere"
     DATABASE_URL = "postgresql://postgres.xxxx:your_password@aws-0-xx.pooler.supabase.com:5432/postgres"
     ```
5. Click **Deploy**. In 2-3 minutes, your app is live on a permanent public URL (e.g. `https://your-business.streamlit.app`).

### Option 2: Render.com / Railway (Docker / PaaS)

1. Connect your GitHub repository to [Render.com](https://render.com/).
2. Select **Web Service** > choose repository.
3. Render automatically detects `render.yaml` or `Dockerfile`.
4. Add environment variables `DATABASE_URL` and `APP_PASSWORD`.
5. Deploy.

> **Note on Vercel**:
> Vercel is built for stateless, short-lived Serverless HTTP functions (10-15s timeout limit). Streamlit is a stateful Python server requiring persistent WebSockets for live interactivity. Deploying Streamlit on Streamlit Community Cloud or Render/Railway provides 100% native stability, unlimited sessions, and zero disconnections.

---

## 🗄️ Database Setup (Supabase / PostgreSQL)

1. Go to [supabase.com](https://supabase.com/) and create a free account.
2. Click **New project**, choose a name, select the nearest region, and set a **Database Password** (use only letters and numbers).
3. Click **Connect** (top bar) > choose **Session pooler** connection string.
4. Replace `[YOUR-PASSWORD]` with your database password.
5. Paste this connection string into `DATABASE_URL` in Streamlit Secrets.
6. The app automatically creates all tables, indexes, and security policies on first boot!

---

## 💻 Running Locally

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Start the application
streamlit run app.py
```
Without `DATABASE_URL`, the app automatically stores data in `local_data.db` (SQLite) on your machine.

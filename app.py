"""
FinSight — Multi-user financial tracker
========================================
Run with:  streamlit run app.py

First run creates finsight.db (SQLite) in this folder automatically.
Optional export features need: pip install openpyxl fpdf2
"""

import base64
import binascii
import hashlib
import html as html_lib
import os
import re
import sqlite3
import urllib.parse
from datetime import date, datetime, timedelta
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# Streamlit's canvas widgets (st.dataframe / st.data_editor / progress) take colours from the Streamlit
# THEME, not from CSS. .streamlit/config.toml is the reliable way to set it; this is a best-effort fallback.
# Old line block:
# for _k, _v in {"base": "light", "primaryColor": "#00D4FF", "backgroundColor": "#FFFFFF", "secondaryBackgroundColor": "#F8FAFC", "textColor": "#0F172A"}.items():

# New Refined Line:
try:
    for _k, _v in {"base": "light", "primaryColor": "#10B981", "backgroundColor": "#FFFFFF", "secondaryBackgroundColor": "#F4F7F6", "textColor": "#111C24"}.items():
        st._config.set_option(f"theme.{_k}", _v)
except Exception:
    pass
# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="FinSight",
    page_icon="◈",
    layout="wide",
    initial_sidebar_state="expanded",
)

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "finsight.db")

CURRENCIES = {
    "INR": "₹",
    "USD": "$",
    "EUR": "€",
    "GBP": "£",
}

PLAN_LIMITS = {
    "free": {"max_workspaces": 1, "max_transactions": 250, "exports": False, "invites": False},
    "pro": {"max_workspaces": 10, "max_transactions": None, "exports": True, "invites": True},
}

# ------------------------------------------------------------
# BUSINESS CATEGORY — fixed list, no free-form text.
# Each category maps to extra transaction categories that matter
# for that kind of business, plus one tailored KPI shown on the
# dashboard. "Other" is always the fallback.
# ------------------------------------------------------------

BUSINESS_CATEGORIES = [
    "Retail",
    "Restaurant / Food",
    "Manufacturing",
    "Wholesale / Distribution",
    "E-commerce",
    "Services",
    "Construction",
    "Healthcare",
    "Education",
    "Transport / Logistics",
    "Professional Services",
    "Agriculture",
    "Hospitality",
    "Beauty / Salon",
    "Repair / Maintenance",
    "Other",
]

CATEGORY_EXTRA_TAGS = {
    "Retail": ["Inventory", "Supplier Purchases", "Store Rent", "Stock Loss"],
    "Restaurant / Food": ["Food Cost", "Staff Wages", "Delivery", "Wastage"],
    "Manufacturing": ["Raw Materials", "Labor", "Equipment Maintenance", "Production Loss"],
    "Wholesale / Distribution": ["Bulk Purchases", "Warehousing", "Freight", "Supplier Purchases"],
    "E-commerce": ["Inventory", "Shipping", "Platform Fees", "Returns"],
    "Services": ["Labor", "Client Delivery Cost", "Subcontractors"],
    "Construction": ["Materials", "Labor", "Equipment Rental", "Site Costs"],
    "Healthcare": ["Medical Supplies", "Staff Wages", "Equipment Maintenance"],
    "Education": ["Staff Wages", "Learning Materials", "Facility Costs"],
    "Transport / Logistics": ["Fuel", "Vehicle Maintenance", "Freight", "Driver Wages"],
    "Professional Services": ["Labor", "Client Delivery Cost", "Subcontractors"],
    "Agriculture": ["Seeds & Inputs", "Labor", "Equipment Maintenance", "Storage"],
    "Hospitality": ["Staff Wages", "Housekeeping", "Food Cost", "Maintenance"],
    "Beauty / Salon": ["Product Stock", "Staff Wages", "Rent"],
    "Repair / Maintenance": ["Spare Parts", "Labor", "Equipment"],
    "Other": [],
}


def compute_category_kpi(business_category, tx_df, income_total):
    """Returns (label, value_str, sub_text) for a business-category-specific KPI, or None."""
    if tx_df.empty:
        return None

    expense_df = tx_df[tx_df["type"] == "expense"]

    def cost_pct(tag_categories, label):
        cost = expense_df[expense_df["category"].isin(tag_categories)]["amount"].sum()
        if income_total <= 0:
            return (label, "—", f"Add revenue to calculate {label.lower()}.")
        pct = cost / income_total * 100
        return (label, f"{pct:.1f}%", f"{cost:,.0f} of revenue this period.")

    mapping = {
        "Restaurant / Food": (["Food Cost"], "Food Cost %"),
        "Hospitality": (["Food Cost"], "Food Cost %"),
        "Retail": (["Inventory", "Supplier Purchases"], "Cost of Goods %"),
        "E-commerce": (["Inventory"], "Cost of Goods %"),
        "Wholesale / Distribution": (["Bulk Purchases", "Supplier Purchases"], "Cost of Goods %"),
        "Manufacturing": (["Raw Materials"], "Raw Material Cost %"),
        "Transport / Logistics": (["Fuel"], "Fuel Cost %"),
        "Construction": (["Materials"], "Materials Cost %"),
        "Agriculture": (["Seeds & Inputs"], "Input Cost %"),
        "Beauty / Salon": (["Product Stock"], "Product Cost %"),
        "Repair / Maintenance": (["Spare Parts"], "Parts Cost %"),
    }

    if business_category not in mapping:
        return None

    tags, label = mapping[business_category]
    return cost_pct(tags, label)

# ============================================================
# PREVENTING REVENUE LEAKS & FUTURE TAX PROJECTIONS
# ============================================================

def forecast_tax_liability(tx_df, gst_rate=18.0):
    """Forecasts upcoming tax liability based on the run-rate of past months."""
    if tx_df.empty or len(tx_df) < 5:
        return None
        
    df = tx_df.copy()
    latest_date = df["date"].max()
    ninety_days_ago = latest_date - pd.Timedelta(days=90)
    recent_df = df[df["date"] >= ninety_days_ago]
    
    # Calculate monthly averages over the last 90 days
    inc_avg = recent_df[recent_df["type"] == "income"]["amount"].sum() / 3.0
    exp_avg = recent_df[recent_df["type"] == "expense"]["amount"].sum() / 3.0
    
    # Estimate monthly GST components
    projected_monthly_gst_collected = inc_avg * (gst_rate / (100 + gst_rate))
    projected_monthly_gst_paid = exp_avg * (gst_rate / (100 + gst_rate))
    net_monthly_gst = max(0.0, projected_monthly_gst_collected - projected_monthly_gst_paid)
    
    # Estimate annual business income tax run-rate (Simplified Indian tax slab reference)
    projected_annual_profit = (inc_avg - exp_avg) * 12
    estimated_income_tax = 0.0
    if projected_annual_profit > 700000:  # New Tax Regime threshold benchmark
        estimated_income_tax = projected_annual_profit * 0.15  # average effective slab estimation
        
    return {
        "monthly_gst_forecast": net_monthly_gst,
        "quarterly_gst_forecast": net_monthly_gst * 3,
        "estimated_annual_income_tax": max(0.0, estimated_income_tax),
        "run_rate_health": "Stable" if inc_avg > exp_avg else "Risk Zone"
    }

def profile_database_leaks(workspace_id):
    """Scans the transaction history for double-entries and pattern leaks."""
    conn = get_conn()
    # Query to catch identical amounts and descriptions logged within 1 day of each other
    query = """
        SELECT t1.date as date1, t2.date as date2, t1.description, t1.amount 
        FROM transactions t1
        JOIN transactions t2 ON t1.workspace_id = t2.workspace_id 
            AND t1.amount = t2.amount 
            AND t1.description = t2.description 
            AND t1.id < t2.id
        WHERE t1.workspace_id = ? AND t1.deleted = 0 AND t2.deleted = 0
            AND abs(julianday(t1.date) - julianday(t2.date)) <= 1
    """
    duplicates = conn.execute(query, (workspace_id,)).fetchall()
    conn.close()
    return [dict(d) for d in duplicates]

# ============================================================
# DATABASE LAYER
# ============================================================

def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_conn()
    c = conn.cursor()

    c.execute("""CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        salt TEXT NOT NULL,
        password_hash TEXT NOT NULL,
        plan TEXT DEFAULT 'free',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )""")

    c.execute("""CREATE TABLE IF NOT EXISTS workspaces (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        currency TEXT DEFAULT 'INR',
        gst_rate REAL DEFAULT 18.0,
        business_category TEXT DEFAULT 'Other',
        owner_id INTEGER NOT NULL,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )""")

    # Migration: older databases won't have business_category yet.
    existing_cols = {row["name"] for row in c.execute("PRAGMA table_info(workspaces)").fetchall()}
    if "business_category" not in existing_cols:
        c.execute("ALTER TABLE workspaces ADD COLUMN business_category TEXT DEFAULT 'Other'")

    c.execute("""CREATE TABLE IF NOT EXISTS workspace_members (
        workspace_id INTEGER NOT NULL,
        user_id INTEGER NOT NULL,
        role TEXT DEFAULT 'owner',
        PRIMARY KEY (workspace_id, user_id)
    )""")

    c.execute("""CREATE TABLE IF NOT EXISTS transactions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        workspace_id INTEGER NOT NULL,
        date TEXT NOT NULL,
        description TEXT,
        amount REAL,
        type TEXT,
        employee TEXT,
        category TEXT,
        gst_applicable INTEGER DEFAULT 0,
        invoice_status TEXT DEFAULT '',
        deleted INTEGER DEFAULT 0,
        created_by INTEGER,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )""")

    # Migration: older databases won't have customer contact fields yet.
    existing_tx_cols = {row["name"] for row in c.execute("PRAGMA table_info(transactions)").fetchall()}
    if "customer_name" not in existing_tx_cols:
        c.execute("ALTER TABLE transactions ADD COLUMN customer_name TEXT DEFAULT ''")
    if "customer_phone" not in existing_tx_cols:
        c.execute("ALTER TABLE transactions ADD COLUMN customer_phone TEXT DEFAULT ''")

    c.execute("""CREATE TABLE IF NOT EXISTS recurring_rules (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        workspace_id INTEGER NOT NULL,
        description TEXT,
        amount REAL,
        type TEXT,
        category TEXT,
        employee TEXT,
        frequency TEXT,
        next_run_date TEXT,
        active INTEGER DEFAULT 1
    )""")

    c.execute("""CREATE TABLE IF NOT EXISTS budgets (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        workspace_id INTEGER NOT NULL,
        category TEXT,
        monthly_limit REAL,
        UNIQUE(workspace_id, category)
    )""")

    c.execute("""CREATE TABLE IF NOT EXISTS goals (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        workspace_id INTEGER NOT NULL,
        name TEXT,
        target_amount REAL,
        target_date TEXT,
        saved_amount REAL DEFAULT 0
    )""")

    c.execute("""CREATE TABLE IF NOT EXISTS category_rules (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        workspace_id INTEGER NOT NULL,
        keyword TEXT,
        category TEXT,
        UNIQUE(workspace_id, keyword)
    )""")

    conn.commit()
    conn.close()


init_db()


# ============================================================
# AUTH HELPERS
# ============================================================

def hash_password(password, salt=None):
    if salt is None:
        salt = os.urandom(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 100_000)
    return binascii.hexlify(salt).decode(), binascii.hexlify(dk).decode()


def verify_password(password, salt_hex, hash_hex):
    salt = binascii.unhexlify(salt_hex)
    _, dk_hex = hash_password(password, salt)
    return dk_hex == hash_hex


def create_user(username, password):
    conn = get_conn()
    salt_hex, hash_hex = hash_password(password)
    try:
        conn.execute(
            "INSERT INTO users (username, salt, password_hash) VALUES (?,?,?)",
            (username.strip(), salt_hex, hash_hex),
        )
        conn.commit()
        return True, "Account created."
    except sqlite3.IntegrityError:
        return False, "That username is already taken."
    finally:
        conn.close()


def authenticate(username, password):
    conn = get_conn()
    row = conn.execute("SELECT * FROM users WHERE username = ?", (username.strip(),)).fetchone()
    conn.close()
    if row is None:
        return None
    if verify_password(password, row["salt"], row["password_hash"]):
        return dict(row)
    return None


def get_user_by_id(user_id):
    conn = get_conn()
    row = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def set_user_plan(user_id, plan):
    conn = get_conn()
    conn.execute("UPDATE users SET plan = ? WHERE id = ?", (plan, user_id))
    conn.commit()
    conn.close()


# ============================================================
# WORKSPACE HELPERS
# ============================================================

def get_user_workspaces(user_id):
    conn = get_conn()
    rows = conn.execute(
        """SELECT w.*, m.role FROM workspaces w
           JOIN workspace_members m ON m.workspace_id = w.id
           WHERE m.user_id = ?
           ORDER BY w.created_at""",
        (user_id,),
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def create_workspace(user_id, name, currency, gst_rate, business_category="Other"):
    conn = get_conn()
    cur = conn.execute(
        "INSERT INTO workspaces (name, currency, gst_rate, business_category, owner_id) VALUES (?,?,?,?,?)",
        (name.strip(), currency, gst_rate, business_category, user_id),
    )
    workspace_id = cur.lastrowid
    conn.execute(
        "INSERT INTO workspace_members (workspace_id, user_id, role) VALUES (?,?, 'owner')",
        (workspace_id, user_id),
    )
    conn.commit()
    conn.close()
    return workspace_id


def invite_member(workspace_id, username, role):
    conn = get_conn()
    user_row = conn.execute("SELECT id FROM users WHERE username = ?", (username.strip(),)).fetchone()
    if user_row is None:
        conn.close()
        return False, "No user with that username exists yet — ask them to sign up first."
    try:
        conn.execute(
            "INSERT INTO workspace_members (workspace_id, user_id, role) VALUES (?,?,?)",
            (workspace_id, user_row["id"], role),
        )
        conn.commit()
        return True, f"{username} added as {role}."
    except sqlite3.IntegrityError:
        return False, "That user is already a member of this workspace."
    finally:
        conn.close()


def get_workspace_members(workspace_id):
    conn = get_conn()
    rows = conn.execute(
        """SELECT u.username, m.role FROM workspace_members m
           JOIN users u ON u.id = m.user_id
           WHERE m.workspace_id = ?""",
        (workspace_id,),
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


# ============================================================
# TRANSACTION HELPERS
# ============================================================

def load_transactions(workspace_id):
    conn = get_conn()
    df = pd.read_sql_query(
        "SELECT * FROM transactions WHERE workspace_id = ? AND deleted = 0 ORDER BY date",
        conn,
        params=(workspace_id,),
    )
    conn.close()
    if df.empty:
        return df
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0.0).abs()
    return df


def insert_transaction(workspace_id, tx_date, description, amount, tx_type, employee,
                        category, gst_applicable, invoice_status, created_by,
                        customer_name="", customer_phone=""):
    conn = get_conn()
    conn.execute(
        """INSERT INTO transactions
           (workspace_id, date, description, amount, type, employee, category,
            gst_applicable, invoice_status, created_by, customer_name, customer_phone)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
        (
            workspace_id,
            pd.to_datetime(tx_date).strftime("%Y-%m-%d"),
            description.strip(),
            abs(float(amount)),
            tx_type,
            employee.strip(),
            category,
            int(gst_applicable),
            invoice_status,
            created_by,
            customer_name.strip(),
            customer_phone.strip(),
        ),
    )
    conn.commit()
    conn.close()


def mark_invoice_paid(tx_id):
    conn = get_conn()
    conn.execute("UPDATE transactions SET invoice_status = 'Paid' WHERE id = ?", (tx_id,))
    conn.commit()
    conn.close()


def get_unpaid_invoices(workspace_id):
    conn = get_conn()
    df = pd.read_sql_query(
        """SELECT * FROM transactions
           WHERE workspace_id = ? AND deleted = 0 AND invoice_status = 'Invoiced — unpaid'
           ORDER BY date""",
        conn, params=(workspace_id,),
    )
    conn.close()
    if not df.empty:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
    return df


def update_transaction(tx_id, tx_date, description, amount, tx_type, employee, category):
    conn = get_conn()
    conn.execute(
        """UPDATE transactions SET date=?, description=?, amount=?, type=?, employee=?, category=?
           WHERE id=?""",
        (
            pd.to_datetime(tx_date).strftime("%Y-%m-%d"),
            description,
            abs(float(amount)),
            tx_type,
            employee,
            category,
            tx_id,
        ),
    )
    conn.commit()
    conn.close()


def soft_delete_transaction(tx_id):
    conn = get_conn()
    conn.execute("UPDATE transactions SET deleted = 1 WHERE id = ?", (tx_id,))
    conn.commit()
    conn.close()


def restore_last_deleted(workspace_id):
    conn = get_conn()
    row = conn.execute(
        "SELECT id FROM transactions WHERE workspace_id = ? AND deleted = 1 ORDER BY id DESC LIMIT 1",
        (workspace_id,),
    ).fetchone()
    if row:
        conn.execute("UPDATE transactions SET deleted = 0 WHERE id = ?", (row["id"],))
        conn.commit()
    conn.close()
    return row is not None


def get_deleted_transactions(workspace_id, limit=5):
    conn = get_conn()
    df = pd.read_sql_query(
        "SELECT * FROM transactions WHERE workspace_id = ? AND deleted = 1 ORDER BY id DESC LIMIT ?",
        conn,
        params=(workspace_id, limit),
    )
    conn.close()
    return df


def transaction_count(workspace_id):
    conn = get_conn()
    n = conn.execute(
        "SELECT COUNT(*) AS n FROM transactions WHERE workspace_id = ? AND deleted = 0", (workspace_id,)
    ).fetchone()["n"]
    conn.close()
    return n


# ============================================================
# CATEGORY RULES (learned + default keyword rules + business-category tags)
# ============================================================

DEFAULT_CATEGORY_RULES = {
    "salary": "Salaries", "wage": "Salaries",
    "supplier": "Suppliers", "inventory": "Suppliers",
    "rent": "Rent",
    "electricity": "Electricity", "power": "Electricity",
    "internet": "Internet",
    "marketing": "Marketing", "advertising": "Marketing", "ads": "Marketing",
    "transport": "Transport", "fuel": "Transport", "travel": "Transport",
    "software": "Software", "subscription": "Software",
    "equipment": "Equipment",
    "office": "Office",
    "food": "Food",
    "tax": "Tax",
    "insurance": "Insurance",
}

BASE_CATEGORIES = sorted(set(DEFAULT_CATEGORY_RULES.values()) | {"Other", "Revenue"})


def get_workspace_categories(business_category):
    """Base category list plus business-category-specific tags, deduplicated."""
    extra = CATEGORY_EXTRA_TAGS.get(business_category, [])
    return sorted(set(BASE_CATEGORIES) | set(extra))


def get_learned_rules(workspace_id):
    conn = get_conn()
    rows = conn.execute(
        "SELECT keyword, category FROM category_rules WHERE workspace_id = ?", (workspace_id,)
    ).fetchall()
    conn.close()
    return {r["keyword"]: r["category"] for r in rows}


def learn_category(workspace_id, description, category):
    keyword = str(description).strip().lower()
    if not keyword:
        return
    conn = get_conn()
    conn.execute(
        """INSERT INTO category_rules (workspace_id, keyword, category) VALUES (?,?,?)
           ON CONFLICT(workspace_id, keyword) DO UPDATE SET category = excluded.category""",
        (workspace_id, keyword, category),
    )
    conn.commit()
    conn.close()


def category_from_text(text, learned_rules):
    raw = str(text).strip().lower()

    if raw in learned_rules:
        return learned_rules[raw]

    for keyword, category in DEFAULT_CATEGORY_RULES.items():
        if keyword in raw:
            return category

    return "Other"


# ============================================================
# RECURRING TRANSACTIONS
# ============================================================

def get_recurring_rules(workspace_id):
    conn = get_conn()
    df = pd.read_sql_query(
        "SELECT * FROM recurring_rules WHERE workspace_id = ? ORDER BY id", conn, params=(workspace_id,)
    )
    conn.close()
    return df


def add_recurring_rule(workspace_id, description, amount, tx_type, category, employee, frequency, start_date):
    conn = get_conn()
    conn.execute(
        """INSERT INTO recurring_rules
           (workspace_id, description, amount, type, category, employee, frequency, next_run_date, active)
           VALUES (?,?,?,?,?,?,?,?,1)""",
        (workspace_id, description.strip(), abs(float(amount)), tx_type, category, employee.strip(),
         frequency, pd.to_datetime(start_date).strftime("%Y-%m-%d")),
    )
    conn.commit()
    conn.close()


def toggle_recurring_rule(rule_id, active):
    conn = get_conn()
    conn.execute("UPDATE recurring_rules SET active = ? WHERE id = ?", (int(active), rule_id))
    conn.commit()
    conn.close()


def _advance_date(current, frequency):
    current = pd.to_datetime(current)
    if frequency == "Weekly":
        return current + pd.DateOffset(weeks=1)
    if frequency == "Monthly":
        return current + pd.DateOffset(months=1)
    if frequency == "Yearly":
        return current + pd.DateOffset(years=1)
    return current + pd.DateOffset(months=1)


def run_due_recurring_transactions(workspace_id, created_by):
    """Auto-generates transactions for any recurring rule whose next run date has passed."""
    rules = get_recurring_rules(workspace_id)
    if rules.empty:
        return 0

    today = pd.Timestamp(date.today())
    generated = 0

    for _, rule in rules[rules["active"] == 1].iterrows():
        next_run = pd.to_datetime(rule["next_run_date"])
        safety_counter = 0
        while next_run <= today and safety_counter < 36:
            insert_transaction(
                workspace_id,
                next_run,
                f"{rule['description']} (auto-recurring)",
                rule["amount"],
                rule["type"],
                rule["employee"] or "",
                rule["category"],
                0,
                "",
                created_by,
            )
            next_run = _advance_date(next_run, rule["frequency"])
            generated += 1
            safety_counter += 1

        conn = get_conn()
        conn.execute(
            "UPDATE recurring_rules SET next_run_date = ? WHERE id = ?",
            (next_run.strftime("%Y-%m-%d"), rule["id"]),
        )
        conn.commit()
        conn.close()

    return generated


# ============================================================
# BUDGETS
# ============================================================

def get_budgets(workspace_id):
    conn = get_conn()
    df = pd.read_sql_query("SELECT * FROM budgets WHERE workspace_id = ?", conn, params=(workspace_id,))
    conn.close()
    return df


def set_budget(workspace_id, category, monthly_limit):
    conn = get_conn()
    conn.execute(
        """INSERT INTO budgets (workspace_id, category, monthly_limit) VALUES (?,?,?)
           ON CONFLICT(workspace_id, category) DO UPDATE SET monthly_limit = excluded.monthly_limit""",
        (workspace_id, category, float(monthly_limit)),
    )
    conn.commit()
    conn.close()


# ============================================================
# GOALS
# ============================================================

def get_goals(workspace_id):
    conn = get_conn()
    df = pd.read_sql_query("SELECT * FROM goals WHERE workspace_id = ?", conn, params=(workspace_id,))
    conn.close()
    return df


def add_goal(workspace_id, name, target_amount, target_date):
    conn = get_conn()
    conn.execute(
        "INSERT INTO goals (workspace_id, name, target_amount, target_date) VALUES (?,?,?,?)",
        (workspace_id, name.strip(), float(target_amount), pd.to_datetime(target_date).strftime("%Y-%m-%d")),
    )
    conn.commit()
    conn.close()


def contribute_to_goal(goal_id, amount):
    conn = get_conn()
    conn.execute("UPDATE goals SET saved_amount = saved_amount + ? WHERE id = ?", (float(amount), goal_id))
    conn.commit()
    conn.close()


# ============================================================
# CSV IMPORT (bank-feed style, flexible column mapping)
# ============================================================

def guess_column(columns, candidates):
    lower_map = {c.lower(): c for c in columns}
    for cand in candidates:
        if cand in lower_map:
            return lower_map[cand]
    for c in columns:
        for cand in candidates:
            if cand in c.lower():
                return c
    return None


def build_stock_dataframe(view):
    graph_data = view.copy()
    graph_data["movement"] = graph_data.apply(
        lambda row: row["amount"] if row["type"] == "income" else -row["amount"], axis=1
    )
    daily = (
        graph_data.groupby(graph_data["date"].dt.date, as_index=False)
        .agg(movement=("movement", "sum"), volume=("amount", "sum"))
        .rename(columns={"date": "date"})
    )
    daily["date"] = pd.to_datetime(daily["date"])
    daily = daily.sort_values("date").reset_index(drop=True)

    opens, running = [], 0.0
    for m in daily["movement"]:
        opens.append(running)
        running += m
    daily["open"] = opens
    daily["close"] = daily["open"] + daily["movement"]

    pad = daily["movement"].abs().clip(lower=1) * 0.15
    daily["high"] = daily[["open", "close"]].max(axis=1) + pad
    daily["low"] = daily[["open", "close"]].min(axis=1) - pad
    return daily


def forecast_next_days(daily, days=7):
    if len(daily) < 3:
        return None
    x = np.arange(len(daily))
    y = daily["close"].values
    slope, intercept = np.polyfit(x, y, 1)
    future_x = np.arange(len(daily), len(daily) + days)
    future_y = slope * future_x + intercept
    last_date = daily["date"].max()
    future_dates = [last_date + pd.Timedelta(days=i + 1) for i in range(days)]
    return pd.DataFrame({"date": future_dates, "close": future_y})


def business_mood(income, expenses, profit, margin):
    if income == 0 and expenses == 0:
        return ("🧭", "Waiting for data", "Add your first transaction.", "neutral")
    if profit > 0 and margin >= 25:
        return ("🚀", "Excellent momentum", "Your business is generating a strong return.", "green")
    if profit > 0 and margin >= 10:
        return ("😎", "Healthy business", "Revenue is comfortably ahead of expenses.", "green")
    if profit >= 0:
        return ("😐", "Watch your margins", "You are profitable, but the buffer is thin.", "yellow")
    return ("📉", "Loss detected", "Expenses currently exceed revenue.", "red")


# ============================================================
# VNEXT INTELLIGENCE LAYER
# Deterministic analytics: anomaly detection, period-over-period
# "what & why" pulse, algorithm-driven health score, and an
# aggregated action center. All pure Python/Pandas over the
# operational dataframes already loaded for the workspace — no
# new tables, no external calls, defensive against empty data.
# ============================================================

def _period_key(ts, granularity):
    """Return a pandas Period for a timestamp, at Month or Quarter granularity."""
    return ts.to_period("Q") if granularity == "Quarter" else ts.to_period("M")

def safe_pdf_str(text):
    """Replaces unsupported currency symbols and dashes with standard ASCII text."""
    return str(text).replace('—', '-').replace('₹', 'Rs. ').replace('€', 'EUR ').replace('£', 'GBP ')

def detect_anomalies(tx_df, budgets_df):
    """Scans transactions (+ budgets) for the four VNext anomaly signals.
    Returns a list of dicts: {type, severity ('red'/'yellow'/'green'), message}.
    Always defensive: any missing/thin data simply skips that check."""
    anomalies = []
    if tx_df is None or tx_df.empty:
        return anomalies

    df = tx_df.copy()
    if df["date"].isna().all():
        return anomalies

    latest_date = df["date"].max()
    expense_df = df[df["type"] == "expense"]

    # ---- 1. Spike Detection ----
    # Compare each category's most recent transaction against that
    # category's trailing distribution (median + std, and median * 1.6).
    if not expense_df.empty:
        for category, group in expense_df.groupby("category"):
            group = group.sort_values("date")
            if len(group) < 4:
                continue  # not enough history to call anything a "spike"
            history = group.iloc[:-1]
            latest_tx = group.iloc[-1]
            median = history["amount"].median()
            std = history["amount"].std()
            std = 0.0 if pd.isna(std) else std
            if median <= 0:
                continue
            std_threshold = median + 1.5 * std
            pct_threshold = median * 1.6
            if latest_tx["amount"] > std_threshold or latest_tx["amount"] > pct_threshold:
                anomalies.append({
                    "type": "spike",
                    "severity": "red",
                    "message": (
                        f"'{category}' expense of {latest_tx['amount']:,.0f} on "
                        f"{latest_tx['date'].strftime('%d %b')} is well above the typical "
                        f"{median:,.0f} for this category."
                    ),
                })

    # ---- 2. Margin Collapse Watcher ----
    try:
        this_month = df[(df["date"].dt.year == latest_date.year) & (df["date"].dt.month == latest_date.month)]
        prev_month_anchor = latest_date.replace(day=1) - pd.Timedelta(days=1)
        prev_month = df[(df["date"].dt.year == prev_month_anchor.year) & (df["date"].dt.month == prev_month_anchor.month)]

        def _margin(d):
            inc = d.loc[d["type"] == "income", "amount"].sum()
            exp = d.loc[d["type"] == "expense", "amount"].sum()
            if inc <= 0:
                return None
            return (inc - exp) / inc * 100

        cur_margin = _margin(this_month)
        prev_margin = _margin(prev_month)
        if cur_margin is not None and prev_margin is not None and (prev_margin - cur_margin) > 5:
            anomalies.append({
                "type": "margin_collapse",
                "severity": "red",
                "message": (
                    f"Net profit margin shrank from {prev_margin:.1f}% last month to "
                    f"{cur_margin:.1f}% this month."
                ),
            })
    except Exception:
        pass

    # ---- 3. Budget Burn Rate Alerts ----
    try:
        if budgets_df is not None and not budgets_df.empty:
            days_in_month = (pd.Timestamp(latest_date.year, latest_date.month, 1) + pd.offsets.MonthEnd(0)).day
            month_fraction = (latest_date.day / days_in_month) if days_in_month else 0
            this_month_expenses = df[
                (df["date"].dt.year == latest_date.year) & (df["date"].dt.month == latest_date.month) & (df["type"] == "expense")
            ]
            spend_by_cat = this_month_expenses.groupby("category")["amount"].sum()
            for _, b in budgets_df.iterrows():
                limit = b["monthly_limit"]
                if not limit or limit <= 0:
                    continue
                spent = spend_by_cat.get(b["category"], 0.0)
                consumed_pct = spent / limit
                if consumed_pct >= 0.8 and month_fraction < 0.5:
                    anomalies.append({
                        "type": "budget_burn",
                        "severity": "red",
                        "message": (
                            f"'{b['category']}' budget is already {consumed_pct * 100:.0f}% spent, but only "
                            f"{month_fraction * 100:.0f}% of the month has passed."
                        ),
                    })
    except Exception:
        pass

    # ---- 4. Velocity Warnings ----
    try:
        expense_only = df[df["type"] == "expense"].copy()
        if not expense_only.empty:
            expense_only["week"] = expense_only["date"].dt.to_period("W")
            weekly = expense_only.groupby("week")["amount"].sum().sort_index()
            if len(weekly) >= 3:
                current_week_total = weekly.iloc[-1]
                history = weekly.iloc[max(0, len(weekly) - 7):-1]
                if len(history) >= 2:
                    rolling_avg = history.mean()
                    if rolling_avg > 0 and current_week_total > rolling_avg * 1.4:
                        anomalies.append({
                            "type": "velocity",
                            "severity": "yellow",
                            "message": (
                                f"This week's expenses ({current_week_total:,.0f}) are "
                                f"{(current_week_total / rolling_avg - 1) * 100:.0f}% above the "
                                f"{len(history)}-week rolling average ({rolling_avg:,.0f})."
                            ),
                        })
    except Exception:
        pass

    return anomalies


def compute_business_pulse(tx_df, granularity="Month"):
    """Period-over-period 'What & Why' pipeline.
    Returns None if there isn't enough data; otherwise a dict with
    'what' (directional movement + values) and 'why' (top contributors
    to the expense change), for Month-vs-previous-Month or
    Quarter-vs-previous-Quarter depending on `granularity`."""
    if tx_df is None or tx_df.empty:
        return None

    df = tx_df.copy()
    if df["date"].isna().all():
        return None

    latest = df["date"].max()
    df["period"] = df["date"].apply(lambda d: _period_key(d, granularity))
    current_period = _period_key(latest, granularity)
    previous_period = current_period - 1

    current_df = df[df["period"] == current_period]
    previous_df = df[df["period"] == previous_period]

    if current_df.empty and previous_df.empty:
        return None

    def _summary(d):
        inc = d.loc[d["type"] == "income", "amount"].sum()
        exp = d.loc[d["type"] == "expense", "amount"].sum()
        profit = inc - exp
        margin = (profit / inc * 100) if inc > 0 else 0.0
        return inc, exp, profit, margin

    cur_inc, cur_exp, cur_profit, cur_margin = _summary(current_df)
    prev_inc, prev_exp, prev_profit, prev_margin = _summary(previous_df)

    def _direction(cur, prev):
        if prev == 0:
            return "→" if cur == 0 else "↑"
        change = (cur - prev) / abs(prev) * 100
        if change > 1:
            return "↑"
        if change < -1:
            return "↓"
        return "→"

    what = {
        "revenue": {"current": cur_inc, "previous": prev_inc, "direction": _direction(cur_inc, prev_inc)},
        "expenses": {"current": cur_exp, "previous": prev_exp, "direction": _direction(cur_exp, prev_exp)},
        "profit": {"current": cur_profit, "previous": prev_profit, "direction": _direction(cur_profit, prev_profit)},
        "margin": {"current": cur_margin, "previous": prev_margin, "direction": _direction(cur_margin, prev_margin)},
    }

    # ---- WHY: which categories drove the expense move ----
    why = []
    expense_change_abs = cur_exp - prev_exp
    if abs(expense_change_abs) > 0:
        cur_by_cat = current_df[current_df["type"] == "expense"].groupby("category")["amount"].sum()
        prev_by_cat = previous_df[previous_df["type"] == "expense"].groupby("category")["amount"].sum()
        all_cats = set(cur_by_cat.index) | set(prev_by_cat.index)
        deltas = {cat: cur_by_cat.get(cat, 0.0) - prev_by_cat.get(cat, 0.0) for cat in all_cats}

        if expense_change_abs > 0:
            contributors = {k: v for k, v in deltas.items() if v > 0}
        else:
            contributors = {k: v for k, v in deltas.items() if v < 0}

        total_contrib = sum(abs(v) for v in contributors.values())
        if total_contrib > 0:
            for cat, delta in sorted(contributors.items(), key=lambda kv: abs(kv[1]), reverse=True)[:3]:
                why.append({
                    "category": cat,
                    "amount": delta,
                    "pct_of_change": abs(delta) / total_contrib * 100,
                })

    return {
        "granularity": granularity,
        "current_period": str(current_period),
        "previous_period": str(previous_period),
        "what": what,
        "why": why,
    }


def compute_health_score(tx_df, budgets_df, granularity="Month"):
    """FinSight Business Health Score, out of 100, split across four
    25-point components. Returns {'total', 'components', 'runway_months',
    'insufficient_data'}. Fully defensive against empty/zero data."""
    if tx_df is None or tx_df.empty:
        return {
            "total": 0,
            "components": {"revenue_momentum": 0, "profitability": 0, "expense_control": 0, "runway": 0},
            "runway_months": None,
            "insufficient_data": True,
        }

    pulse = compute_business_pulse(tx_df, granularity)
    if pulse is None:
        return {
            "total": 0,
            "components": {"revenue_momentum": 0, "profitability": 0, "expense_control": 0, "runway": 0},
            "runway_months": None,
            "insufficient_data": True,
        }

    df = tx_df.copy()
    latest = df["date"].max()

    # ---- Revenue Momentum (25 pts) ----
    rev_cur = pulse["what"]["revenue"]["current"]
    rev_prev = pulse["what"]["revenue"]["previous"]
    if rev_prev > 0:
        growth_pct = (rev_cur - rev_prev) / rev_prev * 100
    else:
        growth_pct = 0.0 if rev_cur == 0 else 100.0
    # Full 25 pts at +10% growth or more; scales linearly either side, floored at 0.
    revenue_momentum = 25 * (growth_pct / 10)
    revenue_momentum = max(0.0, min(25.0, revenue_momentum))

    # ---- Profitability (25 pts) ----
    margin = pulse["what"]["margin"]["current"]
    profitability = 25 * (margin / 25)
    profitability = max(0.0, min(25.0, profitability))

    # ---- Expense Control (25 pts) — budget adherence ----
    expense_control = 25.0
    try:
        if budgets_df is not None and not budgets_df.empty:
            this_month = df[(df["date"].dt.year == latest.year) & (df["date"].dt.month == latest.month)]
            spend_by_cat = this_month[this_month["type"] == "expense"].groupby("category")["amount"].sum()
            n_budgets = len(budgets_df)
            penalty_units = 0.0
            for _, b in budgets_df.iterrows():
                limit = b["monthly_limit"]
                if not limit or limit <= 0:
                    continue
                spent = spend_by_cat.get(b["category"], 0.0)
                consumed_pct = spent / limit
                if consumed_pct > 1.0:
                    penalty_units += 1.0
                elif consumed_pct > 0.8:
                    penalty_units += 0.5
            if n_budgets > 0:
                expense_control = max(0.0, 25.0 - (penalty_units / n_budgets) * 25.0)
    except Exception:
        expense_control = 25.0

    # ---- Runway & Cash Flow (25 pts) ----
    runway_months = None
    runway_score = 25.0
    try:
        span_days = max((latest - df["date"].min()).days, 1)
        span_months = max(span_days / 30.0, 1.0)
        total_expenses_all_time = df.loc[df["type"] == "expense", "amount"].sum()
        avg_monthly_burn = total_expenses_all_time / span_months
        total_income_all_time = df.loc[df["type"] == "income", "amount"].sum()
        current_balance = max(total_income_all_time - total_expenses_all_time, 0.0)
        if avg_monthly_burn > 0:
            runway_months = current_balance / avg_monthly_burn
            runway_score = max(0.0, min(25.0, (runway_months / 6.0) * 25.0))
        else:
            runway_months = None
            runway_score = 25.0
    except Exception:
        runway_score = 12.5

    total = revenue_momentum + profitability + expense_control + runway_score

    return {
        "total": round(total),
        "components": {
            "revenue_momentum": round(revenue_momentum),
            "profitability": round(profitability),
            "expense_control": round(expense_control),
            "runway": round(runway_score),
        },
        "runway_months": runway_months,
        "insufficient_data": False,
    }


def build_action_center(anomalies, health_score):
    """Aggregates anomaly-engine + health-score + budget signals into a
    single prioritized, color-coded list of {severity, icon, message}."""
    items = []

    for a in anomalies:
        icon = "🔴" if a["severity"] == "red" else ("🟡" if a["severity"] == "yellow" else "🟢")
        items.append({"severity": a["severity"], "icon": icon, "message": a["message"]})

    if not health_score.get("insufficient_data"):
        comps = health_score["components"]
        if comps["expense_control"] < 15:
            items.append({
                "severity": "red", "icon": "🔴",
                "message": f"Expense Control score dropped to {comps['expense_control']}/25 — review overspent budget categories or cut overheads.",
            })
        if comps["runway"] < 10:
            items.append({
                "severity": "yellow", "icon": "🟡",
                "message": f"Runway score is low ({comps['runway']}/25) — cash buffer relative to burn rate is thin.",
            })
        if comps["revenue_momentum"] < 8:
            items.append({
                "severity": "yellow", "icon": "🟡",
                "message": f"Revenue Momentum is soft ({comps['revenue_momentum']}/25) — consider a push on sales or collections.",
            })
        if health_score["total"] >= 85:
            items.append({
                "severity": "green", "icon": "🟢",
                "message": f"Business Health Score is excellent at {health_score['total']}/100 — keep up the momentum.",
            })

    order = {"red": 0, "yellow": 1, "green": 2}
    items.sort(key=lambda x: order.get(x["severity"], 3))
    return items


def compute_automatic_highlights(view, previous_view, all_tx, pulse, category_kpi,
                                  business_category, unpaid_df, budgets_df, recurring_df,
                                  income, expenses, workspace_id):
    """Automatic Highlights — 3-5 plain-English observations, ranked by
    usefulness. Pure detection over data FinSight already loads/computes
    (the period `view`, `pulse` from compute_business_pulse, budgets,
    unpaid invoices, recurring rules). No new tables, no external calls.
    Every highlight is only added when its underlying data actually
    exists — nothing here is invented or estimated."""
    candidates = []  # each: (priority, icon, title, text)

    # ---- 1. Revenue momentum (from the What & Why pulse) ----
    if pulse is not None:
        rev = pulse["what"]["revenue"]
        if rev["previous"] > 0:
            rev_pct = (rev["current"] - rev["previous"]) / rev["previous"] * 100
            if abs(rev_pct) >= 5:
                direction_word = "increased" if rev_pct >= 0 else "decreased"
                icon = "✨" if rev_pct >= 0 else "⚠️"
                candidates.append((
                    1 if abs(rev_pct) >= 15 else 3, icon, "Revenue Momentum",
                    f"Revenue {direction_word} {abs(rev_pct):.0f}% compared with the previous {pulse['granularity'].lower()}.",
                ))

        # ---- 4. Profit margin swing ----
        margin_field = pulse["what"]["margin"]
        margin_delta = margin_field["current"] - margin_field["previous"]
        if pulse["what"]["revenue"]["previous"] > 0 and abs(margin_delta) >= 3:
            icon = "💡" if margin_delta >= 0 else "⚠️"
            verb = "improved" if margin_delta >= 0 else "dropped"
            candidates.append((
                1 if abs(margin_delta) >= 8 else 2, icon, "Margin Change",
                f"Profit margin {verb} from {margin_field['previous']:.1f}% to {margin_field['current']:.1f}%.",
            ))

        # ---- 2. Biggest-moving expense category (from pulse "why") ----
        if pulse["why"]:
            top = pulse["why"][0]
            if top["amount"] > 0:
                candidates.append((
                    2, "⚠️", f"{top['category']} Costs",
                    f"{top['category']} expenses grew by {money(abs(top['amount']))} — {top['pct_of_change']:.0f}% of this {pulse['granularity'].lower()}'s expense increase.",
                ))

    # ---- 3. A single expense category dominating total spend ----
    if expenses > 0 and not view.empty:
        expense_data = view[view["type"] == "expense"]
        if not expense_data.empty:
            by_cat = expense_data.groupby("category")["amount"].sum().sort_values(ascending=False)
            top_cat, top_amt = by_cat.index[0], by_cat.iloc[0]
            share = top_amt / expenses * 100
            if share >= 30:
                candidates.append((
                    3, "⚠️", f"{top_cat} Concentration",
                    f"{top_cat} expenses are now {share:.0f}% of total expenses.",
                ))

    # ---- 7. Business-category-specific cost, if unusually high ----
    if category_kpi:
        kpi_label, kpi_value, _ = category_kpi
        if isinstance(kpi_value, str) and kpi_value.endswith("%"):
            try:
                kpi_pct = float(kpi_value.rstrip("%"))
                if kpi_pct >= 40:
                    candidates.append((
                        2, "⚠️", kpi_label,
                        f"{kpi_label} is {kpi_value} of revenue this period — high for a {business_category.lower()} business.",
                    ))
            except ValueError:
                pass

    # ---- 5. Unpaid invoices ----
    if unpaid_df is not None and not unpaid_df.empty:
        total_owed = unpaid_df["amount"].sum()
        candidates.append((
            2, "💰", "Unpaid Invoices",
            f"{len(unpaid_df)} invoice(s) worth {money(total_owed)} are still unpaid.",
        ))

    # ---- 6. Budget being consumed unusually quickly ----
    if budgets_df is not None and not budgets_df.empty and not all_tx.empty:
        latest_date = all_tx["date"].max()
        days_in_month = (pd.Timestamp(latest_date.year, latest_date.month, 1) + pd.offsets.MonthEnd(0)).day
        month_fraction = (latest_date.day / days_in_month) if days_in_month else 0
        this_month_expenses = all_tx[
            (all_tx["date"].dt.year == latest_date.year) & (all_tx["date"].dt.month == latest_date.month) & (all_tx["type"] == "expense")
        ]
        spend_by_cat = this_month_expenses.groupby("category")["amount"].sum()
        for _, b in budgets_df.iterrows():
            limit = b["monthly_limit"]
            if not limit or limit <= 0:
                continue
            consumed_pct = spend_by_cat.get(b["category"], 0.0) / limit
            if consumed_pct >= 0.8 and month_fraction < 0.6:
                candidates.append((
                    1, "⚠️", f"{b['category']} Budget",
                    f"{b['category']} budget is {consumed_pct * 100:.0f}% used with only {month_fraction * 100:.0f}% of the month elapsed.",
                ))
                break  # one budget highlight is enough to avoid crowding the list

    # ---- 8. Top revenue-generating employee, if employee data exists ----
    if not view.empty:
        income_by_employee = view[(view["type"] == "income") & (view["employee"].str.strip() != "")]
        if not income_by_employee.empty:
            by_emp = income_by_employee.groupby("employee")["amount"].sum().sort_values(ascending=False)
            top_emp, top_emp_amt = by_emp.index[0], by_emp.iloc[0]
            if income > 0:
                emp_share = top_emp_amt / income * 100
                if emp_share >= 20:
                    candidates.append((
                        4, "🌟", "Top Contributor",
                        f"{top_emp} is linked to {money(top_emp_amt)} in revenue this period ({emp_share:.0f}% of the total).",
                    ))

    # ---- 9. A recurring expense that's becoming significant ----
    if recurring_df is not None and not recurring_df.empty and expenses > 0:
        active_recurring = recurring_df[(recurring_df["active"] == 1) & (recurring_df["type"] == "expense")]
        if not active_recurring.empty:
            monthly_equiv = active_recurring.apply(
                lambda r: r["amount"] * (4.33 if r["frequency"] == "Weekly" else (1 if r["frequency"] == "Monthly" else 1 / 12)),
                axis=1,
            )
            biggest_idx = monthly_equiv.idxmax()
            biggest_amt = monthly_equiv.loc[biggest_idx]
            biggest_desc = active_recurring.loc[biggest_idx, "description"]
            share = biggest_amt / expenses * 100
            if share >= 15:
                candidates.append((
                    3, "🔁", "Recurring Expense",
                    f"'{biggest_desc}' is a recurring cost of roughly {money(biggest_amt)}/month — about {share:.0f}% of this period's expenses.",
                ))

    # ---- Rank (lower number = higher priority) and cap at 5 ----
    candidates.sort(key=lambda c: c[0])
    return [{"icon": icon, "title": title, "text": text} for _, icon, title, text in candidates[:5]]


def build_whatsapp_link(phone, message):
    """Builds a wa.me click-to-chat link. No API key, no cost — opens WhatsApp
    with the message pre-filled; a human still has to tap send. Works with or
    without a phone number (no number opens WhatsApp's contact picker)."""
    encoded_message = urllib.parse.quote(message)
    digits = re.sub(r"[^0-9]", "", phone or "")
    if digits:
        return f"https://wa.me/{digits}?text={encoded_message}"
    return f"https://wa.me/?text={encoded_message}"


MOOD_STYLES = {
    "green": dict(color="#059669", bg_a="rgba(16,185,129,0.10)", bg_b="rgba(16,185,129,0.015)",
                  border="rgba(16,185,129,0.30)", glow="rgba(16,185,129,0.22)", badge_bg="rgba(16,185,129,0.10)"),
    "yellow": dict(color="#B45309", bg_a="rgba(245,158,11,0.11)", bg_b="rgba(245,158,11,0.015)",
                   border="rgba(245,158,11,0.34)", glow="rgba(245,158,11,0.22)", badge_bg="rgba(245,158,11,0.12)"),
    "red": dict(color="#E11D48", bg_a="rgba(244,63,94,0.09)", bg_b="rgba(244,63,94,0.015)",
                border="rgba(244,63,94,0.30)", glow="rgba(244,63,94,0.20)", badge_bg="rgba(244,63,94,0.09)"),
    "neutral": dict(color="#64748B", bg_a="rgba(17,24,39,0.03)", bg_b="rgba(17,24,39,0.005)",
                     border="rgba(17,24,39,0.10)", glow="rgba(17,24,39,0.08)", badge_bg="rgba(17,24,39,0.05)"),
}


# ============================================================
# STYLE
# ============================================================

def inject_global_css():
    """One light design system for the landing deck AND the workspace.

    Rules this stylesheet enforces:
      * every card / input / button gets ONE 1px #E2E8F0 border, a 12px radius and the same
        soft ambient shadow — no stacked, spread-ring or offset (2px 2px 0) shadows anywhere;
      * inputs are styled on their OUTER Streamlit wrapper only, and every inner layer is
        flattened — that double border is what made fields look boxy / pixelated before.
    """
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=JetBrains+Mono:wght@500;700&display=swap');

        :root {
            color-scheme: light;
            --canvas: #FFFFFF; 
            --canvas-soft: #F9FAFB; 
            --panel: #F4F7F6;
            --ink: #111C24; 
            --ink-2: #2D3D4A; 
            --muted: #6B7280; 
            --faint: #9CA3AF;
            --border: #E5E7EB; 
            --border-soft: #F3F4F6;
            --cyan: #10B981; 
            --cyan-ink: #047857; 
            --cyan-wash: rgba(16, 185, 129, 0.08);
            --green: #10B981; 
            --red: #EF4444; 
            --amber: #F59E0B;
            --radius: 12px;
            --shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.03), 0 2px 4px -1px rgba(0, 0, 0, 0.02);
            --focus-glow: 0 0 14px rgba(16, 185, 129, 0.18);
        }


        /* ---------- canvas & typography ---------- */
        html, body, [data-testid="stAppViewContainer"], [data-testid="stMain"], .stApp {
            font-family: "Inter", -apple-system, "Segoe UI", sans-serif;
            background: var(--canvas) !important; color: var(--ink);
        }
        [data-testid="stHeader"] { background: transparent; }
        [data-testid="stDecoration"], [data-testid="stToolbar"] { display: none; }
        footer, #MainMenu { visibility: hidden; }
        .block-container { max-width: 1550px; padding: 2rem 3rem 5rem 3rem; }

        h1, h2, h3, h4, h5, h6,
        [data-testid="stHeading"] *, [data-testid="stMarkdownContainer"] h1, [data-testid="stMarkdownContainer"] h2,
        [data-testid="stMarkdownContainer"] h3, [data-testid="stMarkdownContainer"] h4, [data-testid="stMarkdownContainer"] h5 {
            color: var(--ink) !important; letter-spacing: -0.6px;
        }
        h1 { font-weight: 900 !important; } h2, h3, h4, h5 { font-weight: 700 !important; }
        [data-testid="stMarkdownContainer"] p, [data-testid="stMarkdownContainer"] li { color: var(--ink-2); }
        [data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] * { color: var(--muted) !important; }
        [data-testid="stWidgetLabel"], [data-testid="stWidgetLabel"] *, .stApp label { color: var(--ink-2) !important; font-weight: 600; }
        hr { border-color: var(--border) !important; }

        /* ---------- left navigation panel & settings boxes ---------- */
        [data-testid="stSidebar"], [data-testid="stSidebar"] > div:first-child {
            background: var(--panel) !important;
        }
        [data-testid="stSidebar"] { border-right: 1px solid var(--border); }
        [data-testid="stSidebar"] [data-testid="stWidgetLabel"] * { color: var(--muted) !important; font-weight: 600; }

        /* ---------- INPUTS: one border + one soft shadow, on the outer wrapper only ---------- */
        .stApp [data-baseweb="input"],
        .stApp [data-baseweb="textarea"],
        .stApp [data-baseweb="select"] > div {
            background: #FFFFFF !important;
            border: 1px solid var(--border) !important;
            border-radius: var(--radius) !important;
            box-shadow: var(--shadow) !important;
            overflow: hidden;
            transition: border-color .2s ease, box-shadow .2s ease;
        }
        /* flatten every inner layer so nothing draws a second box, ring or shadow */
        .stApp [data-baseweb="base-input"],
        .stApp [data-baseweb="input"] > div,
        .stApp [data-baseweb="base-input"] input,
        .stApp [data-baseweb="textarea"] textarea,
        .stApp [data-baseweb="select"] input,
        .stApp input[type="text"], .stApp input[type="password"], .stApp input[type="number"],
        .stApp input[type="email"], .stApp input:not([type]), .stApp textarea {
            background: transparent !important; border: none !important; box-shadow: none !important;
            outline: none !important; color: var(--ink) !important; -webkit-text-fill-color: var(--ink);
        }
        .stApp [data-baseweb="base-input"], .stApp [data-baseweb="input"] > div { border-radius: 0 !important; }
        .stApp [data-testid="stNumberInputContainer"]:has([data-baseweb="input"]) {
            background: transparent !important; border: none !important; box-shadow: none !important;
        }
        .stApp [data-testid="stNumberInputStepUp"], .stApp [data-testid="stNumberInputStepDown"] {
            background: transparent !important; border: none !important; box-shadow: none !important; color: var(--muted) !important;
        }
        .stApp [data-testid="stNumberInputStepUp"]:hover, .stApp [data-testid="stNumberInputStepDown"]:hover { background: var(--panel) !important; }
        .stApp input::placeholder, .stApp textarea::placeholder { color: var(--faint) !important; -webkit-text-fill-color: var(--faint); opacity: 1; }
        .stApp [data-baseweb="select"] [data-baseweb="tag"], .stApp [data-baseweb="select"] div { color: var(--ink); }
        .stApp [data-baseweb="input"]:hover, .stApp [data-baseweb="textarea"]:hover, .stApp [data-baseweb="select"] > div:hover { border-color: #CBD5E1 !important; }
        .stApp [data-baseweb="input"]:focus-within,
        .stApp [data-baseweb="textarea"]:focus-within,
        .stApp [data-baseweb="select"]:focus-within > div {
            border-color: var(--cyan) !important;
            box-shadow: var(--shadow), var(--focus-glow) !important;
        }
        /* dropdown menus */
        [data-baseweb="popover"] > div {
            background: #FFFFFF !important; border: 1px solid var(--border) !important;
            border-radius: var(--radius) !important; box-shadow: var(--shadow) !important; overflow: hidden;
        }
        [data-baseweb="popover"] [data-baseweb="menu"], [data-baseweb="popover"] ul { background: #FFFFFF !important; }
        [data-baseweb="popover"] li { color: var(--ink) !important; }
        [data-baseweb="popover"] li:hover, [data-baseweb="popover"] li[aria-selected="true"] { background: var(--panel) !important; }

        /* ---------- BUTTONS ---------- */
        .stApp .stButton > button, .stApp .stDownloadButton > button,
        .stApp .stFormSubmitButton > button, .stApp .stLinkButton > a {
            min-height: 42px; background: #FFFFFF !important; color: var(--ink) !important;
            border: 1px solid var(--border) !important; border-radius: var(--radius) !important;
            box-shadow: var(--shadow) !important; font-weight: 700 !important;
            transition: border-color .2s ease, background .2s ease, transform .2s ease, box-shadow .2s ease !important;
        }
        .stApp .stButton > button p, .stApp .stDownloadButton > button p,
        .stApp .stFormSubmitButton > button p, .stApp .stLinkButton > a p { color: inherit !important; }
        .stApp .stButton > button:hover, .stApp .stDownloadButton > button:hover,
        .stApp .stFormSubmitButton > button:hover, .stApp .stLinkButton > a:hover {
            border-color: var(--cyan) !important; background: #F5FDFF !important; color: var(--ink) !important; transform: translateY(-1px);
        }
        .stApp .stButton > button:focus-visible, .stApp .stFormSubmitButton > button:focus-visible,
        .stApp .stDownloadButton > button:focus-visible {
            outline: none !important; border-color: var(--cyan) !important; box-shadow: var(--shadow), var(--focus-glow) !important;
        }
        .stApp .stButton > button:disabled { opacity: .5; box-shadow: none !important; transform: none; }
        .stApp .stButton > button[kind="primary"], .stApp .stButton > button[data-testid="stBaseButton-primary"],
        .stApp .stFormSubmitButton > button[kind="primaryFormSubmit"], .stApp .stFormSubmitButton > button[data-testid="stBaseButton-primaryFormSubmit"] {
            background: var(--ink) !important; border-color: var(--ink) !important; color: #FFFFFF !important;
        }
        .stApp .stButton > button[kind="primary"]:hover, .stApp .stButton > button[data-testid="stBaseButton-primary"]:hover,
        .stApp .stFormSubmitButton > button[kind="primaryFormSubmit"]:hover, .stApp .stFormSubmitButton > button[data-testid="stBaseButton-primaryFormSubmit"]:hover {
            background: #1E293B !important; border-color: var(--cyan) !important; color: #FFFFFF !important;
        }

        /* ---------- CONTAINERS: forms, metrics, expanders, tables ---------- */
        [data-testid="stForm"] {
            background: #FFFFFF !important; border: 1px solid var(--border) !important; border-radius: var(--radius) !important;
            box-shadow: var(--shadow) !important; padding: 1.25rem;
        }
        [data-testid="stMetric"] {
            background: #FFFFFF; border: 1px solid var(--border); border-radius: var(--radius);
            padding: 22px; min-height: 130px; box-shadow: var(--shadow);
            transition: transform .25s ease, border-color .25s ease;
        }
        [data-testid="stMetric"]:hover { transform: translateY(-2px); border-color: rgba(0,212,255,0.6); }
        [data-testid="stMetricLabel"], [data-testid="stMetricLabel"] * {
            color: var(--muted) !important; font-size: 11px !important; font-weight: 700 !important; text-transform: uppercase;
        }
        [data-testid="stMetricValue"], [data-testid="stMetricValue"] * {
            color: var(--ink) !important; font-size: 27px !important; font-weight: 800 !important; font-family: "JetBrains Mono", monospace !important;
        }
        [data-testid="stExpander"] { border: none !important; background: transparent !important; box-shadow: none !important; }
        [data-testid="stExpander"] details {
            background: #FFFFFF; border: 1px solid var(--border) !important; border-radius: var(--radius) !important; box-shadow: var(--shadow);
        }
        [data-testid="stExpander"] summary, [data-testid="stExpander"] summary * { color: var(--ink) !important; font-weight: 600; }
        [data-testid="stSidebar"] [data-testid="stExpander"] details { background: #FFFFFF; }
        [data-testid="stDataFrame"], [data-testid="stDataEditor"] {
            border: 1px solid var(--border); border-radius: var(--radius); overflow: hidden; box-shadow: var(--shadow);
        }
        [data-testid="stAlert"] { border-radius: var(--radius); box-shadow: none; }
        [data-baseweb="tab"] { color: var(--muted) !important; font-weight: 600; }
        [data-baseweb="tab"][aria-selected="true"] { color: var(--ink) !important; }
        [data-baseweb="tab-highlight"] { background: var(--cyan) !important; }
        [data-baseweb="tab-border"] { background: var(--border) !important; }
        .stProgress > div > div > div > div { background-color: var(--cyan) !important; }

        /* ---------- workspace components ---------- */
        .hl { background: linear-gradient(transparent 62%, rgba(0,212,255,0.38) 62%); }
        .mood-hero {
            display: flex; align-items: center; gap: 26px; padding: 28px 32px; border-radius: var(--radius);
            border: 1px solid var(--mood-border, var(--border));
            background: linear-gradient(135deg, var(--mood-bg-a, #F8FAFC), var(--mood-bg-b, #FFFFFF));
            box-shadow: var(--shadow); transition: transform .3s ease;
        }
        .mood-hero:hover { transform: translateY(-2px); }
        .mood-emoji { font-size: 84px; line-height: 1; animation: floaty 3.2s ease-in-out infinite; }
        @keyframes floaty { 0%,100% { transform: translateY(0px);} 50% { transform: translateY(-8px);} }
        .mood-text-title { font-size: 22px; font-weight: 800; color: var(--mood-color, var(--ink)); margin-bottom: 4px; }
        .mood-text-sub { color: var(--ink-2); font-size: 14px; }
        .mood-badge {
            display: inline-block; margin-top: 10px; padding: 4px 12px; border-radius: 999px; font-size: 11px; font-weight: 700;
            text-transform: uppercase; background: var(--mood-badge-bg, #F1F5F9); color: var(--mood-color, var(--ink));
            border: 1px solid var(--mood-border, var(--border));
        }
        .ticker-strip { display: flex; gap: 14px; flex-wrap: wrap; margin-top: 14px; }
        .ticker-chip {
            padding: 8px 16px; border-radius: var(--radius); background: #FFFFFF; border: 1px solid var(--border);
            font-family: "JetBrains Mono", monospace; font-size: 13px; font-weight: 600; color: var(--ink-2);
            box-shadow: var(--shadow); transition: border-color .2s ease, transform .2s ease;
        }
        .ticker-chip:hover { border-color: var(--cyan); transform: translateY(-1px); }
        .ticker-up { color: var(--green); } .ticker-down { color: var(--red); }
        .glass-card {
            background: #FFFFFF; border: 1px solid var(--border); border-radius: var(--radius); padding: 20px;
            min-height: 118px; box-shadow: var(--shadow); transition: transform .25s ease, border-color .25s ease;
        }
        .glass-card:hover { transform: translateY(-2px); border-color: rgba(0,212,255,0.6); }
        .glass-label { color: var(--muted); font-size: 11px; font-weight: 700; text-transform: uppercase; }
        .glass-value { font-size: 22px; font-weight: 800; margin: 6px 0 4px 0; font-family: "JetBrains Mono", monospace; color: var(--ink); }
        .glass-sub { color: var(--muted); font-size: 13px; }
        .plan-badge { display:inline-block; padding: 3px 10px; border-radius: 999px; font-size: 11px; font-weight: 800; text-transform: uppercase; letter-spacing: .5px; }
        .plan-free { background: #F1F5F9; color: var(--ink-2); border: 1px solid var(--border); }
        .plan-pro { background: rgba(217,119,6,0.09); color: #B45309; border: 1px solid rgba(217,119,6,0.28); }
        .category-badge {
            display:inline-block; padding: 3px 10px; border-radius: 999px; font-size: 11px; font-weight: 700;
            background: var(--cyan-wash); color: var(--cyan-ink); border: 1px solid rgba(0,212,255,0.4);
        }
        .section { font-size: 19px; font-weight: 800; color: var(--ink); margin: 26px 0 4px 0; letter-spacing: -0.3px; }
        .helper { color: var(--muted); font-size: 13px; margin-bottom: 14px; }
        .insight {
            background: #FFFFFF; border: 1px solid var(--border); border-radius: var(--radius); padding: 14px 16px; margin-bottom: 10px;
            font-size: 14px; color: var(--ink-2); box-shadow: var(--shadow); transition: transform .2s ease, border-color .2s ease;
        }
        .insight:hover { transform: translateY(-1px); border-color: rgba(0,212,255,0.6); }
        .insight span { color: var(--muted); } .insight b { color: var(--ink); }
        .action-item-red { border-left: 3px solid #F43F5E; }
        .action-item-yellow { border-left: 3px solid #F59E0B; }
        .action-item-green { border-left: 3px solid #10B981; }
        .health-score-ring { text-align:center; min-height:170px; display:flex; flex-direction:column; align-items:center; justify-content:center; }
        .why-row { display:flex; justify-content:space-between; align-items:center; padding:8px 0; border-bottom: 1px solid var(--border-soft); font-size: 13px; }
        .why-row:last-child { border-bottom: none; }

        @media (max-width: 900px) {
            .block-container { padding-left: 1rem; padding-right: 1rem; }
            .mood-hero { flex-direction: column; text-align: center; }
            .mood-emoji { font-size: 64px; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


inject_global_css()


# ============================================================
# SESSION STATE DEFAULTS
# ============================================================

for key, default in {
    "user": None,
    "workspace_id": None,
    "auth_mode": "login",
    "pulse_granularity": "Month",
    "tour_frame": 0,  # 0-based index of the landing-deck frame currently shown
}.items():
    if key not in st.session_state:
        st.session_state[key] = default


# ============================================================
# LANDING DECK — 5-frame presentation gateway
# ============================================================

def _html(markup):
    """Collapse indented markup to one line so Markdown never treats it as a code block."""
    return " ".join(line.strip() for line in markup.strip().splitlines())


# ============================================================
# FINBOT — white & blue helmet robot with a tablet
# ============================================================
# The robot body is an SVG delivered as a data-URI <img> (so Streamlit's HTML
# handling can't mangle gradients/filters). The tablet is a real HTML/CSS
# container layered behind the robot's hand, and the dialogue is typed into it.

FB_BADGES = {
    # tiny floating badge hovering beside the helmet — one per tour frame
    "hello": '<path d="M0 -7 L2 -2 L7 0 L2 2 L0 7 L-2 2 L-7 0 L-2 -2 Z" fill="#00D4FF"/>',
    "pulse": '<path d="M-8 1 L-4 1 L-2 -6 L2 7 L4 1 L8 1" fill="none" stroke="#00B8E6" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/>',
    "alert": '<path d="M0 -7 V1.5" stroke="#F43F5E" stroke-width="3" stroke-linecap="round"/><circle cx="0" cy="6" r="1.9" fill="#F43F5E"/>',
    "coin": '<path d="M-4.5 -5.5 H4.5 M-4.5 -2 H4.5 M-3 -5.5 H0.6 Q4 -5.5 4 -2.6 Q4 0.4 0.6 0.4 H-3 M-3 0.4 L3.4 6.6" fill="none" stroke="#B45309" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"/>',
    "shield": '<path d="M0 -8 L7 -5 V1 C7 5 3.5 7.5 0 9 C-3.5 7.5 -7 5 -7 1 V-5 Z" fill="#E0F7FF" stroke="#2F7BF5" stroke-width="1.8" stroke-linejoin="round"/><path d="M-3.2 0.8 L-0.8 3.2 L3.6 -2.2" fill="none" stroke="#0A8FB0" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>',
}


def _capsule(x1, y1, x2, y2, w):
    """A floating limb segment: soft blue edge, white body, tiny specular highlight."""
    return (
        f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="#9CC7FF" stroke-width="{w + 3.4}" stroke-linecap="round"/>'
        f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="#FFFFFF" stroke-width="{w}" stroke-linecap="round"/>'
        f'<line x1="{x1 - 1.2}" y1="{y1 - 1.2}" x2="{x2 - 1.2}" y2="{y2 - 1.2}" stroke="#E2EBF7" stroke-width="{max(w - 6, 2)}" stroke-linecap="round" opacity=".8"/>'
    )


def _joint(cx, cy, r):
    """A blue mechanical joint ring."""
    return (
        f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="url(#fbB)" stroke="#FFFFFF" stroke-width="1.6"/>'
        f'<circle cx="{cx}" cy="{cy}" r="{max(r - 4, 1.6)}" fill="#FFFFFF" opacity=".92"/>'
    )


def finbot_svg_markup(badge="hello"):
    badge_icon = FB_BADGES.get(badge, FB_BADGES["hello"])
    return f"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 220 240" width="220" height="240">
<defs>
  <radialGradient id="fbW" cx=".34" cy=".26" r=".9"><stop offset="0" stop-color="#FFFFFF"/><stop offset=".55" stop-color="#F3F7FC"/><stop offset="1" stop-color="#D3DFEE"/></radialGradient>
  <linearGradient id="fbB" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#6FD3FF"/><stop offset="1" stop-color="#2563EB"/></linearGradient>
  <linearGradient id="fbV" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#1F4287"/><stop offset=".55" stop-color="#0E2150"/><stop offset="1" stop-color="#07122B"/></linearGradient>
  <linearGradient id="fbS" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#fff" stop-opacity=".55"/><stop offset="1" stop-color="#fff" stop-opacity="0"/></linearGradient>
  <filter id="fbG" x="-60%" y="-60%" width="220%" height="220%"><feGaussianBlur stdDeviation="2.6" result="b"/><feMerge><feMergeNode in="b"/><feMergeNode in="b"/><feMergeNode in="SourceGraphic"/></feMerge></filter>
  <filter id="fbBlur" x="-50%" y="-50%" width="200%" height="200%"><feGaussianBlur stdDeviation="5"/></filter>
</defs>
<style>
  .wave{{transform-origin:130px 150px;animation:wave 2.1s ease-in-out infinite}}
  .blink{{transform-box:fill-box;transform-origin:center;animation:blink 5.2s infinite}}
  .jet{{transform-box:fill-box;transform-origin:center;animation:jet 1.6s ease-in-out infinite}}
  .beacon{{animation:beacon 1.8s ease-in-out infinite}}
  .badge{{animation:bob 3s ease-in-out infinite}}
  .hand{{animation:grip 2.8s ease-in-out infinite}}
  @keyframes wave{{0%,100%{{transform:rotate(-2deg)}}50%{{transform:rotate(12deg)}}}}
  @keyframes blink{{0%,92%,100%{{transform:scaleY(1)}}95.5%{{transform:scaleY(.12)}}}}
  @keyframes jet{{0%,100%{{opacity:.55;transform:scale(1)}}50%{{opacity:.95;transform:scale(1.18)}}}}
  @keyframes beacon{{0%,100%{{opacity:.7}}50%{{opacity:1}}}}
  @keyframes bob{{0%,100%{{transform:translate(34px,30px)}}50%{{transform:translate(34px,26px)}}}}
  @keyframes grip{{0%,100%{{transform:translateY(0)}}50%{{transform:translateY(-1.5px)}}}}
</style>

<!-- ground shadow -->
<ellipse cx="118" cy="228" rx="34" ry="5" fill="#0F172A" opacity=".10" filter="url(#fbBlur)"/>

<g transform="translate(10,12)">
  <!-- thruster glow + hull cone -->
  <ellipse class="jet" cx="90" cy="198" rx="17" ry="6.5" fill="#00D4FF" filter="url(#fbBlur)"/>
  <path d="M70 172 Q90 204 110 172 Z" fill="url(#fbB)" stroke="#FFFFFF" stroke-width="1.6"/>
  <path d="M78 176 Q90 192 102 176" fill="none" stroke="#FFFFFF" stroke-width="2" opacity=".55" stroke-linecap="round"/>

  <!-- torso -->
  <rect x="58" y="112" width="64" height="68" rx="30" fill="url(#fbW)" stroke="#D5E1F0" stroke-width="1.6"/>
  <path d="M68 122 Q74 117 84 117" fill="none" stroke="#FFFFFF" stroke-width="3.2" stroke-linecap="round" opacity=".9"/>
  <rect x="75" y="131" width="30" height="24" rx="11" fill="url(#fbV)" stroke="#2F7BF5" stroke-width="1.4"/>
  <circle cx="90" cy="143" r="7.4" fill="#00D4FF" filter="url(#fbG)"/>
  <circle cx="90" cy="143" r="3.2" fill="#FFFFFF"/>
  <rect x="64" y="160" width="52" height="8" rx="4" fill="url(#fbB)"/>
  <rect x="66" y="161.4" width="22" height="2.2" rx="1.1" fill="#FFFFFF" opacity=".6"/>

  <!-- neck ring -->
  <rect x="74" y="105" width="32" height="12" rx="6" fill="url(#fbB)" stroke="#FFFFFF" stroke-width="1.6"/>

  <!-- helmet head -->
  <ellipse cx="90" cy="62" rx="68" ry="55" fill="url(#fbW)" stroke="#D5E1F0" stroke-width="1.8"/>
  <path d="M38 40 Q52 14 82 12" fill="none" stroke="#FFFFFF" stroke-width="6" stroke-linecap="round" opacity=".95"/>
  <path d="M30 56 Q31 46 36 38" fill="none" stroke="#FFFFFF" stroke-width="4" stroke-linecap="round" opacity=".8"/>
  <!-- ear pods -->
  <rect x="12" y="44" width="17" height="34" rx="8.5" fill="url(#fbB)" stroke="#FFFFFF" stroke-width="1.8"/>
  <rect x="151" y="44" width="17" height="34" rx="8.5" fill="url(#fbB)" stroke="#FFFFFF" stroke-width="1.8"/>
  <rect x="16.5" y="51" width="3.6" height="20" rx="1.8" fill="#FFFFFF" opacity=".7"/>
  <rect x="159.5" y="51" width="3.6" height="20" rx="1.8" fill="#FFFFFF" opacity=".7"/>
  <!-- beacon -->
  <rect x="86" y="3" width="8" height="8" rx="3" fill="url(#fbB)"/>
  <circle class="beacon" cx="90" cy="2.5" r="4.6" fill="#00D4FF" filter="url(#fbG)"/>

  <!-- visor -->
  <rect x="34" y="30" width="112" height="66" rx="33" fill="url(#fbV)" stroke="#2F7BF5" stroke-width="2"/>
  <path d="M46 44 Q58 34 82 33" fill="none" stroke="url(#fbS)" stroke-width="5" stroke-linecap="round"/>
  <!-- happy glowing eyes -->
  <g class="blink">
    <path d="M56 70 C56 55 62 48 69 48 C76 48 82 55 82 70 Q69 60 56 70 Z" fill="#26D3FF" filter="url(#fbG)"/>
    <path d="M98 70 C98 55 104 48 111 48 C118 48 124 55 124 70 Q111 60 98 70 Z" fill="#26D3FF" filter="url(#fbG)"/>
    <circle cx="64.5" cy="55.5" r="2.6" fill="#FFFFFF" opacity=".95"/>
    <circle cx="106.5" cy="55.5" r="2.6" fill="#FFFFFF" opacity=".95"/>
  </g>
  <path d="M81 79 Q90 87.5 99 79" fill="none" stroke="#26D3FF" stroke-width="3.6" stroke-linecap="round" filter="url(#fbG)"/>
  <ellipse cx="47" cy="78" rx="6" ry="3.4" fill="#38BDF8" opacity=".5"/>
  <ellipse cx="133" cy="78" rx="6" ry="3.4" fill="#38BDF8" opacity=".5"/>

  <!-- floating badge -->
  <g class="badge"><circle r="13.5" fill="#FFFFFF" stroke="#2F7BF5" stroke-width="2.4"/>{badge_icon}</g>

  <!-- RIGHT arm: floating segments, waving -->
  <g class="wave">
    {_joint(130, 150, 7)}
    {_capsule(141, 146, 153, 136, 9)}
    {_joint(159, 131, 5)}
    {_capsule(164, 125, 172, 109, 8)}
    <ellipse cx="174.5" cy="102" rx="7" ry="3.4" transform="rotate(-30 174.5 102)" fill="url(#fbB)" stroke="#FFFFFF" stroke-width="1.4"/>
    <circle cx="178" cy="91" r="10" fill="url(#fbW)" stroke="#BCD3EE" stroke-width="1.6"/>
    <circle cx="168.5" cy="89.5" r="4.4" fill="url(#fbW)" stroke="#BCD3EE" stroke-width="1.4"/>
    <circle cx="178" cy="92" r="3" fill="#00D4FF" opacity=".85"/>
  </g>

  <!-- LEFT arm: floating segments, holding the tablet -->
  <g class="hand">
    {_joint(50, 150, 7)}
    {_capsule(41, 156, 28, 160, 9)}
    {_joint(21, 161, 5)}
    {_capsule(15, 158, 9, 150, 8)}
    <ellipse cx="6.5" cy="145" rx="7" ry="3.4" transform="rotate(-62 6.5 145)" fill="url(#fbB)" stroke="#FFFFFF" stroke-width="1.4"/>
    <circle cx="4" cy="134" r="10" fill="url(#fbW)" stroke="#BCD3EE" stroke-width="1.6"/>
    <circle cx="12.5" cy="127.5" r="4.2" fill="url(#fbW)" stroke="#BCD3EE" stroke-width="1.4"/>
    <circle cx="4" cy="135" r="3" fill="#00D4FF" opacity=".85"/>
  </g>
</g>
</svg>"""


def finbot_data_uri(badge="hello"):
    raw = finbot_svg_markup(badge).encode("utf-8")
    return "data:image/svg+xml;base64," + base64.b64encode(raw).decode("ascii")


def _typed_words(text, anim_id):
    """Turns 'Hi **FinBot**!' into per-word spans that 'print' one after another.
    Punctuation stuck to a bold phrase stays in the same span so it never wraps onto its own line."""
    items = []  # dicts: word, bold, gap (space after), tail (glued plain text)
    for part in re.split(r"(\*\*.+?\*\*)", text):
        if not part:
            continue
        bold = part.startswith("**") and part.endswith("**")
        body = part[2:-2] if bold else part
        for n, token in enumerate(re.findall(r"\S+\s*", body)):
            word = token.strip()
            gap = " " if token != word else ""
            glued = (not bold) and n == 0 and items and not body[:1].isspace() and items[-1]["gap"] == ""
            if glued:
                items[-1]["tail"] = word
                items[-1]["gap"] = gap
            else:
                items.append({"word": word, "bold": bold, "gap": gap, "tail": ""})
    spans = []
    for i, it in enumerate(items):
        inner = html_lib.escape(it["word"])
        if it["bold"]:
            inner = f'<b class="fbb">{inner}</b>'
        inner += html_lib.escape(it["tail"])
        delay = 0.35 + i * 0.065
        spans.append(f'<span class="fb-w" style="animation-name:fbtype{anim_id};animation-delay:{delay:.2f}s">{inner}</span>{it["gap"]}')
    return "".join(spans)


def finbot_stage_html(message, badge="hello", mode="hang", anim_id=0):
    """FinBot + tablet. mode='hang' anchors to a slide corner; mode='dock' is the fixed workspace overlay."""
    return _html(f"""
    <div class="fb-stage fb-{mode}" aria-label="FinBot, your finance helper">
        <div class="fb-float">
            <div class="fb-tablet" role="status" aria-live="polite">
                <div class="fb-screen">
                    <div class="fb-bar"><i class="fb-cam"></i><span class="fb-name">FinBot</span><span class="fb-live">&#9679; live</span></div>
                    <div class="fb-text">{_typed_words(message, anim_id)}<i class="fb-cursor"></i></div>
                </div>
                <i class="fb-home"></i>
            </div>
            <img class="fb-bot" alt="FinBot waving hello" src="{finbot_data_uri(badge)}"/>
        </div>
    </div>
    """)


def finbot_css():
    """Shared FinBot styling (tablet, typing animation, hang + dock placement)."""
    typing = "".join(
        f"@keyframes fbtype{i}{{from{{opacity:0;transform:translateY(3px)}}to{{opacity:1;transform:none}}}}"
        for i in range(8)
    )
    return (
        "<style>"
        + typing
        + """
        .fb-stage { position: relative; width: 430px; height: 240px; pointer-events: none; }
        .fb-float { position: absolute; inset: 0; animation: fb-bob 3.6s ease-in-out infinite; }
        .fb-bot { position: absolute; right: 0; top: 0; width: 220px; height: 240px; z-index: 2; display: block; }

        .fb-tablet {
            position: absolute; left: 0; bottom: 20px; width: 236px; z-index: 1; padding: 8px 8px 14px;
            background: linear-gradient(150deg, #1E293B, #0F172A); border-radius: 22px;
            box-shadow: 0 10px 15px -3px rgba(0,0,0,0.12), 0 4px 6px -2px rgba(0,0,0,0.06);
            transform: rotate(-2.5deg); transform-origin: 90% 55%;
        }
        .fb-screen { background: linear-gradient(180deg, #FFFFFF, #EEF9FF); border-radius: 15px; padding: 9px 26px 12px 12px; min-height: 108px; }
        .fb-bar { display: flex; align-items: center; gap: 7px; font-size: 9.5px; font-weight: 700; letter-spacing: .4px; color: #64748B; margin-bottom: 8px; }
        .fb-cam { width: 5px; height: 5px; border-radius: 50%; background: #CBD5E1; flex: none; }
        .fb-name { color: #0F172A; text-transform: uppercase; letter-spacing: 1px; }
        .fb-live { margin-left: auto; color: #00B8E6; }
        .fb-text { font-size: 12.5px; line-height: 1.5; color: #0F172A; font-weight: 500; }
        .fb-w { display: inline-block; opacity: 0; animation-duration: .2s; animation-fill-mode: both; animation-timing-function: ease-out; }
        .fb-w .fbb { font-weight: 800; color: #0A8FB0; }
        .fb-cursor { display: inline-block; width: 6px; height: 13px; margin-left: 2px; vertical-align: -2px; border-radius: 2px; background: #00D4FF; animation: fb-caret 1s steps(1) infinite; }
        .fb-home { position: absolute; left: 50%; bottom: 4px; width: 46px; height: 3px; margin-left: -23px; border-radius: 2px; background: #475569; }

        /* hang: pinned to the top-right corner of a slide (parent must be position:relative) */
        .fb-hang { position: absolute; top: -184px; right: 22px; z-index: 6; }
        /* dock: fixed overlay on the right-hand corner of the workspace */
        .fb-dock { position: fixed; right: 14px; bottom: 4px; z-index: 90; transform: scale(.8); transform-origin: right bottom; }
        .fb-dock .fb-text { font-size: 13.5px; }

        @keyframes fb-bob { 0%,100% { transform: translateY(0); } 50% { transform: translateY(-7px); } }
        @keyframes fb-caret { 0%,49% { opacity: 1; } 50%,100% { opacity: 0; } }

        @media (max-width: 1100px) { .fb-dock { display: none; } }
        @media (max-width: 820px) {
            .fb-hang { position: relative; top: auto; right: auto; margin: 0 auto 14px; transform: scale(.88); transform-origin: center top; }
        }
        @media (prefers-reduced-motion: reduce) {
            .fb-float, .fb-cursor { animation: none; }
            .fb-w { opacity: 1; animation: none; }
        }
        </style>
        """
    )


TOUR_FRAMES = [
    {
        "tag": "Welcome", "badge": "hello",
        "title": "Know how your business is <em>really</em> doing.",
        "sub": "FinSight turns everyday sales and expenses into plain-English signals. "
               "No spreadsheets to decode and no accounting degree needed.",
        "say": "Hi, I'm **FinBot**! I read your sales and expenses so you don't have to. Ready for a 30-second tour?",
        "visual": """
            <div class="pv-label">How it works</div>
            <div class="pv-step"><span class="pv-num">1</span><div><b>Add a sale or a bill</b><small>One line, about five seconds.</small></div></div>
            <div class="pv-step"><span class="pv-num">2</span><div><b>I sort and read it</b><small>Categories and totals update instantly.</small></div></div>
            <div class="pv-step"><span class="pv-num">3</span><div><b>You get clear answers</b><small>Healthy? Overspending? Tax coming up?</small></div></div>
        """,
    },
    {
        "tag": "Real-Time Mood Radar", "badge": "pulse",
        "title": "A mood ring for your money.",
        "sub": "The moment you log a transaction, FinSight compares revenue with expenses "
               "and tells you how your business feels: strong, steady, thin, or losing money.",
        "say": "This is my favorite part! Every time you add a number I check your **profit margin** and show a mood, "
               "from **Excellent momentum** to **Loss detected**. It updates live.",
        "visual": """
            <div class="pv-label">Sample preview</div>
            <div class="pv-mood">
                <div class="pv-emoji">&#128640;</div>
                <div><b>Excellent momentum</b><small>Your business is generating a strong return.</small></div>
            </div>
            <div class="pv-chips">
                <span>REVENUE <i class="up">&#8377;124,000</i></span>
                <span>EXPENSES <i class="down">&#8377;89,500</i></span>
                <span>MARGIN <i class="up">27.8%</i></span>
            </div>
        """,
    },
    {
        "tag": "Autonomous Alerts &amp; Anomaly Spotting", "badge": "alert",
        "title": "Problems spotted before they grow.",
        "sub": "FinSight quietly watches your ledger and budgets. When something looks unusual, "
               "it lands in your Action Center with a plain explanation of what changed.",
        "say": "I never take a day off! A bill much bigger than usual? A budget blown? I'll tell you **what happened and why**.",
        "visual": """
            <div class="pv-label">Sample alerts</div>
            <div class="pv-alert red"><b>Spike</b><small>Rent payment is far above its usual range.</small></div>
            <div class="pv-alert yellow"><b>Budget</b><small>Marketing has used 112% of this month's limit.</small></div>
            <div class="pv-alert green"><b>Milestone</b><small>Profit margin improved for the second month.</small></div>
        """,
    },
    {
        "tag": "Premium Tax &amp; Leak Profiling", "badge": "coin",
        "title": "Set tax money aside. Plug the leaks.",
        "sub": "Premium forecasts your upcoming GST and yearly income tax, and scans your ledger "
               "for accidental double entries that quietly drain cash.",
        "say": "With Premium I forecast **next month's GST** so you can save for it, and I hunt for **duplicate payments** in your ledger.",
        "visual": """
            <div class="pv-label">Sample premium insights <span class="pv-pill">PREMIUM</span></div>
            <div class="pv-kv"><small>Next 30-day GST forecast</small><b>&#8377;18,400</b></div>
            <div class="pv-kv"><small>Estimated annual income tax</small><b>&#8377;96,000</b></div>
            <div class="pv-alert red"><b>Leak scan</b><small>2 possible double entries found.</small></div>
        """,
    },
    {
        "tag": "Secure Workspace Creation", "badge": "shield",
        "title": "Create your secure workspace.",
        "sub": "Log in or sign up in seconds. Each business gets its own private workspace, "
               "and you decide who joins it.",
        "say": "You made it! Create a free account and I'll start reading your numbers. Passwords are **salted and hashed**.",
        "visual": "",
    },
]
TOUR_LAST = len(TOUR_FRAMES) - 1


def inject_landing_css():
    """Deck-only layout. Colors, borders, radius and shadow all come from the global tokens."""
    st.markdown(finbot_css(), unsafe_allow_html=True)
    st.markdown(
        """
        <style>
        [data-testid="stHeader"] { display: none; }
        [data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"] { display: none; }
        .block-container { max-width: 1060px; padding-top: 1.6rem; }

        .deck-brand { font-size: 22px; font-weight: 900; letter-spacing: -0.6px; color: var(--ink); padding-top: 6px; }
        .deck-brand span { color: var(--cyan); }

        /* Skip Tour: prominent outlined pill */
        .st-key-skip_tour { display: flex; justify-content: flex-end; }
        .st-key-skip_tour button { min-height: 0 !important; padding: 8px 22px !important; border-radius: 999px !important; border: 1.5px solid var(--ink) !important; }
        .st-key-skip_tour button p { font-size: 14px; font-weight: 800; text-decoration: underline; text-underline-offset: 4px; text-decoration-color: var(--cyan); }
        .st-key-skip_tour button:hover { background: var(--ink) !important; color: #FFFFFF !important; border-color: var(--ink) !important; }

        /* the slide: FinBot hangs off its top-right corner, so overflow must stay visible */
        .st-key-deck_frame {
            position: relative; overflow: visible !important; background: #FFFFFF; border: 1px solid var(--border);
            border-radius: 20px; padding: 46px 52px 26px 52px; min-height: 470px; margin: 150px 0 40px 0; box-shadow: var(--shadow);
            gap: .4rem;
        }
        .st-key-deck_frame [data-testid="stElementContainer"], .st-key-deck_frame [data-testid="stMarkdown"],
        .st-key-deck_frame [data-testid="stMarkdownContainer"] { overflow: visible !important; }
        .st-key-deck_frame > [data-testid="stElementContainer"]:last-child { margin-top: auto; }  /* pins the dots to the slide bottom */

        .deck-grid { display: grid; grid-template-columns: 1.05fr 1fr; gap: 44px; align-items: start; }
        .deck-copy { margin-top: 30px; }
        .deck-grid.solo { grid-template-columns: 1fr; padding-top: 76px; text-align: center; }
        .deck-grid.solo .deck-copy { margin-top: 0; }
        .deck-eyebrow { font-family: "JetBrains Mono", monospace; font-size: 11px; font-weight: 700; letter-spacing: 1.4px; text-transform: uppercase; color: var(--muted); margin-bottom: 14px; }
        .deck-eyebrow b { color: var(--ink); }
        .deck-title { font-size: 42px; line-height: 1.06; font-weight: 900; letter-spacing: -1.4px; color: var(--ink); margin: 0 0 16px 0; }
        .deck-title em { font-style: normal; background: linear-gradient(transparent 62%, rgba(0,212,255,0.38) 62%); }
        .deck-sub { font-size: 16px; line-height: 1.6; color: var(--muted); max-width: 470px; }
        .deck-grid.solo .deck-sub { margin: 0 auto; }

        .deck-visual { margin-top: 100px; background: var(--canvas-soft); border: 1px solid var(--border); border-radius: var(--radius); padding: 20px; }
        .pv-label { font-size: 10.5px; font-weight: 800; letter-spacing: 1.2px; text-transform: uppercase; color: var(--faint); margin-bottom: 12px; }
        .pv-pill { margin-left: 6px; padding: 2px 8px; border-radius: 999px; background: var(--ink); color: #FFFFFF; letter-spacing: .8px; }
        .pv-step, .pv-kv, .pv-alert, .pv-mood { background: #FFFFFF; border: 1px solid var(--border); border-radius: var(--radius); box-shadow: var(--shadow); }
        .pv-step { display: flex; gap: 14px; align-items: center; padding: 12px 14px; margin-bottom: 10px; }
        .pv-step:last-child, .pv-alert:last-child, .pv-kv:last-child { margin-bottom: 0; }
        .pv-num { width: 28px; height: 28px; flex: none; border-radius: 999px; background: var(--ink); color: #FFFFFF; font-weight: 800; font-size: 13px; display: flex; align-items: center; justify-content: center; }
        .deck-visual b { display: block; color: var(--ink); font-size: 14px; font-weight: 700; }
        .deck-visual small { display: block; color: var(--muted); font-size: 12.5px; margin-top: 2px; }
        .pv-mood { display: flex; gap: 16px; align-items: center; padding: 16px; border-color: rgba(16,185,129,0.35); }
        .pv-emoji { font-size: 44px; line-height: 1; }
        .pv-chips { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 12px; }
        .pv-chips span { padding: 6px 12px; background: #FFFFFF; border: 1px solid var(--border); border-radius: 10px; font-family: "JetBrains Mono", monospace; font-size: 11.5px; font-weight: 600; color: var(--ink-2); }
        .pv-chips i { font-style: normal; } .pv-chips .up { color: var(--green); } .pv-chips .down { color: var(--red); }
        .pv-alert { padding: 11px 14px; margin-bottom: 10px; border-left-width: 3px; }
        .pv-alert.red { border-left-color: #F43F5E; } .pv-alert.yellow { border-left-color: #F59E0B; } .pv-alert.green { border-left-color: #10B981; }
        .pv-kv { display: flex; justify-content: space-between; align-items: center; padding: 12px 14px; margin-bottom: 10px; }
        .pv-kv small { margin: 0; } .pv-kv b { font-family: "JetBrains Mono", monospace; font-size: 15px; }

        /* pagination dots: the active frame glows cyan */
        .deck-dots { display: flex; justify-content: center; align-items: center; gap: 10px; margin-top: 26px; }
        .deck-dot { width: 8px; height: 8px; border-radius: 999px; background: #CBD5E1; transition: all .3s ease; }
        .deck-dot.active { width: 26px; background: var(--cyan); box-shadow: 0 0 12px rgba(0,212,255,0.7); }
        .st-key-tour_next button { min-width: 170px; }

        @media (max-width: 820px) {
            .st-key-deck_frame { padding: 28px 20px 22px 20px; margin: 16px 0 30px 0; }
            .deck-grid { grid-template-columns: 1fr; gap: 24px; }
            .deck-grid.solo { padding-top: 0; }
            .deck-visual { margin-top: 0; }
            .deck-title { font-size: 32px; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _tour_go(step=0, to=None):
    """Button callback: runs before the next rerun, so no st.rerun() loop is needed."""
    current = st.session_state.tour_frame
    target = to if to is not None else current + step
    st.session_state.tour_frame = max(0, min(TOUR_LAST, target))


def _slide_html(idx):
    frame = TOUR_FRAMES[idx]
    stage = finbot_stage_html(frame["say"], frame["badge"], "hang", idx)
    eyebrow = f'<div class="deck-eyebrow"><b>Frame {idx + 1:02d} / {len(TOUR_FRAMES):02d}</b> &nbsp;·&nbsp; {frame["tag"]}</div>'
    copy = f'<div class="deck-copy">{eyebrow}<div class="deck-title">{frame["title"]}</div><div class="deck-sub">{frame["sub"]}</div></div>'
    if idx == TOUR_LAST:
        body = f'<div class="deck-grid solo">{copy}</div>'
    else:
        body = f'<div class="deck-grid">{copy}<div class="deck-visual">{frame["visual"]}</div></div>'
    return _html(stage + body)


def _dots_html(idx):
    dots = "".join(f'<span class="deck-dot{" active" if j == idx else ""}"></span>' for j in range(len(TOUR_FRAMES)))
    return f'<div class="deck-dots" role="presentation" aria-label="Frame {idx + 1} of {len(TOUR_FRAMES)}">{dots}</div>'


def render_finbot_dock(message, badge="hello"):
    """Fixed bottom-right FinBot overlay for the workspace (pointer-events are off, so it never blocks clicks)."""
    st.markdown(finbot_css(), unsafe_allow_html=True)
    st.markdown(finbot_stage_html(message, badge, "dock", 5), unsafe_allow_html=True)


def render_auth_forms():
    """Log in / create account. Logic is unchanged from the original gateway."""
    tab_login, tab_signup = st.tabs(["Log in", "Create account"])

    with tab_login:
        with st.form("login_form"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Log in", use_container_width=True, type="primary")

        if submitted:
            user = authenticate(username, password)
            if user:
                st.session_state.user = user
                workspaces = get_user_workspaces(user["id"])
                if workspaces:
                    st.session_state.workspace_id = workspaces[0]["id"]
                st.rerun()
            else:
                st.error("Incorrect username or password.")

    with tab_signup:
        with st.form("signup_form"):
            new_username = st.text_input("Choose a username")
            new_password = st.text_input("Choose a password", type="password")
            confirm_password = st.text_input("Confirm password", type="password")
            submitted_signup = st.form_submit_button("Create account", use_container_width=True, type="primary")

        if submitted_signup:
            if not new_username.strip() or not new_password:
                st.warning("Username and password are required.")
            elif new_password != confirm_password:
                st.warning("Passwords do not match.")
            elif len(new_password) < 6:
                st.warning("Password should be at least 6 characters.")
            else:
                ok, message = create_user(new_username, new_password)
                if ok:
                    st.success(message + " You can log in now.")
                else:
                    st.error(message)


def render_gateway():
    """Single-pass gateway: header -> frame (+ forms on the last frame) -> nav row."""
    inject_landing_css()
    idx = max(0, min(TOUR_LAST, st.session_state.tour_frame))
    st.session_state.tour_frame = idx

    head_left, head_right = st.columns([5, 1.6], vertical_alignment="center")
    with head_left:
        st.markdown('<div class="deck-brand"><span>◈</span> FinSight</div>', unsafe_allow_html=True)
    with head_right:
        if idx < TOUR_LAST:
            st.button("Skip Tour ➔", key="skip_tour", on_click=_tour_go, kwargs={"to": TOUR_LAST})

    with st.container(key="deck_frame"):
        st.markdown(_slide_html(idx), unsafe_allow_html=True)
        if idx == TOUR_LAST:  # registration / login only appear on Frame 5
            _, form_col, _ = st.columns([1, 2.2, 1])
            with form_col:
                render_auth_forms()
        st.markdown(_dots_html(idx), unsafe_allow_html=True)

    with st.container(key="deck_nav"):
        prev_col, _, next_col = st.columns([1.3, 4, 1.7])
        with prev_col:
            st.button("← Previous", key="tour_prev", disabled=(idx == 0), on_click=_tour_go, kwargs={"step": -1})
        with next_col:
            if idx < TOUR_LAST:
                st.button("Next Frame →", key="tour_next", type="primary", on_click=_tour_go, kwargs={"step": 1})

if st.session_state.user is None:
    render_gateway()
    st.stop()

current_user = get_user_by_id(st.session_state.user["id"])
plan = current_user["plan"]
limits = PLAN_LIMITS[plan]


# ============================================================
# WORKSPACE SETUP
# ============================================================

workspaces = get_user_workspaces(current_user["id"])

if not workspaces:
    st.title("◈ FinSight")
    st.subheader("Create your first workspace")
    render_finbot_dock("Let's set up your first workspace! Add your **business name** and I'll start tracking.", "hello")
    st.caption("A workspace holds one business's transactions, budgets, and goals.")

    with st.form("create_workspace_form"):
        ws_name = st.text_input("Business name", placeholder="ABC Traders")
        ws_category = st.selectbox("Business category", BUSINESS_CATEGORIES, index=BUSINESS_CATEGORIES.index("Other"))
        ws_currency = st.selectbox("Currency", list(CURRENCIES.keys()))
        ws_gst = st.number_input("GST / tax rate (%)", min_value=0.0, max_value=100.0, value=18.0, step=0.5)
        create_submit = st.form_submit_button("Create workspace", use_container_width=True)

    if create_submit:
        if not ws_name.strip():
            st.warning("Please give your workspace a name.")
        else:
            new_id = create_workspace(current_user["id"], ws_name, ws_currency, ws_gst, ws_category)
            st.session_state.workspace_id = new_id
            st.rerun()
    st.stop()

if st.session_state.workspace_id not in [w["id"] for w in workspaces]:
    st.session_state.workspace_id = workspaces[0]["id"]

workspace = next(w for w in workspaces if w["id"] == st.session_state.workspace_id)
role = workspace["role"]
can_edit = role in ("owner", "accountant")
is_owner = role == "owner"
currency_symbol = CURRENCIES.get(workspace["currency"], "₹")
business_category = workspace["business_category"] or "Other"
workspace_categories = get_workspace_categories(business_category)


def money(value):
    return f"{currency_symbol}{value:,.0f}"


# Auto-generate any due recurring transactions for this workspace.
generated_count = run_due_recurring_transactions(workspace["id"], current_user["id"])


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.title("◈ FinSight")
    plan_class = "plan-pro" if plan == "pro" else "plan-free"
    st.markdown(f"<span class='plan-badge {plan_class}'>{plan} plan</span>", unsafe_allow_html=True)
    st.caption(f"Logged in as **{current_user['username']}**")

    st.toggle("Show FinBot", value=True, key="show_finbot")

    if st.button("Log out", use_container_width=True):
        st.session_state.user = None
        st.session_state.workspace_id = None
        st.session_state.tour_frame = TOUR_LAST  # skip the tour for returning users
        st.rerun()

    st.divider()
    st.subheader("Workspace")

    ws_names = {w["id"]: f"{w['name']} ({w['role']})" for w in workspaces}
    selected_ws = st.selectbox(
        "Active workspace",
        options=list(ws_names.keys()),
        format_func=lambda wid: ws_names[wid],
        index=list(ws_names.keys()).index(workspace["id"]),
    )
    if selected_ws != st.session_state.workspace_id:
        st.session_state.workspace_id = selected_ws
        st.rerun()

    if len(workspaces) < (limits["max_workspaces"] or 999) or plan == "pro":
        with st.expander("＋ New workspace"):
            with st.form("new_ws_form"):
                nw_name = st.text_input("Business name")
                nw_category = st.selectbox("Business category", BUSINESS_CATEGORIES, index=BUSINESS_CATEGORIES.index("Other"), key="nw_category")
                nw_currency = st.selectbox("Currency", list(CURRENCIES.keys()), key="nw_currency")
                nw_gst = st.number_input("GST rate (%)", value=18.0, key="nw_gst")
                nw_submit = st.form_submit_button("Create")
            if nw_submit and nw_name.strip():
                nid = create_workspace(current_user["id"], nw_name, nw_currency, nw_gst, nw_category)
                st.session_state.workspace_id = nid
                st.rerun()
    else:
        st.caption("Free plan is limited to 1 workspace. Upgrade to add more.")

    if plan != "pro":
        st.info("🔓 Free plan: 1 workspace, 250 transactions, no exports/invites.")
        if st.button("Simulate upgrade to Pro", use_container_width=True):
            set_user_plan(current_user["id"], "pro")
            st.rerun()

    st.divider()
    st.subheader("Team")

    members = get_workspace_members(workspace["id"])
    for m in members:
        st.caption(f"• {m['username']} — {m['role']}")

    if is_owner:
        if limits["invites"]:
            with st.expander("＋ Invite member"):
                with st.form("invite_form"):
                    inv_username = st.text_input("Username")
                    inv_role = st.selectbox("Role", ["accountant", "viewer"])
                    inv_submit = st.form_submit_button("Add to workspace")
                if inv_submit:
                    ok, message = invite_member(workspace["id"], inv_username, inv_role)
                    (st.success if ok else st.error)(message)
        else:
            st.caption("Inviting teammates needs the Pro plan.")

    st.divider()
    st.subheader("Filters")
    period = st.selectbox("Time window", ["All time", "This month", "Last 7 days"])

    st.divider()
    st.subheader("Data center")

    uploaded_file = st.file_uploader("Import bank/transactions CSV", type=["csv"])
    if uploaded_file is not None:
        try:
            raw = pd.read_csv(uploaded_file)
            st.caption("Map your columns:")
            cols = list(raw.columns)
            col_date = st.selectbox("Date column", cols, index=cols.index(guess_column(cols, ["date"]) or cols[0]))
            col_desc = st.selectbox("Description column", cols, index=cols.index(guess_column(cols, ["description", "narration", "details"]) or cols[0]))
            col_amount = st.selectbox("Amount column", cols, index=cols.index(guess_column(cols, ["amount", "value"]) or cols[0]))
            col_type = st.selectbox("Type column (income/expense)", ["(infer from sign)"] + cols)

            if st.button("Import mapped rows", use_container_width=True):
                learned = get_learned_rules(workspace["id"])
                imported_n = 0
                for _, r in raw.iterrows():
                    amt = pd.to_numeric(r[col_amount], errors="coerce")
                    if pd.isna(amt):
                        continue
                    if col_type == "(infer from sign)":
                        tx_type = "income" if amt >= 0 else "expense"
                    else:
                        tx_type = str(r[col_type]).strip().lower()
                        if tx_type not in ("income", "expense"):
                            continue
                    desc = str(r[col_desc]).strip()
                    insert_transaction(
                        workspace["id"], r[col_date], desc, abs(amt), tx_type, "",
                        category_from_text(desc, learned), 0, "", current_user["id"],
                    )
                    imported_n += 1
                st.success(f"Imported {imported_n} transactions.")
                st.rerun()
        except Exception as error:
            st.error(f"Import failed: {error}")

    st.divider()
    st.success("FinSight engine online")
    if generated_count:
        st.info(f"⏱ {generated_count} recurring transaction(s) auto-generated.")


# ============================================================
# LOAD DATA
# ============================================================

all_tx = load_transactions(workspace["id"])
learned_rules = get_learned_rules(workspace["id"])

if not all_tx.empty and (all_tx["category"].isna() | (all_tx["category"] == "")).any():
    mask = all_tx["category"].isna() | (all_tx["category"] == "")
    all_tx.loc[mask, "category"] = all_tx.loc[mask, "description"].apply(
        lambda d: category_from_text(d, learned_rules)
    )


def filter_data(tx_df, period_label):
    if tx_df.empty:
        return tx_df.copy()
    latest_date = tx_df["date"].max()
    if period_label == "This month":
        return tx_df[(tx_df["date"].dt.year == latest_date.year) & (tx_df["date"].dt.month == latest_date.month)].copy()
    if period_label == "Last 7 days":
        start = latest_date - pd.Timedelta(days=6)
        return tx_df[tx_df["date"] >= start].copy()
    return tx_df.copy()


def previous_period_df(tx_df, period_label):
    if tx_df.empty:
        return tx_df.copy()

    latest = tx_df["date"].max()

    if period_label == "This month":
        current_start = latest.replace(day=1)
        previous_end = current_start - pd.Timedelta(days=1)
        previous_start = previous_end.replace(day=1)
        return tx_df[(tx_df["date"] >= previous_start) & (tx_df["date"] <= previous_end)].copy()

    if period_label == "Last 7 days":
        current_start = latest - pd.Timedelta(days=6)
        previous_end = current_start - pd.Timedelta(days=1)
        previous_start = previous_end - pd.Timedelta(days=6)
        return tx_df[(tx_df["date"] >= previous_start) & (tx_df["date"] <= previous_end)].copy()

    return tx_df.iloc[0:0].copy()


def pct_change(current, previous):
    if previous == 0:
        return None
    return (current - previous) / previous * 100


# ---- single source of truth for the current + previous period ----

view = filter_data(all_tx, period)
previous_view = previous_period_df(all_tx, period)

income = view.loc[view["type"] == "income", "amount"].sum()
expenses = view.loc[view["type"] == "expense", "amount"].sum()
profit = income - expenses
margin = (profit / income * 100) if income > 0 else 0

prev_income = previous_view.loc[previous_view["type"] == "income", "amount"].sum()
prev_expenses = previous_view.loc[previous_view["type"] == "expense", "amount"].sum()
prev_profit = prev_income - prev_expenses

revenue_change = pct_change(income, prev_income)
expense_change = pct_change(expenses, prev_expenses)
profit_change = pct_change(profit, prev_profit) if prev_profit != 0 else None

mood_emoji, mood_text, mood_description, mood_key = business_mood(income, expenses, profit, margin)
mood_style = MOOD_STYLES[mood_key]


# ============================================================
# EXECUTIVE DASHBOARD — premium first impression
# ============================================================

nav1, nav2, nav3 = st.columns([5.5, 1.3, 1.3])
with nav1:
    st.caption("● LIVE FINANCIAL OVERVIEW")
with nav2:
    st.markdown(
        f"<div class='plan-badge plan-{'pro' if plan == 'pro' else 'free'}'>{plan.upper()}</div>",
        unsafe_allow_html=True,
    )
with nav3:
    st.caption(f"{workspace['currency']} · {role}")

st.markdown(
    f"""
    <div style='display:flex;justify-content:space-between;gap:20px;align-items:flex-end;margin-bottom:12px;'>
        <div>
            <h1 style='font-size:42px;margin:0;line-height:1.08;'>Your business,<br><span class='hl'>seen clearly.</span></h1>
            <p style='color:#64748B;margin-top:10px;font-size:14px;'>
                {workspace['name']} · <span class='category-badge'>{business_category}</span> · {period} · a control room for the money behind the business.
            </p>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------- action bar ----------
if can_edit:
    st.markdown("##### ⚡ Quick actions")
    ac1, ac2, ac3, ac4 = st.columns([2.7, 1.35, 1.35, 1.1])
    with ac1:
        quick_desc = st.text_input(
            "Description",
            key="dashboard_quick_desc",
            label_visibility="collapsed",
            placeholder="Customer payment, rent, supplier...",
        )
    with ac2:
        quick_amount = st.number_input(
            "Amount",
            key="dashboard_quick_amount",
            min_value=0.0,
            step=100.0,
            label_visibility="collapsed",
        )
    with ac3:
        quick_type = st.selectbox(
            "Type",
            ["income", "expense"],
            key="dashboard_quick_type",
            label_visibility="collapsed",
        )
    with ac4:
        if st.button("＋ Add", use_container_width=True, key="dashboard_quick_add"):
            if not quick_desc.strip() or quick_amount <= 0:
                st.warning("Enter a description and amount.")
            elif limits["max_transactions"] and transaction_count(workspace["id"]) >= limits["max_transactions"]:
                st.error("Free plan transaction limit reached.")
            else:
                cat = category_from_text(quick_desc, learned_rules)
                insert_transaction(
                    workspace["id"], date.today(), quick_desc, quick_amount, quick_type,
                    "", cat, 0, "", current_user["id"],
                )
                st.success("Transaction added.")
                st.rerun()

st.divider()

# ---------- financial mood hero ----------
profit_arrow = "▲" if profit >= 0 else "▼"
profit_class = "ticker-up" if profit >= 0 else "ticker-down"
margin_class = "ticker-up" if margin >= 0 else "ticker-down"

st.markdown(
    f"""
    <div class='mood-hero' style='--mood-color:{mood_style['color']};--mood-bg-a:{mood_style['bg_a']};
        --mood-bg-b:{mood_style['bg_b']};--mood-border:{mood_style['border']};--mood-glow:{mood_style['glow']};
        --mood-badge-bg:{mood_style['badge_bg']};'>
        <div class='mood-emoji'>{mood_emoji}</div>
        <div style='flex:1;'>
            <div class='mood-text-title'>{mood_text}</div>
            <div class='mood-text-sub'>{mood_description}</div>
            <div class='mood-badge'>Net position · {money(profit)} · {margin:.1f}% margin</div>
        </div>
        <div style='text-align:right;min-width:180px;'>
            <div style='color:#64748B;font-size:11px;font-weight:700;text-transform:uppercase;'>Today's signal</div>
            <div style='font-size:30px;font-weight:900;color:{mood_style['color']};'>{profit_arrow} {abs(profit):,.0f}</div>
        </div>
    </div>
    <div class='ticker-strip'>
        <div class='ticker-chip'>REVENUE&nbsp; <span class='ticker-up'>{money(income)}</span></div>
        <div class='ticker-chip'>EXPENSES&nbsp; <span class='ticker-down'>{money(expenses)}</span></div>
        <div class='ticker-chip'>NET&nbsp; <span class='{profit_class}'>{profit_arrow} {money(abs(profit))}</span></div>
        <div class='ticker-chip'>MARGIN&nbsp; <span class='{margin_class}'>{margin:.1f}%</span></div>
        <div class='ticker-chip'>TRANSACTIONS&nbsp; <span>{len(view):,}</span></div>
    </div>
    """,
    unsafe_allow_html=True,
)

digest_message = (
    f"{workspace['name']} — {period} update\n"
    f"Revenue: {money(income)}\n"
    f"Expenses: {money(expenses)}\n"
    f"Net: {money(profit)} ({margin:.1f}% margin)\n"
    f"Mood: {mood_text}"
)
st.link_button("💬 Share this update on WhatsApp", build_whatsapp_link("", digest_message))

st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)


# ============================================================
# ACTION CENTER + BUSINESS HEALTH SCORE
# "FinSight VNext — The Autonomous Intelligence & Decision Layer"
# All computed here, on the fly, from all_tx + budgets — no new
# tables, fully defensive against thin/empty data.
# ============================================================

budgets_df_full = get_budgets(workspace["id"])
anomalies = detect_anomalies(all_tx, budgets_df_full)
health_score = compute_health_score(all_tx, budgets_df_full, st.session_state.pulse_granularity)
action_items = build_action_center(anomalies, health_score)

ac_header_col, ac_toggle_col = st.columns([3, 1.2])
with ac_header_col:
    st.markdown('<div class="section" style="margin-top:10px;">🧭 Action Center</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="helper">Automated alerts from the anomaly engine, health score, and budgets — highest priority first.</div>',
        unsafe_allow_html=True,
    )
with ac_toggle_col:
    new_granularity = st.radio(
        "Pulse period",
        ["Month", "Quarter"],
        index=["Month", "Quarter"].index(st.session_state.pulse_granularity),
        horizontal=True,
        key="pulse_granularity_radio",
        label_visibility="collapsed",
    )
    if new_granularity != st.session_state.pulse_granularity:
        st.session_state.pulse_granularity = new_granularity
        st.rerun()

if all_tx.empty:
    st.info("Not enough historical data yet to calculate trends — add more transactions to activate insights!")
elif not action_items:
    st.markdown(
        "<div class='insight action-item-green'>🟢 <b>All clear.</b> "
        "<span>No critical anomalies detected right now — keep adding transactions to sharpen these insights.</span></div>",
        unsafe_allow_html=True,
    )
else:
    for item in action_items[:6]:
        label = "Action Required" if item["severity"] == "red" else ("Attention Needed" if item["severity"] == "yellow" else "Positive Milestone")
        st.markdown(
            f"<div class='insight action-item-{item['severity']}'>{item['icon']} <b>{label}:</b> <span>{item['message']}</span></div>",
            unsafe_allow_html=True,
        )

st.markdown('<div class="section" style="margin-top:22px;">🩺 Business Health Score</div>', unsafe_allow_html=True)
st.markdown(
    f'<div class="helper">{st.session_state.pulse_granularity}-over-{st.session_state.pulse_granularity.lower()} view — four weighted components, 25 points each.</div>',
    unsafe_allow_html=True,
)

if health_score.get("insufficient_data"):
    st.info("Not enough historical data yet to calculate a health score — add more transactions to activate this.")
else:
    hs_col1, hs_col2 = st.columns([1, 2.3], gap="large")
    with hs_col1:
        score_color = "#059669" if health_score["total"] >= 75 else ("#D97706" if health_score["total"] >= 50 else "#E11D48")
        st.markdown(
            f"""
            <div class='glass-card health-score-ring'>
                <div class='glass-label'>Overall Score</div>
                <div style='font-size:48px;font-weight:900;color:{score_color};font-family:"JetBrains Mono",monospace;'>{health_score['total']}</div>
                <div class='glass-sub'>out of 100</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with hs_col2:
        comps = health_score["components"]
        for label, val in [
            ("Revenue Momentum", comps["revenue_momentum"]),
            ("Profitability", comps["profitability"]),
            ("Expense Control", comps["expense_control"]),
            ("Runway & Cash Flow", comps["runway"]),
        ]:
            st.markdown(f"**{label}** — {val}/25")
            st.progress(min(max(val / 25, 0.0), 1.0))
        if health_score.get("runway_months") is not None:
            st.caption(f"Estimated runway: {health_score['runway_months']:.1f} months of cash buffer at the current burn rate.")

# ---------- What & Why business pulse ----------
pulse = compute_business_pulse(all_tx, st.session_state.pulse_granularity)
st.markdown('<div class="section" style="margin-top:22px;">📊 What &amp; Why</div>', unsafe_allow_html=True)

if pulse is None:
    st.info("Not enough historical data yet to compare periods — add more transactions to activate this.")
else:
    what = pulse["what"]
    pulse_cols = st.columns(4, gap="medium")
    pulse_fields = [
        ("Revenue", what["revenue"], False),
        ("Expenses", what["expenses"], False),
        ("Net Profit", what["profit"], False),
        ("Profit Margin", what["margin"], True),
    ]
    for col, (label, field, is_pct) in zip(pulse_cols, pulse_fields):
        with col:
            display_val = f"{field['current']:.1f}%" if is_pct else money(field["current"])
            arrow_color = "#059669" if field["direction"] == "↑" else ("#E11D48" if field["direction"] == "↓" else "#64748B")
            st.markdown(
                f"""
                <div class='glass-card'>
                    <div class='glass-label'>{label}</div>
                    <div class='glass-value' style='color:{arrow_color};'>{field['direction']} {display_val}</div>
                    <div class='glass-sub'>vs previous {st.session_state.pulse_granularity.lower()}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    if pulse["why"]:
        st.markdown(f"<div class='helper' style='margin-top:14px;'>Why expenses moved this {st.session_state.pulse_granularity.lower()}:</div>", unsafe_allow_html=True)
        why_rows = "".join(
            f"<div class='why-row'><span>{w['category']}</span>"
            f"<span><b>{money(abs(w['amount']))}</b> · {w['pct_of_change']:.0f}% of the change</span></div>"
            for w in pulse["why"]
        )
        st.markdown(f"<div class='glass-card' style='min-height:auto;'>{why_rows}</div>", unsafe_allow_html=True)

st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# ---------- key metric cards ----------
metric_cols = st.columns(4, gap="medium")
metric_data = [
    ("↗ Revenue", income, revenue_change, "up"),
    ("↘ Expenses", expenses, expense_change, "down"),
    ("◇ Net result", profit, profit_change, "up" if profit >= 0 else "down"),
    ("◌ Margin", margin, None, "up" if margin >= 0 else "down"),
]

for col, (label, value, change, direction) in zip(metric_cols, metric_data):
    with col:
        display_value = f"{value:.1f}%" if label == "◌ Margin" else money(value)
        if change is None:
            change_text = "—"
        else:
            sign = "↑" if change >= 0 else "↓"
            change_text = f"{sign} {abs(change):.1f}%"
        st.markdown(
            f"""
            <div class='glass-card' style='min-height:112px;'>
                <div class='glass-label'>{label}</div>
                <div class='glass-value'>{display_value}</div>
                <div class='glass-sub'><span class='{'ticker-up' if direction == 'up' else 'ticker-down'}'>{change_text}</span> vs previous period</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

st.markdown('<div class="section" style="margin-top:28px;">📈 Financial Pulse</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="helper">The line is the running financial position. Green means the position improved; red means it moved lower. Purple is the trend projection.</div>',
    unsafe_allow_html=True,
)

if view.empty:
    st.info("No transactions for this period. Add your first transaction above.")
else:
    daily = build_stock_dataframe(view)
    forecast = forecast_next_days(daily, days=7)

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=daily["date"],
            y=daily["close"],
            mode="lines+markers",
            line=dict(color="#10B981", width=3),  # Updated to Neo-Mint
            marker=dict(size=7, color="#FFFFFF", line=dict(width=2, color="#10B981")),
            fill="tozeroy",
            fillcolor="rgba(16, 185, 129, 0.04)",  # Soft clean mint glow tint
            name="Position",
            hovertemplate="<b>%{x|%d %b}</b><br>Position: " + currency_symbol + "%{y:,.0f}<extra></extra>",
        )
    )


    if forecast is not None:
        forecast_x = pd.concat([daily[["date", "close"]].tail(1), forecast], ignore_index=True)
        fig.add_trace(
            go.Scatter(
                x=forecast_x["date"],
                y=forecast_x["close"],
                mode="lines",
                line=dict(color="#64748B", width=2.5, dash="dot"),
                name="7-day trend",
                hovertemplate="<b>%{x|%d %b}</b><br>Projected: " + currency_symbol + "%{y:,.0f}<extra></extra>",
            )
        )

    fig.add_hline(y=0, line_width=1, line_dash="dot", line_color="rgba(17,24,39,0.15)")
    fig.update_layout(
        height=500,
        margin=dict(l=5, r=5, t=10, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color="#64748B"),
        hovermode="x unified",
        showlegend=False,
        hoverlabel=dict(bgcolor="#FFFFFF", bordercolor="#E2E8F0", font=dict(family="Inter", color="#0F172A")),
    )
    fig.update_xaxes(showgrid=False, zeroline=False, fixedrange=True)
    fig.update_yaxes(
        showgrid=True,
        gridcolor="rgba(17,24,39,0.06)",
        zeroline=False,
        fixedrange=True,
        tickprefix=currency_symbol,
        separatethousands=True,
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False, "responsive": True})

# ---------- insight command center ----------
st.markdown('<div class="section">🧠 What changed?</div>', unsafe_allow_html=True)

insight_cards = []
if revenue_change is not None:
    revenue_arrow = "📈" if revenue_change >= 0 else "📉"
    insight_cards.append(
        f"{revenue_arrow} <b>Revenue</b> {'increased' if revenue_change >= 0 else 'decreased'} <span>{abs(revenue_change):.1f}%</span> versus the previous period."
    )
if expense_change is not None:
    expense_arrow = "⚠️" if expense_change > 0 else "✅"
    insight_cards.append(
        f"{expense_arrow} <b>Expenses</b> {'rose' if expense_change > 0 else 'fell'} <span>{abs(expense_change):.1f}%</span> versus the previous period."
    )
if not view.empty:
    expense_data = view[view["type"] == "expense"].copy()
    if not expense_data.empty:
        top_expense = expense_data.groupby("category")["amount"].sum().sort_values(ascending=False)
        cat = top_expense.index[0]
        val = top_expense.iloc[0]
        insight_cards.append(
            f"🔥 <b>Biggest cost</b> is <span>{cat} · {money(val)}</span>."
        )
if margin >= 25:
    insight_cards.append("🚀 <b>Margin signal:</b> <span>strong profitability buffer.</span>")
elif margin >= 10:
    insight_cards.append("😎 <b>Margin signal:</b> <span>healthy, but keep watching costs.</span>")
elif margin >= 0:
    insight_cards.append("😐 <b>Margin signal:</b> <span>thin buffer — small costs matter.</span>")
else:
    insight_cards.append("🚨 <b>Margin signal:</b> <span>expenses currently exceed revenue.</span>")

icols = st.columns(2)
for i, text in enumerate(insight_cards[:4]):
    with icols[i % 2]:
        st.markdown(f"<div class='insight'>{text}</div>", unsafe_allow_html=True)

# ---------- compact business overview ----------
st.markdown('<div class="section">🎯 Business at a glance</div>', unsafe_allow_html=True)

category_kpi = compute_category_kpi(business_category, view, income)
mini_cols = st.columns(4 if category_kpi else 3, gap="medium")

with mini_cols[0]:
    st.markdown(
        f"""
        <div class='glass-card'>
            <div class='glass-label'>Cash pressure</div>
            <div class='glass-value'>{'LOW' if expenses < income else 'HIGH'}</div>
            <div class='glass-sub'>{money(expenses)} spent in the selected window</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with mini_cols[1]:
    avg_tx = view["amount"].mean() if not view.empty else 0
    st.markdown(
        f"""
        <div class='glass-card'>
            <div class='glass-label'>Average transaction</div>
            <div class='glass-value'>{money(avg_tx)}</div>
            <div class='glass-sub'>{len(view):,} transactions tracked</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with mini_cols[2]:
    latest_activity = view["date"].max().strftime("%d %b %Y") if not view.empty else "—"
    st.markdown(
        f"""
        <div class='glass-card'>
            <div class='glass-label'>Latest activity</div>
            <div class='glass-value'>{latest_activity}</div>
            <div class='glass-sub'>Most recent transaction in this view</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

if category_kpi:
    kpi_label, kpi_value, kpi_sub = category_kpi
    with mini_cols[3]:
        st.markdown(
            f"""
            <div class='glass-card'>
                <div class='glass-label'>{kpi_label} <span style='opacity:.6'>· {business_category}</span></div>
                <div class='glass-value'>{kpi_value}</div>
                <div class='glass-sub'>{kpi_sub}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

# ---------- ✨ Automatic Highlights ----------
st.markdown('<div class="section" style="margin-top:22px;">✨ Automatic Highlights</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="helper">3–5 of the most useful things happening in your numbers right now — detected automatically, nothing invented.</div>',
    unsafe_allow_html=True,
)

highlights_unpaid_df = get_unpaid_invoices(workspace["id"])
highlights_recurring_df = get_recurring_rules(workspace["id"])
automatic_highlights = compute_automatic_highlights(
    view, previous_view, all_tx, pulse, category_kpi, business_category,
    highlights_unpaid_df, budgets_df_full, highlights_recurring_df,
    income, expenses, workspace["id"],
)

if not automatic_highlights:
    st.info("No major changes detected yet. Keep adding transactions to build stronger insights.")
else:
    hl_cols = st.columns(min(len(automatic_highlights), 3), gap="medium")
    for i, h in enumerate(automatic_highlights):
        with hl_cols[i % len(hl_cols)]:
            st.markdown(
                f"""
                <div class='glass-card' style='min-height:128px;'>
                    <div class='glass-label'>{h['icon']} {h['title']}</div>
                    <div class='glass-sub' style='margin-top:8px;font-size:13.5px;color:#111827;'>{h['text']}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

st.divider()


# ============================================================
# BUDGETS
# ============================================================

st.header("Budgets")
st.caption("Set a monthly cap per category and see how this month is tracking.")

budgets_df = get_budgets(workspace["id"])
latest_all = all_tx["date"].max() if not all_tx.empty else pd.Timestamp(date.today())
this_month_tx = all_tx[
    (all_tx["date"].dt.year == latest_all.year) & (all_tx["date"].dt.month == latest_all.month)
] if not all_tx.empty else all_tx

if budgets_df.empty:
    st.info("No budgets set yet. Add one below.")
else:
    spend_by_cat = this_month_tx[this_month_tx["type"] == "expense"].groupby("category")["amount"].sum()
    b_cols = st.columns(min(3, len(budgets_df)) or 1)
    for i, (_, b) in enumerate(budgets_df.iterrows()):
        spent = spend_by_cat.get(b["category"], 0.0)
        pct = min(spent / b["monthly_limit"], 1.5) if b["monthly_limit"] > 0 else 0
        color = "🟢" if pct < 0.75 else ("🟡" if pct < 1.0 else "🔴")
        with b_cols[i % len(b_cols)]:
            st.markdown(f"**{color} {b['category']}**")
            st.progress(min(pct, 1.0))
            st.caption(f"{money(spent)} of {money(b['monthly_limit'])} this month")
            if pct >= 1.0:
                st.warning(f"Over budget by {money(spent - b['monthly_limit'])}.")

if can_edit:
    with st.expander("＋ Set / update a budget"):
        with st.form("budget_form"):
            b_cat = st.selectbox("Category", workspace_categories)
            b_limit = st.number_input(f"Monthly limit ({currency_symbol})", min_value=0.0, step=500.0)
            b_submit = st.form_submit_button("Save budget")
        if b_submit:
            set_budget(workspace["id"], b_cat, b_limit)
            st.success("Budget saved.")
            st.rerun()

st.divider()


# ============================================================
# GOALS
# ============================================================

st.header("Savings Goals")
st.caption("Track progress toward a target amount.")

goals_df = get_goals(workspace["id"])

if goals_df.empty:
    st.info("No goals yet. Add one below.")
else:
    g_cols = st.columns(min(3, len(goals_df)) or 1)
    for i, (_, g) in enumerate(goals_df.iterrows()):
        pct = min(g["saved_amount"] / g["target_amount"], 1.0) if g["target_amount"] > 0 else 0
        with g_cols[i % len(g_cols)]:
            st.markdown(f"**🎯 {g['name']}**")
            st.progress(pct)
            st.caption(f"{money(g['saved_amount'])} of {money(g['target_amount'])} · by {g['target_date']}")
            if can_edit:
                add_amt = st.number_input("Add contribution", min_value=0.0, step=100.0, key=f"goal_{g['id']}")
                if st.button("Contribute", key=f"goal_btn_{g['id']}"):
                    if add_amt > 0:
                        contribute_to_goal(g["id"], add_amt)
                        st.rerun()

if can_edit:
    with st.expander("＋ New goal"):
        with st.form("goal_form"):
            g_name = st.text_input("Goal name", placeholder="Emergency fund")
            g_target = st.number_input(f"Target amount ({currency_symbol})", min_value=0.0, step=1000.0)
            g_date = st.date_input("Target date", value=date.today() + timedelta(days=90))
            g_submit = st.form_submit_button("Create goal")
        if g_submit and g_name.strip():
            add_goal(workspace["id"], g_name, g_target, g_date)
            st.success("Goal created.")
            st.rerun()

st.divider()


# ============================================================
# TAX / GST CENTER
# ============================================================

st.header("Tax Center")
st.caption(f"GST-applicable transactions at the workspace rate of {workspace['gst_rate']}%.")

gst_tx = all_tx[all_tx.get("gst_applicable", 0) == 1] if not all_tx.empty else all_tx

if gst_tx.empty:
    st.info("No transactions marked as GST-applicable yet. Toggle it when adding a transaction below.")
else:
    rate = workspace["gst_rate"]
    gst_tx = gst_tx.copy()
    gst_tx["quarter"] = gst_tx["date"].dt.to_period("Q").astype(str)
    gst_tx["gst_component"] = gst_tx["amount"] * rate / (100 + rate)

    collected = gst_tx.loc[gst_tx["type"] == "income", "gst_component"].sum()
    paid = gst_tx.loc[gst_tx["type"] == "expense", "gst_component"].sum()

    t1, t2, t3 = st.columns(3)
    t1.metric("GST collected (output)", money(collected))
    t2.metric("GST paid (input)", money(paid))
    t3.metric("Net GST payable", money(collected - paid))

    quarterly = (
        gst_tx.groupby(["quarter", "type"])["gst_component"].sum().unstack(fill_value=0).reset_index()
    )
    st.dataframe(quarterly, use_container_width=True, hide_index=True)

st.divider()


# ============================================================
# BUSINESS HIGHLIGHTS
# ============================================================

st.header("Business Highlights")

highlight_columns = st.columns(3, gap="medium")

with highlight_columns[0]:
    if expenses > 0:
        expense_data = view[view["type"] == "expense"]
        categories = expense_data.groupby("category")["amount"].sum().sort_values(ascending=False)
        largest_category = categories.index[0]
        largest_amount = categories.iloc[0]
        share = largest_amount / expenses * 100
        st.markdown(
            f"""<div class="glass-card"><div class="glass-label">↘ Largest cost bucket</div>
            <div class="glass-value" style="color:#E11D48;">{largest_category}</div>
            <div class="glass-sub">{share:.1f}% of expenses • {money(largest_amount)}</div></div>""",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            """<div class="glass-card"><div class="glass-label">↘ Largest cost bucket</div>
            <div class="glass-value">—</div><div class="glass-sub">Add expenses to generate this insight.</div></div>""",
            unsafe_allow_html=True,
        )

income_data = view[view["type"] == "income"]
with highlight_columns[1]:
    if not income_data.empty:
        biggest_income = income_data.loc[income_data["amount"].idxmax()]
        st.markdown(
            f"""<div class="glass-card"><div class="glass-label">↗ Largest revenue event</div>
            <div class="glass-value" style="color:#059669;">{money(biggest_income['amount'])}</div>
            <div class="glass-sub">{biggest_income['description']}</div></div>""",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            """<div class="glass-card"><div class="glass-label">↗ Largest revenue event</div>
            <div class="glass-value">—</div><div class="glass-sub">Add income to generate this insight.</div></div>""",
            unsafe_allow_html=True,
        )

with highlight_columns[2]:
    if income <= 0 and expenses <= 0:
        p_color, p_label, p_sub = "#64748B", "No data", "Add transactions to activate insights."
    elif profit < 0:
        p_color, p_label, p_sub = "#E11D48", f"-{money(abs(profit))}", f"Expenses exceed revenue by {money(abs(profit))}."
    elif margin < 10:
        p_color, p_label, p_sub = "#D97706", f"{margin:.1f}%", "Positive, but the margin is thin."
    else:
        p_color, p_label, p_sub = "#059669", f"{margin:.1f}%", "Healthy share of revenue retained."
    st.markdown(
        f"""<div class="glass-card"><div class="glass-label">◆ Profitability</div>
        <div class="glass-value" style="color:{p_color};">{p_label}</div><div class="glass-sub">{p_sub}</div></div>""",
        unsafe_allow_html=True,
    )

st.divider()


# ============================================================
# PAYMENT REMINDERS
# No API keys, no cost: uses WhatsApp's own click-to-chat links
# (wa.me). It opens WhatsApp with the message pre-filled — you still
# tap send yourself. Fully automatic sending (no tap) would need a
# paid provider like Twilio or Meta's WhatsApp Business API, which
# needs business verification — a real next step, not built here.
# ============================================================

st.header("💬 Payment Reminders")
st.caption("Unpaid invoices, ready to chase over WhatsApp in one tap.")

unpaid_df = get_unpaid_invoices(workspace["id"])

if unpaid_df.empty:
    st.info("No unpaid invoices. Mark an income transaction as \"Invoiced — unpaid\" to track it here.")
else:
    total_owed = unpaid_df["amount"].sum()
    st.markdown(
        f"""<div class="glass-card" style="margin-bottom:16px;">
            <div class="glass-label">Total outstanding</div>
            <div class="glass-value" style="color:#D97706;">{money(total_owed)}</div>
            <div class="glass-sub">{len(unpaid_df)} unpaid invoice(s)</div>
        </div>""",
        unsafe_allow_html=True,
    )

    for _, inv in unpaid_df.sort_values("date").iterrows():
        rc1, rc2, rc3 = st.columns([3, 1.3, 1.3])
        customer_label = inv["customer_name"] or "Unnamed customer"
        with rc1:
            st.markdown(f"**{customer_label}** — {money(inv['amount'])}")
            st.caption(f"{inv['description']} · invoiced {inv['date'].strftime('%d %b %Y')}")
        with rc2:
            reminder_message = (
                f"Hi {customer_label}, this is a reminder that {money(inv['amount'])} "
                f"for \"{inv['description']}\" is still pending from {workspace['name']}. "
                f"Please let us know when we can expect payment. Thank you!"
            )
            wa_link = build_whatsapp_link(inv["customer_phone"], reminder_message)
            st.link_button("💬 Send reminder", wa_link, use_container_width=True)
        with rc3:
            if can_edit and st.button("Mark paid", key=f"paid_{inv['id']}", use_container_width=True):
                mark_invoice_paid(int(inv["id"]))
                st.rerun()

    if (unpaid_df["customer_phone"].str.strip() == "").any():
        st.caption("⚠️ Some unpaid invoices have no customer phone number — the reminder link will open WhatsApp's contact picker instead of a specific chat.")

st.divider()


# ============================================================
# ADD TRANSACTION (full form: GST, invoicing, recurring)
# ============================================================

st.header("Add Transaction")

if can_edit:
    with st.expander("＋ Open transaction form", expanded=False):
        form_col1, form_col2 = st.columns(2)

        with form_col1:
            description = st.text_input("Description", placeholder="Customer payment, rent, salary...")
            amount = st.number_input(f"Amount ({currency_symbol})", min_value=0.0, step=100.0)
            transaction_type = st.selectbox("Type", ["income", "expense"])
            category_choice = st.selectbox("Category (auto-suggested, editable)", ["(auto)"] + workspace_categories)

        with form_col2:
            transaction_date = st.date_input("Date", value=date.today())
            employee = st.text_input("Employee", placeholder="Optional — e.g. Kumar")
            gst_applicable = st.checkbox("GST applicable")
            invoice_status = ""
            customer_name = ""
            customer_phone = ""
            if transaction_type == "income":
                invoice_status = st.selectbox("Invoice status", ["Paid", "Invoiced — unpaid"])
                if invoice_status == "Invoiced — unpaid":
                    customer_name = st.text_input("Customer name", placeholder="Who owes this?")
                    customer_phone = st.text_input("Customer WhatsApp number", placeholder="e.g. 9198765xxxxx (with country code)")

        make_recurring = st.checkbox("Make this a recurring transaction")
        recurring_freq = None
        if make_recurring:
            recurring_freq = st.selectbox("Frequency", ["Weekly", "Monthly", "Yearly"])

        save_transaction = st.button("Save transaction", use_container_width=True)

        if save_transaction:
            if not description.strip():
                st.warning("Please enter a description.")
            elif amount <= 0:
                st.warning("Amount must be greater than zero.")
            elif limits["max_transactions"] and transaction_count(workspace["id"]) >= limits["max_transactions"]:
                st.error("Free plan transaction limit reached — upgrade to Pro to keep adding.")
            else:
                final_category = category_choice if category_choice != "(auto)" else category_from_text(description, learned_rules)
                insert_transaction(
                    workspace["id"], transaction_date, description, amount, transaction_type,
                    employee, final_category, gst_applicable, invoice_status, current_user["id"],
                    customer_name, customer_phone,
                )
                if make_recurring and recurring_freq:
                    add_recurring_rule(workspace["id"], description, amount, transaction_type, final_category,
                                        employee, recurring_freq, transaction_date)
                st.success("Transaction added.")
                st.rerun()
else:
    st.info("Your role is **viewer** — you can browse and export, but not add or edit transactions.")

if can_edit:
    recurring_df = get_recurring_rules(workspace["id"])
    if not recurring_df.empty:
        with st.expander(f"🔁 Recurring rules ({len(recurring_df)})"):
            for _, r in recurring_df.iterrows():
                rc1, rc2 = st.columns([4, 1])
                rc1.write(f"**{r['description']}** — {money(r['amount'])} · {r['frequency']} · next: {r['next_run_date']}")
                new_state = rc2.checkbox("Active", value=bool(r["active"]), key=f"rec_{r['id']}")
                if new_state != bool(r["active"]):
                    toggle_recurring_rule(r["id"], new_state)
                    st.rerun()

st.divider()


# ============================================================
# EXPLORE BUSINESS
# ============================================================

st.header("Explore Your Business")

left, right = st.columns(2, gap="large")

with left:
    st.subheader("Spending breakdown")
    expense_data = view[view["type"] == "expense"]
    if expense_data.empty:
        st.info("No expenses for this period.")
    else:
        category_totals = expense_data.groupby("category")["amount"].sum().sort_values()
        expense_chart = go.Figure()
        expense_chart.add_trace(
            go.Bar(x=category_totals.values, y=category_totals.index, orientation="h",
                   marker=dict(color=category_totals.values, colorscale=[[0, "rgba(16, 185, 129, 0.2)"], [1, "#10B981"]], line=dict(width=0)),
                   hovertemplate="<b>%{y}</b><br>Spent: " + currency_symbol + "%{x:,.0f}<extra></extra>")
        )
        expense_chart.update_layout(
            height=340, margin=dict(l=10, r=10, t=15, b=15), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Inter", color="#64748B"),
            hoverlabel=dict(bgcolor="#FFFFFF", bordercolor="#E2E8F0", font=dict(family="Inter", color="#0F172A")),
            xaxis=dict(showgrid=True, gridcolor="rgba(17,24,39,0.06)", tickprefix=currency_symbol),
            yaxis=dict(showgrid=False, tickfont=dict(size=10, color="#64748B")),
        )
        st.plotly_chart(expense_chart, use_container_width=True, config={"displayModeBar": False})

with right:
    st.subheader("Transaction activity by employee")
    st.caption("Volume of transactions tagged to each person — not a performance ranking. Revenue generated, deals closed, or hours worked would need to be tracked separately to mean that.")
    people = view[view["employee"].str.strip() != ""]
    if people.empty:
        st.info("Add an employee name to transactions to unlock this view.")
    else:
        team = people.groupby("employee").agg(
            Transactions=("amount", "count"),
            Total_Movement=("amount", "sum"),
        )
        team["Avg. transaction"] = team["Total_Movement"] / team["Transactions"]
        team = team.sort_values("Total_Movement", ascending=False)
        team = team.rename(columns={"Total_Movement": f"Total Movement ({currency_symbol})", "Avg. transaction": f"Avg. Transaction ({currency_symbol})"})
        st.dataframe(team, use_container_width=True, height=300)

st.divider()


# ============================================================
# TRANSACTION LEDGER (with soft delete + undo)
# ============================================================

st.header("Transaction Ledger")
st.caption("Edit records directly. Deleted rows can be restored below.")

if all_tx.empty:
    st.info("No transactions yet.")
else:
    editable = all_tx[["id", "date", "description", "amount", "type", "category", "employee"]].copy()
    editable["date"] = editable["date"].dt.date

    edited_data = st.data_editor(
        editable,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic" if can_edit else "fixed",
        disabled=not can_edit,
        column_config={
            "id": st.column_config.NumberColumn("ID", disabled=True),
            "date": st.column_config.DateColumn("Date"),
            "description": st.column_config.TextColumn("Description"),
            "amount": st.column_config.NumberColumn(f"Amount ({currency_symbol})", min_value=0, step=100),
            "type": st.column_config.SelectboxColumn("Type", options=["income", "expense"]),
            "category": st.column_config.SelectboxColumn("Category", options=workspace_categories),
            "employee": st.column_config.TextColumn("Employee"),
        },
        key="ledger_editor",
    )

    if can_edit and st.button("Save table changes", use_container_width=True):
        original_ids = set(editable["id"].tolist())
        edited_ids = set(edited_data["id"].dropna().tolist())

        deleted_ids = original_ids - edited_ids
        for tx_id in deleted_ids:
            soft_delete_transaction(int(tx_id))

        for _, row in edited_data.iterrows():
            if pd.isna(row["id"]):
                if row["description"] and row["amount"] and row["amount"] > 0:
                    insert_transaction(workspace["id"], row["date"], row["description"], row["amount"],
                                        row["type"], row.get("employee", ""), row["category"], 0, "", current_user["id"])
            else:
                original_row = editable[editable["id"] == row["id"]].iloc[0]
                if not original_row.equals(row):
                    update_transaction(int(row["id"]), row["date"], row["description"], row["amount"], row["type"], row.get("employee", ""), row["category"])
                    if row["category"] != original_row["category"]:
                        learn_category(workspace["id"], row["description"], row["category"])

        st.success("Changes saved.")
        st.rerun()

    if can_edit:
        deleted_recent = get_deleted_transactions(workspace["id"])
        if not deleted_recent.empty:
            with st.expander(f"🗑 Recently deleted ({len(deleted_recent)})"):
                st.dataframe(deleted_recent[["date", "description", "amount", "type"]], use_container_width=True, hide_index=True)
                if st.button("Restore most recently deleted"):
                    if restore_last_deleted(workspace["id"]):
                        st.success("Restored.")
                        st.rerun()

st.divider()


# ============================================================
# RECENT ACTIVITY
# ============================================================

st.header("Recent Activity")

if view.empty:
    st.info("No recent activity.")
else:
    recent = view.sort_values("date", ascending=False).copy()
    recent["date"] = recent["date"].dt.strftime("%d %b %Y")
    recent["amount"] = recent["amount"].map(lambda v: money(v))
    recent = recent.rename(columns={"date": "Date", "description": "Description", "amount": "Amount", "type": "Type", "employee": "Employee", "category": "Category"})
    st.dataframe(recent[["Date", "Description", "Amount", "Type", "Category", "Employee"]], use_container_width=True, hide_index=True)

st.divider()

# ============================================================
# PREMIUM COMMAND INTEL — PAYWALL TERMINAL
# ============================================================
st.divider()
st.header("🧠 FinSight Command Intelligence")
st.caption("Unlock deeply profiled database metrics and predictive tax runway forecasting.")

# Create the secret key input in the UI
token_input = st.text_input("🔑 Enter your Premium Access Token to unlock", type="password", help="Contact admin to get your access key.")

# Define your master token phrase here (Change 'FinsightPro2026' to whatever you want!)
# Streamlit will securely pull this word from the server dashboard hidden settings
MASTER_TOKEN = st.secrets["PREMIUM_KEY"]


if token_input == MASTER_TOKEN:
    st.success("🔓 Premium Intel Active.")
    
    intel_col1, intel_col2 = st.columns(2, gap="large")
    
    with intel_col1:
        st.subheader("🔍 Deep Database Profiling")
        leaks = profile_database_leaks(workspace["id"])
        if not leaks:
            st.markdown("<div class='insight action-item-green'>🟢 <b>No Leaks Found:</b> No duplicate transaction patterns detected in the database schema.</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div class='insight action-item-red'>🔴 <b>Potential Double-Entries Found:</b> Review the ledger to verify these are not accidental duplicates:</div>", unsafe_allow_html=True)
            for leak in leaks:
                st.warning(f"⚠️ {leak['description']} — {money(leak['amount'])} logged multiple times between {leak['date1']} and {leak['date2']}.")
                
    with intel_col2:
        st.subheader("📉 AI Tax Run-Rate Forecasting")
        tax_data = forecast_tax_liability(all_tx, workspace["gst_rate"])
        if tax_data is None:
            st.info("Add at least 5 different transaction logs over past cycles to generate predictive runways.")
        else:
            health_color = "#059669" if tax_data["run_rate_health"] == "Stable" else "#E11D48"
            st.markdown(
                f"""
                <div class='glass-card'>
                    <div class='glass-label'>Next 30-Day GST Outflow Forecast</div>
                    <div class='glass-value'>{money(tax_data['monthly_gst_forecast'])}</div>
                    <div class='glass-label' style='margin-top:10px;'>Estimated Annual Income Tax Liability</div>
                    <div class='glass-value' style='color:#0F172A;'>{money(tax_data['estimated_annual_income_tax'])}</div>
                    <div class='glass-sub' style='margin-top:8px;'>Operation Run-Rate Health: <b style='color:{health_color};'>{tax_data['run_rate_health']}</b></div>
                </div>
                """,
                unsafe_allow_html=True,
            )
else:
    # This is the locked paywall state that shows your pricing and payment prompt
    st.info("🔒 Advanced Intelligence Locked.")
    
    pay_col1, pay_col2 = st.columns([2.5, 1.5])
    with pay_col1:
        st.markdown(
            """
            ### What's included in Premium Access:
            * **Automated AI Tax Liability Projections:** Know exactly how much GST and annual business taxes you need to put aside before the quarter ends.
            * **Deep Ledger Leak Profiling:** Automated system scans to flag accidental duplicate customer billings or double-entered supplier expenses.
            """
        )
    with pay_col2:
        st.markdown(
            f"""
            <div class='glass-card' style='text-align:center;'>
                <div class='glass-label'>Premium Plan</div>
                <div class='glass-value' style='color:#D97706;'>₹299 / Mo</div>
                <div class='glass-sub'>Instant Activation via UPI</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        # Change this text to your actual UPI string or contact layout!
        st.markdown("💬 **How to buy:** Send payment to your UPI ID or drop a WhatsApp message to the admin. Paste the code you receive above!")


# ============================================================
# EXPORTS
# ============================================================

st.header("Data Management")
st.caption("Export your complete FinSight financial dataset.")

export_df = all_tx.drop(columns=["deleted"], errors="ignore")
csv_data = export_df.to_csv(index=False)

e1, e2, e3 = st.columns(3)

with e1:
    st.download_button("Export as CSV", csv_data, "finsight_transactions.csv", "text/csv", use_container_width=True)

with e2:
    if limits["exports"]:
        try:
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                export_df.to_excel(writer, index=False, sheet_name="Transactions")
            st.download_button("Export as Excel", buffer.getvalue(), "finsight_transactions.xlsx",
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        except ImportError:
            st.button("Export as Excel", disabled=True, use_container_width=True, help="Run: pip install openpyxl")
    else:
        st.button("Export as Excel (Pro)", disabled=True, use_container_width=True, help="Upgrade to Pro to unlock Excel export.")

with e3:
    if limits["exports"]:
        try:
            from fpdf import FPDF

            # 1. Clean helper function to safely swap out unsupported symbols
            def clean_txt(text):
                if not text:
                    return ""
                return str(text).replace('—', '-').replace('₹', 'Rs.').replace('€', 'EUR').replace('£', 'GBP')

            # 2. Setup the document layout
            pdf = FPDF()
            pdf.add_page()
            
            # --- TITLE & SUMMARY ---
            pdf.set_font("Helvetica", "B", 16)
            pdf.cell(0, 10, clean_txt(f"FinSight Report - {workspace['name']}"), ln=True)
            
            pdf.set_font("Helvetica", "B", 10)
            pdf.cell(0, 8, clean_txt(f"Period: {period}"), ln=True)
            
            summary_text = f"Revenue: {money(income)}   Expenses: {money(expenses)}   Net: {money(profit)}   Margin: {margin:.1f}%"
            pdf.cell(0, 8, clean_txt(summary_text), ln=True)
            pdf.ln(6)
            
            # --- TABLE HEADER ---
            pdf.set_font("Helvetica", "B", 10)
            pdf.cell(30, 8, "Date")
            pdf.cell(70, 8, "Description")
            pdf.cell(30, 8, "Amount")
            pdf.cell(30, 8, "Type")
            pdf.ln()
            
            # --- TABLE BODY ---
            pdf.set_font("Helvetica", "", 9)
            for _, r in view.sort_values("date", ascending=False).head(60).iterrows():
                # Safely format the date string
                date_str = r["date"].strftime("%d %b %Y") if hasattr(r["date"], "strftime") else str(r["date"])
                
                pdf.cell(30, 7, clean_txt(date_str))
                pdf.cell(70, 7, clean_txt(str(r["description"])[:38]))
                pdf.cell(30, 7, clean_txt(money(r["amount"])))  # Clean the currency sign here!
                pdf.cell(30, 7, clean_txt(r["type"]))
                pdf.ln()
                
            # 3. Export data smoothly to byte array output
            pdf_raw_data = pdf.output()
            pdf_bytes = bytes(pdf_raw_data)  # Converts bytearray into standard bytes
            
            st.download_button(
                label="Export PDF report", 
                data=pdf_bytes, 
                file_name="finsight_report.pdf", 
                mime="application/pdf", 
                use_container_width=True
            )

        except ImportError:
            st.button("Export PDF report", disabled=True, use_container_width=True, help="Run: pip install fpdf2")
    else:
        st.button("Export PDF report (Pro)", disabled=True, use_container_width=True, help="Upgrade to Pro to unlock PDF reports.")


# ============================================================
# FOOTER
# ============================================================

st.divider()
st.caption("FinSight • Financial intelligence workspace")

if st.session_state.get("show_finbot", True):
    render_finbot_dock(f"**{mood_text}.** {mood_description}", "alert" if mood_key in ("red", "yellow") else "hello")

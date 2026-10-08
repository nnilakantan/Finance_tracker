import datetime
import json
import os
import re
from pathlib import Path

import google.generativeai as genai
import pandas as pd
import pdfplumber
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Financial Dashboard", layout="wide", page_icon="📈")
st.title("Personal Finance & Equity Dashboard")

# -- FILE PATHS --
BASE_DIR = Path(__file__).resolve().parent
DB_FILE = BASE_DIR / "transactions_db.csv"
CATEGORY_CACHE_FILE = BASE_DIR / "category_cache.json"
CUSTOM_CATEGORIES_FILE = BASE_DIR / "custom_categories.json"
USER_RULES_FILE = BASE_DIR / "user_rules.json"
TRACKER_FILE = BASE_DIR / "Tracker_2024.xlsx"

# -- AI SETUP --
GEMINI_CONFIGURED = False
if "GEMINI_API_KEY" in st.secrets:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    GEMINI_CONFIGURED = True
elif os.getenv("GEMINI_API_KEY"):
    genai.configure(api_key=os.getenv("GEMINI_API_KEY"))
    GEMINI_CONFIGURED = True

# -- BASE CATEGORY RULES --
CATEGORY_RULES = {
    "Grocery": [
        "safeway", "whole foods", "ayala farms", "fulton family farms", "r. farms", "yee yang farm",
        "draper girls country farm", "costco whse", "india supermarket", "trader joe", "fred meyer",
        "new seasons", "grocery", "groceries",
    ],
    "Restaurants": [
        "restaurant", "cafe", "pizza", "thai", "mexican", "sangam indian cuisine", "virasat indian cuisine",
        "desi pizza house", "mumbai street food", "sabor mexicano", "the hammond", "phurr",
        "grand central bakery", "julia bakery", "brewed cafe", "tst*", "starbucks", "doordash",
        "uber eats",
    ],
    "Utilities": [
        "pge", "pacific power", "electric", "nw natural", "gas", "water bureau", "utility",
        "waste connections", "garbage", "sewer",
    ],
    "Mortgage": ["mortgage", "home loan", "unitedwholesale", "loan paymt"],
    "HOA": ["hoa", "homeowners association", "assn dues", "granite highland"],
    "Car Lease / Auto": ["toyota financial", "honda financial", "lease", "dmv", "auto"],
    "Fuel": ["costco gas", "shell", "chevron", "76", "exxon", "gas station"],
    "Daycare / Nanny": ["daycare", "nanny", "childcare", "learning adventures"],
    "Telephone / Internet": ["comcast", "xfinity", "at&t", "verizon", "t-mobile", "tmobile"],
    "Insurance": ["state farm", "insurance", "northwestern mu", "american heritag"],
    "Travel": ["hotel", "lodging", "airlines", "airways", "klm", "delta", "united airlines", "lyft", "uber"],
    "Shopping": ["amazon", "target", "walmart", "costco.com", "home depot", "lowes", "nursery"],
    "Subscriptions / Software": ["adobe", "netflix", "spotify", "apple.com", "google", "microsoft"],
    "Healthcare": ["pharmacy", "doctor", "dental", "medical", "hospital", "clinic"],
    "Education": ["school", "tuition", "529"],
    "India transfer": ["remitly", "xoom", "wise"],
    "Investments / Savings": ["vanguard", "my529", "fidelity", "schwab", "td ameritrade", "investment"],
    "Credit Card Payment": [
        "american express des:ach pmt", "amex", "citi card online", "chase credit crd",
        "jpmorgan chase des:chase ach", "comenity pay", "online payment, thank you",
    ],
    "Transfer": ["zelle payment", "paypal des:transfer", "digital federal", "transfer"],
    "Income": ["payroll", "interest earned", "primus realty"],
}

SPENDING_EXCLUSIONS = {"Income", "Transfer", "Credit Card Payment", "Investments / Savings"}
MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


# -- PERSISTENT USER RULES & CATEGORIES --
def load_user_rules():
    if USER_RULES_FILE.exists():
        try:
            return json.loads(USER_RULES_FILE.read_text())
        except Exception:
            return {}
    return {}


def save_user_rules(rules_dict):
    USER_RULES_FILE.write_text(json.dumps(rules_dict, indent=2, sort_keys=True))


def get_all_categories():
    base_cats = list(CATEGORY_RULES.keys())
    custom_cats = []
    if CUSTOM_CATEGORIES_FILE.exists():
        try:
            custom_cats = json.loads(CUSTOM_CATEGORIES_FILE.read_text())
        except Exception:
            pass
    all_cats = sorted(list(set(base_cats + custom_cats)))
    if "Uncategorized" in all_cats:
        all_cats.remove("Uncategorized")
    return all_cats + ["Uncategorized"]


# -- DATABASE DISK WRITER (AUTO-SYNC ENGINE) --
def load_transactions_db():
    if not DB_FILE.exists():
        return pd.DataFrame(columns=["Date", "Description", "Amount", "Source", "Category"])
    db_df = pd.read_csv(DB_FILE)
    if "Category" not in db_df.columns:
        db_df["Category"] = "Uncategorized"
    if "Source" not in db_df.columns:
        db_df["Source"] = "Unknown"
    db_df["Date"] = pd.to_datetime(db_df["Date"], errors="coerce")
    db_df["Amount"] = pd.to_numeric(db_df["Amount"], errors="coerce")
    return db_df.dropna(subset=["Date", "Amount"])


def write_transactions_db(df):
    """Physically writes the DataFrame to transactions_db.csv on disk."""
    combined_df = df.copy()
    combined_df["Date"] = pd.to_datetime(combined_df["Date"], errors="coerce").dt.strftime("%Y-%m-%d")
    combined_df["Description"] = combined_df["Description"].astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
    combined_df["Amount"] = pd.to_numeric(combined_df["Amount"], errors="coerce").round(2)
    combined_df["Category"] = combined_df["Category"].fillna("Uncategorized")
    combined_df["Source"] = combined_df["Source"].fillna("Unknown")
    combined_df = combined_df.dropna(subset=["Date", "Amount"])
    combined_df = combined_df.drop_duplicates(subset=["Date", "Description", "Amount", "Source"], keep="last")
    combined_df.sort_values(["Date", "Source", "Description"]).to_csv(DB_FILE, index=False)
    return combined_df


def save_transactions(df):
    existing_df = load_transactions_db()
    combined_df = pd.concat([existing_df, df], ignore_index=True)
    return write_transactions_db(combined_df)


# -- HELPER FUNCTIONS --
def clean_money(value):
    if pd.isna(value):
        return None
    text = str(value).strip().replace("$", "").replace(",", "")
    if not text:
        return None
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]
    try:
        return float(text)
    except ValueError:
        return None


def parse_date_value(value):
    if pd.isna(value):
        return None
    if isinstance(value, (datetime.datetime, datetime.date, pd.Timestamp)):
        return pd.to_datetime(value).date()
    if isinstance(value, (int, float)):
        return None
    value_text = str(value).strip()
    if not re.search(r"\d{1,2}[/-]\d{1,2}|20\d{2}|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec", value_text, re.I):
        return None
    try:
        parsed = pd.to_datetime(value_text, errors="coerce")
        if pd.notna(parsed):
            return parsed.date()
    except Exception:
        return None
    return None


# -- EXCEL EXTRACTION --
def find_net_worth_column(raw_df):
    for r_idx in range(min(15, len(raw_df))):
        for c_idx in range(len(raw_df.columns)):
            cell_val = str(raw_df.iloc[r_idx, c_idx]).strip().lower()
            if "net" in cell_val and "worth" in cell_val:
                return c_idx
    return None


def load_net_worth_data(file_path):
    xls = pd.ExcelFile(file_path)
    records = []
    for sheet_name in xls.sheet_names:
        raw_df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        nw_col_idx = find_net_worth_column(raw_df)
        if nw_col_idx is None:
            continue
        sheet_year_match = re.search(r"(20\d{2})", str(sheet_name))
        if not sheet_year_match:
            continue
        sheet_year = int(sheet_year_match.group(1))

        for row_idx, row in raw_df.iterrows():
            month_date = None
            for col_idx in range(min(4, len(row))):
                month_date = parse_date_value(row.iloc[col_idx])
                if month_date:
                    break
            if not month_date:
                continue

            net_worth = None
            for detail_idx in range(row_idx + 1, min(row_idx + 5, len(raw_df))):
                for col_offset in [0, 1, -1]:
                    target_col = nw_col_idx + col_offset
                    if 0 <= target_col < len(raw_df.columns):
                        parsed_value = clean_money(raw_df.iloc[detail_idx, target_col])
                        if parsed_value is not None and parsed_value > 0:
                            net_worth = max(net_worth or 0, parsed_value)

            if net_worth:
                records.append({
                    "Year": sheet_year,
                    "Month": month_date.strftime("%b"),
                    "MonthNumber": month_date.month,
                    "DateYear": month_date.year,
                    "Date": datetime.date(sheet_year, month_date.month, min(month_date.day, 28)),
                    "Net Worth": net_worth,
                    "Source Sheet": str(sheet_name),
                    "Source Row": row_idx,
                })

    if not records:
        return pd.DataFrame(columns=["Year", "Month", "MonthNumber", "DateYear", "Date", "Net Worth"])
    nw_df = pd.DataFrame(records)
    nw_df = nw_df.sort_values(["Year", "MonthNumber", "Source Row"])
    nw_df = nw_df.drop_duplicates(["Year", "MonthNumber"], keep="last")
    return nw_df


def december_net_worth(nw_df, current_year):
    prior_df = nw_df[(nw_df["Year"] < current_year) & (nw_df["Net Worth"] > 0)].copy()
    if prior_df.empty:
        return prior_df
    annual_rows = []
    for year, year_df in prior_df.groupby("Year"):
        december_df = year_df[year_df["MonthNumber"] == 12]
        if not december_df.empty:
            row = december_df.sort_values("Source Row").iloc[-1].copy()
            row["Snapshot"] = "December"
        else:
            row = year_df.sort_values("MonthNumber").iloc[-1].copy()
            row["Snapshot"] = f"Latest available ({row['Month']})"
        annual_rows.append(row)
    return pd.DataFrame(annual_rows).sort_values("Year")


def current_year_net_worth(nw_df, current_year):
    current_month = datetime.date.today().month
    month_df = nw_df[(nw_df["Year"] == current_year) & (nw_df["MonthNumber"] <= current_month)].copy()
    month_df = month_df.sort_values("MonthNumber")
    if month_df.empty:
        return month_df
    all_months = pd.DataFrame({"MonthNumber": list(range(1, current_month + 1)), "Month": MONTH_ORDER[:current_month]})
    return all_months.merge(month_df, on=["MonthNumber", "Month"], how="left")


# -- STATEMENT PARSERS --
def get_file_name(file_object):
    return Path(getattr(file_object, "name", str(file_object))).name


def infer_year_for_mmdd(month, statement_end_date):
    if not statement_end_date:
        return datetime.date.today().year
    year = statement_end_date.year
    if month > statement_end_date.month:
        year -= 1
    return year


def extract_statement_end_date(text):
    patterns = [
        r"Billing Period:\s*\d{2}/\d{2}/\d{2}\s*-\s*(\d{2}/\d{2}/\d{2})",
        r"for\s+\w+\s+\d{1,2},\s+\d{4}\s+to\s+(\w+\s+\d{1,2},\s+\d{4})",
        r"New balance as of\s+(\d{2}/\d{2}/\d{2})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            parsed = pd.to_datetime(match.group(1), errors="coerce")
            if pd.notna(parsed):
                return parsed.date()
    return None


def normalize_transactions(df, source):
    if df is None or df.empty:
        return pd.DataFrame(columns=["Date", "Description", "Amount", "Source"])
    clean_df = df[["Date", "Description", "Amount"]].copy()
    clean_df["Date"] = pd.to_datetime(clean_df["Date"], errors="coerce")
    clean_df["Description"] = clean_df["Description"].astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
    clean_df["Amount"] = pd.to_numeric(clean_df["Amount"], errors="coerce")
    clean_df["Source"] = source
    clean_df = clean_df.dropna(subset=["Date", "Amount"])
    clean_df = clean_df[clean_df["Description"] != ""]
    return clean_df


def parse_amex_tabular(file_object):
    file_name = get_file_name(file_object)
    df = pd.read_excel(file_object) if file_name.lower().endswith(".xlsx") else pd.read_csv(file_object)
    header_row_idx = None
    for idx, row in df.iterrows():
        row_values = [str(v).strip().lower() for v in row.tolist()]
        if "date" in row_values and "description" in row_values and "amount" in row_values:
            header_row_idx = idx
            break
    if header_row_idx is not None:
        if hasattr(file_object, "seek"):
            file_object.seek(0)
        df = pd.read_excel(file_object, header=header_row_idx + 1) if file_name.lower().endswith(".xlsx") else pd.read_csv(file_object, header=header_row_idx + 1)

    df.columns = [str(c).strip() for c in df.columns]
    date_col = next((c for c in df.columns if c.lower() == "date"), None)
    desc_col = next((c for c in df.columns if c.lower() == "description"), None)
    amount_col = next((c for c in df.columns if c.lower() == "amount"), None)
    if not all([date_col, desc_col, amount_col]):
        return pd.DataFrame()

    clean_df = df[[date_col, desc_col, amount_col]].copy()
    clean_df.columns = ["Date", "Description", "Amount"]
    clean_df["Amount"] = -clean_df["Amount"].apply(clean_money)
    return normalize_transactions(clean_df, file_name)


def parse_citi_pdf(pdf_file):
    transactions = []
    pattern = re.compile(r"^\s*(?:(\d{2}/\d{2})\s+)?(\d{2}/\d{2})\s+(.+?)\s+(-?\$[\d,]+\.\d{2})\s*$")
    with pdfplumber.open(pdf_file) as pdf:
        first_text = pdf.pages[0].extract_text() or ""
        statement_end_date = extract_statement_end_date(first_text)
        for page in pdf.pages:
            for line in (page.extract_text() or "").split("\n"):
                match = pattern.match(line)
                if not match:
                    continue
                month, day = [int(p) for p in match.group(2).split("/")]
                year = infer_year_for_mmdd(month, statement_end_date)
                amount_val = clean_money(match.group(4))
                if amount_val is not None:
                    transactions.append({"Date": datetime.date(year, month, day), "Description": match.group(3).strip(), "Amount": -amount_val})
    return normalize_transactions(pd.DataFrame(transactions), get_file_name(pdf_file))


def parse_bofa_checking(pdf_file):
    transactions = []
    line_pattern = re.compile(r"^(\d{2}/\d{2}/\d{2})\s+(.+?)\s+(-?[\d,]+\.\d{2})$")
    in_transactions = False
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            for line in (page.extract_text() or "").split("\n"):
                line = line.strip()
                if "Deposits and other additions" in line or "Withdrawals and other subtractions" in line:
                    in_transactions = True
                    continue
                if line.startswith("Total ") or line.startswith("Page ") or not in_transactions:
                    continue
                match = line_pattern.match(line)
                if not match:
                    continue
                amount_val = clean_money(match.group(3))
                if amount_val is not None:
                    transactions.append({"Date": pd.to_datetime(match.group(1), format="%m/%d/%y").date(), "Description": match.group(2).strip(), "Amount": amount_val})
    return normalize_transactions(pd.DataFrame(transactions), get_file_name(pdf_file))


def identify_bank(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        first_page_text = (pdf.pages[0].extract_text() or "").lower()
        if "citi" in first_page_text or "citicard" in first_page_text:
            return "citi"
        if "bank of america" in first_page_text or "deposits" in first_page_text:
            return "bofa"
        if "chase" in first_page_text:
            return "chase"
    return "unknown"


def process_statement(file_object):
    ext = get_file_name(file_object).split(".")[-1].lower()
    if ext in {"csv", "xlsx"}:
        return parse_amex_tabular(file_object)
    if ext == "pdf":
        bank_id = identify_bank(file_object)
        if hasattr(file_object, "seek"):
            file_object.seek(0)
        if bank_id == "citi":
            return parse_citi_pdf(file_object)
        if bank_id == "bofa":
            return parse_bofa_checking(file_object)
    return pd.DataFrame()


# -- CATEGORIZATION ENGINE --
def rule_based_category(description):
    desc_lower = str(description).lower()
    user_rules = load_user_rules()
    for rule_pattern, category in user_rules.items():
        if rule_pattern.lower() in desc_lower:
            return category
    for category, keywords in CATEGORY_RULES.items():
        if any(keyword in desc_lower for keyword in keywords):
            return category
    return "Uncategorized"


def gemini_categorize(description, amount=None):
    if not GEMINI_CONFIGURED:
        return "Uncategorized"
    valid_categories = get_all_categories()
    prompt = f"""
Categorize this personal finance transaction.
Description: {description}
Amount: {amount}

Return exactly one category from this list:
{valid_categories}

Use "Uncategorized" only when the merchant or purpose is genuinely unclear.
"""
    try:
        model = genai.GenerativeModel("gemini-1.5-flash")
        response = model.generate_content(prompt)
        ai_guess = response.text.strip()
        return ai_guess if ai_guess in valid_categories else "Uncategorized"
    except Exception:
        return "Uncategorized"


def smart_categorize(description, amount=None, use_gemini=True):
    category = rule_based_category(description)
    if category != "Uncategorized":
        return category

    cache_key = str(description).strip().lower()
    cache = {}
    if CATEGORY_CACHE_FILE.exists():
        try:
            cache = json.loads(CATEGORY_CACHE_FILE.read_text())
        except Exception:
            pass

    if cache_key in cache:
        return cache[cache_key]

    category = gemini_categorize(description, amount) if use_gemini else "Uncategorized"
    cache[cache_key] = category
    CATEGORY_CACHE_FILE.write_text(json.dumps(cache, indent=2, sort_keys=True))
    return category


def categorize_transactions(df, use_gemini=True):
    if df.empty:
        return df
    categorized = df.copy()
    categorized["Category"] = categorized.apply(
        lambda row: smart_categorize(row["Description"], row["Amount"], use_gemini=use_gemini),
        axis=1,
    )
    return categorized


def discover_statement_files():
    statement_files = []
    for path in BASE_DIR.iterdir():
        if path.suffix.lower() not in {".pdf", ".csv", ".xlsx"}:
            continue
        if path.name in {TRACKER_FILE.name, DB_FILE.name}:
            continue
        statement_files.append(path)
    return sorted(statement_files, key=lambda p: p.name.lower())


def prepare_expenses(db_df, excluded_categories):
    expenses_df = db_df[db_df["Amount"] < 0].copy()
    if excluded_categories:
        expenses_df = expenses_df[~expenses_df["Category"].isin(excluded_categories)]
    expenses_df["Amount"] = expenses_df["Amount"].abs()
    expenses_df["MonthDate"] = expenses_df["Date"].dt.to_period("M").dt.to_timestamp()
    expenses_df["Month"] = expenses_df["MonthDate"].dt.strftime("%b %Y")
    return expenses_df.sort_values("MonthDate")


# -- START OF UI --
try:
    nw_df = load_net_worth_data(str(TRACKER_FILE))
    data_loaded = True
except FileNotFoundError:
    st.error(f"Could not find '{TRACKER_FILE.name}'. Please ensure it is in the same folder as app.py.")
    data_loaded = False


if data_loaded:
    st.sidebar.header("Dashboard Controls")
    view_selection = st.sidebar.radio("Navigation", ["Net Worth Analysis", "Cash Flow Analysis", "Statement Importer"])

    # Live Database File Status Monitor
    db_df = load_transactions_db()
    if DB_FILE.exists():
        mtime_str = datetime.datetime.fromtimestamp(DB_FILE.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        st.sidebar.markdown("---")
        st.sidebar.markdown(f"💾 **Disk Database Sync**\n* File: `{DB_FILE.name}`\n* Total Rows: **{len(db_df):,}**\n* Last Written: `{mtime_str}`")

    # -- VIEW 1: NET WORTH --
    if view_selection == "Net Worth Analysis":
        st.subheader("Historical & Current Net Worth Analysis")

        if not nw_df.empty:
            current_year = datetime.date.today().year
            annual_df = december_net_worth(nw_df, current_year)
            current_year_df = current_year_net_worth(nw_df, current_year)

            col1, col2 = st.columns(2)

            with col1:
                st.markdown("### Prior-Year Net Worth")
                if not annual_df.empty:
                    fig_annual = px.bar(
                        annual_df,
                        x="Year",
                        y="Net Worth",
                        color="Snapshot",
                        text_auto=".3s",
                        title="December Net Worth, or Latest Available Month if December Is Blank",
                    )
                    fig_annual.update_layout(xaxis_type="category", yaxis_tickprefix="$")
                    st.plotly_chart(fig_annual, width="stretch")
                    st.dataframe(annual_df[["Year", "Snapshot", "Month", "Net Worth"]].style.format({"Net Worth": "${:,.0f}"}), width="stretch")
                else:
                    st.info("No positive prior-year net-worth entries were found.")

            with col2:
                st.markdown(f"### {current_year} Monthly Net Worth")
                if not current_year_df.empty and current_year_df["Net Worth"].notna().any():
                    fig_monthly = px.line(current_year_df, x="Month", y="Net Worth", markers=True, title=f"{current_year} Net Worth by Month")
                    min_val = current_year_df["Net Worth"].min(skipna=True) * 0.95
                    max_val = current_year_df["Net Worth"].max(skipna=True) * 1.05
                    fig_monthly.update_yaxes(range=[min_val, max_val], tickprefix="$")
                    fig_monthly.update_xaxes(categoryorder="array", categoryarray=MONTH_ORDER)
                    st.plotly_chart(fig_monthly, width="stretch")
                    st.dataframe(current_year_df[["Month", "Net Worth"]].style.format({"Net Worth": "${:,.0f}"}), width="stretch")
                else:
                    st.info(f"No Net Worth data found yet for {current_year}.")
        else:
            st.info(f"No Net Worth data could be extracted from {TRACKER_FILE.name}.")

    # -- VIEW 2: CASH FLOW ANALYSIS --
    elif view_selection == "Cash Flow Analysis":
        st.subheader("Monthly Spending Analysis")

        # Expander: Manage Categories & Rules
        with st.expander("⚙️ Manage Categories & Auto-Rules"):
            cat_tab1, cat_tab2 = st.tabs(["Add Categories", "Active Auto-Rules"])
            with cat_tab1:
                new_cat = st.text_input("New Category Name (e.g., 'Pet Supplies')")
                if st.button("Add Category"):
                    new_cat_clean = new_cat.strip()
                    if new_cat_clean:
                        existing = get_all_categories()
                        if new_cat_clean not in existing:
                            custom_cats = []
                            if CUSTOM_CATEGORIES_FILE.exists():
                                try:
                                    custom_cats = json.loads(CUSTOM_CATEGORIES_FILE.read_text())
                                except Exception:
                                    pass
                            custom_cats.append(new_cat_clean)
                            CUSTOM_CATEGORIES_FILE.write_text(json.dumps(custom_cats))
                            st.success(f"Added '{new_cat_clean}' to categories!")
                            st.rerun()
                        else:
                            st.warning("That category already exists.")
            with cat_tab2:
                current_rules = load_user_rules()
                if current_rules:
                    st.dataframe(
                        pd.DataFrame(list(current_rules.items()), columns=["Keyword / Pattern", "Assigned Category"]),
                        width="stretch",
                    )
                    if st.button("Clear All Custom Rules"):
                        save_user_rules({})
                        st.success("Cleared custom rules.")
                        st.rerun()
                else:
                    st.caption("No custom learned rules yet. Categorize recurring expenses below to build rules automatically.")

        if db_df.empty:
            st.warning("No transaction data found. Import statements first to generate spending charts.")
        else:
            uncategorized_mask = db_df["Category"] == "Uncategorized"
            uncategorized_count = int(uncategorized_mask.sum())

            if uncategorized_count:
                st.info(f"{uncategorized_count} transactions are currently Uncategorized.")

                # SECTION A: RECURRING EXPENSE DETECTOR (AUTO-SAVES TO CSV)
                uncat_descriptions = (
                    db_df[uncategorized_mask]
                    .groupby("Description")
                    .agg(Count=("Amount", "count"), TotalSpend=("Amount", lambda x: x.abs().sum()))
                    .reset_index()
                )
                recurring_uncat = uncat_descriptions[uncat_descriptions["Count"] >= 2].sort_values("Count", ascending=False)

                if not recurring_uncat.empty:
                    with st.expander("🔁 Recurring Uncategorized Expenses Detected (2+ occurrences)", expanded=True):
                        st.write("Assigning a category below updates all matching rows in `transactions_db.csv` immediately and saves an auto-rule:")
                        col_r1, col_r2, col_r3 = st.columns([2, 1, 1])
                        with col_r1:
                            selected_vendor = st.selectbox("Select recurring merchant:", recurring_uncat["Description"].tolist())
                        with col_r2:
                            vendor_cat = st.selectbox("Assign Category:", [c for c in get_all_categories() if c != "Uncategorized"], key="rec_cat")
                        with col_r3:
                            st.write("")
                            st.write("")
                            if st.button("Apply to All & Auto-Save", type="primary"):
                                rules = load_user_rules()
                                rules[selected_vendor] = vendor_cat
                                save_user_rules(rules)

                                db_df.loc[db_df["Description"] == selected_vendor, "Category"] = vendor_cat
                                write_transactions_db(db_df)
                                st.toast(f"✅ Auto-saved {selected_vendor} to transactions_db.csv", icon="💾")
                                st.rerun()

                        st.dataframe(recurring_uncat.style.format({"TotalSpend": "${:,.2f}"}), width="stretch")

                # SECTION B: GEMINI BULK RUNNER (AUTO-SAVES TO CSV)
                if st.button("Ask Gemini to categorize remaining uncategorized transactions", disabled=not GEMINI_CONFIGURED):
                    with st.spinner("Categorizing with Gemini..."):
                        db_df.loc[uncategorized_mask, "Category"] = db_df.loc[uncategorized_mask].apply(
                            lambda row: smart_categorize(row["Description"], row["Amount"], use_gemini=True),
                            axis=1,
                        )
                    write_transactions_db(db_df)
                    st.toast("✅ Auto-saved Gemini categories to transactions_db.csv", icon="💾")
                    st.rerun()

                # SECTION C: INSTANT TABLE AUTO-SAVE
                with st.expander("Manually categorize individual transactions", expanded=False):
                    if "editor_version" not in st.session_state:
                        st.session_state["editor_version"] = 0

                    manual_df = db_df[uncategorized_mask].copy().reset_index(names="RowId")
                    manual_df["Date"] = manual_df["Date"].dt.strftime("%Y-%m-%d")
                    manual_df = manual_df[["RowId", "Date", "Description", "Amount", "Source", "Category"]]

                    auto_learn_toggle = st.checkbox("Automatically remember selections as rules for future statements", value=True)

                    editor_key = f"manual_uncategorized_editor_{st.session_state['editor_version']}"
                    edited_manual_df = st.data_editor(
                        manual_df,
                        key=editor_key,
                        hide_index=True,
                        width="stretch",
                        disabled=["RowId", "Date", "Description", "Amount", "Source"],
                        column_config={
                            "RowId": None,
                            "Amount": st.column_config.NumberColumn("Amount", format="$%.2f"),
                            "Category": st.column_config.SelectboxColumn("Category", options=get_all_categories(), required=True),
                        },
                    )

                    # Intercept any category edit and immediately flush to transactions_db.csv on disk
                    diff_mask = edited_manual_df["Category"] != manual_df["Category"]
                    if diff_mask.any():
                        changed_rows = edited_manual_df[diff_mask]
                        rules = load_user_rules()
                        for _, ch_row in changed_rows.iterrows():
                            r_id = int(ch_row["RowId"])
                            new_cat = ch_row["Category"]
                            desc = ch_row["Description"]
                            if new_cat in get_all_categories() and new_cat != "Uncategorized":
                                db_df.loc[r_id, "Category"] = new_cat
                                if auto_learn_toggle:
                                    rules[desc] = new_cat
                                    db_df.loc[db_df["Description"] == desc, "Category"] = new_cat
                        if auto_learn_toggle:
                            save_user_rules(rules)
                        write_transactions_db(db_df)
                        st.session_state["editor_version"] += 1
                        st.toast("✅ Auto-saved changes directly to transactions_db.csv!", icon="💾")
                        st.rerun()

            # --- FILTERS: CATEGORIES & DATE RANGE ---
            st.markdown("---")
            filter_col1, filter_col2, filter_col3 = st.columns([2, 1, 1])

            with filter_col1:
                default_exclusions = sorted([c for c in SPENDING_EXCLUSIONS if c in set(db_df["Category"])])
                excluded_categories = st.multiselect(
                    "Exclude non-spending categories",
                    sorted(db_df["Category"].dropna().unique()),
                    default=default_exclusions,
                )

            expenses_df = prepare_expenses(db_df, excluded_categories)
            start_date, end_date = None, None

            if not expenses_df.empty:
                min_avail = expenses_df["Date"].min().date()
                max_avail = expenses_df["Date"].max().date()
                today = datetime.date.today()

                with filter_col2:
                    time_preset = st.selectbox(
                        "Date Filter",
                        ["All Time", "Last 30 Days", "Last 90 Days", "Year to Date (YTD)", "Custom Range"],
                        index=0,
                    )

                if time_preset == "Last 30 Days":
                    start_date = today - datetime.timedelta(days=30)
                elif time_preset == "Last 90 Days":
                    start_date = today - datetime.timedelta(days=90)
                elif time_preset == "Year to Date (YTD)":
                    start_date = datetime.date(today.year, 1, 1)
                elif time_preset == "Custom Range":
                    with filter_col3:
                        selected_range = st.date_input("Select date range", value=(min_avail, max_avail), min_value=min_avail, max_value=max_avail)
                        if isinstance(selected_range, (tuple, list)) and len(selected_range) == 2:
                            start_date, end_date = selected_range

                if start_date:
                    expenses_df = expenses_df[expenses_df["Date"].dt.date >= start_date]
                if end_date:
                    expenses_df = expenses_df[expenses_df["Date"].dt.date <= end_date]

            # --- VISUALIZATIONS ---
            if expenses_df.empty:
                st.info("No expense data found for the selected categories and date range.")
            else:
                monthly_totals = expenses_df.groupby("MonthDate", as_index=False)["Amount"].sum()
                monthly_totals["Month"] = monthly_totals["MonthDate"].dt.strftime("%b %Y")

                st.markdown("### Spending Visualizations")
                tab_bar, tab_tree, tab_line = st.tabs(["📊 Stacked Bar (Totals)", "🌲 Treemap (Proportions)", "📈 Line Chart (Trends)"])

                with tab_bar:
                    col1, col2 = st.columns([2, 1])
                    with col1:
                        fig_bar = px.bar(expenses_df, x="Month", y="Amount", color="Category", title="Monthly Spending by Category")
                        fig_bar.update_layout(barmode="stack", yaxis_tickprefix="$")
                        fig_bar.update_xaxes(categoryorder="array", categoryarray=monthly_totals["Month"].tolist())
                        st.plotly_chart(fig_bar, width="stretch")
                    with col2:
                        fig_total = px.line(monthly_totals, x="Month", y="Amount", markers=True, title="Total Monthly Spending")
                        fig_total.update_layout(yaxis_tickprefix="$")
                        fig_total.update_xaxes(categoryorder="array", categoryarray=monthly_totals["Month"].tolist())
                        st.plotly_chart(fig_total, width="stretch")

                with tab_tree:
                    cat_totals = expenses_df.groupby("Category", as_index=False)["Amount"].sum()
                    cat_totals = cat_totals[cat_totals["Amount"] > 0]
                    fig_tree = px.treemap(cat_totals, path=["Category"], values="Amount", title=f"Spending Breakdown ({time_preset})")
                    fig_tree.update_traces(textinfo="label+value+percent root", texttemplate="%{label}<br>$%{value:,.2f}<br>%{percentRoot:.1%}")
                    st.plotly_chart(fig_tree, width="stretch")

                with tab_line:
                    trend_df = expenses_df.groupby(["MonthDate", "Month", "Category"], as_index=False)["Amount"].sum().sort_values("MonthDate")
                    fig_line = px.line(trend_df, x="Month", y="Amount", color="Category", markers=True, title="Category Trends Over Time")
                    fig_line.update_layout(yaxis_tickprefix="$")
                    fig_line.update_xaxes(categoryorder="array", categoryarray=monthly_totals["Month"].tolist())
                    st.plotly_chart(fig_line, width="stretch")

                pivot_df = pd.pivot_table(expenses_df, values="Amount", index="Category", columns="Month", aggfunc="sum", fill_value=0)
                sorted_cols = sorted(list(pivot_df.columns), key=lambda d: datetime.datetime.strptime(d, "%b %Y"))

                fig_heatmap = px.imshow(
                    pivot_df[sorted_cols], aspect="auto", labels=dict(x="Month", y="Category", color="Spend"), title="Category Spend Heatmap"
                )
                fig_heatmap.update_layout(coloraxis_colorbar_tickprefix="$")
                st.plotly_chart(fig_heatmap, width="stretch")

                st.markdown("### Expense Matrix")
                st.dataframe(pivot_df[sorted_cols].style.format("${:,.0f}"), width="stretch")

    # -- VIEW 3: STATEMENT IMPORTER --
    elif view_selection == "Statement Importer":
        st.subheader("Automated Statement Processing")

        folder_files = discover_statement_files()
        selected_folder_files = st.multiselect("Statements found in this app folder", folder_files, default=folder_files, format_func=lambda p: p.name)
        uploaded_files = st.file_uploader("Upload additional PDF, CSV, or Excel statements", type=["pdf", "csv", "xlsx"], accept_multiple_files=True)

        use_gemini = st.checkbox("Use Gemini for descriptions not matched by local rules", value=GEMINI_CONFIGURED, disabled=not GEMINI_CONFIGURED)
        if not GEMINI_CONFIGURED:
            st.caption("Add GEMINI_API_KEY to Streamlit secrets or your environment to enable Gemini categorization.")

        files_to_process = list(selected_folder_files) + list(uploaded_files or [])
        if st.button("Parse selected statements", type="primary", disabled=not files_to_process):
            all_extracted_dfs = []
            skipped_files = []

            with st.spinner("Parsing documents..."):
                for statement_file in files_to_process:
                    processed_df = process_statement(statement_file)
                    if processed_df is not None and not processed_df.empty:
                        all_extracted_dfs.append(processed_df)
                    else:
                        skipped_files.append(get_file_name(statement_file))

            if all_extracted_dfs:
                combined_df = pd.concat(all_extracted_dfs, ignore_index=True)
                with st.spinner("Categorizing transactions..."):
                    combined_df = categorize_transactions(combined_df, use_gemini=use_gemini)

                saved_df = save_transactions(combined_df)
                st.success(f"Saved {len(combined_df):,} parsed transactions. Database now has {len(saved_df):,} rows.")
                st.write("### Latest Parsed Transactions")
                st.dataframe(combined_df.sort_values("Date", ascending=False), width="stretch")

            if skipped_files:
                st.warning(f"Could not parse: {', '.join(skipped_files)}")

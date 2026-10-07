import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import pdfplumber
import re
import datetime
import openpyxl
import google.generativeai as genai

# -- PAGE CONFIG & AI SETUP --
st.set_page_config(page_title="Financial Dashboard", layout="wide")
st.title("Personal Finance & Equity Dashboard")

# IMPORTANT: Paste your copied API key here
genai.configure(api_key="AQ.Ab8RN6KO6Sd08P_rXG_1Ob5BNFgWZUVAg9B67bIoMspzvatX5A")

# -- CATEGORIZATION RULES --
CATEGORY_RULES = {
    "Grocery": ["safeway", "whole foods", "ayala farms", "fulton family farms", "r. farms", "yee yang farm",
                "draper girls country farm", "costco"],
    "Restaurants": ["sangam indian cuisine", "virasat indian cuisine", "desi pizza house", "mumbai street food",
                    "sabor mexicano", "the hammond", "phurr", "grand central bakery", "julia bakery", "brewed cafe"],
    "Electricity": ["pge", "pacific power", "electric"],
    "Gas": ["nw natural", "gas"],
    "Water": ["water bureau", "utility"],
    "Mortgage": ["mortgage", "home loan"],
    "HOA": ["hoa", "homeowners association"],
    "Car lease": ["toyota financial", "honda financial", "lease"],
    "Telephone, Internet": ["comcast", "xfinity", "at&t", "verizon"],
    "India transfer": ["remitly", "xoom", "wise"]
}


def smart_categorize(description):
    # 1. Fast path: Check hardcoded dictionary
    desc_lower = str(description).lower()
    for category, keywords in CATEGORY_RULES.items():
        if any(keyword in desc_lower for keyword in keywords):
            return category

    # 2. Smart path: Fallback to Gemini for unknown vendors
    try:
        model = genai.GenerativeModel("gemini-1.5-flash")
        valid_categories = list(CATEGORY_RULES.keys())

        prompt = f"""
        You are a financial categorizer. 
        Map this bank transaction description: "{description}"
        To EXACTLY ONE of these categories: {valid_categories}
        If you cannot confidently guess, reply with "Uncategorized".
        Reply with ONLY the category name, nothing else.
        """

        response = model.generate_content(prompt)
        ai_guess = response.text.strip()

        if ai_guess in valid_categories:
            return ai_guess
    except Exception as e:
        pass  # If API fails or rate limits, fail gracefully

    return "Uncategorized"


# -- DATA INGESTION --
@st.cache_data
def load_data(file_path):
    xls = pd.ExcelFile(file_path)
    current_year = str(datetime.datetime.now().year)

    historical_net_worth = []
    current_nw_df = None
    current_exp_df = None

    for sheet_name in xls.sheet_names:
        # Skip the raw data log tab if it exists
        if sheet_name == "Transaction Log":
            continue

        raw_df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

        # Locate the Net Worth / Expense split
        mask = raw_df.apply(lambda col: col.astype(str).str.contains(r'expense|spending', case=False, na=False))

        if mask.any().any():
            expense_start_idx = mask.any(axis=1).idxmax()

            # Isolate Net Worth
            nw_df = raw_df.iloc[:expense_start_idx - 1].dropna(how='all').copy()
            nw_df.columns = nw_df.columns.fillna("").astype(str)

            if sheet_name != current_year:
                # Consolidate historical years to End-of-Year value (assuming value is in last column of last row)
                try:
                    end_of_year_value = float(str(nw_df.iloc[-1, -1]).replace(',', '').replace('$', ''))
                    historical_net_worth.append({
                        "Timeframe": f"{sheet_name}",
                        "Net Worth": end_of_year_value
                    })
                except:
                    pass
            else:
                # Keep current year detailed
                current_nw_df = nw_df

                # Extract active monthly expense grid
                exp_df = raw_df.iloc[expense_start_idx + 1:].copy()
                try:
                    header_idx = exp_df.dropna(how='all').index[0]
                    exp_df.columns = exp_df.loc[header_idx].fillna("Category").astype(str)
                    exp_df = exp_df.loc[header_idx + 1:].dropna(subset=[exp_df.columns[0]])
                    current_exp_df = exp_df.fillna(0)
                except IndexError:
                    pass

    historical_nw_df = pd.DataFrame(historical_net_worth)
    return historical_nw_df, current_nw_df, current_exp_df


try:
    historical_nw_df, current_nw_df, current_exp_df = load_data("Tracker_2024.xlsx")
    data_loaded = True
except FileNotFoundError:
    st.error("Could not find 'Tracker_2024.xlsx'. Please ensure it is in the same folder as app.py.")
    data_loaded = False


# -- SPECIFIC STATEMENT PARSERS --
def parse_amex_csv(file_object):
    df = pd.read_csv(file_object)
    clean_df = df[['Date', 'Description', 'Amount']].copy()
    if 'Category' in df.columns:
        clean_df['Amex_Category'] = df['Category']
    clean_df['Amount'] = -clean_df['Amount']
    return clean_df


def parse_citi_pdf(pdf_file):
    transactions = []
    pattern = re.compile(r'^\s*(?:(\d{2}/\d{2})\s+)?(\d{2}/\d{2})\s+(.+?)\s+(-?\$[\d,]+\.\d{2})\s*$')
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                for line in text.split('\n'):
                    match = pattern.match(line)
                    if match:
                        post_date = match.group(2)
                        desc = match.group(3).strip()
                        raw_amount = match.group(4)
                        amount_val = float(raw_amount.replace('$', '').replace(',', ''))
                        amount_val = -amount_val
                        transactions.append({'Date': post_date, 'Description': desc, 'Amount': amount_val})
    return pd.DataFrame(transactions)


def parse_bofa_checking(pdf_file):
    transactions = []
    line_pattern = re.compile(r'^(\d{2}/\d{2}/\d{2})\s+(.+?)\s+([0-9,]+\.\d{2})$')
    current_section = None
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                for line in text.split('\n'):
                    if "Deposits and other additions" in line:
                        current_section = "deposit"
                        continue
                    elif "Withdrawals and other subtractions" in line or "Checks" in line:
                        current_section = "withdrawal"
                        continue

                    match = line_pattern.match(line.strip())
                    if match:
                        date = match.group(1)
                        desc = match.group(2).strip()
                        raw_amount = match.group(3)
                        amount_val = float(raw_amount.replace(',', ''))
                        if current_section == "withdrawal":
                            amount_val = -amount_val
                        transactions.append({'Date': date, 'Description': desc, 'Amount': amount_val})
    return pd.DataFrame(transactions)


# -- PDF PARSING ROUTER --
def identify_bank(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        first_page_text = pdf.pages[0].extract_text().lower()
        if "citi" in first_page_text or "citicard" in first_page_text:
            return "citi"
        elif "bank of america" in first_page_text or "deposits and other additions" in first_page_text:
            return "bofa"
        elif "chase" in first_page_text:
            return "chase"
    return "unknown"


def process_statement(uploaded_file):
    file_extension = uploaded_file.name.split('.')[-1].lower()

    if file_extension == 'csv':
        df = pd.read_csv(uploaded_file)
        if 'Card Member' in df.columns or 'Appears On Your Statement As' in df.columns:
            st.info(f"Detected: American Express ({uploaded_file.name})")
            uploaded_file.seek(0)
            return parse_amex_csv(uploaded_file)
        else:
            st.error("CSV layout not recognized.")
            return None

    elif file_extension == 'pdf':
        bank_id = identify_bank(uploaded_file)
        uploaded_file.seek(0)

        if bank_id == "citi":
            st.info(f"Detected: Citi Credit Card ({uploaded_file.name})")
            return parse_citi_pdf(uploaded_file)
        elif bank_id == "bofa":
            st.info(f"Detected: Checking Account ({uploaded_file.name})")
            return parse_bofa_checking(uploaded_file)
        else:
            st.error(f"No parser built for this PDF signature: {uploaded_file.name}")
            return None


# -- UI LAYOUT --
if data_loaded:
    st.sidebar.header("Dashboard Controls")
    view_selection = st.sidebar.radio("Select View",
                                      ["Net Worth & Equity", "Cash Flow Analysis", "PDF Statement Importer"])

    # -- VIEW 1: NET WORTH & EQUITY --
    if view_selection == "Net Worth & Equity":
        st.subheader("Asset & Liability Breakdown")

        # Plot Consolidated Historical Net Worth
        if not historical_nw_df.empty:
            st.markdown("### Historical Net Worth Growth")
            fig_hist = px.bar(historical_nw_df, x='Timeframe', y='Net Worth',
                              title="End of Year Net Worth (Consolidated)",
                              labels={'Timeframe': 'Year', 'Net Worth': 'Total Value ($)'})
            st.plotly_chart(fig_hist, width="stretch")

        col1, col2 = st.columns(2)
        with col1:
            st.markdown("### Real Estate Equity (WA vs. NC)")
            categories = ['WA Real Estate', 'NC Real Estate']
            values = [854353, 406362]
            mortgages = [374949, 280375]

            fig_equity = go.Figure(data=[
                go.Bar(name='Mortgage Debt', x=categories, y=mortgages, marker_color='#ef553b'),
                go.Bar(name='Home Equity', x=categories, y=[v - m for v, m in zip(values, mortgages)],
                       marker_color='#00cc96')
            ])
            fig_equity.update_layout(barmode='stack', title="Current Real Estate Equity")
            st.plotly_chart(fig_equity, width="stretch")

        with col2:
            st.markdown("### Liquid vs Locked Assets")
            labels = ['Liquid & Amex', '401k & ESOP', 'Brokerage (Vanguard/TD)']
            values = [58842, 242229, 130795]
            fig_assets = px.pie(values=values, names=labels, hole=0.4, title="Asset Distribution")
            st.plotly_chart(fig_assets, width="stretch")

    # -- VIEW 2: CASH FLOW ANALYSIS --
    elif view_selection == "Cash Flow Analysis":
        st.subheader("Monthly Expense Trends (Current Year)")
        if current_exp_df is not None:
            st.dataframe(current_exp_df, width="stretch")
        else:
            st.info("No expense table found for the current year.")

    # -- VIEW 3: PDF / CSV IMPORTER --
    elif view_selection == "PDF Statement Importer":
        st.subheader("Automated Statement Processing")

        uploaded_files = st.file_uploader("Upload PDF or CSV statements", type=['pdf', 'csv'],
                                          accept_multiple_files=True)

        if uploaded_files:
            all_extracted_dfs = []

            with st.spinner("Routing and parsing documents..."):
                for uploaded_file in uploaded_files:
                    processed_df = process_statement(uploaded_file)
                    if processed_df is not None:
                        all_extracted_dfs.append(processed_df)

            if all_extracted_dfs:
                combined_df = pd.concat(all_extracted_dfs, ignore_index=True)

                # Apply AI Smart Categorization
                with st.spinner("Asking Gemini to categorize unknown transactions..."):
                    combined_df['Category'] = combined_df['Description'].apply(smart_categorize)

                st.write("### All Extracted Transactions")
                st.dataframe(combined_df, width="stretch")

                st.write("### Combined Category Totals")
                spending_df = combined_df[combined_df['Amount'] < 0].copy()
                category_totals = spending_df.groupby('Category')['Amount'].sum().abs().reset_index()
                category_totals.columns = ['Category', 'Total Spent']
                st.dataframe(category_totals, width="stretch")

                # Append to Excel Log
                if st.button("Categorize & Save to Tracker", type="primary"):
                    try:
                        with pd.ExcelWriter("Tracker_2024.xlsx", mode="a", engine="openpyxl",
                                            if_sheet_exists="overlay") as writer:
                            # Check if the log sheet exists to append properly
                            if 'Transaction Log' in writer.sheets:
                                start_row = writer.sheets['Transaction Log'].max_row
                                write_header = False
                            else:
                                start_row = 0
                                write_header = True

                            combined_df.to_excel(
                                writer,
                                sheet_name='Transaction Log',
                                index=False,
                                header=write_header,
                                startrow=start_row
                            )
                        st.success("Successfully appended new transactions to 'Transaction Log' in Tracker_2024.xlsx!")
                        st.cache_data.clear()  # Clears cache so next load sees the new data
                    except PermissionError:
                        st.error("Permission Denied: Please close Tracker_2024.xlsx in Excel before saving.")
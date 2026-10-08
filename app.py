import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import pdfplumber
import re
import datetime
import os
import google.generativeai as genai

# -- PAGE CONFIG & AI SETUP --
st.set_page_config(page_title="Financial Dashboard", layout="wide", page_icon="📈")
st.title("Personal Finance & Equity Dashboard")

# Safely load the API key from Streamlit Secrets
if "GEMINI_API_KEY" in st.secrets:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])

DB_FILE = "transactions_db.csv"

# -- CATEGORIZATION RULES --
CATEGORY_RULES = {
    "Grocery": ["safeway", "whole foods", "ayala farms", "fulton family farms", "r. farms", "yee yang farm", "draper girls country farm", "costco"],
    "Restaurants": ["sangam indian cuisine", "virasat indian cuisine", "desi pizza house", "mumbai street food", "sabor mexicano", "the hammond", "phurr", "grand central bakery", "julia bakery", "brewed cafe"],
    "Electricity": ["pge", "pacific power", "electric"],
    "Gas": ["nw natural", "gas"],
    "Water": ["water bureau", "utility"],
    "Mortgage": ["mortgage", "home loan"],
    "HOA": ["hoa", "homeowners association"],
    "Car lease": ["toyota financial", "honda financial", "lease"],
    "Daycare / Nanny": ["daycare", "nanny", "childcare"],
    "Telephone, Internet": ["comcast", "xfinity", "at&t", "verizon"],
    "India transfer": ["remitly", "xoom", "wise"]
}

def smart_categorize(description):
    desc_lower = str(description).lower()
    for category, keywords in CATEGORY_RULES.items():
        if any(keyword in desc_lower for keyword in keywords):
            return category
            
    try:
        model = genai.GenerativeModel("gemini-1.5-flash")
        valid_categories = list(CATEGORY_RULES.keys())
        prompt = f"""
        Map this bank transaction description: "{description}"
        To EXACTLY ONE of these categories: {valid_categories}
        If you cannot confidently guess, reply with "Uncategorized".
        Reply with ONLY the category name.
        """
        response = model.generate_content(prompt)
        ai_guess = response.text.strip()
        if ai_guess in valid_categories:
            return ai_guess
    except Exception:
        pass
        
    return "Uncategorized"

# -- NET WORTH EXCEL PARSER (CACHE REMOVED FOR LIVE UPDATES) --
def load_net_worth_data(file_path):
    xls = pd.ExcelFile(file_path)
    nw_data = []
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    
    for sheet_name in xls.sheet_names:
        raw_df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        
        # 1. Fuzzy search for the exact "Net Worth" column index
        nw_col_idx = -1
        for r_idx in range(min(15, len(raw_df))):
            for c_idx in range(len(raw_df.columns)):
                cell_val = str(raw_df.iloc[r_idx, c_idx]).strip().lower()
                if 'net' in cell_val and 'worth' in cell_val:
                    nw_col_idx = c_idx
                    break
            if nw_col_idx != -1:
                break
        
        for idx, row in raw_df.iterrows():
            month_str = None
            
            # 2. Bulletproof Date/Month detection
            for col_idx in range(min(4, len(row))):
                val = row.iloc[col_idx]
                if pd.notna(val):
                    if isinstance(val, (datetime.datetime, datetime.date, pd.Timestamp)):
                        month_str = val.strftime('%b')
                        break
                    
                    val_str = str(val).strip()
                    if any(m in val_str for m in months) and (val_str[0].isdigit() or '-' in val_str):
                        month_str = [m for m in months if m in val_str][0]
                        break
                        
                    try:
                        parsed_date = pd.to_datetime(val_str)
                        month_str = parsed_date.strftime('%b')
                        break
                    except:
                        pass
                            
            if month_str:
                net_worth = 0
                
                # 3. Look down the specific "Net Worth" column (checking adjacent columns for merged cells)
                if nw_col_idx != -1:
                    for current_idx in range(idx + 1, min(idx + 7, len(raw_df))):
                        for col_offset in [-1, 0, 1]:
                            target_col = nw_col_idx + col_offset
                            if 0 <= target_col < len(raw_df.columns):
                                cell_val = raw_df.iloc[current_idx, target_col]
                                if pd.notna(cell_val):
                                    try:
                                        clean = str(cell_val).replace('$', '').replace(',', '').strip()
                                        num = float(clean)
                                        if num > net_worth:
                                            net_worth = num
                                    except:
                                        pass
                                        
                if net_worth > 0:
                    nw_data.append({
                        'Year': str(sheet_name),
                        'Month': month_str,
                        'Net Worth': net_worth
                    })
                        
    return pd.DataFrame(nw_data)

try:
    nw_df = load_net_worth_data("Tracker_2024.xlsx")
    data_loaded = True
except FileNotFoundError:
    st.error("Could not find 'Tracker_2024.xlsx'. Please ensure it is in the same folder as app.py.")
    data_loaded = False


# -- SPECIFIC STATEMENT PARSERS --
def parse_amex_csv(file_object):
    df = pd.read_csv(file_object)
    clean_df = df[['Date', 'Description', 'Amount']].copy()
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
                        post_date = match.group(2) + f"/{datetime.datetime.now().year}"
                        amount_val = -float(match.group(4).replace('$', '').replace(',', ''))
                        transactions.append({'Date': post_date, 'Description': match.group(3).strip(), 'Amount': amount_val})
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
                        amount_val = float(match.group(3).replace(',', ''))
                        if current_section == "withdrawal":
                            amount_val = -amount_val
                        transactions.append({'Date': match.group(1), 'Description': match.group(2).strip(), 'Amount': amount_val})
    return pd.DataFrame(transactions)

def identify_bank(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        first_page_text = pdf.pages[0].extract_text().lower()
        if "citi" in first_page_text or "citicard" in first_page_text: return "citi"
        elif "bank of america" in first_page_text or "deposits" in first_page_text: return "bofa"
        elif "chase" in first_page_text: return "chase"
    return "unknown"

def process_statement(uploaded_file):
    file_extension = uploaded_file.name.split('.')[-1].lower()
    if file_extension == 'csv':
        return parse_amex_csv(uploaded_file)
    elif file_extension == 'pdf':
        bank_id = identify_bank(uploaded_file)
        uploaded_file.seek(0) 
        if bank_id == "citi": return parse_citi_pdf(uploaded_file)
        elif bank_id == "bofa": return parse_bofa_checking(uploaded_file)
    return None


# -- UI LAYOUT --
if data_loaded:
    st.sidebar.header("Dashboard Controls")
    view_selection = st.sidebar.radio("Navigation", ["Net Worth Analysis", "Cash Flow Analysis", "PDF Statement Importer"])

    # -- VIEW 1: NET WORTH ANALYSIS --
    if view_selection == "Net Worth Analysis

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

# IMPORTANT: Paste your copied API key here
genai.configure(api_key="YOUR_GEMINI_API_KEY")

# Local database file
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

# -- NET WORTH EXCEL PARSER --
@st.cache_data
def load_net_worth_data(file_path):
    xls = pd.ExcelFile(file_path)
    nw_data = []
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    
    for sheet_name in xls.sheet_names:
        raw_df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        
        for idx, row in raw_df.iterrows():
            month_str = None
            
            # 1. Search the first 4 columns for a Date/Month
            for col_idx in range(min(4, len(row))):
                val = row.iloc[col_idx]
                if pd.notna(val):
                    if isinstance(val, (datetime.datetime, datetime.date, pd.Timestamp)):
                        month_str = val.strftime('%b')
                        break
                    elif isinstance(val, str):
                        val_str = val.strip()
                        if any(m in val_str for m in months) and (val_str[0].isdigit() or '-' in val_str):
                            month_str = [m for m in months if m in val_str][0]
                            break
                            
            if month_str:
                # 2. Find the "Value" header column in this block
                value_col_idx = -1
                for search_idx in range(max(0, idx-2), min(len(raw_df), idx+3)):
                    for i, cell in enumerate(raw_df.iloc[search_idx]):
                        if pd.notna(cell) and 'value' in str(cell).lower() and i > 4:
                            value_col_idx = i
                            break
                    if value_col_idx != -1: break
                    
                # Fallback if "Value" header is missing
                if value_col_idx == -1:
                    value_col_idx = len(row) - 4
                    
                net_worth = 0
                
                # 3. Only scan the immediate next 4 rows (to avoid bleeding into the next month)
                for current_idx in range(idx + 1, min(idx + 5, len(raw_df))):
                    b_row = raw_df.iloc[current_idx]
                    
                    # Stop if we accidentally hit the next month's block
                    hit_next_month = False
                    for c_idx in range(min(4, len(b_row))):
                        cell_val = b_row.iloc[c_idx]
                        if pd.notna(cell_val):
                            val_str = str(cell_val).strip()
                            if any(m in val_str for m in months) and (val_str[0].isdigit() or '-' in val_str):
                                hit_next_month = True
                    if hit_next_month: break
                    
                    # Ensure this is an N, R, or NR row
                    is_data_row = False
                    for c_idx in range(min(4, len(b_row))):
                        cell_val = b_row.iloc[c_idx]
                        if pd.notna(cell_val):
                            clean_val = str(cell_val).strip().upper()
                            if clean_val in ['N', 'R', 'NR']:
                                is_data_row = True
                                break
                                
                    if is_data_row:
                        # 4. Check the cells in and around the Value column (spanning right) to find the absolute max total
                        for col_offset in range(-2, 6): 
                            target_col = value_col_idx + col_offset
                            if 0 <= target_col < len(b_row):
                                cell_val = b_row.iloc[target_col]
                                if pd.notna(cell_val):
                                    try:
                                        clean = str(cell_val).replace('$', '').replace(',', '').strip()
                                        num = float(clean)
                                        if num > net_worth:
                                            net_worth = num # Grabs the largest aggregate total
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
    if view_selection == "Net Worth Analysis":
        st.subheader("Historical & Current Net Worth Analysis")
        
        if not nw_df.empty:
            current_year = str(datetime.datetime.now().year)
            available_years = nw_df['Year'].unique().tolist()
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("### Historical Annual Comparison")
                historical_years = [y for y in available_years if y != current_year]
                
                if historical_years:
                    selected_years = st.multiselect("Select years to compare:", historical_years, default=historical_years)
                    
                    if selected_years:
                        annual_data = []
                        for y in selected_years:
                            year_data = nw_df[nw_df['Year'] == y]
                            if not year_data.empty:
                                last_month_val = year_data.iloc[-1]['Net Worth']
                                annual_data.append({'Year': y, 'End of Year Net Worth': last_month_val})
                                
                        if annual_data:
                            annual_df = pd.DataFrame(annual_data)
                            fig_annual = px.bar(annual_df, x='Year', y='End of Year Net Worth', text_auto='.3s')
                            fig_annual.update_layout(xaxis_type='category') 
                            st.plotly_chart(fig_annual, width="stretch")
                else:
                    st.info("No historical years found in tracker.")
                    
            with col2:
                st.markdown(f"### {current_year} Month-to-Month Progression")
                current_year_data = nw_df[nw_df['Year'] == current_year]
                
                if not current_year_data.empty:
                    fig_monthly = px.line(current_year_data, x='Month', y='Net Worth', markers=True)
                    st.plotly_chart(fig_monthly, width="stretch")
                else:
                    st.info(f"No Net Worth data found yet for {current_year}.")
        else:
            st.info("No Net Worth data could be extracted from Tracker_2024.xlsx. Ensure the 'Value' column and 'N/R/NR' rows exist.")

    # -- VIEW 2: CASH FLOW ANALYSIS (ON-SITE COMPARSION) --
    elif view_selection == "Cash Flow Analysis":
        st.subheader("Month-Over-Month Expense Comparison")
        
        if os.path.exists(DB_FILE):
            db_df = pd.read_csv(DB_FILE)
            db_df['Date'] = pd.to_datetime(db_df['Date'])
            
            expenses_df = db_df[db_df['Amount'] < 0].copy()
            expenses_df['Amount'] = expenses_df['Amount'].abs()
            expenses_df['Month'] = expenses_df['Date'].dt.strftime('%b %Y')
            
            if not expenses_df.empty:
                fig = px.bar(expenses_df, x="Month", y="Amount", color="Category", title="Monthly Outflows by Category")
                fig.update_layout(barmode='stack', xaxis={'categoryorder': 'category ascending'})
                st.plotly_chart(fig, width="stretch")
                
                st.markdown("### Expense Matrix")
                pivot_df = pd.pivot_table(expenses_df, values='Amount', index='Category', columns='Month', aggfunc='sum', fill_value=0)
                sorted_cols = sorted(list(pivot_df.columns), key=lambda d: datetime.datetime.strptime(d, "%b %Y"))
                st.dataframe(pivot_df[sorted_cols], width="stretch")
            else:
                st.info("No expenses logged yet.")
        else:
            st.warning("No data found. Please upload your statements in the PDF Statement Importer tab to generate your charts.")

    # -- VIEW 3: PDF / CSV IMPORTER --
    elif view_selection == "PDF Statement Importer":
        st.subheader("Automated Statement Processing")
        
        uploaded_files = st.file_uploader("Upload PDF or CSV statements", type=['pdf', 'csv'], accept_multiple_files=True)
        
        if uploaded_files:
            all_extracted_dfs = []
            with st.spinner("Parsing documents..."):
                for uploaded_file in uploaded_files:
                    processed_df = process_statement(uploaded_file)
                    if processed_df is not None:
                        all_extracted_dfs.append(processed_df)
                
            if all_extracted_dfs:
                combined_df = pd.concat(all_extracted_dfs, ignore_index=True)
                
                with st.spinner("Asking Gemini to categorize unfamiliar transactions..."):
                    combined_df['Category'] = combined_df['Description'].apply(smart_categorize)
                
                st.write("### Review Extracted Transactions")
                st.dataframe(combined_df, width="stretch")
                
                if st.button("Save to Local Database", type="primary"):
                    if os.path.exists(DB_FILE):
                        combined_df.to_csv(DB_FILE, mode='a', header=False, index=False)
                    else:
                        combined_df.to_csv(DB_FILE, index=False)
                    st.success("Successfully saved! Head to the 'Cash Flow Analysis' tab to see your updated charts.")

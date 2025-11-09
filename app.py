import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# --- Page Configuration ---
st.set_page_config(
    page_title="DCF Model",
    page_icon="💹",
    layout="wide"
)

# --- Helper Functions ---

def format_value(value, value_type="currency"):
    """Formats a number as currency or percentage."""
    if value is None or not isinstance(value, (int, float)):
        return value
    if value_type == "currency":
        return f"${value:,.2f}M"
    elif value_type == "percentage":
        return f"{value:.2%}"
    elif value_type == "multiple":
        return f"{value:.1f}x"
    else:
        return f"{value:,.2f}"

def to_excel(df_dict):
    """Exports a dictionary of DataFrames to an Excel file in memory."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in df_dict.items():
            # Slice multi-index dataframes for clean export
            if isinstance(df.index, pd.MultiIndex):
                try:
                    df.loc['Income Statement'].to_excel(writer, sheet_name='Income Statement')
                    df.loc['Balance Sheet'].to_excel(writer, sheet_name='Balance Sheet')
                    df.loc['Cash Flow Statement'].to_excel(writer, sheet_name='Cash Flow Statement')
                except KeyError:
                    # Handle cases where one of the statements might be missing in a slice
                    df.to_excel(writer, sheet_name=sheet_name)
            else:
                df.to_excel(writer, sheet_name=sheet_name)
    processed_data = output.getvalue()
    return processed_data

def get_template_excel(historical_years):
    """Generates a blank Excel template in memory for data upload."""
    output = BytesIO()
    
    # Define the required rows
    income_statement_items = ['Revenue', 'COGS', 'SG&A', 'D&A', 'Interest Expense']
    balance_sheet_items = [
        'Cash', 'Accounts Receivable', 'Inventory', 'PP&E', 'Accounts Payable',
        'Accrued Liabilities', 'Long-Term Debt', 'Common Stock', 'Retained Earnings'
    ]
    
    # Create blank DataFrames
    df_is = pd.DataFrame(0.0, index=income_statement_items, columns=historical_years)
    df_bs = pd.DataFrame(0.0, index=balance_sheet_items, columns=historical_years)
    
    df_is.index.name = "Metric"
    df_bs.index.name = "Metric"

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_is.to_excel(writer, sheet_name='Income Statement')
        df_bs.to_excel(writer, sheet_name='Balance Sheet')
        
    processed_data = output.getvalue()
    return processed_data

# --- Core Modeling Functions ---

def build_financial_statements(assumptions, historical_data, projection_years_list):
    """Builds the three financial statements based on historicals and assumptions."""
    # Ensure historical_years is sorted
    historical_years = sorted(historical_data.keys())
    all_years = historical_years + projection_years_list

    # Initialize DataFrame
    idx = pd.MultiIndex.from_tuples([
        ('Income Statement', 'Revenue'), ('Income Statement', 'COGS'), ('Income Statement', 'Gross Profit'),
        ('Income Statement', 'SG&A'), ('Income Statement', 'EBITDA'), ('Income Statement', 'D&A'),
        ('Income Statement', 'EBIT'), ('Income Statement', 'Interest Expense'), ('Income Statement', 'EBT'),
        ('Income Statement', 'Taxes'), ('Income Statement', 'Net Income'),
        ('Balance Sheet', 'Cash & Cash Equivalents'), ('Balance Sheet', 'Accounts Receivable'),
        ('Balance Sheet', 'Inventory'), ('Balance Sheet', 'Total Current Assets'),
        ('Balance Sheet', 'PP&E, Net'), ('Balance Sheet', 'Total Assets'),
        ('Balance Sheet', 'Accounts Payable'), ('Balance Sheet', 'Accrued Liabilities'),
        ('Balance Sheet', 'Total Current Liabilities'), ('Balance Sheet', 'Long-Term Debt'),
        ('Balance Sheet', 'Total Liabilities'), ('Balance Sheet', 'Common Stock'),
        ('Balance Sheet', 'Retained Earnings'), ('Balance Sheet', 'Total Equity'),
        ('Balance Sheet', 'Total Liabilities & Equity'), ('Balance Sheet', 'Balance Check'),
        ('Cash Flow Statement', 'Net Income'), ('Cash Flow Statement', 'D&A'),
        ('Cash Flow Statement', 'Change in Accounts Receivable'), ('Cash Flow Statement', 'Change in Inventory'),
        ('Cash Flow Statement', 'Change in Accounts Payable'), ('Cash Flow Statement', 'Change in Accrued Liabilities'),
        ('Cash Flow Statement', 'Cash Flow from Operations (CFO)'),
        ('Cash Flow Statement', 'Capital Expenditures (Capex)'), ('Cash Flow Statement', 'Cash Flow from Investing (CFI)'),
        ('Cash Flow Statement', 'Debt Issuance / (Repayment)'), ('Cash Flow Statement', 'Cash Flow from Financing (CFF)'),
        ('Cash Flow Statement', 'Net Change in Cash'),
    ])
    model = pd.DataFrame(0.0, index=idx, columns=all_years)

    # Populate Historical Data
    for year in historical_years:
        try:
            model.loc[('Income Statement', 'Revenue'), year] = historical_data[year]['Revenue']
            model.loc[('Income Statement', 'COGS'), year] = historical_data[year]['COGS']
            model.loc[('Income Statement', 'SG&A'), year] = historical_data[year]['SG&A']
            model.loc[('Income Statement', 'D&A'), year] = historical_data[year]['D&A']
            model.loc[('Income Statement', 'Interest Expense'), year] = historical_data[year]['Interest Expense']
            model.loc[('Balance Sheet', 'Cash & Cash Equivalents'), year] = historical_data[year]['Cash']
            model.loc[('Balance Sheet', 'Accounts Receivable'), year] = historical_data[year]['Accounts Receivable']
            model.loc[('Balance Sheet', 'Inventory'), year] = historical_data[year]['Inventory']
            model.loc[('Balance Sheet', 'PP&E, Net'), year] = historical_data[year]['PP&E']
            model.loc[('Balance Sheet', 'Accounts Payable'), year] = historical_data[year]['Accounts Payable']
            model.loc[('Balance Sheet', 'Accrued Liabilities'), year] = historical_data[year]['Accrued Liabilities']
            model.loc[('Balance Sheet', 'Long-Term Debt'), year] = historical_data[year]['Long-Term Debt']
            model.loc[('Balance Sheet', 'Common Stock'), year] = historical_data[year]['Common Stock']
            model.loc[('Balance Sheet', 'Retained Earnings'), year] = historical_data[year]['Retained Earnings']
        except KeyError as e:
            st.error(f"Error populating historicals: Missing key {e} for year {year}. Check your inputs.")
            st.stop()

    # Projection Loop
    for i, year in enumerate(all_years):
        prev_year = all_years[i-1] if i > 0 else None

        # --- Income Statement ---
        if year in projection_years_list:
            model.loc[('Income Statement', 'Revenue'), year] = model.loc[('Income Statement', 'Revenue'), prev_year] * (1 + assumptions['revenue_growth_rate'])
            model.loc[('Income Statement', 'COGS'), year] = model.loc[('Income Statement', 'Revenue'), year] * assumptions['cogs_percent_revenue']
            model.loc[('Income Statement', 'SG&A'), year] = model.loc[('Income Statement', 'Revenue'), year] * assumptions['sga_percent_revenue']
            model.loc[('Income Statement', 'D&A'), year] = model.loc[('Balance Sheet', 'PP&E, Net'), prev_year] * assumptions['depreciation_rate']
            model.loc[('Income Statement', 'Interest Expense'), year] = model.loc[('Balance Sheet', 'Long-Term Debt'), prev_year] * assumptions['cost_of_debt']

        model.loc[('Income Statement', 'Gross Profit'), year] = model.loc[('Income Statement', 'Revenue'), year] - model.loc[('Income Statement', 'COGS'), year]
        model.loc[('Income Statement', 'EBITDA'), year] = model.loc[('Income Statement', 'Gross Profit'), year] - model.loc[('Income Statement', 'SG&A'), year]
        model.loc[('Income Statement', 'EBIT'), year] = model.loc[('Income Statement', 'EBITDA'), year] - model.loc[('Income Statement', 'D&A'), year]
        model.loc[('Income Statement', 'EBT'), year] = model.loc[('Income Statement', 'EBIT'), year] - model.loc[('Income Statement', 'Interest Expense'), year]
        
        # Ensure EBT is not negative before applying tax (prevents tax *benefit* in this simple model)
        ebt_for_tax = model.loc[('Income Statement', 'EBT'), year]
        model.loc[('Income Statement', 'Taxes'), year] = max(0, ebt_for_tax) * assumptions['tax_rate']
        model.loc[('Income Statement', 'Net Income'), year] = model.loc[('Income Statement', 'EBT'), year] - model.loc[('Income Statement', 'Taxes'), year]

        # --- Balance Sheet Drivers (Projections only) ---
        if year in projection_years_list:
            model.loc[('Balance Sheet', 'Accounts Receivable'), year] = model.loc[('Income Statement', 'Revenue'), year] * (assumptions['ar_days'] / 365)
            model.loc[('Balance Sheet', 'Inventory'), year] = model.loc[('Income Statement', 'COGS'), year] * (assumptions['inventory_days'] / 365)
            model.loc[('Balance Sheet', 'Accounts Payable'), year] = model.loc[('Income Statement', 'COGS'), year] * (assumptions['ap_days'] / 365)
            model.loc[('Balance Sheet', 'Accrued Liabilities'), year] = model.loc[('Income Statement', 'Revenue'), year] * assumptions['accrued_liabilities_percent_revenue']

        # --- Cash Flow Statement (Projections only) ---
        if year in projection_years_list:
            model.loc[('Cash Flow Statement', 'Net Income'), year] = model.loc[('Income Statement', 'Net Income'), year]
            model.loc[('Cash Flow Statement', 'D&A'), year] = model.loc[('Income Statement', 'D&A'), year]
            model.loc[('Cash Flow Statement', 'Change in Accounts Receivable'), year] = model.loc[('Balance Sheet', 'Accounts Receivable'), prev_year] - model.loc[('Balance Sheet', 'Accounts Receivable'), year]
            model.loc[('Cash Flow Statement', 'Change in Inventory'), year] = model.loc[('Balance Sheet', 'Inventory'), prev_year] - model.loc[('Balance Sheet', 'Inventory'), year]
            model.loc[('Cash Flow Statement', 'Change in Accounts Payable'), year] = model.loc[('Balance Sheet', 'Accounts Payable'), year] - model.loc[('Balance Sheet', 'Accounts Payable'), prev_year]
            model.loc[('Cash Flow Statement', 'Change in Accrued Liabilities'), year] = model.loc[('Balance Sheet', 'Accrued Liabilities'), year] - model.loc[('Balance Sheet', 'Accrued Liabilities'), prev_year]

            cfo_rows = [
                ('Cash Flow Statement', 'Net Income'), ('Cash Flow Statement', 'D&A'),
                ('Cash Flow Statement', 'Change in Accounts Receivable'), ('Cash Flow Statement', 'Change in Inventory'),
                ('Cash Flow Statement', 'Change in Accounts Payable'), ('Cash Flow Statement', 'Change in Accrued Liabilities')
            ]
            model.loc[('Cash Flow Statement', 'Cash Flow from Operations (CFO)'), year] = model.loc[cfo_rows, year].sum()

            model.loc[('Cash Flow Statement', 'Capital Expenditures (Capex)'), year] = -(model.loc[('Income Statement', 'Revenue'), year] * assumptions['capex_percent_revenue'])
            model.loc[('Cash Flow Statement', 'Cash Flow from Investing (CFI)'), year] = model.loc[('Cash Flow Statement', 'Capital Expenditures (Capex)'), year]
            
            repayment = model.loc[('Balance Sheet', 'Long-Term Debt'), prev_year] * assumptions['debt_repayment_percent']
            model.loc[('Cash Flow Statement', 'Debt Issuance / (Repayment)'), year] = -repayment
            model.loc[('Cash Flow Statement', 'Cash Flow from Financing (CFF)'), year] = model.loc[('Cash Flow Statement', 'Debt Issuance / (Repayment)'), year]

            net_cash_change_rows = [
                ('Cash Flow Statement', 'Cash Flow from Operations (CFO)'),
                ('Cash Flow Statement', 'Cash Flow from Investing (CFI)'),
                ('Cash Flow Statement', 'Cash Flow from Financing (CFF)')
            ]
            model.loc[('Cash Flow Statement', 'Net Change in Cash'), year] = model.loc[net_cash_change_rows, year].sum()

        # --- Link CFS to Balance Sheet (Projections only) ---
        if year in projection_years_list:
            model.loc[('Balance Sheet', 'Cash & Cash Equivalents'), year] = model.loc[('Balance Sheet', 'Cash & Cash Equivalents'), prev_year] + model.loc[('Cash Flow Statement', 'Net Change in Cash'), year]
            model.loc[('Balance Sheet', 'PP&E, Net'), year] = model.loc[('Balance Sheet', 'PP&E, Net'), prev_year] + abs(model.loc[('Cash Flow Statement', 'Capital Expenditures (Capex)'), year]) - model.loc[('Income Statement', 'D&A'), year]
            model.loc[('Balance Sheet', 'Long-Term Debt'), year] = model.loc[('Balance Sheet', 'Long-Term Debt'), prev_year] + model.loc[('Cash Flow Statement', 'Debt Issuance / (Repayment)'), year]
            model.loc[('Balance Sheet', 'Retained Earnings'), year] = model.loc[('Balance Sheet', 'Retained Earnings'), prev_year] + model.loc[('Income Statement', 'Net Income'), year]
            model.loc[('Balance Sheet', 'Common Stock'), year] = model.loc[('Balance Sheet', 'Common Stock'), prev_year]

        # --- Final Balance Sheet Calculations (for all years) ---
        tca_rows = [('Balance Sheet', 'Cash & Cash Equivalents'), ('Balance Sheet', 'Accounts Receivable'), ('Balance Sheet', 'Inventory')]
        model.loc[('Balance Sheet', 'Total Current Assets'), year] = model.loc[tca_rows, year].sum()
        model.loc[('Balance Sheet', 'Total Assets'), year] = model.loc[('Balance Sheet', 'Total Current Assets'), year] + model.loc[('Balance Sheet', 'PP&E, Net'), year]
        
        tcl_rows = [('Balance Sheet', 'Accounts Payable'), ('Balance Sheet', 'Accrued Liabilities')]
        model.loc[('Balance Sheet', 'Total Current Liabilities'), year] = model.loc[tcl_rows, year].sum()
        model.loc[('Balance Sheet', 'Total Liabilities'), year] = model.loc[('Balance Sheet', 'Total Current Liabilities'), year] + model.loc[('Balance Sheet', 'Long-Term Debt'), year]
        
        model.loc[('Balance Sheet', 'Total Equity'), year] = model.loc[('Balance Sheet', 'Common Stock'), year] + model.loc[('Balance Sheet', 'Retained Earnings'), year]
        model.loc[('Balance Sheet', 'Total Liabilities & Equity'), year] = model.loc[('Balance Sheet', 'Total Liabilities'), year] + model.loc[('Balance Sheet', 'Total Equity'), year]
        model.loc[('Balance Sheet', 'Balance Check'), year] = model.loc[('Balance Sheet', 'Total Assets'), year] - model.loc[('Balance Sheet', 'Total Liabilities & Equity'), year]

    return model

def build_dcf_model(statements, assumptions, projection_years_list, historical_data):
    """Builds the DCF model."""
    dcf_idx = [
        'EBIT', 'Taxes on EBIT', 'NOPAT', 'D&A', 'Capital Expenditures (Capex)',
        'Change in Net Working Capital', 'Unlevered Free Cash Flow (UFCF)',
        'Discount Factor', 'PV of UFCF'
    ]
    dcf = pd.DataFrame(0.0, index=dcf_idx, columns=projection_years_list)

    # Calculate UFCF
    nwc = (statements.loc[('Balance Sheet', 'Accounts Receivable')] +
           statements.loc[('Balance Sheet', 'Inventory')] -
           statements.loc[('Balance Sheet', 'Accounts Payable')] -
           statements.loc[('Balance Sheet', 'Accrued Liabilities')])
    change_in_nwc = nwc.diff()

    dcf.loc['EBIT'] = statements.loc[('Income Statement', 'EBIT'), projection_years_list]
    # Use max(0, EBIT) for tax calculation to avoid tax shield on negative EBIT
    dcf.loc['Taxes on EBIT'] = dcf.loc['EBIT'].apply(lambda x: max(0, x)) * assumptions['tax_rate']
    dcf.loc['NOPAT'] = dcf.loc['EBIT'] - dcf.loc['Taxes on EBIT']
    dcf.loc['D&A'] = statements.loc[('Income Statement', 'D&A'), projection_years_list]
    dcf.loc['Capital Expenditures (Capex)'] = statements.loc[('Cash Flow Statement', 'Capital Expenditures (Capex)'), projection_years_list]
    dcf.loc['Change in Net Working Capital'] = -change_in_nwc[projection_years_list]
    dcf.loc['Unlevered Free Cash Flow (UFCF)'] = dcf.loc[['NOPAT', 'D&A', 'Capital Expenditures (Capex)', 'Change in Net Working Capital']].sum(axis=0)

    # Discount UFCF
    discount_periods = np.arange(1, len(projection_years_list) + 1)
    dcf.loc['Discount Factor'] = 1 / (1 + assumptions['wacc']) ** discount_periods
    dcf.loc['PV of UFCF'] = dcf.loc['Unlevered Free Cash Flow (UFCF)'] * dcf.loc['Discount Factor']

    # Terminal Value
    last_proj_year = projection_years_list[-1]
    last_ufcf = dcf.loc['Unlevered Free Cash Flow (UFCF)', last_proj_year]
    
    if (assumptions['wacc'] - assumptions['perpetual_growth_rate']) == 0:
        st.sidebar.error("WACC and Perpetual Growth Rate cannot be equal.")
        st.stop()
        
    terminal_value = (last_ufcf * (1 + assumptions['perpetual_growth_rate'])) / (assumptions['wacc'] - assumptions['perpetual_growth_rate'])
    pv_terminal_value = terminal_value * dcf.loc['Discount Factor', last_proj_year]

    # Enterprise and Equity Value
    enterprise_value = dcf.loc['PV of UFCF'].sum() + pv_terminal_value
    latest_historical_year = sorted(historical_data.keys())[-1]
    net_debt = statements.loc[('Balance Sheet', 'Long-Term Debt'), latest_historical_year] - statements.loc[('Balance Sheet', 'Cash & Cash Equivalents'), latest_historical_year]
    equity_value = enterprise_value - net_debt
    implied_share_price = equity_value / assumptions['shares_outstanding'] if assumptions['shares_outstanding'] > 0 else 0

    metrics = {
        'Enterprise Value': enterprise_value, 'Net Debt': net_debt,
        'Equity Value': equity_value, 'Implied Share Price': implied_share_price,
        'Terminal Value': terminal_value, 'PV of Terminal Value': pv_terminal_value,
        'Sum of PV of UFCF': dcf.loc['PV of UFCF'].sum()
    }
    return dcf, metrics

# --- UI & App Layout ---

# Get Company Name from User
company_name = st.text_input("Enter Company Name", "My Company")

st.title(f"DCF Model: {company_name}")
st.markdown("*An interactive tool for company valuation based on your own historical data and assumptions.*")

# --- Sidebar for Inputs ---
st.sidebar.header("Control Panel")

# --- Historical Data Input (UPDATED WITH UPLOAD OPTION) ---
with st.sidebar.expander("📈 Historical Data Input", expanded=True):
    # --- Year Selection Logic ---
    last_full_year = pd.to_datetime('today').year - 1
    num_years_options = {2: "2 years", 3: "3 years", 4: "4 years", 5: "5 years"}
    num_years = st.selectbox(
        "Select number of historical years",
        options=list(num_years_options.keys()),
        format_func=lambda x: num_years_options[x],
        index=1  # Defaults to 3 years
    )
    latest_possible_start_year = last_full_year - num_years + 1
    year_range = range(2010, latest_possible_start_year + 1)
    start_year = st.selectbox(
        "Select the first historical year",
        options=year_range,
        index=len(year_range) - 1
    )
    historical_years = [str(start_year + i) for i in range(num_years)]
    st.info(f"Please enter data for: {', '.join(historical_years)}")

    # --- File Uploader ---
    st.markdown("---")
    st.markdown("**Option 1: Upload Data**")

    # Generate template for download
    excel_template = get_template_excel(historical_years)
    st.download_button(
        label="📥 Download Template",
        data=excel_template,
        file_name="dcf_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheet.sheet"
    )

    uploaded_file = st.file_uploader(
        "Upload your completed Excel template",
        type=['xlsx'],
        help="Upload the template with 'Income Statement' and 'Balance Sheet' tabs."
    )
    
    # --- Data Initialization Logic ---
    income_statement_items = ['Revenue', 'COGS', 'SG&A', 'D&A', 'Interest Expense']
    balance_sheet_items = [
        'Cash', 'Accounts Receivable', 'Inventory', 'PP&E', 'Accounts Payable',
        'Accrued Liabilities', 'Long-Term Debt', 'Common Stock', 'Retained Earnings'
    ]

    # Initialize session state if it doesn't exist or if years change
    if 'historical_is' not in st.session_state or list(st.session_state.historical_is.columns) != historical_years:
        st.session_state.historical_is = pd.DataFrame(0.0, index=income_statement_items, columns=historical_years)
        st.session_state.upload_success = False # Reset upload flag
    if 'historical_bs' not in st.session_state or list(st.session_state.historical_bs.columns) != historical_years:
        st.session_state.historical_bs = pd.DataFrame(0.0, index=balance_sheet_items, columns=historical_years)
        st.session_state.upload_success = False # Reset upload flag

    # --- Logic to process uploaded file ---
    if uploaded_file is not None:
        try:
            # Read the sheets from the uploaded file
            df_is_uploaded = pd.read_excel(uploaded_file, sheet_name="Income Statement", index_col=0)
            df_bs_uploaded = pd.read_excel(uploaded_file, sheet_name="Balance Sheet", index_col=0)

            # --- Validation ---
            # 1. Check if columns match selected historical years (as strings)
            uploaded_is_cols = [str(col) for col in df_is_uploaded.columns]
            uploaded_bs_cols = [str(col) for col in df_bs_uploaded.columns]

            if uploaded_is_cols != historical_years or uploaded_bs_cols != historical_years:
                st.error(f"Upload Error: The columns in your file ({uploaded_is_cols}) do not match the selected years ({historical_years}). Please re-download the template or adjust your year selection.")
            
            # 2. Check if all required rows are present
            elif not all(item in df_is_uploaded.index for item in income_statement_items) or \
                 not all(item in df_bs_uploaded.index for item in balance_sheet_items):
                st.error("Upload Error: Your file is missing one or more required metric rows. Please re-download the template.")
            
            # --- Success ---
            else:
                # Update session state with the *exact* required rows and columns
                st.session_state.historical_is = df_is_uploaded.loc[income_statement_items, historical_years].astype(float)
                st.session_state.historical_bs = df_bs_uploaded.loc[balance_sheet_items, historical_years].astype(float)
                if not st.session_state.get('upload_success', False):
                     st.success("File uploaded successfully! Data editors updated below.")
                     st.session_state.upload_success = True # Prevents success message from re-appearing

        except Exception as e:
            st.error(f"Error reading file: {e}. Make sure the file contains 'Income Statement' and 'Balance Sheet' tabs with the first column as metrics.")
            st.stop()
            
    # --- Data Editor Tabs ---
    st.markdown("---")
    st.markdown("**Option 2: Edit Data Manually**")
    tab1, tab2 = st.tabs(["Income Statement", "Balance Sheet"])
    with tab1:
        st.subheader("Income Statement History")
        st.session_state.historical_is = st.data_editor(st.session_state.historical_is, key="is_editor")
    with tab2:
        st.subheader("Balance Sheet History")
        st.session_state.historical_bs = st.data_editor(st.session_state.historical_bs, key="bs_editor")

    # --- Data Loading ---
    try:
        combined_df = pd.concat([st.session_state.historical_is, st.session_state.historical_bs])
        if combined_df.isnull().values.any() or (combined_df == 0).all().all():
            st.sidebar.warning("Please fill in or upload the historical data.")
            st.stop() # Stop execution if data is empty
        historical_data = combined_df.to_dict()
        st.sidebar.success("Historical data loaded.")
    except Exception as e:
        st.sidebar.error(f"Error processing data: {e}")
        st.stop()

# --- Calculated Historical Assumptions ---
first_hist_year = historical_years[0]
last_hist_year = historical_years[-1]

try:
    # Use dynamic years for calculations
    cagr = (historical_data[last_hist_year]['Revenue'] / historical_data[first_hist_year]['Revenue']) ** (1 / (num_years - 1)) - 1 if (num_years > 1 and historical_data[first_hist_year]['Revenue'] > 0) else 0
    cogs_percent_avg = np.mean([historical_data[y]['COGS'] / historical_data[y]['Revenue'] for y in historical_years if historical_data[y]['Revenue'] != 0])
    sga_percent_avg = np.mean([historical_data[y]['SG&A'] / historical_data[y]['Revenue'] for y in historical_years if historical_data[y]['Revenue'] != 0])
except (ZeroDivisionError, KeyError):
    cagr, cogs_percent_avg, sga_percent_avg = 0.0, 0.0, 0.0

# --- Operational Assumptions ---
with st.sidebar.expander("⚙️ Operational Assumptions"):
    st.markdown("**Growth & Profitability**")
    rev_growth_override = st.slider("Revenue Growth Rate (%)", 0.0, 25.0, cagr * 100, 0.5, help="Override the historically calculated CAGR.") / 100
    cogs_percent_override = st.slider("COGS (% of Revenue)", 0.0, 100.0, cogs_percent_avg * 100, 1.0, help="Override the historical average.") / 100
    sga_percent_override = st.slider("SG&A (% of Revenue)", 0.0, 100.0, sga_percent_avg * 100, 1.0, help="Override the historical average.") / 100

    st.markdown("**Balance Sheet & Cash Flow**")
    ar_days = st.slider("A/R Days", 0, 90, 30)
    inventory_days = st.slider("Inventory Days", 0, 90, 45)
    ap_days = st.slider("A/P Days", 0, 90, 25)
    accrued_liabilities_percent_revenue = st.slider("Accrued Liabilities (% of Revenue)", 0.0, 10.0, 3.0, 0.1) / 100
    depreciation_rate = st.slider("Depreciation (% of prior PP&E)", 0.0, 25.0, 10.0, 0.5) / 100
    capex_percent_revenue = st.slider("Capex (% of Revenue)", 0.0, 25.0, 12.0, 0.5) / 100
    debt_repayment_percent = st.slider("Annual Debt Repayment (%)", 0.0, 10.0, 2.0, 0.5) / 100

    st.markdown("**General**")
    tax_rate = st.slider("Corporate Tax Rate (%)", 0.0, 50.0, 21.0, 1.0) / 100
    shares_outstanding = st.number_input("Shares Outstanding (in millions)", value=500.0, min_value=0.1, help="Value must be greater than 0.")

# --- WACC Inputs ---
with st.sidebar.expander("⚖️ WACC Inputs"):
    risk_free_rate = st.slider("Risk-Free Rate (%)", 0.0, 10.0, 4.5, 0.1, help="Typically the yield on a 10-year government bond.") / 100
    market_risk_premium = st.slider("Market Risk Premium (%)", 0.0, 15.0, 5.5, 0.1, help="The excess return that investing in the stock market provides over the risk-free rate.") / 100
    company_beta = st.slider("Company Beta", 0.0, 3.0, 1.2, 0.1, help="A measure of the stock's volatility in relation to the overall market.")
    market_cap = st.number_input("Market Cap (in millions)", value=10000.0, min_value=0.0, help="Market Value of Equity.")

    cost_of_equity = risk_free_rate + company_beta * market_risk_premium
    try:
        # Use latest historical year for cost of debt
        cost_of_debt = historical_data[last_hist_year]['Interest Expense'] / historical_data[last_hist_year]['Long-Term Debt']
    except (ZeroDivisionError, KeyError, TypeError):
        cost_of_debt = 0.05 # Fallback value
    
    market_value_of_debt = historical_data[last_hist_year]['Long-Term Debt']
    total_capital = market_cap + market_value_of_debt

    try:
        wacc = ((market_cap / total_capital) * cost_of_equity) + \
               ((market_value_of_debt / total_capital) * cost_of_debt * (1 - tax_rate))
    except ZeroDivisionError:
        wacc = cost_of_equity # Fallback if total_capital is zero

    st.markdown("---")
    st.markdown(f"**Calculated Cost of Equity:** `{format_value(cost_of_equity, 'percentage')}`")
    st.markdown(f"**Calculated Cost of Debt:** `{format_value(cost_of_debt, 'percentage')}`")
    st.markdown(f"**Calculated WACC:** `{format_value(wacc, 'percentage')}`")

# --- Terminal Value Inputs ---
with st.sidebar.expander("♾️ Terminal Value Inputs"):
    perpetual_growth_rate = st.slider("Perpetual Growth Rate (%)", 0.0, 5.0, 2.5, 0.1, help="The long-term growth rate of cash flows beyond the projection period.") / 100
    if perpetual_growth_rate >= wacc:
        st.sidebar.error("Perpetual growth rate must be less than WACC.")
        st.stop()

# --- Final Assumptions Dictionary ---
assumptions = {
    'revenue_growth_rate': rev_growth_override, 'cogs_percent_revenue': cogs_percent_override,
    'sga_percent_revenue': sga_percent_override, 'ar_days': ar_days, 'inventory_days': inventory_days,
    'ap_days': ap_days, 'accrued_liabilities_percent_revenue': accrued_liabilities_percent_revenue,
    'depreciation_rate': depreciation_rate, 'capex_percent_revenue': capex_percent_revenue,
    'debt_repayment_percent': debt_repayment_percent, 'tax_rate': tax_rate,
    'shares_outstanding': shares_outstanding, 'wacc': wacc, 'cost_of_equity': cost_of_equity,
    'cost_of_debt': cost_of_debt, 'perpetual_growth_rate': perpetual_growth_rate
}

# --- Model Execution ---
projection_years_list = [str(int(last_hist_year) + i) for i in range(1, 6)]
statements = build_financial_statements(assumptions, historical_data, projection_years_list)
dcf_model, metrics = build_dcf_model(statements, assumptions, projection_years_list, historical_data)

# --- Main Dashboard Display ---
st.header("Valuation Summary")
col1, col2, col3 = st.columns(3)
col1.metric("Implied Share Price", f"${metrics['Implied Share Price']:.2f}")
col2.metric("Enterprise Value", format_value(metrics['Enterprise Value'], "currency"))
col3.metric("WACC", format_value(assumptions['wacc'], "percentage"))

# --- Tabs for Detailed Analysis ---
tab1, tab2, tab3 = st.tabs(["📊 DCF Analysis", "🧾 Financial Statements", "⚖️ WACC Breakdown"])

with tab1:
    st.subheader("Enterprise Value Bridge")
    value_bridge_data = pd.DataFrame({
        'Component': ['PV of Forecasted UFCF', 'PV of Terminal Value'],
        'Value': [metrics['Sum of PV of UFCF'], metrics['PV of Terminal Value']]
    }).set_index('Component')
    st.bar_chart(value_bridge_data)
    st.markdown(f"The analysis implies an Enterprise Value of **{format_value(metrics['Enterprise Value'], 'currency')}**.")

    st.subheader("Discounted Cash Flow (DCF) Calculation")
    st.dataframe(dcf_model.style.format("{:,.2f}"))

with tab2:
    st.subheader("Projected Income Statement")
    st.dataframe(statements.loc['Income Statement'].style.format("{:,.2f}"))

    st.subheader("ProjectV Balance Sheet")
    st.dataframe(statements.loc['Balance Sheet'].style.format("{:,.2f}"))
    # Highlight the balance check
    st.write("Balance Check (Assets - L&E):")
    st.dataframe(statements.loc[('Balance Sheet', 'Balance Check')].to_frame().T.style.format("{:,.2f}"))


    st.subheader("Projected Cash Flow Statement")
    st.dataframe(statements.loc['Cash Flow Statement'].style.format("{:,.2f}"))

with tab3:
    st.subheader("WACC Calculation Breakdown")
    st.markdown("The Weighted Average Cost of Capital (WACC) is the average rate of return a company is expected to provide to all its different investors.")

    wacc_data = {
        'Component': ['Risk-Free Rate', 'Market Risk Premium', 'Company Beta', '**Cost of Equity (Ke)**',
                      f'Interest Expense ({last_hist_year})', f'L/T Debt ({last_hist_year})', '**Cost of Debt (Kd)**',
                      'Market Cap (E)', 'Market Value of Debt (D)', 'Corporate Tax Rate', '**WACC**'],
        'Value': [format_value(risk_free_rate, 'percentage'), format_value(market_risk_premium, 'percentage'),
                  f"{company_beta:.2f}", format_value(cost_of_equity, 'percentage'),
                  format_value(historical_data[last_hist_year]['Interest Expense'], 'currency'),
                  format_value(historical_data[last_hist_year]['Long-Term Debt'], 'currency'),
                  format_value(cost_of_debt, 'percentage'),
                  format_value(market_cap, 'currency'), format_value(market_value_of_debt, 'currency'),
                  format_value(tax_rate, 'percentage'), format_value(wacc, 'percentage')]
    }
    wacc_df = pd.DataFrame(wacc_data).set_index('Component')
    st.table(wacc_df)

    with st.expander("See Formulas"):
        st.markdown("**Cost of Equity ($K_e$)**")
        st.latex(r'''K_e = R_f + \beta \times (R_m - R_f)''')
        st.markdown(r'''Where: $R_f$ = Risk-Free Rate, $\beta$ = Company Beta, $(R_m - R_f)$ = Market Risk Premium''')

        st.markdown("**Weighted Average Cost of Capital (WACC)**")
        st.latex(r'''WACC = \left(\frac{E}{E+D}\right)K_e + \left(\frac{D}{E+D}\right)K_d(1-t)''')
        st.markdown(r'''Where: $E$ = Market Value of Equity, $D$ = Market Value of Debt, $K_e$ = Cost of Equity, $K_d$ = Cost of Debt, $t$ = Corporate Tax Rate''')


# --- Download Button ---
safe_company_name = "".join([c for c in company_name if c.isalpha() or c.isdigit() or c==' ']).rstrip().replace(" ", "_")
excel_file = to_excel({
    "DCF Analysis": dcf_model,
    "Financial Statements": statements,
    "WACC Breakdown": wacc_df
})
st.sidebar.download_button(
    label="📥 Download Full Model to Excel",
    data=excel_file,
    file_name=f"dcf_model_{safe_company_name}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

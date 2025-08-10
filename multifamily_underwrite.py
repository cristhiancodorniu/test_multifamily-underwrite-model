import streamlit as st
import pandas as pd
import numpy as np
import numpy_financial as npf
from datetime import datetime
from dateutil.relativedelta import relativedelta
from scipy.stats import norm
import io

# --- Streamlit App Configuration ---
st.set_page_config(layout="wide")
st.title("Multifamily Development Financial Underwriting Model")

# --- Helper Functions ---
def generate_s_curve_distribution(periods):
    """Generates a monthly spend percentage based on an S-curve distribution."""
    if periods <= 1: return np.array([1.0])
    x = np.arange(periods + 1)
    mean, std_dev = periods / 2, periods / 4
    cumulative_pct = norm.cdf(x, loc=mean, scale=std_dev)
    normalized_cumulative_pct = (cumulative_pct - cumulative_pct.min()) / (cumulative_pct.max() - cumulative_pct.min())
    monthly_pct = np.diff(normalized_cumulative_pct)
    return monthly_pct

def to_excel(df_dict):
    """Takes a dictionary of DataFrames and returns an Excel file in memory."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, sheet_name=sheet_name)
    return output.getvalue()

# --- Input Sections ---
st.sidebar.header("General & Timing")
project_start_date = st.sidebar.date_input("Project Start Date", value=datetime(2025, 8, 1))
num_units = st.sidebar.number_input("Number of Units", min_value=1, value=100)
construction_period = st.sidebar.number_input("Construction Period (months)", min_value=1, value=24)
exit_cap_rate = st.sidebar.number_input("Exit Capitalization Rate (%)", min_value=0.1, value=5.0, step=0.1) / 100

st.sidebar.header("Construction Financing")
loan_to_cost_ratio = st.sidebar.number_input("Loan to Cost Ratio (%)", min_value=0.0, max_value=100.0, value=65.0) / 100
index_rate = st.sidebar.number_input("Index Rate (e.g., SOFR) (%)", min_value=0.0, value=3.0, step=0.1) / 100
spread_margin = st.sidebar.number_input("Spread / Margin (bps)", min_value=0, value=250) / 10000

st.sidebar.header("Closing Costs")
origination_fee_pct = st.sidebar.number_input("Origination Fee (% of Loan)", min_value=0.0, value=1.0, step=0.1) / 100
debt_procurement_fee_pct = st.sidebar.number_input("Debt Procurement Fee (% of Loan)", min_value=0.0, value=0.0) / 100
equity_procurement_fee_pct = st.sidebar.number_input("Equity Procurement Fee (% of Equity)", min_value=0.0, value=0.0) / 100
legal_fees = st.sidebar.number_input("Legal & Professional Fees ($)", min_value=0.0, value=150000.0)
ir_cap_costs = st.sidebar.number_input("Interest Rate Cap Cost ($)", min_value=0.0, value=50000.0)
other_closing_costs = st.sidebar.number_input("Other Closing Costs ($)", min_value=0.0, value=25000.0)

st.sidebar.header("Project Costs")
land_cost_per_unit = st.sidebar.number_input("Land Cost per Unit ($)", min_value=0.0, value=50000.0)
hard_cost_per_unit = st.sidebar.number_input("Hard Cost per Unit ($)", min_value=0.0, value=150000.0)
escalation_contingency = st.sidebar.number_input("Escalation Contingency (% of Hard Costs)", min_value=0.0, value=5.0) / 100
soft_cost_percentage = st.sidebar.number_input("Soft Costs (% of Hard Costs)", min_value=0.0, value=30.0) / 100
hard_cost_contingency = st.sidebar.number_input("Hard Cost Contingency (% of Hard Costs)", min_value=0.0, value=5.0) / 100
soft_cost_contingency = st.sidebar.number_input("Soft Cost Contingency (% of Soft Costs)", min_value=0.0, value=5.0) / 100
dev_mgmt_fee = st.sidebar.number_input("Dev. Mgmt Fee (% of Project Costs less Land)", min_value=0.0, value=3.0) / 100
const_mgmt_fee = st.sidebar.number_input("Const. Mgmt Fee (% of Hard Costs)", min_value=0.0, value=2.0) / 100

st.sidebar.header("Unit Mix & Rents")
studio_rent = st.sidebar.number_input("Studio Rent ($/month)", min_value=0.0, value=1500.0)
studio_units = st.sidebar.number_input("Number of Studio Units", min_value=0, value=20)
one_bed_rent = st.sidebar.number_input("1-Bedroom Rent ($/month)", min_value=0.0, value=2000.0)
one_bed_units = st.sidebar.number_input("Number of 1-Bedroom Units", min_value=0, value=40)
two_bed_rent = st.sidebar.number_input("2-Bedroom Rent ($/month)", min_value=0.0, value=2500.0)
two_bed_units = st.sidebar.number_input("Number of 2-Bedroom Units", min_value=0, value=30)
three_bed_rent = st.sidebar.number_input("3-Bedroom Rent ($/month)", min_value=0.0, value=3000.0)
three_bed_units = st.sidebar.number_input("Number of 3-Bedroom Units", min_value=0, value=10)

st.sidebar.header("Revenue & Operations")
min_vacancy_rate = st.sidebar.number_input("Stabilized Vacancy Rate (%)", min_value=0.0, value=5.0) / 100
credit_loss = st.sidebar.number_input("Credit Loss (% of GPR)", min_value=0.0, value=1.0) / 100
other_income_pct = st.sidebar.number_input("Other Income (% of EGI)", min_value=0.0, value=5.0) / 100
revenue_escalation = st.sidebar.number_input("Revenue Escalation (% p.a.)", min_value=0.0, value=3.0) / 100
opex_per_unit_year = st.sidebar.number_input("Opex per Unit per Year ($)", min_value=0.0, value=5000.0)
re_taxes_per_unit_year = st.sidebar.number_input("RE Taxes per Unit per Year ($)", min_value=0.0, value=2000.0)
insurance_per_unit_year = st.sidebar.number_input("Insurance per Unit per Year ($)", min_value=0.0, value=1000.0)
mgmt_fee_pct = st.sidebar.number_input("Mgmt Fee (% of EGI)", min_value=0.0, value=3.0) / 100
opex_escalation = st.sidebar.number_input("Opex Escalation (% p.a.)", min_value=0.0, value=3.0) / 100

st.sidebar.header("Lease-Up Schedule")
leased_units_per_month = st.sidebar.number_input("Leased Units per Month", min_value=1, value=10)

# --- Main Calculation Block ---
main_container = st.container()

total_input_units = studio_units + one_bed_units + two_bed_units + three_bed_units
if total_input_units != num_units:
    st.error(f"Total units in unit mix ({total_input_units}) must equal Number of Units ({num_units}). Please adjust values in the sidebar.")
    st.stop()

# Base costs that don't depend on financing
land_costs = land_cost_per_unit * num_units
base_hard_costs = hard_cost_per_unit * num_units
total_hard_costs = base_hard_costs * (1 + escalation_contingency + hard_cost_contingency)
total_soft_costs = total_hard_costs * soft_cost_percentage * (1 + soft_cost_contingency)
total_project_costs_less_land_fees = total_hard_costs + total_soft_costs
dev_mgmt_fee_cost = total_project_costs_less_land_fees * dev_mgmt_fee
const_mgmt_fee_cost = total_hard_costs * const_mgmt_fee
total_fees = dev_mgmt_fee_cost + const_mgmt_fee_cost
base_costs = land_costs + total_hard_costs + total_soft_costs + total_fees

# Hold Period & Lease-Up
stabilized_occupancy_rate = 1 - min_vacancy_rate
stabilized_occupied_units = int(num_units * stabilized_occupancy_rate)
months_to_stabilize = construction_period + int(np.ceil(stabilized_occupied_units / leased_units_per_month)) if leased_units_per_month > 0 else construction_period
hold_period = months_to_stabilize + 12
dates = [project_start_date + relativedelta(months=i) for i in range(hold_period)]

# ADDITION: Use st.spinner for a better user experience during long calculations
with st.spinner('Running iterative financing calculations... Please wait.'):
    # Initialize loop variables
    total_capitalized_interest = 0.0
    total_operating_reserve = 0.0
    total_closing_costs = 0.0
    interest_rate = index_rate + spread_margin
    last_total_costs = 0.0 # For convergence check

    for i in range(30): # Iterate up to 30 times
        total_costs = base_costs + total_closing_costs + total_capitalized_interest + total_operating_reserve
        
        # ADDITION: Check for convergence to exit loop early
        if abs(total_costs - last_total_costs) < 1.00:
            break
        last_total_costs = total_costs

        total_debt = total_costs * loan_to_cost_ratio
        total_equity = total_costs - total_debt
        
        closing_cost_origination = total_debt * origination_fee_pct
        closing_cost_debt_proc = total_debt * debt_procurement_fee_pct
        closing_cost_equity_proc = total_equity * equity_procurement_fee_pct
        total_closing_costs = closing_cost_origination + closing_cost_debt_proc + closing_cost_equity_proc + legal_fees + ir_cap_costs + other_closing_costs
        
        cf = pd.DataFrame(index=pd.to_datetime(dates), dtype=np.float64)
        cf['Month'] = range(1, hold_period + 1)
        
        construction_costs_monthly = np.zeros(hold_period)
        if construction_period > 0:
            hard_cost_pct = generate_s_curve_distribution(construction_period)
            construction_costs_monthly[:construction_period] += total_hard_costs * hard_cost_pct
            construction_costs_monthly[:construction_period] += (total_soft_costs + total_fees) / construction_period
        
        cf['Base Project Spend'] = construction_costs_monthly
        cf.iloc[0, cf.columns.get_loc('Base Project Spend')] += (land_costs + total_closing_costs + total_operating_reserve)
        
        cf['Occupied Units'] = [min(stabilized_occupied_units, leased_units_per_month * (m - construction_period)) if m > construction_period else 0 for m in cf['Month']]
        cf['Occupancy Rate'] = cf['Occupied Units'] / num_units if num_units > 0 else 0

        weighted_rent = (studio_rent * studio_units + one_bed_rent * one_bed_units + two_bed_rent * two_bed_units + three_bed_rent * three_bed_units) / num_units if num_units > 0 else 0
        opex_escalation_factor = (1 + opex_escalation) ** (cf['Month'] / 12)
        revenue_escalation_factor = (1 + revenue_escalation) ** (cf['Month'] / 12)
        cf['GPR'] = cf['Occupied Units'] * weighted_rent * revenue_escalation_factor
        cf['Vacancy Loss'] = cf['GPR'] * min_vacancy_rate
        cf['Credit Loss'] = cf['GPR'] * credit_loss
        cf['Effective Billed Revenue'] = cf['GPR'] - cf['Vacancy Loss'] - cf['Credit Loss']
        cf['EGI'] = cf['Effective Billed Revenue'] / (1 - other_income_pct) if other_income_pct < 1 else cf['Effective Billed Revenue']
        cf['Other Income'] = cf['EGI'] * other_income_pct
        cf['Opex'] = cf['Occupied Units'] * (opex_per_unit_year / 12) * opex_escalation_factor
        cf['RE Taxes'] = cf['Occupied Units'] * (re_taxes_per_unit_year / 12) * opex_escalation_factor
        cf['Insurance'] = cf['Occupied Units'] * (insurance_per_unit_year / 12) * opex_escalation_factor
        cf['Management Fee'] = cf['EGI'] * mgmt_fee_pct
        cf['Total Opex'] = cf['Opex'] + cf['RE Taxes'] + cf['Insurance'] + cf['Management Fee']
        cf['NOI'] = cf['EGI'] - cf['Total Opex']
        
        cf['Net Funding Requirement'] = cf['Base Project Spend'] - cf['NOI']
        
        outstanding_balance = np.zeros(hold_period + 1, dtype=np.float64)
        monthly_interest = np.zeros(hold_period, dtype=np.float64)
        capitalized_interest_monthly = np.zeros(hold_period, dtype=np.float64)
        
        cf['Equity Draw'] = 0.0
        cf['Debt Draw'] = 0.0

        for m in range(hold_period):
            monthly_interest[m] = outstanding_balance[m] * interest_rate / 12
            funding_req_this_month = cf['Net Funding Requirement'].iloc[m]

            if cf['Month'].iloc[m] <= construction_period:
                funding_req_this_month += monthly_interest[m]
                capitalized_interest_monthly[m] = monthly_interest[m]
            
            equity_drawn_so_far = cf['Equity Draw'].sum()
            equity_to_draw = max(0, min(funding_req_this_month, total_equity - equity_drawn_so_far))
            
            cf.loc[cf.index[m], 'Equity Draw'] = equity_to_draw
            cf.loc[cf.index[m], 'Debt Draw'] = funding_req_this_month - equity_to_draw
            
            outstanding_balance[m+1] = outstanding_balance[m] + cf.loc[cf.index[m], 'Debt Draw']

        total_capitalized_interest = capitalized_interest_monthly.sum()
        
        post_construction_interest = monthly_interest * (cf['Month'] > construction_period)
        post_construction_noi = cf['NOI'] * (cf['Month'] > construction_period)
        operating_shortfall = np.maximum(0, post_construction_interest - post_construction_noi)
        total_operating_reserve = operating_shortfall.sum()

# --- Post-Loop Final Calculations and Data Prep ---
cf['Monthly Interest'] = monthly_interest
cf['Capitalized Interest'] = capitalized_interest_monthly
cf['Interest Paid from Operations'] = np.minimum(cf['NOI'], post_construction_interest)
cf['Investor Cash Flow'] = -cf['Equity Draw'] + cf['NOI'] - cf['Interest Paid from Operations']

outstanding_debt = cf['Debt Draw'].sum()
final_noi_trended = cf['NOI'].iloc[-1] * 12
exit_value = final_noi_trended / exit_cap_rate if exit_cap_rate > 0 else 0
sale_costs = exit_value * 0.02
net_sale_proceeds_levered = exit_value - sale_costs - outstanding_debt
cf.iloc[-1, cf.columns.get_loc('Investor Cash Flow')] += net_sale_proceeds_levered
levered_irr = npf.irr(cf['Investor Cash Flow'].values)
levered_annual_irr = (1 + levered_irr) ** 12 - 1 if levered_irr is not None and not np.isnan(levered_irr) else 0
equity_multiple = cf['Investor Cash Flow'].sum() / total_equity if total_equity > 0 else 0

# Unlevered Metrics
unlevered_cash_flow = cf['NOI'] - cf['Base Project Spend']
net_sale_proceeds_unlevered = exit_value - sale_costs
unlevered_cash_flow.iloc[-1] += net_sale_proceeds_unlevered
unlevered_irr = npf.irr(unlevered_cash_flow.values)
unlevered_annual_irr = (1 + unlevered_irr) ** 12 - 1 if unlevered_irr is not None and not np.isnan(unlevered_irr) else 0

# Untrended Yield Calculations
stabilized_gpr_untr = weighted_rent * num_units * stabilized_occupancy_rate * 12
stabilized_ebr_untr = stabilized_gpr_untr * (1 - min_vacancy_rate - credit_loss)
stabilized_egi_untr = stabilized_ebr_untr / (1 - other_income_pct) if other_income_pct < 1 else stabilized_ebr_untr
stabilized_base_opex_untr = (opex_per_unit_year + re_taxes_per_unit_year + insurance_per_unit_year) * stabilized_occupied_units
stabilized_mgmt_fee_untr = stabilized_egi_untr * mgmt_fee_pct
stabilized_total_opex_untr = stabilized_base_opex_untr + stabilized_mgmt_fee_untr
stabilized_noi_untr = stabilized_egi_untr - stabilized_total_opex_untr
untrended_yoc = stabilized_noi_untr / total_costs if total_costs > 0 else 0
untrended_debt_yield = stabilized_noi_untr / total_debt if total_debt > 0 else 0

# Trended Yield Calculations
start_stabilization_index = months_to_stabilize
end_stabilization_index = start_stabilization_index + 12
if end_stabilization_index <= len(cf):
    trended_noi_stabilized = cf['NOI'].iloc[start_stabilization_index:end_stabilization_index].sum()
    trended_yoc = trended_noi_stabilized / total_costs if total_costs > 0 else 0
    trended_debt_yield = trended_noi_stabilized / total_debt if total_debt > 0 else 0
else:
    trended_yoc = 0
    trended_debt_yield = 0

# Prepare DataFrames for Display and Export
sources_df = pd.DataFrame({"Sources": ["Construction Loan", "Total Common Equity"], "Amount": [total_debt, total_equity]})
uses_df = pd.DataFrame({
    "Uses": ["Land Costs", "Hard Costs", "Soft Costs", "Fees", "Closing Costs", "Capitalized Interest", "Operating Reserve"],
    "Amount": [land_costs, total_hard_costs, total_soft_costs, total_fees, total_closing_costs, total_capitalized_interest, total_operating_reserve]
})
annual_cols = ['GPR', 'Total Opex', 'NOI', 'Capitalized Interest', 'Interest Paid from Operations']
annual_income = cf.groupby(cf.index.year)[annual_cols].sum()
bs = pd.DataFrame(index=cf.index)
bs['Total Assets'] = (cf['Base Project Spend'].cumsum() - cf['NOI'].cumsum()).cumsum()
bs['Debt'] = cf['Debt Draw'].cumsum()
bs['Equity'] = cf['Equity Draw'].cumsum()
bs['Total Liabilities & Equity'] = bs['Debt'] + bs['Equity']
excel_export_dict = {'Monthly Pro-Forma': cf, 'Annual Summary': annual_income, 'Sources and Uses': uses_df, 'Balance Sheet': bs}
excel_data = to_excel(excel_export_dict)

# --- Display Results ---
with main_container:
    st.header("Project Overview")
    col1, col2, col3 = st.columns(3)
    col1.metric("Start Date", project_start_date.strftime("%b %Y"))
    col1.metric("Number of Units", f"{num_units}")
    col2.metric("Total Hard Cost / Unit", f"${total_hard_costs/num_units:,.0f}" if num_units > 0 else "$0")
    col2.metric("Total Project Cost / Unit", f"${total_costs/num_units:,.0f}" if num_units > 0 else "$0")
    col3.metric("Construction Period", f"{construction_period} Months")
    col3.metric("Projected Hold Period", f"{hold_period} Months")
    st.markdown("---")
    
    st.header("Financial Outputs")
    col1, col2, col3 = st.columns(3)
    col1.metric("Unlevered Project IRR", f"{unlevered_annual_irr*100:.2f}%")
    col2.metric("Levered Project IRR", f"{levered_annual_irr*100:.2f}%")
    col3.metric("Equity Multiple", f"{equity_multiple:.2f}x")

    col1, col2, col3 = st.columns(3)
    col1.metric("Untrended Yield on Cost", f"{untrended_yoc*100:.2f}%")
    col2.metric("Trended Yield on Cost", f"{trended_yoc*100:.2f}%")
    col3.write("")

    col1, col2, col3 = st.columns(3)
    col1.metric("Untrended Debt Yield", f"{untrended_debt_yield*100:.2f}%")
    col2.metric("Trended Debt Yield", f"{trended_debt_yield*100:.2f}%")
    with col3:
        st.download_button(label="📥 Download Model as Excel", data=excel_data, file_name="multifamily_financial_model.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Sources")
        st.dataframe(sources_df.style.format({"Amount": "${:,.0f}"}), use_container_width=True, hide_index=True)
        st.write(f"**Total Sources: ${total_costs:,.2f}**")
    with col2:
        st.subheader("Uses")
        st.dataframe(uses_df.style.format({"Amount": "${:,.0f}"}), use_container_width=True, hide_index=True)
        st.write(f"**Total Uses: ${uses_df['Amount'].sum():,.2f}**")
        
    st.markdown("---")
    
    st.subheader("Lease-Up Schedule")
    lease_up_df = cf[['Month', 'Occupied Units', 'Occupancy Rate']].copy()
    lease_up_df = lease_up_df[lease_up_df['Month'] > construction_period].reset_index(drop=True)
    st.dataframe(lease_up_df.style.format({'Occupancy Rate': "{:.2%}", "Occupied Units": "{:,.0f}"}), hide_index=True, use_container_width=True)
    st.markdown("---")
    
    st.subheader("Monthly Pro-Forma Cash Flow (Detailed)")
    display_cols = ['Month', 'NOI', 'Base Project Spend', 'Monthly Interest', 'Capitalized Interest', 'Interest Paid from Operations', 'Equity Draw', 'Debt Draw', 'Investor Cash Flow']
    display_cf = cf[display_cols].copy()
    st.dataframe(display_cf.style.format(formatter="${:,.0f}", subset=[c for c in display_cols if c != 'Month']), use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("Annual Summary")
    st.dataframe(annual_income.style.format("${:,.0f}"), use_container_width=True)

    st.markdown("---")
    st.subheader("Balance Sheet (Simplified)")
    st.dataframe(bs.style.format("${:,.0f}"), use_container_width=True, hide_index=True)
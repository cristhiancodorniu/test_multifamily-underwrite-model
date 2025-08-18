import streamlit as st
import pandas as pd
import numpy as np
import numpy_financial as npf
import altair as alt
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

# --- Input Sections (Sidebar) ---
st.sidebar.header("General & Timing")
project_start_date = st.sidebar.date_input("Project Start Date", value=datetime(2025, 8, 1))
num_units = st.sidebar.number_input("Number of Units", min_value=1, value=100)
construction_period = st.sidebar.number_input("Construction Period (months)", min_value=1, value=24)
exit_cap_rate = st.sidebar.number_input("Exit Capitalization Rate (%)", min_value=0.1, value=5.0, step=0.1) / 100

st.sidebar.header("Building Metrics")
gsf = st.sidebar.number_input("Total Gross Square Footage (GSF)", min_value=1, value=125000)
st.sidebar.subheader("Rentable Square Feet (RSF) per Unit")
studio_sf = st.sidebar.number_input("Studio RSF", min_value=1, value=500)
one_bed_sf = st.sidebar.number_input("1-Bedroom RSF", min_value=1, value=750)
two_bed_sf = st.sidebar.number_input("2-Bedroom RSF", min_value=1, value=1000)
three_bed_sf = st.sidebar.number_input("3-Bedroom RSF", min_value=1, value=1250)

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
with st.sidebar.expander("Hard Cost Contingencies"):
    escalation_contingency = st.number_input("Escalation Contingency (% of Hard Costs)", min_value=0.0, value=5.0) / 100
    hard_cost_contingency = st.number_input("Hard Cost Contingency (% of Hard Costs)", min_value=0.0, value=5.0) / 100

st.sidebar.subheader("Detailed Soft Costs ($)")
sc_municipal = st.sidebar.number_input("Municipal", min_value=0.0, value=100000.0)
sc_arch_eng = st.sidebar.number_input("Architectural & Engineering", min_value=0.0, value=1200000.0)
sc_inspection = st.sidebar.number_input("Inspection", min_value=0.0, value=150000.0)
sc_survey = st.sidebar.number_input("Testing/Survey/Consultants", min_value=0.0, value=200000.0)
sc_insurance = st.sidebar.number_input("Project Insurance", min_value=0.0, value=300000.0)
sc_ga = st.sidebar.number_input("G&A", min_value=0.0, value=250000.0)
sc_ffe = st.sidebar.number_input("FF&E", min_value=0.0, value=50000.0)
sc_marketing = st.sidebar.number_input("Marketing & Start-Up", min_value=0.0, value=250000.0)
with st.sidebar.expander("Soft Cost Contingency"):
    soft_cost_contingency = st.number_input("Soft Cost Contingency (% of Soft Costs)", min_value=0.0, value=5.0) / 100

st.sidebar.subheader("Fees")
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

st.sidebar.header("Revenue")
with st.sidebar.expander("Annual Rent Escalation (%)"):
    rev_esc_rates = [st.number_input(f"Year {i+1} Rent Escalation", value=val, key=f"rev_esc_{i}") / 100 for i, val in enumerate([5.0, 4.0, 3.0, 3.0, 3.0])]
min_vacancy_rate = st.sidebar.number_input("Stabilized Vacancy Rate (%)", min_value=0.0, value=5.0) / 100
credit_loss = st.sidebar.number_input("Credit Loss (% of GPR)", min_value=0.0, value=1.0) / 100
other_income_pct = st.sidebar.number_input("Other Income (% of EGI)", min_value=0.0, value=5.0) / 100

st.sidebar.header("Operating Expenses")
with st.sidebar.expander("Lease-Up Concessions"):
    concession_pct_new_leases = st.number_input("% of New Leases w/ Concession", min_value=0.0, max_value=100.0, value=50.0) / 100
    concession_months_free = st.number_input("# Months Free Rent", min_value=0.0, value=1.0)
with st.sidebar.expander("Annual Opex Escalation (%)"):
    opex_esc_rates = [st.number_input(f"Year {i+1} Opex Escalation", value=val, key=f"opex_esc_{i}") / 100 for i, val in enumerate([3.0, 3.0, 2.5, 2.5, 2.5])]

st.sidebar.subheader("Controllable Expenses ($/unit/year)")
exp_admin = st.sidebar.number_input("Administration", value=500.0)
exp_marketing = st.sidebar.number_input("Leasing & Marketing", value=300.0)
exp_repairs = st.sidebar.number_input("Repairs & Maintenance", value=800.0)
exp_grounds = st.sidebar.number_input("Grounds & Landscaping", value=250.0)
exp_common_area = st.sidebar.number_input("Common Area Expense", value=400.0)
exp_turnover = st.sidebar.number_input("Turnover Costs", value=600.0)
exp_security = st.sidebar.number_input("Security", value=350.0)
exp_employee = st.sidebar.number_input("Employee Expenses", value=1500.0)
exp_other_controllable = st.sidebar.number_input("Other Controllable", value=100.0)

st.sidebar.subheader("Non-Controllable Expenses ($/unit/year)")
exp_util_gas = st.sidebar.number_input("Utility - Gas", value=300.0)
exp_util_common = st.sidebar.number_input("Utility - Common", value=400.0)
exp_util_electric = st.sidebar.number_input("Utility - Electric", value=500.0)
exp_internet = st.sidebar.number_input("Internet/Cable", value=200.0)
exp_insurance = st.sidebar.number_input("Operating Insurance", value=1000.0)
exp_other_noncontrollable = st.sidebar.number_input("Other Non-Controllable", value=50.0)
re_taxes_per_unit_year = st.sidebar.number_input("Real Estate Taxes", min_value=0.0, value=2000.0)
mgmt_fee_pct = st.sidebar.number_input("Management Fee (% of EGI)", min_value=0.0, value=3.0) / 100
replacement_reserve_per_unit = st.sidebar.number_input("Replacement Reserve per Unit ($)", value=250.0)

st.sidebar.header("Lease-Up Schedule")
leased_units_per_month = st.sidebar.number_input("Leased Units per Month", min_value=1, value=10)
pre_leased_pct = st.sidebar.number_input("Pre-Leased Units at Delivery (%)", min_value=0.0, max_value=100.0, value=10.0) / 100
stabilization_threshold = st.sidebar.number_input("Stabilization Threshold (%)", min_value=1.0, max_value=100.0, value=95.0) / 100

# --- Main Calculation Block ---
main_container = st.container()

total_input_units = studio_units + one_bed_units + two_bed_units + three_bed_units
if total_input_units != num_units:
    st.error(f"Total units in unit mix ({total_input_units}) must equal Number of Units ({num_units}). Please adjust values in the sidebar.")
    st.stop()

# --- Calculate Total Rentable Square Footage ---
total_rsf = (studio_sf * studio_units) + (one_bed_sf * one_bed_units) + (two_bed_sf * two_bed_units) + (three_bed_sf * three_bed_units)

# --- Pre-Calculation Setup ---
detailed_controllable_exp = { "Administration": exp_admin, "Leasing & Marketing": exp_marketing, "Repairs & Maintenance": exp_repairs, "Grounds & Landscaping": exp_grounds, "Common Area Expense": exp_common_area, "Turnover Costs": exp_turnover, "Security": exp_security, "Employee Expenses": exp_employee, "Other Controllable": exp_other_controllable }
detailed_non_controllable_exp = { "Utility - Gas": exp_util_gas, "Utility - Common": exp_util_common, "Utility - Electric": exp_util_electric, "Internet/Cable": exp_internet, "Operating Insurance": exp_insurance, "Other Non-Controllable": exp_other_noncontrollable }
total_controllable_exp_per_unit = sum(detailed_controllable_exp.values())
total_non_controllable_exp_per_unit = sum(detailed_non_controllable_exp.values())

soft_costs_upfront = sum([sc_municipal, sc_arch_eng, sc_inspection, sc_survey, sc_insurance, sc_ga])
soft_costs_spread = sc_ffe + sc_marketing
base_soft_costs = soft_costs_upfront + soft_costs_spread
total_soft_costs = base_soft_costs * (1 + soft_cost_contingency)

land_costs = land_cost_per_unit * num_units
base_hard_costs = hard_cost_per_unit * num_units
total_hard_costs = base_hard_costs * (1 + escalation_contingency + hard_cost_contingency)
total_project_costs_less_land_fees = total_hard_costs + total_soft_costs
dev_mgmt_fee_cost = total_project_costs_less_land_fees * dev_mgmt_fee
const_mgmt_fee_cost = total_hard_costs * const_mgmt_fee
total_fees = dev_mgmt_fee_cost + const_mgmt_fee_cost
base_costs = land_costs + total_hard_costs + total_soft_costs + total_fees

# Determine Timing
stabilized_occupied_units_target = int(num_units * stabilization_threshold)
pre_leased_units = int(num_units * pre_leased_pct)
units_to_lease_at_delivery = stabilized_occupied_units_target - pre_leased_units
months_to_stabilize = construction_period + (int(np.ceil(units_to_lease_at_delivery / leased_units_per_month)) if leased_units_per_month > 0 and units_to_lease_at_delivery > 0 else 0)
sale_month = months_to_stabilize + 1
hold_period = sale_month + 12
dates = [project_start_date + relativedelta(months=i) for i in range(hold_period)]

# Create annual step-up escalation factors
annual_rev_rates = [1.0] + [1 + r for r in rev_esc_rates]
cumulative_annual_rev_factors = np.cumprod(annual_rev_rates)
cumulative_rev_factor = np.array([cumulative_annual_rev_factors[min(i // 12, 5)] for i in range(hold_period)])

annual_opex_rates = [1.0] + [1 + o for o in opex_esc_rates]
cumulative_annual_opex_factors = np.cumprod(annual_opex_rates)
cumulative_opex_factor = np.array([cumulative_annual_opex_factors[min(i // 12, 5)] for i in range(hold_period)])

with st.spinner('Running iterative financing calculations... Please wait.'):
    total_capitalized_interest, total_operating_reserve, total_closing_costs = 0.0, 0.0, 0.0
    interest_rate = index_rate + spread_margin
    last_total_costs = 0.0

    for i in range(30):
        total_costs = base_costs + total_closing_costs + total_capitalized_interest + total_operating_reserve
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
        
        cf['Cumulative Revenue Factor'] = cumulative_rev_factor
        cf['Cumulative Opex Factor'] = cumulative_opex_factor
        
        construction_costs_monthly = np.zeros(hold_period)
        if construction_period > 0:
            hard_cost_pct = generate_s_curve_distribution(construction_period)
            construction_costs_monthly[:construction_period] += total_hard_costs * hard_cost_pct
            construction_costs_monthly[:construction_period] += total_fees / construction_period
            
            construction_costs_monthly[0] += soft_costs_upfront * (1 + soft_cost_contingency)
            
            spread_duration = min(construction_period, 3)
            if spread_duration > 0:
                spread_cost_monthly = (soft_costs_spread * (1 + soft_cost_contingency)) / spread_duration
                for m in range(spread_duration):
                    construction_costs_monthly[construction_period - 1 - m] += spread_cost_monthly

        cf['Base Project Spend'] = construction_costs_monthly
        cf.iloc[0, cf.columns.get_loc('Base Project Spend')] += (land_costs + total_closing_costs + total_operating_reserve)
        
        cf['Occupied Units'] = [min(stabilized_occupied_units_target, pre_leased_units + (leased_units_per_month * (m - construction_period))) if m >= construction_period else 0 for m in cf['Month']]
        cf['Occupancy Rate'] = cf['Occupied Units'] / num_units if num_units > 0 else 0
        
        new_leases_monthly = cf['Occupied Units'].diff().fillna(0).clip(lower=0)
        if hold_period > construction_period:
            new_leases_monthly.iloc[construction_period] = pre_leased_units
        cf['New Leases'] = new_leases_monthly
        
        unit_mix_pct = {'Studio': studio_units / num_units, '1BR': one_bed_units / num_units, '2BR': two_bed_units / num_units, '3BR': three_bed_units / num_units} if num_units > 0 else {}
        cf['Rent Studio'] = studio_rent * cumulative_rev_factor
        cf['Rent 1BR'] = one_bed_rent * cumulative_rev_factor
        cf['Rent 2BR'] = two_bed_rent * cumulative_rev_factor
        cf['Rent 3BR'] = three_bed_rent * cumulative_rev_factor
        cf['Weighted Average Rent'] = (cf['Rent Studio'] * studio_units + cf['Rent 1BR'] * one_bed_units + cf['Rent 2BR'] * two_bed_units + cf['Rent 3BR'] * three_bed_units) / num_units if num_units > 0 else 0
        
        for unit_type in ['Studio', '1BR', '2BR', '3BR']:
            cf[f'Occupied {unit_type}'] = cf['Occupied Units'] * unit_mix_pct.get(unit_type, 0)
        
        cf['GPR Studio'] = cf[f'Occupied Studio'] * cf['Rent Studio']
        cf['GPR 1BR'] = cf[f'Occupied 1BR'] * cf['Rent 1BR']
        cf['GPR 2BR'] = cf[f'Occupied 2BR'] * cf['Rent 2BR']
        cf['GPR 3BR'] = cf[f'Occupied 3BR'] * cf['Rent 3BR']
        cf['GPR'] = cf['GPR Studio'] + cf['GPR 1BR'] + cf['GPR 2BR'] + cf['GPR 3BR']
        
        cf['Vacancy Loss'] = cf['GPR'] * min_vacancy_rate
        cf['Credit Loss'] = cf['GPR'] * credit_loss
        cf['Effective Billed Revenue'] = cf['GPR'] - cf['Vacancy Loss'] - cf['Credit Loss']
        cf['Other Income'] = (cf['Effective Billed Revenue'] / (1 - other_income_pct) - cf['Effective Billed Revenue']) if other_income_pct < 1 else 0
        cf['EGI'] = cf['Effective Billed Revenue'] + cf['Other Income']
        
        for name, amount in detailed_controllable_exp.items():
            cf[f'Opex - {name}'] = cf['Occupied Units'] * (amount / 12) * cumulative_opex_factor
        for name, amount in detailed_non_controllable_exp.items():
            cf[f'Opex - {name}'] = cf['Occupied Units'] * (amount / 12) * cumulative_opex_factor
        
        cf['RE Taxes'] = cf['Occupied Units'] * (re_taxes_per_unit_year / 12) * cumulative_opex_factor
        cf['Management Fee'] = cf['EGI'] * mgmt_fee_pct
        concession_period_mask = (cf['Month'] > construction_period) & (cf['Month'] <= construction_period + 12)
        cf['Concession Cost'] = cf['New Leases'] * concession_pct_new_leases * concession_months_free * cf['Weighted Average Rent'] * concession_period_mask

        cf['Total Opex'] = sum([cf[f'Opex - {name}'] for name in detailed_controllable_exp] + [cf[f'Opex - {name}'] for name in detailed_non_controllable_exp] + [cf['RE Taxes'], cf['Management Fee'], cf['Concession Cost']])
        cf['NOI'] = cf['EGI'] - cf['Total Opex']
        
        cf['Replacement Reserve'] = cf['Occupied Units'] * (replacement_reserve_per_unit / 12)
        cf['NOI After Reserves'] = cf['NOI'] - cf['Replacement Reserve']
        
        cf['Net Funding Requirement'] = cf['Base Project Spend'] - cf['NOI'] + cf['Replacement Reserve']
        
        outstanding_balance, monthly_interest, capitalized_interest_monthly = np.zeros(hold_period + 1), np.zeros(hold_period), np.zeros(hold_period)
        cf['Equity Draw'], cf['Debt Draw'] = 0.0, 0.0

        for m in range(hold_period):
            monthly_interest[m] = outstanding_balance[m] * interest_rate / 12
            funding_req = cf['Net Funding Requirement'].iloc[m]
            if cf['Month'].iloc[m] <= construction_period:
                funding_req += monthly_interest[m]
                capitalized_interest_monthly[m] = monthly_interest[m]
            
            equity_drawn = cf['Equity Draw'].sum()
            equity_to_draw = max(0, min(funding_req, total_equity - equity_drawn))
            
            cf.loc[cf.index[m], 'Equity Draw'] = equity_to_draw
            debt_draw_pre_clip = funding_req - equity_to_draw
            cf.loc[cf.index[m], 'Debt Draw'] = max(0, debt_draw_pre_clip)
            
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
cf['Operating Reserve Drawdown'] = operating_shortfall
post_construction_mask = (cf['Month'] > construction_period)
excess_cash_flow = -np.minimum(0, cf['Net Funding Requirement'] - cf['Equity Draw'] - cf['Debt Draw']) * post_construction_mask
cf['Excess Cash Flow'] = excess_cash_flow.values

cf['Investor Cash Flow'] = -cf['Equity Draw'] + cf['NOI'] - cf['Interest Paid from Operations'] - cf['Replacement Reserve'] + excess_cash_flow + cf['Operating Reserve Drawdown']


cf['NOI for Exit'] = cf['EGI'] - (cf['Total Opex'] - cf['Concession Cost'])
start_noi_slice = sale_month
end_noi_slice = sale_month + 12
forward_noi = cf['NOI for Exit'].iloc[start_noi_slice:end_noi_slice].sum()
exit_value = forward_noi / exit_cap_rate if exit_cap_rate > 0 else 0
sale_costs = exit_value * 0.02
outstanding_debt = cf['Debt Draw'].iloc[:sale_month].sum()
net_sale_proceeds_levered = exit_value - sale_costs - outstanding_debt
cf.loc[cf.index[sale_month - 1], 'Investor Cash Flow'] += net_sale_proceeds_levered
cf.loc[cf.index[sale_month]:, 'Investor Cash Flow'] = 0
levered_irr = npf.irr(cf['Investor Cash Flow'].values[:sale_month])
levered_annual_irr = (1 + levered_irr) ** 12 - 1 if levered_irr is not None and not np.isnan(levered_irr) else 0
total_distributions = cf['Investor Cash Flow'].sum() + total_equity
equity_multiple = total_distributions / total_equity if total_equity > 0 else 0

unlevered_cash_flow = -cf['Base Project Spend'] + cf['NOI']
net_sale_proceeds_unlevered = exit_value - sale_costs
unlevered_cash_flow.iloc[sale_month - 1] += net_sale_proceeds_unlevered
unlevered_cash_flow.iloc[sale_month:] = 0
unlevered_irr = npf.irr(unlevered_cash_flow.values[:sale_month])
unlevered_annual_irr = (1 + unlevered_irr) ** 12 - 1 if unlevered_irr is not None and not np.isnan(unlevered_irr) else 0

# Untrended Yield Calculations
weighted_rent_untr = (studio_rent * studio_units + one_bed_rent * one_bed_units + two_bed_rent * two_bed_units + three_bed_rent * three_bed_units) / num_units if num_units > 0 else 0
stabilized_gpr_untr = weighted_rent_untr * num_units * 12
stabilized_vacancy_untr = stabilized_gpr_untr * min_vacancy_rate
stabilized_credit_loss_untr = stabilized_gpr_untr * credit_loss
stabilized_ebr_untr = stabilized_gpr_untr - stabilized_vacancy_untr - stabilized_credit_loss_untr
stabilized_egi_untr = stabilized_ebr_untr / (1 - other_income_pct) if other_income_pct < 1 else stabilized_ebr_untr
stabilized_other_income_untr = stabilized_egi_untr - stabilized_ebr_untr
stabilized_controllable_untr = total_controllable_exp_per_unit * stabilized_occupied_units_target
stabilized_non_controllable_untr = total_non_controllable_exp_per_unit * stabilized_occupied_units_target
stabilized_re_taxes_untr = re_taxes_per_unit_year * stabilized_occupied_units_target
stabilized_mgmt_fee_untr = stabilized_egi_untr * mgmt_fee_pct
stabilized_noi_untr = stabilized_egi_untr - (stabilized_controllable_untr + stabilized_non_controllable_untr + stabilized_re_taxes_untr + stabilized_mgmt_fee_untr)
untrended_yoc = stabilized_noi_untr / total_costs if total_costs > 0 else 0
untrended_debt_yield = stabilized_noi_untr / total_debt if total_debt > 0 else 0

start_stabilization_index, end_stabilization_index = months_to_stabilize, months_to_stabilize + 12
if end_stabilization_index <= len(cf):
    trended_noi_stabilized = cf['NOI for Exit'].iloc[start_stabilization_index:end_stabilization_index].sum()
    trended_yoc = trended_noi_stabilized / total_costs if total_costs > 0 else 0
    trended_debt_yield = trended_noi_stabilized / total_debt if total_debt > 0 else 0
else:
    trended_yoc, trended_debt_yield = 0, 0

stabilized_period_start = dates[months_to_stabilize].strftime('%b-%Y')
stabilized_period_end = dates[end_stabilization_index-1].strftime('%b-%Y')
stabilized_cf_title = f"Untrended Stabilized Annual Cash Flow ({stabilized_period_start} - {stabilized_period_end})"

stabilized_cf_data = { "Line Item": [], "Annual Amount": [] }
stabilized_cf_data["Line Item"].extend(["Gross Potential Rent", "(-) Vacancy Loss", "(-) Credit Loss", "(+) Other Income", "Effective Gross Income"])
stabilized_cf_data["Annual Amount"].extend([stabilized_gpr_untr, -stabilized_vacancy_untr, -stabilized_credit_loss_untr, stabilized_other_income_untr, stabilized_egi_untr])
stabilized_cf_data["Line Item"].append("(-) Controllable Expenses:")
stabilized_cf_data["Annual Amount"].append(np.nan)
for name, amount in detailed_controllable_exp.items():
    stabilized_cf_data["Line Item"].append(f"    - {name}")
    stabilized_cf_data["Annual Amount"].append(-(amount * stabilized_occupied_units_target))
stabilized_cf_data["Line Item"].append("(-) Non-Controllable Expenses:")
stabilized_cf_data["Annual Amount"].append(np.nan)
for name, amount in detailed_non_controllable_exp.items():
    stabilized_cf_data["Line Item"].append(f"    - {name}")
    stabilized_cf_data["Annual Amount"].append(-(amount * stabilized_occupied_units_target))
stabilized_cf_data["Line Item"].extend(["(-) Real Estate Taxes", "(-) Management Fee", "Net Operating Income", "", "Total Project Costs", "Untrended Yield on Cost"])
stabilized_cf_data["Annual Amount"].extend([-stabilized_re_taxes_untr, -stabilized_mgmt_fee_untr, stabilized_noi_untr, np.nan, total_costs, untrended_yoc])
stabilized_cf_table = pd.DataFrame(stabilized_cf_data)

sources_df = pd.DataFrame({"Sources": ["Construction Loan", "Total Common Equity"], "Amount": [total_debt, total_equity]})
development_budget_df = pd.DataFrame({
    "Uses": ["Land Costs", "Hard Costs", "Soft Costs", "Fees", "Closing Costs", "Capitalized Interest", "Operating Reserve"],
    "Amount": [land_costs, total_hard_costs, total_soft_costs, total_fees, total_closing_costs, total_capitalized_interest, total_operating_reserve]
})

annual_summary = cf.groupby(cf.index.year).sum()

unit_types = ["Studio", "1-Bedroom", "2-Bedroom", "3-Bedroom"]
base_rents = [studio_rent, one_bed_rent, two_bed_rent, three_bed_rent]
num_years = int(np.ceil(hold_period / 12))
rent_schedule_data = {"Unit Type": unit_types, "Untrended Monthly Rent": base_rents}
for year in range(1, num_years + 1):
    year_end_month_index = min(year * 12, hold_period) - 1
    avg_escalation_factor = cumulative_rev_factor[year_end_month_index]
    rent_schedule_data[f"Year {year} Avg. Rent"] = [r * avg_escalation_factor for r in base_rents]
unit_mix_rent_schedule = pd.DataFrame(rent_schedule_data)

# --- Final Prep for Excel Export ---
excel_cf = cf.copy()
# Format as negative numbers
excel_cf['Vacancy Loss'] *= -1
excel_cf['Credit Loss'] *= -1
for name in detailed_controllable_exp: excel_cf[f'Opex - {name}'] *= -1
for name in detailed_non_controllable_exp: excel_cf[f'Opex - {name}'] *= -1
excel_cf['RE Taxes'] *= -1
excel_cf['Management Fee'] *= -1
excel_cf['Concession Cost'] *= -1
excel_cf['Total Opex'] *= -1
excel_cf['Replacement Reserve'] *= -1
# Add reconciliation sections
excel_cf['Spacer_1'] = ''
excel_cf['ICF Check: Equity Draw (-)'] = -cf['Equity Draw']
excel_cf['ICF Check: NOI (+)'] = cf['NOI']
excel_cf['ICF Check: Interest Paid (-)'] = -cf['Interest Paid from Operations']
excel_cf['ICF Check: Reserves (-)'] = -cf['Replacement Reserve']
excel_cf['ICF Check: Excess Cash (+)'] = excess_cash_flow
excel_cf['ICF Check: Op Reserve Draw (+)'] = cf['Operating Reserve Drawdown'] # Add to check
excel_cf['ICF Check: Sale Proceeds (+)'] = 0
excel_cf.loc[excel_cf.index[sale_month - 1], 'ICF Check: Sale Proceeds (+)'] = net_sale_proceeds_levered
excel_cf['ICF Check: Sum'] = excel_cf[['ICF Check: Equity Draw (-)', 'ICF Check: NOI (+)', 'ICF Check: Interest Paid (-)', 'ICF Check: Reserves (-)', 'ICF Check: Excess Cash (+)', 'ICF Check: Op Reserve Draw (+)', 'ICF Check: Sale Proceeds (+)']].sum(axis=1)
excel_cf['ICF Check: Difference'] = cf['Investor Cash Flow'] - excel_cf['ICF Check: Sum']

excel_cf['Spacer_2'] = ''
excel_cf['Recon Check: Equity Draw (+)'] = cf['Equity Draw']
excel_cf['Recon Check: Debt Draw (+)'] = cf['Debt Draw']
excel_cf['Recon Check: NOI (+)'] = cf['NOI']
excel_cf['Recon Check: Project Spend (-)'] = -cf['Base Project Spend']
excel_cf['Recon Check: Interest (-)'] = -cf['Monthly Interest']
excel_cf['Recon Check: Sum'] = excel_cf[['Recon Check: Equity Draw (+)', 'Recon Check: Debt Draw (+)', 'Recon Check: NOI (+)', 'Recon Check: Project Spend (-)', 'Recon Check: Interest (-)']].sum(axis=1)
excel_cf['Cash Reconciliation'] = excel_cf['Recon Check: Sum']
excel_cf['Recon Check: Difference'] = excel_cf['Cash Reconciliation'] - excel_cf['Recon Check: Sum']

# Define final column order for Excel
excel_cols = [
    'Month', 'Occupied Units', 'Occupied Studio', 'Occupied 1BR', 'Occupied 2BR', 'Occupied 3BR', 'New Leases', 'Cumulative Revenue Factor', 'Cumulative Opex Factor',
    'Rent Studio', 'Rent 1BR', 'Rent 2BR', 'Rent 3BR', 'Weighted Average Rent',
    'GPR Studio', 'GPR 1BR', 'GPR 2BR', 'GPR 3BR', 'GPR', 'Vacancy Loss', 'Credit Loss',
    'Effective Billed Revenue', 'Other Income', 'EGI'
]
excel_cols.extend([f'Opex - {name}' for name in detailed_controllable_exp])
excel_cols.extend([f'Opex - {name}' for name in detailed_non_controllable_exp])
excel_cols.extend(['RE Taxes', 'Management Fee', 'Concession Cost', 'Total Opex', 'NOI', 'Replacement Reserve', 'NOI After Reserves', 'Net Funding Requirement',
    'Equity Draw', 'Debt Draw', 'Monthly Interest', 'Capitalized Interest', 'Operating Reserve Drawdown', 'Interest Paid from Operations', 'Excess Cash Flow', 
    'Sale Proceeds', 'Debt Repayment', 'Net Reversion Cash Flow', 'Investor Cash Flow',
    'Spacer_1', 'ICF Check: Equity Draw (-)', 'ICF Check: NOI (+)', 'ICF Check: Interest Paid (-)', 'ICF Check: Reserves (-)', 'ICF Check: Excess Cash (+)', 'ICF Check: Op Reserve Draw (+)', 'ICF Check: Sale Proceeds (+)', 'ICF Check: Sum', 'ICF Check: Difference',
    'Spacer_2', 'Recon Check: Equity Draw (+)', 'Recon Check: Debt Draw (+)', 'Recon Check: NOI (+)', 'Recon Check: Project Spend (-)', 'Recon Check: Interest (-)', 'Recon Check: Sum', 'Cash Reconciliation', 'Recon Check: Difference'
])

excel_cf['Sale Proceeds'], excel_cf['Debt Repayment'], excel_cf['Net Reversion Cash Flow'] = 0, 0, 0
excel_cf.loc[excel_cf.index[sale_month - 1], 'Sale Proceeds'] = exit_value - sale_costs
excel_cf.loc[excel_cf.index[sale_month - 1], 'Debt Repayment'] = -outstanding_debt
excel_cf.loc[excel_cf.index[sale_month - 1], 'Net Reversion Cash Flow'] = net_sale_proceeds_levered

excel_export_dict = {'Monthly Pro-Forma (Detailed)': excel_cf[excel_cols], 'Annual Summary': annual_summary, 'Development Budget': development_budget_df}
excel_data = to_excel(excel_export_dict)

# --- Display Results ---
with main_container:
    st.header("Project Overview")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Start Date", project_start_date.strftime("%b %Y"))
    col1.metric("Number of Units", f"{num_units}")
    col2.metric("Total Gross SF", f"{gsf:,.0f} SF")
    col2.metric("Total Rentable SF", f"{total_rsf:,.0f} SF")
    col3.metric("Total Project Cost / Unit", f"${total_costs/num_units:,.0f}" if num_units > 0 else "$0")
    col3.metric("Sale Month", dates[sale_month-1].strftime("%b-%Y") if sale_month <= hold_period else "N/A")
    col4.metric("Construction Period", f"{construction_period} Months")
    col4.metric("Projected Hold Period (Months)", f"{sale_month}")
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

    with st.expander("Model Health & Audit Checks"):
        col1, col2, col3 = st.columns(3)
        total_sources_check = total_debt + total_equity
        total_uses_check = development_budget_df['Amount'].sum()
        s_u_diff = total_sources_check - total_uses_check
        col1.metric("Total Sources vs. Uses", f"${s_u_diff:,.0f}", "Difference")
        if abs(s_u_diff) < 1: col1.success("Balanced")
        else: col1.error("Unbalanced")
        
        loan_balance_check = outstanding_debt
        sum_draws_check = cf['Debt Draw'].sum()
        loan_diff = loan_balance_check - sum_draws_check
        col2.metric("Final Loan Balance vs. Sum of Draws", f"${loan_diff:,.0f}", "Difference")
        if abs(loan_diff) < 1: col2.success("Balanced")
        else: col2.error("Unbalanced")

        sum_equity_draws_check = cf['Equity Draw'].sum()
        equity_diff = total_equity - sum_equity_draws_check
        col3.metric("Required Equity vs. Sum of Draws", f"${equity_diff:,.0f}", "Difference")
        if abs(equity_diff) < 1: col3.success("Balanced")
        else: col3.error("Unbalanced")

    st.markdown("---")
    
    st.header("Summaries & Schedules")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Sources & Uses")
        sources_df.loc['Total'] = sources_df.sum(numeric_only=True)
        st.dataframe(sources_df.style.format({"Amount": "${:,.0f}"}), use_container_width=True, hide_index=True)
        
        st.subheader("Development Budget")
        development_budget_df.loc['Total'] = development_budget_df.sum(numeric_only=True)
        st.dataframe(development_budget_df.style.format({"Amount": "${:,.0f}"}), use_container_width=True, hide_index=True)
        
    with col2:
        st.subheader(stabilized_cf_title)
        format_mapping = { "Annual Amount": lambda x: f"{x*100:.2f}%" if (isinstance(x, (int, float)) and abs(x) < 1 and x != 0 and not pd.isna(x)) else f"${x:,.0f}" }
        st.dataframe(stabilized_cf_table.style.format(format_mapping, na_rep=""), use_container_width=True, hide_index=True)

    st.subheader("Unit Mix Rent Schedule")
    st.dataframe(unit_mix_rent_schedule.style.format("${:,.2f}", subset=[col for col in unit_mix_rent_schedule.columns if col != "Unit Type"]), use_container_width=True, hide_index=True)
    st.markdown("---")

    st.header("Visualizations & Audits")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Net Equity Cash Flow (Monthly)")
        st.bar_chart(cf.loc[cf.index < pd.to_datetime(dates[sale_month]), 'Investor Cash Flow'])
    with col2:
        st.subheader("Project Cash Usage (During Construction)")
        construction_spend = cf['Base Project Spend'].iloc[:construction_period]
        st.bar_chart(construction_spend)

    st.subheader("Lease-Up Velocity")
    lease_up_df = cf.loc[(cf['Month'] > construction_period) & (cf['Month'] <= months_to_stabilize + 1), ['Month', 'Occupied Units']].copy()
    lease_up_df['Monthly New Leases'] = lease_up_df['Occupied Units'].diff().fillna(lease_up_df['Occupied Units'].iloc[0] if not lease_up_df.empty else 0)
    lease_up_df_melted = lease_up_df.melt(id_vars=['Month'], value_vars=['Monthly New Leases', 'Occupied Units'], var_name='Metric', value_name='Units')
    
    base = alt.Chart(lease_up_df_melted).encode(x=alt.X('Month:O', title='Month of Operation'))
    bars = base.transform_filter(alt.datum.Metric == 'Monthly New Leases').mark_bar(color='#4c78a8').encode(y=alt.Y('Units:Q', title='Monthly New Leases'))
    line = base.transform_filter(alt.datum.Metric == 'Occupied Units').mark_line(color='#e45756', strokeWidth=3).encode(y=alt.Y('Units:Q', title='Total Occupied Units'))
    chart = alt.layer(bars, line).resolve_scale(y='independent').properties(height=300)
    st.altair_chart(chart, use_container_width=True)
    st.markdown("---")

    st.header("Detailed Financial Statements")
    st.subheader("Annual Cash Flow Summary")
    annual_display_cols_ui = ['GPR', 'EGI', 'Total Opex', 'NOI', 'Replacement Reserve', 'NOI After Reserves', 'Capitalized Interest', 'Interest Paid from Operations', 'Investor Cash Flow']
    st.dataframe(annual_summary[annual_display_cols_ui].style.format("${:,.0f}"), use_container_width=True)
    
    st.subheader("Monthly Pro-Forma Cash Flow")
    display_cols = ['Month', 'NOI', 'Replacement Reserve', 'NOI After Reserves', 'Base Project Spend', 'Monthly Interest', 'Capitalized Interest', 'Interest Paid from Operations', 'Equity Draw', 'Debt Draw', 'Investor Cash Flow']
    display_cf = cf[display_cols].copy()
    st.dataframe(display_cf.style.format(formatter="${:,.0f}", subset=[c for c in display_cols if c != 'Month']), use_container_width=True, hide_index=True)
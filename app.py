import numpy as np
import numpy_financial as npf
import pandas as pd
import streamlit as st

st.set_page_config(page_title="CRE Underwriting Calculator", page_icon="🧮", layout="wide")
st.title("🧮 Real Estate Deal Underwriting Calculator")
st.caption(
    "Enter a deal's basic numbers to see its cap rate, DSCR, cash-on-cash return, and a full "
    "multi-year levered IRR -- the same math behind a typical acquisition underwriting model."
)

# ---------------------------------------------------------------------------
# Amortization helpers
# ---------------------------------------------------------------------------

def annual_debt_service(loan_amount, annual_rate, amort_years):
    """Standard fixed-rate, monthly-amortizing mortgage payment, annualized."""
    if loan_amount <= 0:
        return 0.0
    monthly_rate = annual_rate / 12
    n_periods = amort_years * 12
    if monthly_rate == 0:
        return loan_amount / amort_years
    monthly_payment = loan_amount * monthly_rate / (1 - (1 + monthly_rate) ** (-n_periods))
    return monthly_payment * 12


def remaining_loan_balance(loan_amount, annual_rate, amort_years, years_elapsed):
    """Outstanding principal after `years_elapsed` years of amortization."""
    if loan_amount <= 0:
        return 0.0
    monthly_rate = annual_rate / 12
    n_periods = amort_years * 12
    k_periods = years_elapsed * 12
    if monthly_rate == 0:
        return max(loan_amount * (1 - years_elapsed / amort_years), 0.0)
    balance = loan_amount * (
        (1 + monthly_rate) ** n_periods - (1 + monthly_rate) ** k_periods
    ) / ((1 + monthly_rate) ** n_periods - 1)
    return max(balance, 0.0)


# ---------------------------------------------------------------------------
# Sidebar inputs
# ---------------------------------------------------------------------------

def format_money(value):
    """Shorthand dollar formatting so KPI tiles never overflow -- $1.75M instead
    of $1,750,000, which was getting cut off with "..." in narrower windows."""
    if abs(value) >= 1_000_000:
        return f"${value / 1_000_000:.2f}M"
    if abs(value) >= 1_000:
        return f"${value / 1_000:.0f}K"
    return f"${value:,.0f}"


st.sidebar.header("Deal Inputs")

purchase_price = st.sidebar.number_input("Purchase Price ($)", min_value=100_000, value=5_000_000, step=50_000)
year1_noi = st.sidebar.number_input("Year 1 NOI ($)", min_value=0, value=300_000, step=5_000)
noi_growth = st.sidebar.slider("Annual NOI Growth Rate (%)", 0.0, 8.0, 3.0, 0.25) / 100

st.sidebar.divider()
st.sidebar.subheader("Financing")
ltv = st.sidebar.slider("Loan-to-Value (%)", 0, 90, 65, 5) / 100
interest_rate = st.sidebar.slider("Interest Rate (%)", 2.0, 10.0, 6.5, 0.125) / 100
amort_years = st.sidebar.selectbox("Amortization (years)", [20, 25, 30], index=2)

st.sidebar.divider()
st.sidebar.subheader("Hold & Exit")
hold_period = st.sidebar.slider("Hold Period (years)", 1, 10, 5)
exit_cap_rate = st.sidebar.slider("Exit Cap Rate (%)", 3.0, 10.0, 5.5, 0.25) / 100
selling_costs_pct = st.sidebar.slider("Selling Costs (%)", 0.0, 5.0, 2.0, 0.25) / 100

# ---------------------------------------------------------------------------
# Core underwriting metrics
# ---------------------------------------------------------------------------

loan_amount = purchase_price * ltv
equity_invested = purchase_price - loan_amount
going_in_cap_rate = year1_noi / purchase_price if purchase_price else 0
debt_service = annual_debt_service(loan_amount, interest_rate, amort_years)
dscr = year1_noi / debt_service if debt_service else float("inf")
year1_cash_flow = year1_noi - debt_service
cash_on_cash = year1_cash_flow / equity_invested if equity_invested else 0

# ---------------------------------------------------------------------------
# KPI tiles
# ---------------------------------------------------------------------------

col1, col2, col3, col4 = st.columns(4)
col1.metric("Going-in Cap Rate", f"{going_in_cap_rate * 100:.2f}%")
col2.metric("Year 1 DSCR", f"{dscr:.2f}x", delta_color="off")
col3.metric("Year 1 Cash-on-Cash", f"{cash_on_cash * 100:.2f}%")
col4.metric("Equity Invested", format_money(equity_invested))

if dscr < 1.20:
    st.warning(
        f"DSCR of {dscr:.2f}x is below the 1.20x-1.25x minimum most lenders require -- this loan "
        "amount or rate combination likely wouldn't qualify as structured."
    )

# ---------------------------------------------------------------------------
# Multi-year proforma + levered IRR
# ---------------------------------------------------------------------------

years = list(range(1, hold_period + 1))
noi_by_year = [year1_noi * (1 + noi_growth) ** (y - 1) for y in years]
cash_flow_by_year = [noi - debt_service for noi in noi_by_year]

exit_year_noi = year1_noi * (1 + noi_growth) ** hold_period  # NOI in the year after sale
exit_value = exit_year_noi / exit_cap_rate if exit_cap_rate else 0
selling_costs = exit_value * selling_costs_pct
loan_payoff = remaining_loan_balance(loan_amount, interest_rate, amort_years, hold_period)
net_sale_proceeds = exit_value - selling_costs - loan_payoff

proforma = pd.DataFrame({
    "Year": years,
    "NOI": noi_by_year,
    "Debt Service": [debt_service] * hold_period,
    "Cash Flow": cash_flow_by_year,
})
proforma.loc[proforma["Year"] == hold_period, "Cash Flow"] += net_sale_proceeds

levered_cash_flows = [-equity_invested] + proforma["Cash Flow"].tolist()
irr = npf.irr(levered_cash_flows)
total_distributions = sum(proforma["Cash Flow"])
equity_multiple = total_distributions / equity_invested if equity_invested else 0

st.divider()
st.header("📈 Hold-Period Returns")

col1, col2, col3 = st.columns(3)
col1.metric("Levered IRR", f"{irr * 100:.1f}%" if irr is not None else "n/a")
col2.metric("Equity Multiple", f"{equity_multiple:.2f}x")
col3.metric("Exit Value", format_money(exit_value))

st.dataframe(
    proforma,
    column_config={
        "NOI": st.column_config.NumberColumn(format="$%d"),
        "Debt Service": st.column_config.NumberColumn(format="$%d"),
        "Cash Flow": st.column_config.NumberColumn(format="$%d"),
    },
    hide_index=True,
    use_container_width=True,
)

st.line_chart(proforma, x="Year", y="Cash Flow")

# ---------------------------------------------------------------------------
# Sensitivity analysis: IRR across exit cap rate x NOI growth rate
# ---------------------------------------------------------------------------

st.divider()
st.header("🔬 Sensitivity Analysis: Levered IRR")
st.caption(
    "How the deal's IRR shifts under different exit cap rate and NOI growth assumptions -- the same "
    "kind of sensitivity table used to stress-test a deal before committing capital."
)


def irr_for(exit_cap, growth):
    noi_by_yr = [year1_noi * (1 + growth) ** (y - 1) for y in years]
    cf_by_yr = [noi - debt_service for noi in noi_by_yr]
    exit_noi = year1_noi * (1 + growth) ** hold_period
    ev = exit_noi / exit_cap if exit_cap else 0
    sc = ev * selling_costs_pct
    payoff = remaining_loan_balance(loan_amount, interest_rate, amort_years, hold_period)
    proceeds = ev - sc - payoff
    cfs = cf_by_yr.copy()
    cfs[-1] += proceeds
    return npf.irr([-equity_invested] + cfs)


exit_cap_grid = [exit_cap_rate + delta for delta in (-0.01, -0.005, 0, 0.005, 0.01)]
growth_grid = [noi_growth + delta for delta in (-0.02, -0.01, 0, 0.01, 0.02)]

sensitivity = pd.DataFrame(
    {f"{g * 100:.1f}% growth": [irr_for(ec, g) for ec in exit_cap_grid] for g in growth_grid},
    index=[f"{ec * 100:.2f}% exit cap" for ec in exit_cap_grid],
)

st.dataframe(
    sensitivity.style.format("{:.1%}").background_gradient(cmap="RdYlGn", axis=None),
    use_container_width=True,
)

st.divider()
with st.expander("How these numbers are calculated"):
    st.markdown(
        "**Going-in cap rate** = Year 1 NOI ÷ Purchase Price.\n\n"
        "**Debt service** uses a standard fixed-rate, monthly-amortizing loan payment, annualized.\n\n"
        "**DSCR** = NOI ÷ annual debt service -- most lenders require at least 1.20x-1.25x.\n\n"
        "**Cash-on-cash return** = (NOI − debt service) ÷ equity invested, for the first year.\n\n"
        "**Exit value** = the NOI in the year *after* the sale, divided by the exit cap rate (the "
        "standard convention, since a buyer is underwriting the property's *next* year of income).\n\n"
        "**Levered IRR** discounts the initial equity outflow against every year's cash flow plus "
        "net sale proceeds (exit value, minus selling costs, minus the remaining loan balance) in "
        "the final year."
    )

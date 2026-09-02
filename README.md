# CRE Underwriting Calculator

An interactive Streamlit tool for underwriting a commercial real estate acquisition -- the same
math behind a typical Excel acquisition model, built as a live, shareable web app instead.

## What it calculates

- Going-in cap rate, loan sizing, and equity required
- Debt Service Coverage Ratio (DSCR) against a standard fixed-rate amortizing loan
- Year 1 cash-on-cash return
- A full multi-year proforma with levered cash flows
- Exit value, net sale proceeds, and levered IRR / equity multiple over a chosen hold period
- A sensitivity table showing how IRR shifts across a grid of exit cap rate and NOI growth assumptions

## Stack

Python, Streamlit, pandas, numpy_financial

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

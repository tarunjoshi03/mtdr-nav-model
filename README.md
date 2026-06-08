# Matador Resources (MTDR) — Upstream NAV Model

**Built by Tarun Joshi | FY2025 10-K Data | Filed Feb 26, 2026**

---

## What This Is

A full upstream Net Asset Value (NAV) model for Matador Resources (NYSE: MTDR), built from scratch using public SEC filings. Covers the full NAV bridge from proved reserves to equity value per share, with oil price and discount rate sensitivity.

This is the same framework used by oil & gas investment banks and A&D advisory teams to value upstream E&P companies — PDP cashflows, PUD development economics, midstream asset valuation, and cap structure bridge to equity NAV.

---

## Model Structure

### Python Engine (`nav_model.py`)
- **Arps Hyperbolic Decline Curves** — calibrated to Delaware Basin Wolfcamp/Bone Spring parameters (Di = 40%, b = 1.30, Dmin = 6%)
- **Annual Cashflow Engine** — 20-year projection, mid-year discounting convention
- **PDP NAV** — NPV of proved developed producing reserves at base case pricing
- **PUD NAV** — NPV of proved undeveloped reserves net of $3.8B development capex phased over 5 years
- **Midstream NAV** — San Mateo 51% stake valued at 8.0x EV/EBITDA
- **Sensitivity Tables** — 7 oil prices × 4 discount rates for both equity NAV/share and total asset NAV

### Excel Deliverable (`MTDR_NAV_Model_TarunJoshi.xlsx`)
Four tabs:
- **Summary** — NAV bridge + key metrics, all inputs sourced to 10-K
- **Decline Curves & CF** — Year-by-year PDP and PUD cashflow tables
- **Sensitivity** — NAV/share matrix + premium/discount to current price
- **Assumptions** — Every hardcoded input sourced with filing reference

---

## Base Case Results

| Component | Value |
|---|---|
| PDP NPV10 | $7.56B |
| PUD NPV10 (net capex) | $1.90B |
| Midstream (San Mateo 51%) | $2.45B |
| **Total Asset NAV** | **$11.91B** |
| Less: Net Debt | ($3.40B) |
| **Equity NAV** | **$8.51B** |
| **NAV / Share** | **$68.52** |
| Current Price (ref.) | ~$41.31 |
| **Implied Upside** | **+65.9%** |

*Base case: $65/Bbl WTI, $2.25/Mcf gas, 10% discount rate*

---

## Key Inputs (All from MTDR FY2025 10-K)

| Input | Value | Source |
|---|---|---|
| PDP Reserves | 406.5 MMBoe | 10-K Item 1, Reserves Table |
| PUD Reserves | 260.5 MMBoe | 10-K Item 1, Reserves Table |
| Avg Daily Production | 207,070 BOE/d | 10-K Operating Summary |
| Realized Oil Price | $64.99/Bbl | 10-K Operating Summary |
| LOE | $5.50/BOE | 10-K Operating Summary |
| G&A | $1.81/BOE | 10-K Operating Summary |
| PV-10 (SEC) | $8.24B | 10-K Item 1 |
| Total Debt | $3.41B | 10-K Balance Sheet |
| Shares Outstanding | 124.25mm | 10-K Cover Page |

---

## Why This Matters

MTDR is a pure-play Delaware Basin operator — 99% of reserves in the Wolfcamp and Bone Spring plays. The model shows the stock was trading at a ~39% discount to NAV at the buyback reference price (~$41), largely because SEC pricing used a 14% lower oil price deck than the forward strip, compressing the reported PV-10.

The midstream contribution (~$2.45B from San Mateo) is often underappreciated in simple reserve-based valuations — this model isolates it explicitly.

---

## Requirements

```bash
pip install numpy scipy openpyxl
python nav_model.py
```

Outputs: `MTDR_NAV_Model_TarunJoshi.xlsx`

---

## Related Work

[PR / APA Deal Teardown](https://github.com/tarunjoshi03/pr-apa-deal-teardown) — Permian Resources acquisition of APA's Northern Delaware Basin assets: decline curve fitting through IRR sensitivity, all costs from 10-K filings.

---

*All data sourced from public SEC filings. This model is for educational and portfolio purposes only — not investment advice.*

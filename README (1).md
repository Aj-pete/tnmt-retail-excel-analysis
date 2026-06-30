# TNMT Retail — Business & Sales Performance Analysis
### Excel Portfolio Project | Business Analytics

---

## 📌 Project Overview

This project delivers a full-scope business analysis for **TNMT**, a mid-size US retail company operating across 367 stores, 45 states, and 4 sales regions. Working from a raw transactional dataset of 20,000 sales orders, the analysis was built to answer critical performance questions ahead of a leadership board meeting.

The deliverable is a **multi-tab Excel workbook** covering data preparation, sales performance, trend analysis, underperformance diagnostics, revenue segmentation, an interactive executive dashboard with slicers, and a structured recommendations report.

> **Role:** Business Analyst  
> **Tools:** Microsoft Excel (PivotTables, XLOOKUP, Slicers, Conditional Formatting, Combo Charts, Sparklines, VBA Macro)  
> **Dataset:** Internal sales transaction data — January to May 2020  
> **Audience:** Executive leadership and board of directors

---

## 📁 Repository Contents

```
├── TNMT_Retail_Analysis.xlsx        # Main Excel workbook (8 tabs)
├── README.md                        # This file
└── data/
    └── TNMT_Retail_raw_data.xlsx    # Original source file (unmodified)
```

---

## 📊 Dataset Summary

| Dimension | Detail |
|---|---|
| **Total Orders** | 20,000 transactions |
| **Period Covered** | January 2020 – May 2020 |
| **Total Revenue** | $44,365,028 |
| **Total Gross Profit** | $12,675,722 |
| **Overall Gross Margin** | 28.6% |
| **Average Order Value** | $2,218 |
| **Products** | 47 products across 10 categories |
| **Sales Reps** | 28 representatives |
| **Regions** | 4 (Midwest, Northeast, South, West) |
| **Sales Channels** | 4 (Online, In-Store, Wholesale, Distributor) |
| **Store Locations** | 367 stores across 45 US states |

### Source Tables

| Table | Rows | Description |
|---|---|---|
| Sales Table | 20,000 | Core transaction data — orders, dates, channels, quantity, price, cost |
| Sales Team | 28 | Sales rep names mapped to regions |
| Products | 47 | Product names mapped to categories |
| Store Locations | 367 | Store cities with state, population, and median income data |

---

## ❓ Business Questions Answered

The analysis was structured around 7 core questions posed by leadership:

1. **Which products and categories are driving revenue and profit?**
   → Category and product performance ranked by revenue and gross profit

2. **Are our four sales channels equally efficient?**
   → Revenue per order calculated by channel to measure efficiency, not just volume

3. **Which regions and sales reps are over- or underperforming?**
   → Benchmarked against team averages on both revenue AND margin simultaneously

4. **Where is revenue growth accelerating or decelerating month over month?**
   → Month-over-month growth rates with directional indicators and trend sparklines

5. **Is our product catalogue too large for the revenue it generates?**
   → Pareto analysis across all 47 products to identify long-tail drag

6. **Where is value concentrated — which 20% of the portfolio drives the most return?**
   → Revenue segmentation, channel × region matrix, margin banding

7. **What should leadership do differently in the next half-year?**
   → 7 prioritised recommendations with owner, timeline, and expected impact

---

## 🗂️ Workbook Structure

| Tab | Type | Description |
|---|---|---|
| `RAW DATA` | 🔵 Data | Enriched transaction table — 24 columns including calculated KPIs and XLOOKUP joins |
| `TEAMS` | ⚫ Reference | Sales rep index with region mapping |
| `PRODUCTS` | ⚫ Reference | Product index with category mapping |
| `STORES` | ⚫ Reference | Store locations with demographic data |
| `SALES PERFORMANCE` | 🟠 Analysis | Revenue and profit by category, product, channel, region, and rep |
| `TREND ANALYSIS` | 🟠 Analysis | Monthly trends, MoM growth, channel and category heatmaps, sparklines |
| `UNDERPERFORMANCE` | 🟠 Analysis | Bottom products, flagged reps, underperforming regions |
| `SEGMENTATION` | 🟠 Analysis | Pareto analysis, channel × region matrix, margin banding |
| `DASHBOARD` | 🟢 Output | Interactive executive dashboard with 4 slicers and 5 charts |
| `RECOMMENDATIONS` | ⚪ Output | 7 data-backed recommendations with owner and timeline |

---

## 🔍 Key Findings

### Revenue Performance
- **Decoratives is the dominant category** at $13.3M — representing 30% of total portfolio revenue, more than double the second-ranked category (Furniture at $6.5M)
- **Top 3 products** — Clocks ($1.10M), Vases ($1.06M), and Cookware ($1.02M) — lead by revenue but operate within a tightly compressed range, suggesting no single hero product
- **Midwest leads regional revenue** at $13.6M (30.6% of total), with South, Northeast, and West all within $3.5M of each other

### Channel Analysis
- All four channels show remarkably similar revenue — Online ($11.2M), In-Store ($11.1M), Distributor ($11.0M), Wholesale ($11.0M)
- **Revenue per order** is the differentiating efficiency metric — channel strategy should be guided by this figure rather than total revenue volume alone

### Trend Analysis
- **May was the strongest revenue month** at $9.18M, followed by January ($9.07M)
- February showed the sharpest monthly dip at $8.56M — worth investigating for seasonal or operational causes
- Revenue range across the 5-month period is relatively tight ($8.56M–$9.18M), suggesting stable but not growing demand

### Pareto / Segmentation Finding
- **37 of 47 products are required to reach 80% of total revenue** — an anti-Pareto distribution
- Unlike a classic 80/20 portfolio, TNMT's revenue is spread relatively evenly across the catalogue, meaning there are no dominant star products and no obvious long-tail cuts
- This finding reframes the product strategy question: rather than cutting the bottom, the priority is **elevating the middle** — building stronger heroes from existing mid-tier products

### Sales Rep Concentration
- **Top 5 reps account for 20.2% of total team revenue**: Nicholas Cunningham, Donald Reynolds, Joe Price, Shawn Wallace, and Carl Nguyen
- While less extreme than typical retail concentration (often 80/20), the gap between top and bottom performers still represents a material coaching and retention risk

### Data Limitation — Gross Margin
- Unit cost across the dataset reflects a fixed ratio to unit price (71.4% cost : 100% price), resulting in a uniform 28.6% gross margin across all products and orders
- Margin % comparisons by product or category are therefore not differentiated in this dataset
- **Gross Profit ($)** is used as the primary profitability metric in place of margin % for all product and category-level analysis

---

## 💡 Recommendations

All 7 recommendations follow the framework: **Observation → Insight → Action → Expected Impact**

---

### 01 — Product Portfolio Rationalisation
**Priority: 🔴 HIGH | Owner: Category Manager + Procurement | Timeline: Q3 2020**

The TNMT catalogue contains 47 products with relatively even revenue distribution. While this avoids dependency on any single SKU, it creates operational complexity — inventory management, sales team attention, and warehouse space are consumed across the full range without proportional return from the bottom tier.

**Action:** Initiate a formal SKU rationalisation review for the lowest-revenue products. Set a 90-day threshold: each flagged product either hits a minimum revenue floor or exits the catalogue. Redirect freed operational capacity toward building stronger hero products from the current mid-tier.

---

### 02 — Gross Profit Recovery Through Pricing Strategy
**Priority: 🔴 HIGH | Owner: Pricing Lead + Finance | Timeline: Q2–Q3 2020**

With a uniform 28.6% gross margin, TNMT has no differentiated high-margin products to subsidise lower-margin lines. Pricing strategy becomes the single most accessible lever for GP improvement without requiring volume growth.

**Action:** Commission a pricing benchmarking exercise across all 47 products against category-level market rates. Model the GP impact of a 5% price increase on the top 20 revenue-generating products. Simultaneously renegotiate unit costs on the three highest-volume products — a 5% COGS reduction on these alone materially improves overall gross profit.

---

### 03 — Channel Investment Rebalancing
**Priority: 🟡 MEDIUM | Owner: Sales Director + Channel Managers | Timeline: Q3 2020**

Four channels generating near-identical revenue is unusual and warrants scrutiny. Equal volume does not mean equal value — fulfilment costs, return rates, and sales effort per channel vary materially. Without a net margin by channel view, resource allocation decisions are flying blind.

**Action:** Conduct a full channel cost analysis inclusive of fulfilment, returns, and sales team time to calculate true net margin per channel. Shift resource allocation toward the highest net-margin channel. Review Distributor channel terms — if net margin after distributor fees falls below the portfolio threshold, renegotiate or reduce volume commitment.

---

### 04 — Sales Team Performance Management
**Priority: 🔴 HIGH | Owner: Sales Director + HR | Timeline: Immediate (Q2 2020)**

The top 5 reps generate 20.2% of team revenue — a concentration that creates retention risk. Simultaneously, reps flagged as below average on both revenue and margin represent genuine performance drag that is not offset by volume contribution.

**Action:**
1. Implement quarterly targets segmented by both revenue AND gross profit — not revenue alone, which incentivises discounting
2. Assign the top 3 performing reps as structured peer mentors for the bottom quartile over 60 days
3. Initiate retention conversations — compensation review, recognition, and career development — with the top 5 reps within 30 days
4. Conduct territory reviews for dual underperformers before any performance management action, to distinguish execution gaps from structural territory disadvantages

---

### 05 — Regional Resource Reallocation
**Priority: 🟡 MEDIUM | Owner: Regional Sales Managers + Finance | Timeline: Q3 2020**

The Midwest leads total revenue at $13.6M, but this reflects its larger rep count rather than superior efficiency. Revenue per rep — a normalised performance metric — may tell a different story and should govern how regional targets and resources are set.

**Action:** Rebase regional performance targets on revenue per rep, not total revenue. Assess the lowest revenue-per-rep region for structural causes (territory size, store density, product mix) before making headcount decisions. Consider adding one additional rep to the highest revenue-per-rep region to capture demonstrably available capacity.

---

### 06 — Channel × Geography Pilot Expansion
**Priority: 🟡 MEDIUM | Owner: Strategy + Regional Sales Lead | Timeline: Q3–Q4 2020**

The channel × region revenue matrix reveals geographic variation in channel performance. Where a channel over-performs in one region but is underweight in another with comparable store counts, the cause is almost always execution rather than market demand — making it a low-risk, high-return expansion opportunity.

**Action:** Identify the highest-performing channel-region combination and the equivalent underweight region. Launch a 90-day pilot replicating the successful channel strategy in the underweight region. Measure revenue and gross profit outcomes at 30-day intervals against a control baseline.

---

### 07 — State-Level Performance Activation via Dashboard Slicer
**Priority: 🟡 MEDIUM | Owner: Strategy + Regional Managers | Timeline: Q3–Q4 2020**

The executive dashboard includes a State slicer connected to all pivot-driven charts, enabling real-time drill-down to state-level performance. Cross-referencing revenue per store against state median household income (available in the Store Locations table) will surface high-income states where TNMT has store presence but underperforms revenue potential.

**Action:**
1. Use the dashboard State slicer to identify states with above-average median income but below-average revenue per store — these are the highest-priority activation markets
2. Commission a 30-day store-level diagnostic in the top 3 identified states — product range audit, mystery shopping, and local competitor benchmarking
3. Test a targeted local marketing activation in one identified state in Q3, measuring revenue per store uplift against a control group over 60 days

**Impact:** Bringing underperforming high-income states to portfolio-average revenue per store represents incremental revenue using existing infrastructure — no capital expenditure or new store openings required.

---

## 🛠️ Technical Skills Demonstrated

| Skill | Application |
|---|---|
| **Data Modelling** | Joined 4 relational tables using XLOOKUP across 20,000 rows |
| **Calculated Metrics** | Revenue, COGS, Gross Profit, Avg Order Value, Revenue per Rep |
| **PivotTables** | 15+ pivots across 5 analysis tabs with custom grouping and sorting |
| **Conditional Formatting** | Color scales, icon sets, data bars, text-based rule formatting |
| **Slicer Architecture** | 4 cross-tab slicers (Channel, Category, Month, State) connected to all dashboard pivots simultaneously |
| **Advanced Charts** | Combo charts, stacked area, horizontal bar, scatter, doughnut, sparklines |
| **VBA Macro** | Clear All Filters button resetting all slicer caches in one click |
| **Dashboard Design** | 10-second rule layout — KPI cards, chart grid, insight strip, dark header |
| **Business Communication** | Observation → Insight → Action → Impact recommendation framework |

---

## ⚠️ Data Limitations & Assumptions

- **Period:** Analysis covers January–May 2020 only (5 months). True year-over-year comparisons are not possible. All trend analysis is month-over-month within this window.
- **Gross Margin:** Unit cost reflects a fixed 71.4% ratio to unit price across all records. Gross margin % is uniform at 28.6% portfolio-wide and is therefore not used as a differentiating metric at product or category level. Gross Profit ($) is used instead.
- **Currency:** All transactions recorded in USD. No currency conversion required.
- **Impact Estimates:** All financial impact figures in recommendations are directional estimates based on available data. Full validation requires operational cost modelling and full-year revenue data.

---

## 👤 About This Project

This project was built as a portfolio capstone to demonstrate end-to-end business analysis capability in Excel — from raw data preparation through to executive presentation. It mirrors the workflow of a real BA engagement: receiving a business brief, interrogating the data, surfacing insights, and delivering actionable recommendations to leadership.

The dataset was provided as a synthetic retail transaction file. All analysis, metric design, chart construction, and recommendations are original work.

---

*Confidential — TNMT Retail Business Analytics | Prepared: May 2020 | Tool: Microsoft Excel*

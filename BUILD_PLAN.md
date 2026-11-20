# BUILD PLAN: Customer Acquisition Cost (CAC) Analysis Platform

**Project:** Customer Acquisition Cost (CAC) Analysis  
**Author:** Gourav Kumar (Roll: 2500320140055) | MCA 2nd Year (Data Analytics), ABES Engineering College  
**Faculty Mentor:** Prof. Alpna Lodhi  
**Repository Working Directory:** `C:\Users\asus\Downloads\MERN-Expense-Tracker-main`  
**Target Delivery:** Streamlit Community Cloud + ABES Academic Submission  

---

## 1. Project Summary & Baselined Artifact Directory

1. **Strategic Purpose:** Deliver an enterprise-grade Customer Acquisition Cost (CAC) and marketing unit economics analytics dashboard.
2. **Business Problem:** Digital marketing organizations suffer from opaque attribution and unmonitored CAC inflation when relying on blended metrics.
3. **Core Functionality:** Calculates channel-specific CAC, models capital payback against a 12-month solvency ceiling, and computes LTV:CAC ratios.
4. **Data Science & Statistics:** Incorporates One-Way ANOVA and Kruskal-Wallis hypothesis testing at $\alpha=0.05$ to prove channel cost divergence.
5. **Interactive Simulation:** Enables leadership to test budget reallocation vectors and churn sensitivity on marginal ROAS curves.
6. **Technology Stack:** Python 3.12/3.10, Pandas, NumPy, SciPy, Statsmodels, Plotly, Streamlit multipage, SQLite (`data/app.db`), and Tailwind CSS.
7. **Quality Assurance:** 14 automated pytest test cases verifying 100% mathematical parity, referential integrity, and sub-second query latency.
8. **UI/UX Excellence:** Executive Slate/Navy theme with Tailwind CSS metric cards and responsive flexbox layouts rendered via iframes.
9. **Deployment Architecture:** Public zero-downtime hosting on Streamlit Community Cloud continuously synced to GitHub `main`.
10. **Academic Compliance:** 100% compliant with ABES MCA major project rubrics, verified through end-to-end Requirements Traceability (RTM).

### Artifact Links (Steps 1–13):
- **Step 1 (Project Understanding):** [`docs/01_Customer_Acquisition_Cost_Project_Understanding.docx`](file:///c:/Users/asus/Downloads/MERN-Expense-Tracker-main/docs/01_Customer_Acquisition_Cost_Project_Understanding.docx)
- **Step 2 (HMW Statements):** [`docs/02_HMW_Statements.docx`](file:///c:/Users/asus/Downloads/MERN-Expense-Tracker-main/docs/02_HMW_Statements.docx)
- **Step 3 (Needs Analysis / BRs):** [`docs/03_Needs_Analysis.docx`](file:///c:/Users/asus/Downloads/MERN-Expense-Tracker-main/docs/03_Needs_Analysis.docx)
- **Step 4 (Traceability Matrix):** [`docs/04_Traceability_Matrix.xlsx`](file:///c:/Users/asus/Downloads/MERN-Expense-Tracker-main/docs/04_Traceability_Matrix.xlsx)
- **Step 5 (BRD):** [`docs/05_BRD.docx`](file:///c:/Users/asus/Downloads/MERN-Expense-Tracker-main/docs/05_BRD.docx)
- **Step 7 (TRD / Architecture):** [`docs/07_TRD.docx`](file:///c:/Users/asus/Downloads/MERN-Expense-Tracker-main/docs/07_TRD.docx)
- **Step 8 (Database Design):** [`docs/08_Database_Design.xlsx`](file:///c:/Users/asus/Downloads/MERN-Expense-Tracker-main/docs/08_Database_Design.xlsx)
- **Step 9 (ER Diagram):** [`docs/09_ER_Diagram.docx`](file:///c:/Users/asus/Downloads/MERN-Expense-Tracker-main/docs/09_ER_Diagram.docx) | [`docs/09_ER_Diagram.png`](file:///c:/Users/asus/Downloads/MERN-Expense-Tracker-main/docs/09_ER_Diagram.png)
- **Step 10 (Synthetic Dataset):** [`data/cac_analysis_data.xlsx`](file:///c:/Users/asus/Downloads/MERN-Expense-Tracker-main/data/cac_analysis_data.xlsx) | [`data/app.db`](file:///c:/Users/asus/Downloads/MERN-Expense-Tracker-main/data/app.db) | [`docs/10_Validation_Report.docx`](file:///c:/Users/asus/Downloads/MERN-Expense-Tracker-main/docs/10_Validation_Report.docx)
- **Step 11 (User Stories):** [`docs/11_User_Stories.docx`](file:///c:/Users/asus/Downloads/MERN-Expense-Tracker-main/docs/11_User_Stories.docx)
- **Step 12 (UI Framework):** [`docs/12_UI_Framework.docx`](file:///c:/Users/asus/Downloads/MERN-Expense-Tracker-main/docs/12_UI_Framework.docx) | [`src/ui.py`](file:///c:/Users/asus/Downloads/MERN-Expense-Tracker-main/src/ui.py)
- **Step 13 (Test Cases & RTM):** [`docs/13_Test_Cases.docx`](file:///c:/Users/asus/Downloads/MERN-Expense-Tracker-main/docs/13_Test_Cases.docx) | [`docs/04_Traceability_Matrix.xlsx`](file:///c:/Users/asus/Downloads/MERN-Expense-Tracker-main/docs/04_Traceability_Matrix.xlsx)

---

## 2. Final Production Folder Structure

```text
MERN-Expense-Tracker-main/
├── app.py                          # Entrypoint: Executive Overview, Tailwind Hero, KPI cards (UI-01)
├── pages/                          # Multipage Streamlit analytical routing
│   ├── 1_Channel_CAC.py            # Channel vs Blended CAC comparative visualizer (UI-02)
│   ├── 2_Funnel_Analysis.py        # Multi-stage conversion funnel & CPC/CPL leakage (UI-03)
│   ├── 3_Compare.py                # Volume vs Cost-Efficiency 2x2 scatter matrix (UI-04)
│   ├── 4_Statistical_Testing.py    # One-Way ANOVA & Kruskal-Wallis inferential testing (UI-05)
│   ├── 5_Budget_Simulator.py       # Unit economics, marginal ROAS & churn sensitivity (UI-06)
│   ├── 6_Data_Explorer.py          # Raw transactional table explorer & Excel/CSV download (UI-07)
│   └── 7_Formula_Catalog.py        # Auditable LaTeX formulas & data quality diagnostics (UI-08)
├── src/                            # Modular calculation & data access packages
│   ├── __init__.py
│   ├── db.py                       # SQLite connection pool & read-only access manager
│   ├── data_loader.py              # Cached query extractors (@st.cache_data)
│   ├── data_cleaner.py             # Automated deduplication & median null imputation
│   ├── cac_calculator.py           # Blended/Channel CAC, LTV, and Payback mathematical logic
│   ├── stats_engine.py             # Shapiro-Wilk, One-Way ANOVA, and Kruskal-Wallis engine
│   ├── simulator.py                # Diminishing marginal ROAS logarithmic simulation curves
│   └── ui.py                       # Tailwind CSS components (Hero, KPI cards, Alerts)
├── data/                           # Relational persistence & exports
│   ├── app.db                      # Embedded SQLite relational database
│   ├── cac_analysis_data.xlsx      # Master 5-sheet dataset workbook
│   ├── dim_channels.csv            # Master channels CSV
│   ├── dim_subscription_tiers.csv  # Master pricing tiers CSV
│   ├── fact_marketing_campaigns.csv# Campaign spend & performance fact CSV
│   ├── fact_customer_cohorts.csv   # Customer subscription cohort fact CSV
│   └── view_cac_full_analysis.csv  # 34-column joined analytical view CSV
├── tests/                          # Automated regression test suites
│   ├── __init__.py
│   ├── test_data.py                # Schema, row counts, PK/FK, and cleaner tests (TC-T-01..05)
│   ├── test_analysis.py            # Unit tests for CAC, LTV, Payback, ANOVA (TC-T-06..10)
│   └── test_app.py                 # UI module, config, latency benchmarks (TC-T-11..14)
├── docs/                           # 18-step engineering & academic deliverables
├── .streamlit/
│   └── config.toml                 # Pinned Navy theme & server settings
├── BUILD_PLAN.md                   # Machine & human execution plan
├── requirements.txt                # Pinned production dependencies
└── README.md                       # Comprehensive operations runbook
```

---

## 3. Sequential Build Tasks (BUILD-01 to BUILD-14)

### BUILD-01: Environment Configuration & Pinned Dependencies
- **Goal:** Pin production library dependencies and baseline Streamlit server theming.
- **Files Involved:** `requirements.txt`, `.streamlit/config.toml`
- **Linked Requirements & User Stories:** NFR-01, US-01 | F-01
- **Exact Prompt for AI:**  
  *"Verify requirements.txt contains pinned versions for streamlit, pandas, numpy, scipy, statsmodels, plotly, openpyxl, pytest, and pytest-cov. Ensure .streamlit/config.toml has primaryColor='#1E3A8A', backgroundColor='#FFFFFF', and secondaryBackgroundColor='#F8FAFC'."*
- **How to Check:** Run `pip check` and inspect `.streamlit/config.toml`.
- **Tests That Must Pass:** `tests/test_app.py::test_tc_t_11_streamlit_config_existence`

---

### BUILD-02: Database Access, Caching & Data Cleaning Engine
- **Goal:** Build the data access layer connecting to `data/app.db` with `@st.cache_data` and automated data cleaning for nulls and duplicates.
- **Files Involved:** `src/db.py`, `src/data_loader.py`, `src/data_cleaner.py`
- **Linked Requirements & User Stories:** BR-01, FR-01 | US-01, US-03 | F-01
- **Exact Prompt for AI:**  
  *"Create src/db.py to provide read-only SQLite connections. Create src/data_cleaner.py to remove duplicate customer codes and impute null servicing_cost with median values. Create src/data_loader.py with @st.cache_data wrappers to query view_cac_full_analysis in <50ms."*
- **How to Check:** Execute Python interactive shell and assert `df = load_clean_data()` has 0 nulls and 0 duplicate customer IDs.
- **Tests That Must Pass:** `tests/test_data.py::test_tc_t_01`, `test_tc_t_04`, `test_tc_t_05`

---

### BUILD-03: Analytical Calculation & Statistical Modeling Engine
- **Goal:** Implement pure mathematical calculation modules for CAC, LTV, Payback Period, ANOVA, and marginal ROAS curves.
- **Files Involved:** `src/cac_calculator.py`, `src/stats_engine.py`, `src/simulator.py`
- **Linked Requirements & User Stories:** BR-02, BR-04, BR-05, BR-07, BR-08 | US-04, US-07, US-08, US-11, US-12 | F-02, F-05, F-06, F-08, F-09
- **Exact Prompt for AI:**  
  *"Create src/cac_calculator.py computing blended CAC, channel CAC, payback months, and LTV:CAC health ratings with zero-division guards. Create src/stats_engine.py computing Shapiro-Wilk, One-Way ANOVA, and Kruskal-Wallis tests returning F/H stats and p-values. Create src/simulator.py modeling budget reallocation on logarithmic ROAS curves."*
- **How to Check:** Execute unit tests against known static vectors.
- **Tests That Must Pass:** `tests/test_analysis.py::test_tc_t_06`, `test_tc_t_07`, `test_tc_t_08`, `test_tc_t_09`, `test_tc_t_10`

---

### BUILD-04: Reusable UI Component Primitives & Tailwind Integration
- **Goal:** Provide responsive Tailwind CSS hero banners, metric cards, and alert callouts rendered via `st.components.v1.html`.
- **Files Involved:** `src/ui.py`
- **Linked Requirements & User Stories:** NFR-02 | US-02 | F-02
- **Exact Prompt for AI:**  
  *"Enhance src/ui.py to support render_header(), render_kpi_cards() with positive/negative pill badges, and render_insight_box() with warning/critical alerts using Tailwind CDN."*
- **How to Check:** Inspect HTML output string and verify no script errors.
- **Tests That Must Pass:** `tests/test_app.py::test_tc_t_12_ui_components_module`

---

### BUILD-05: Page 1 — Executive Overview & Cockpit
- **Goal:** Build main landing page `app.py` with global sidebar filters, 4 top KPI cards, monthly CAC trendline, and channel spend distribution.
- **Files Involved:** `app.py`
- **Linked Requirements & User Stories:** BR-01, BR-02 | US-01, US-02, US-03 | F-02, F-03 | UI-01
- **Exact Prompt for AI:**  
  *"Implement app.py using Streamlit. Configure set_page_config(page_title='CAC Cockpit', layout='wide'). Add sidebar date-range and channel multiselect filters saving to st.session_state. Render Tailwind header, 4 KPI cards (Blended CAC, Total Spend, Customers, Avg Payback), a Plotly monthly CAC line chart, and a spend share donut chart."*
- **How to Check:** Run `streamlit run app.py` and inspect browser rendering at `http://localhost:8501`.
- **Tests That Must Pass:** `tests/test_app.py::test_tc_t_13`, `test_tc_t_14`, TC-F-01, TC-F-02

---

### BUILD-06: Page 2 — Channel CAC & Payback Trends
- **Goal:** Implement `pages/1_Channel_CAC.py` comparing channel CAC vs blended reference lines and flagging Q4 seasonal price surges.
- **Files Involved:** `pages/1_Channel_CAC.py`
- **Linked Requirements & User Stories:** BR-02, BR-04, BR-09 | US-04, US-07, US-13 | F-03, F-05, F-10 | UI-02
- **Exact Prompt for AI:**  
  *"Create pages/1_Channel_CAC.py displaying comparative bar charts of Channel CAC vs Blended CAC dashed line, a payback period status table highlighting Paid Social in bold red, and a monthly seasonal price index area chart flagging Q4 holiday surges."*
- **How to Check:** Verify visual alert renders on channels with payback $> 12$ months.
- **Tests That Must Pass:** TC-F-03, TC-F-06

---

### BUILD-07: Page 3 — Multi-Stage Funnel & Leakage Analyzer
- **Goal:** Implement `pages/2_Funnel_Analysis.py` rendering Plotly funnel chart with stage drop-off % and CPC/CPL cards.
- **Files Involved:** `pages/2_Funnel_Analysis.py`
- **Linked Requirements & User Stories:** BR-03 | US-05, US-06 | F-04 | UI-03
- **Exact Prompt for AI:**  
  *"Create pages/2_Funnel_Analysis.py rendering CPC and CPL cards, followed by a Plotly Funnel chart (Impressions -> Clicks -> Leads -> Customers) with stage conversion percentages and a callout highlighting checkout leakage in Paid Social."*
- **How to Check:** Change channel filter to 'Referral' and assert conversion rate changes from $<4\%$ to $>20\%$.
- **Tests That Must Pass:** TC-F-04, TC-F-05

---

### BUILD-08: Page 4 — Volume vs Efficiency Quadrant Matrix
- **Goal:** Implement `pages/3_Compare.py` plotting 2x2 scatter matrix of acquisition volume versus cost per conversion.
- **Files Involved:** `pages/3_Compare.py`
- **Linked Requirements & User Stories:** BR-10 | US-14 | F-11 | UI-04
- **Exact Prompt for AI:**  
  *"Create pages/3_Compare.py rendering a Plotly 2x2 quadrant scatter chart (X=Total Customers Acquired, Y=Average CAC). Add quadrant threshold lines splitting channels into Scale Winners, Niche Champions, High-Cost Scalers, and Underperformers."*
- **How to Check:** Verify all 5 channels plot in their expected quadrant zones.
- **Tests That Must Pass:** TC-F-03

---

### BUILD-09: Page 5 — Inferential Statistics & ANOVA Testing
- **Goal:** Implement `pages/4_Statistical_Testing.py` executing Shapiro-Wilk, One-Way ANOVA, and Kruskal-Wallis with narrative reporting.
- **Files Involved:** `pages/4_Statistical_Testing.py`
- **Linked Requirements & User Stories:** BR-07 | US-11 | F-08 | UI-05
- **Exact Prompt for AI:**  
  *"Create pages/4_Statistical_Testing.py with an interactive significance slider (default alpha=0.05). Execute Shapiro-Wilk, One-Way ANOVA, and Kruskal-Wallis on campaign CAC distributions. Output F/H stats, p-values, a Plotly boxplot with outlier whiskers, and formal narrative conclusions."*
- **How to Check:** Run statistical engine and verify $p < 0.001$ rejects null hypothesis.
- **Tests That Must Pass:** TC-F-09, `tests/test_analysis.py::test_tc_t_09`

---

### BUILD-10: Page 6 — Unit Economics & Budget Reallocation Sandbox
- **Goal:** Implement `pages/5_Budget_Simulator.py` featuring tiered LTV:CAC matrix, churn sensitivity slider, and marginal ROAS curves.
- **Files Involved:** `pages/5_Budget_Simulator.py`
- **Linked Requirements & User Stories:** BR-05, BR-06, BR-08, BR-12 | US-08, US-09, US-10, US-12 | F-06, F-07, F-09, F-13 | UI-06
- **Exact Prompt for AI:**  
  *"Create pages/5_Budget_Simulator.py displaying the LTV:CAC health matrix across Basic, Pro, SMB, and Enterprise tiers. Add a churn sensitivity slider dynamically recalculating LTV. Add budget reallocation sliders simulating shift of ad budget from Social to Referral, displaying projected customer growth and marginal ROAS curves."*
- **How to Check:** Move churn slider from 0% to +20% and verify LTV updates reactively.
- **Tests That Must Pass:** TC-F-07, TC-F-08, TC-F-10

---

### BUILD-11: Page 7 — Data Explorer & Audit Downloads
- **Goal:** Implement `pages/6_Data_Explorer.py` providing interactive data table filtering and instant CSV/Excel export buttons.
- **Files Involved:** `pages/6_Data_Explorer.py`
- **Linked Requirements & User Stories:** BR-01 | US-01, US-15 | UI-07
- **Exact Prompt for AI:**  
  *"Create pages/6_Data_Explorer.py allowing users to select tables (dim_channels, dim_subscription_tiers, fact_marketing_campaigns, fact_customer_cohorts, view_cac_full_analysis). Render interactive searchable tables and Streamlit download buttons for CSV and Excel formats."*
- **How to Check:** Click download button and verify generated Excel workbook opens cleanly.
- **Tests That Must Pass:** TC-F-12

---

### BUILD-12: Page 8 — Formula Catalog & Data Quality Diagnostics
- **Goal:** Implement `pages/7_Formula_Catalog.py` rendering LaTeX KaTeX formula documentation and data quality audit logs.
- **Files Involved:** `pages/7_Formula_Catalog.py`
- **Linked Requirements & User Stories:** BR-11 | US-15 | F-12 | UI-08
- **Exact Prompt for AI:**  
  *"Create pages/7_Formula_Catalog.py rendering mathematical formula definitions using st.latex for all 12 metrics. Include an automated Data Quality Diagnostics card reporting null counts, duplicates removed, and outlier boundaries, alongside pytest pass status badges."*
- **How to Check:** Verify all LaTeX formulas render KaTeX mathematical equations cleanly without syntax warnings.
- **Tests That Must Pass:** `tests/test_app.py::test_tc_t_12`

---

### BUILD-13: Comprehensive QA Regression & Test Execution
- **Goal:** Execute full automated test suite using pytest and verify 100% test pass rate with coverage reporting.
- **Files Involved:** `tests/test_data.py`, `tests/test_analysis.py`, `tests/test_app.py`
- **Linked Requirements & User Stories:** NFR-03 | All User Stories (US-01 to US-15)
- **Exact Prompt for AI:**  
  *"Run python -m pytest tests/ -v --tb=short. Verify all 14 unit and integration tests pass with 0 failures. Verify query execution on view_cac_full_analysis completes in under 15ms."*
- **How to Check:** Output must display `14 passed in <20s`.
- **Tests That Must Pass:** All 14 technical tests (`TC-T-01` through `TC-T-14`).

---

### BUILD-14: Operations Runbook & Cloud Deployment
- **Goal:** Create professional `README.md` with step-by-step setup instructions and deploy app to Streamlit Community Cloud.
- **Files Involved:** `README.md`, Git commit & push to GitHub
- **Linked Requirements & User Stories:** NFR-04 | US-15
- **Exact Prompt for AI:**  
  *"Create a comprehensive README.md with project overview, architecture diagram, local installation commands, pytest instructions, and Streamlit Community Cloud deployment steps."*
- **How to Check:** Verify README renders cleanly on GitHub and public deployment link is live.
- **Tests That Must Pass:** Public URL returns HTTP 200.

---

## 4. Definition of Done (DoD)

The project is officially complete when all of the following criteria are met:
1. **Business Requirements Coverage:** 100% of Must-Have and Should-Have Business Requirements (`BR-01` to `BR-12`) are implemented.
2. **Statistical Validity:** One-Way ANOVA and Kruskal-Wallis engines execute accurately with statistically defensible conclusions at $\alpha=0.05$.
3. **Automated Testing:** 100% of automated `pytest` suites (`TC-T-01` to `TC-T-14`) pass with zero regressions.
4. **Performance & Memory:** Application cold start $< 3.0\text{ seconds}$, cached query latency $< 50\text{ms}$, and memory usage $< 400\text{ MB}$ (within 1 GB cloud envelope).
5. **UI & Cross-Device Compatibility:** Responsive layout verified across desktop and mobile screen viewports with zero text truncation or clipping.
6. **Production Documentation:** Complete ABES Major Project Report Word document and defense presentation slide deck ready.
7. **Cloud Live Status:** App publicly accessible on Streamlit Community Cloud with green health status.

---

## 5. Development CLI Commands

```bash
# 1. Clone repository and navigate to root
git clone <repo-url>
cd MERN-Expense-Tracker-main

# 2. Create and activate Python virtual environment
python -m venv venv
# Windows:
.\venv\Scripts\activate
# Linux/Mac:
source venv/bin/activate

# 3. Install pinned dependencies
pip install -r requirements.txt

# 4. Generate synthetic dataset and SQLite database
python src/generate_data.py

# 5. Run full automated test suite
python -m pytest tests/ -v

# 6. Launch local Streamlit application
streamlit run app.py
```

---

## 6. Viva Defense Notes: 10 Likely Questions & Authoritative Answers

1. **Q: What is the business rationale for isolating Channel CAC instead of tracking only Blended CAC?**  
   *Answer:* Blended CAC averages total marketing spend across all new customers, allowing low-cost organic traffic to mask unprofitable paid ad campaigns. Isolating channel-specific CAC reveals true unit economics and prevents capital misallocation.
2. **Q: Why is the CAC Payback Period critical for financial governance, and what is the benchmark threshold?**  
   *Answer:* Payback Period measures the months required for customer gross profit to recoup initial acquisition outlays. In SaaS and digital commerce, a payback under 12 months is essential to prevent working capital depletion.
3. **Q: Why is an LTV:CAC ratio below 3.0 considered sub-optimal even if it exceeds 1.0?**  
   *Answer:* An LTV:CAC $> 1.0$ only recovers direct acquisition spend. It leaves insufficient margin to cover ongoing research and development, hosting infrastructure, customer support, and administrative overhead.
4. **Q: What statistical test was used to prove that channel acquisition costs differ significantly?**  
   *Answer:* We executed One-Way ANOVA ($F$-test) for parametric variance testing and verified results with the non-parametric Kruskal-Wallis $H$-test to ensure robustness against skewed cost distributions at $\alpha=0.05$.
5. **Q: Why was a Log-normal distribution chosen over a Normal distribution for synthetic ad spend generation?**  
   *Answer:* Advertising costs in commercial business are strictly non-negative and right-skewed, characterized by a baseline spend with occasional promotional spikes. Normal Gaussian curves allow negative values, which are physically impossible.
6. **Q: How does `@st.cache_data` prevent the app from exhausting Streamlit Community Cloud's 1 GB RAM limit?**  
   *Answer:* `@st.cache_data` stores serialized query results in an in-memory Apache Arrow cache. Successive user filter interactions hit memory in microseconds without querying the database or duplicating DataFrame memory allocations.
7. **Q: How did you embed Tailwind CSS styling inside a native Streamlit application?**  
   *Answer:* Custom KPI cards and hero headers are authored in semantic HTML5 styled with Tailwind CSS via CDN and rendered inside `st.components.v1.html(..., height=...)`, providing full CSS flexbox responsiveness.
8. **Q: How does the platform guard against division-by-zero errors when calculating conversion rates?**  
   *Answer:* Vectorized calculations use `np.where(denominator > 0, numerator / denominator, 0.0)`, guaranteeing that zero clicks or zero conversions return `0.0` rather than throwing fatal runtime exceptions.
9. **Q: What is the purpose of the 2x2 Volume vs Cost-Efficiency Quadrant Matrix (UI-04)?**  
   *Answer:* It categorizes acquisition channels into four strategic quadrants: Scale Winners (low CAC, high volume), Niche Champions (low CAC, low volume), High-Cost Scalers, and Underperformers, guiding capital allocation.
10. **Q: How does the Requirements Traceability Matrix (RTM) prove the academic and architectural integrity of your project?**  
    *Answer:* The RTM provides bidirectional forward and backward traceability, proving that every initial business operational question (OQ) maps directly to an approved user story, database entity, UI screen, and automated test case.

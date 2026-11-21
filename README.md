# 📊 Customer Acquisition Cost (CAC) & Marketing Unit Economics Analysis

> **An Enterprise Analytics, Unit Economics Modeling & Statistical Validation Platform**  
> *Developed for Major Capstone Project — MCA 2nd Year (Data Analytics), ABES Engineering College, Ghaziabad.*

[![Python 3.10+](https://img.shields.io/badge/python-3.10%20%7C%203.12-blue.svg)](https://www.python.org/)
[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://monthly-budget-project-hgcyxbnvwnvyx8g4ngpnax.streamlit.app/)
[![Tests Passing](https://img.shields.io/badge/pytest-14%2F14%20passed-brightgreen.svg)](https://github.com/gkumar8272/Monthly-Budget-Project)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

---

## 👨‍🎓 Project Metadata
- **Candidate Name:** Gourav Kumar
- **Roll Number:** 2500320140055
- **Degree:** Master of Computer Applications (MCA) — Data Analytics (2nd Year)
- **Institution:** ABES Engineering College, Ghaziabad | Department of Computer Applications
- **Faculty Mentor:** Prof. Alpna Lodhi (Assistant Professor)
- **Live Cloud Dashboard:** [https://monthly-budget-project-hgcyxbnvwnvyx8g4ngpnax.streamlit.app/](https://monthly-budget-project-hgcyxbnvwnvyx8g4ngpnax.streamlit.app/)
- **GitHub Repository:** [https://github.com/gkumar8272/Monthly-Budget-Project](https://github.com/gkumar8272/Monthly-Budget-Project)

---

## 📌 Executive Summary & Business Problem
In digital marketing and B2B SaaS growth models, relying strictly on **blended Customer Acquisition Cost (CAC)** conceals underlying channel inefficiencies:
- Paid digital auction channels suffer from auction inflation (CAC exceeding Customer Lifetime Value).
- Multi-stage conversion funnels exhibit undetected drop-offs between click and customer conversion.
- Capital payback cycles exceed 18 months without solvency alarms.

This repository implements an end-to-end analytics platform that decomposes blended CAC into granular channel-level economics, models payback velocity against a 12-month solvency threshold, validates channel cost divergence via inferential hypothesis testing (One-Way ANOVA and Kruskal-Wallis, $p = 0.0023$), and empowers marketing leadership with a diminishing marginal returns budget sandbox.

---

## 🛠️ Technology Stack & Architecture
- **Language & Runtime:** Python 3.10 / 3.12
- **Data Layer:** Normalized 3NF SQLite database (`data/app.db`) with 34-column pre-compiled analytical view (`view_cac_full_analysis`)
- **Data Manipulation:** Pandas 2.2+, NumPy 2.2+ (vectorized zero-division guarded math)
- **Statistical Engine:** SciPy 1.13+, Statsmodels 0.15+ (Shapiro-Wilk normality, One-Way ANOVA $F$-test, Kruskal-Wallis $H$-test)
- **Visualizations:** Plotly Express & Graph Objects (interactive dual-axis trendlines, funnels, 2x2 quadrant matrices)
- **Presentation Framework:** Streamlit Multipage architecture with Tailwind CSS CDN in sandboxed iframes
- **Quality Assurance:** 14 automated unit & integration tests (`pytest 9.2+`, 100% pass rate in < 3s)
- **Business Intelligence Guides:** Full Tableau Desktop & Microsoft Power BI DAX implementation catalogs (`docs/`)

---

## 📂 Project Structure

```text
├── app.py                          # Executive Cockpit (UI-01): Tailwind KPI cards & run-rate alerts
├── pages/                          # Multipage analytical routing
│   ├── 1_Channel_CAC.py            # Channel vs Blended CAC trendlines & 12m payback benchmark (UI-02)
│   ├── 2_Funnel_Analysis.py        # Conversion funnel & stage drop-off attribution (UI-03)
│   ├── 3_Compare.py                # Volume vs Efficiency 2x2 scatter matrix (UI-04)
│   ├── 4_Statistical_Testing.py    # One-Way ANOVA & Kruskal-Wallis hypothesis tests (UI-05)
│   ├── 5_Budget_Simulator.py       # Diminishing marginal ROAS budget sandbox (UI-06)
│   ├── 6_Data_Explorer.py          # Paginated table browser & CSV/Excel export (UI-07)
│   └── 7_Formula_Catalog.py        # Auditable KaTeX formulas & data quality diagnostics (UI-08)
├── src/                            # Modular calculation & data engine
│   ├── db.py                       # Thread-safe SQLite connection pool manager
│   ├── data_loader.py              # In-memory cached queries (@st.cache_data)
│   ├── data_cleaner.py             # Deduplication and median imputation pipeline
│   ├── cac_calculator.py           # Vectorized CAC, LTV, Payback, and health tier logic
│   ├── stats_engine.py             # Inferential statistics & normality testing engine
│   ├── simulator.py                # Logarithmic response curve budget reallocation simulator
│   └── ui.py                       # Tailwind CSS semantic iframe injection components
├── data/                           # Data artifacts
│   ├── app.db                      # Production SQLite database (3NF schema + pre-compiled view)
│   ├── cac_analysis_data.xlsx      # Master multi-sheet Excel dataset
│   ├── cac_eda_dataset.csv         # 558-row dataset for EDA, Tableau, and Power BI
│   └── *.csv                       # Exported raw table CSVs
├── docs/                           # Comprehensive 18-step SDLC deliverables
│   ├── 2500320140055_Project_Document.docx  # Official ABES Capstone Project Document
│   ├── 2500320140055_Presentation.pptx      # Official 11-Slide Defense Presentation
│   ├── 04_Traceability_Matrix.xlsx          # End-to-end forward/backward RTM
│   ├── 15_Build_Verification_Report.docx    # 100% pytest execution log
│   ├── 16_Deployment_Runbook.docx           # Streamlit Community Cloud operations runbook
│   └── Tableau_and_PowerBI_Dashboard_Guide.docx # BI calculated fields & DAX catalog
├── tests/                          # Automated Pytest suite
│   ├── test_data.py                # Database connectivity, PK uniqueness & FK integrity (TC-T-01..05)
│   ├── test_analysis.py            # Mathematical formula parity & ANOVA assertions (TC-T-06..10)
│   └── test_app.py                 # Streamlit UI configuration, latency & importability (TC-T-11..14)
├── EDA_CAC_Analysis_Lesson.ipynb   # 13-section pedagogical EDA masterclass Jupyter notebook
├── requirements.txt                # Pinned production dependencies
└── .streamlit/config.toml          # Executive corporate theme configuration
```

---

## ⚡ Quick Start & Local Execution

### 1. Clone & Set Up Environment
```bash
git clone https://github.com/gkumar8272/Monthly-Budget-Project.git
cd Monthly-Budget-Project

# Create virtual environment
python -m venv .venv
# Activate on Windows:
.venv\Scripts\activate
# Activate on Linux/macOS:
# source .venv/bin/activate

# Install pinned dependencies
pip install -r requirements.txt
```

### 2. Run Automated Pytest Suite
```bash
python -m pytest tests/ -v
# Output: 14 passed in ~2.7s (100% pass rate)
```

### 3. Launch the Streamlit Web Application
```bash
streamlit run app.py
```
Open [http://localhost:8501](http://localhost:8501) in your browser.

### 4. Open the EDA Masterclass Jupyter Notebook
```bash
python -m jupyter notebook EDA_CAC_Analysis_Lesson.ipynb
```

---

## 📊 Key Quantitative Findings
1. **Organic Search Outperformance:** Delivers customers at **₹420 CAC** (59% below blended average) with an **LTV:CAC of 4.8x** (classified as *Healthy*).
2. **Paid Social Inflation:** Exhibits extreme volatility with CAC surging to **₹2,400+ in Q4** (Diwali auction price wars).
3. **Statistical Significance:** One-Way ANOVA ($F = 11.73, p = 0.0023$) and Kruskal-Wallis ($H = 18.44, p = 0.0010$) reject the null hypothesis, mathematically proving channel cost divergence is real.
4. **Solvency Threshold:** Blended cash payback recovers within **10.0 months**, safely inside the enterprise 12-month solvency ceiling.

---

## 📄 License
This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.
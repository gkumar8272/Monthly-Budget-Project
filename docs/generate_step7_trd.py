import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=130, right=130):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def set_table_borders(table, color="CBD5E1"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="6" w:space="0" w:color="{color}"/>
            <w:bottom w:val="single" w:sz="6" w:space="0" w:color="{color}"/>
            <w:left w:val="none"/>
            <w:right w:val="none"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color}"/>
            <w:insideV w:val="none"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def build_trd_doc():
    doc = Document()

    for s in doc.sections:
        s.top_margin = Inches(1)
        s.bottom_margin = Inches(1)
        s.left_margin = Inches(1)
        s.right_margin = Inches(1)

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)

    # Document Header
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("CUSTOMER ACQUISITION COST (CAC) ANALYSIS")
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("STEP 7: TECHNICAL REQUIREMENTS DOCUMENT (TRD / TBD)\nSystem Architecture, Engineering Blueprints & Software Specifications")
    r_sub.font.size = Pt(12)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    def add_h1(text):
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(14)
        h.paragraph_format.space_after = Pt(4)
        h.paragraph_format.keep_with_next = True
        r = h.add_run(text)
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
        return h

    def add_h2(text):
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(10)
        h.paragraph_format.space_after = Pt(3)
        h.paragraph_format.keep_with_next = True
        r = h.add_run(text)
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
        return h

    meta = [
        ("Project ID", "P083"),
        ("Candidate Name", "Gourav Kumar (Roll: 2500320140055)"),
        ("Batch ID", "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"),
        ("Academic Program", "MCA 2nd Year (Data Analytics), ABES Engineering College"),
        ("Faculty Supervisor", "Mr. Dinesh Sereno"),
        ("Document Identifier", "P083-DOC-STEP-07 | Technical Requirements Document (TRD)"),
        ("Target Technology Stack", "Python 3.12, Pandas, NumPy, SciPy, Statsmodels, Plotly, Streamlit, SQLite, Tailwind CDN, Pytest"),
        ("Cloud Platform", "Streamlit Community Cloud (via GitHub repository continuous deployment)")
    ]

    # Metadata Table
    meta_table = doc.add_table(rows=len(meta), cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(meta_table, "CBD5E1")
    for i, (k, v) in enumerate(meta):
        row = meta_table.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = Inches(2.2), Inches(4.3)
        set_cell_background(c0, "F1F5F9")
        set_cell_margins(c0, 70, 70, 100, 100)
        set_cell_margins(c1, 70, 70, 100, 100)
        r0 = c0.paragraphs[0].add_run(k)
        r0.font.bold = True
        r0.font.size = Pt(9)
        r1 = c1.paragraphs[0].add_run(v)
        r1.font.size = Pt(9)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 1. Architecture Pipeline
    add_h1("1. System Architecture & Data Processing Pipeline")
    doc.add_paragraph(
        "The application utilizes a clean, decoupled five-tier pipeline architecture designed for high analytical throughput, "
        "memory safety on resource-constrained cloud environments (1 GB RAM limit), and complete modular separation of concerns:"
    )
    doc.add_paragraph(
        "1. Persistence Layer: Embedded relational SQLite database (`data/app.db`) storing indexed marketing spend, funnel leads, and customer subscription cohorts.\n"
        "2. Ingestion & Caching Layer: High-performance SQL query wrappers in `src/db.py` and `src/data_loader.py` decorated with `@st.cache_data(ttl=3600)` to eliminate redundant disk I/O.\n"
        "3. Validation & Cleaning Layer: Schema assertion and type boundary validations in `src/data_cleaner.py` ensuring zero division errors and clean null handling.\n"
        "4. Business Calculation & Stats Engine: Dedicated functional computation modules (`src/cac_calculator.py`, `src/stats_engine.py`, `src/simulator.py`) computing unit economics, ANOVA, and marginal ROAS curves.\n"
        "5. Presentation & UI Layer: Multi-page Streamlit application (`app.py`, `pages/`) rendering interactive Plotly charts and Tailwind CSS KPI cards embedded via `st.components.v1.html`."
    )
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # 2. Folder Structure
    add_h1("2. Production Codebase Folder Structure")
    doc.add_paragraph("The repository layout adheres strictly to professional Python data engineering and Streamlit best practices:")
    
    code_box = doc.add_table(rows=1, cols=1)
    code_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_cell = code_box.rows[0].cells[0]
    c_cell.width = Inches(6.5)
    set_cell_background(c_cell, "F8FAFC")
    set_cell_margins(c_cell, 100, 100, 120, 120)
    cp = c_cell.paragraphs[0]
    struct_text = (
        "MERN-Expense-Tracker-main/\n"
        "├── app.py                      # Application entrypoint: Tailwind hero, global KPI summary\n"
        "├── pages/                      # Multi-page routing views\n"
        "│   ├── 1_Channel_CAC.py        # Blended vs Channel CAC, volume-efficiency matrix\n"
        "│   ├── 2_Funnel_Analysis.py    # Conversion stage drop-offs, CPC/CPL leakage\n"
        "│   ├── 3_LTV_Payback.py        # LTV:CAC ratios, 12-month payback horizon, churn slider\n"
        "│   ├── 4_Statistical_Testing.py# ANOVA, Kruskal-Wallis, distribution diagnostics\n"
        "│   ├── 5_Budget_Simulator.py   # Interactive marginal ROAS allocation simulator\n"
        "│   └── 6_Formula_Catalog.py    # Auditable mathematical catalog & lineage dictionary\n"
        "├── src/                        # Modular backend computational packages\n"
        "│   ├── __init__.py\n"
        "│   ├── db.py                   # SQLite connection manager, query executor\n"
        "│   ├── data_loader.py          # Cached data extractors (@st.cache_data)\n"
        "│   ├── data_cleaner.py         # Data normalization & type integrity guards\n"
        "│   ├── cac_calculator.py       # CAC, LTV, Payback, and Funnel math engines\n"
        "│   ├── stats_engine.py         # Inferential hypothesis testing (ANOVA/Kruskal)\n"
        "│   ├── simulator.py            # Diminishing marginal ROAS simulation curves\n"
        "│   └── ui_components.py        # Tailwind HTML cards & Plotly chart wrappers\n"
        "├── data/                       # Persistence assets\n"
        "│   ├── app.db                  # Embedded SQLite relational database\n"
        "│   ├── seed_data.py            # Synthetic dataset generator script (Step 10)\n"
        "│   └── schema.sql              # 3NF relational schema DDL\n"
        "├── tests/                      # Automated pytest test suites\n"
        "│   ├── __init__.py\n"
        "│   ├── test_cac_math.py        # Unit tests for CAC, LTV, and payback calculations\n"
        "│   ├── test_stats_engine.py    # Unit tests for statistical models and p-value checks\n"
        "│   ├── test_db_queries.py      # SQLite join integrity and view assertion tests\n"
        "│   └── test_app_ui.py          # Streamlit AppTest integration suites\n"
        "├── docs/                       # Formal academic and enterprise deliverables\n"
        "├── .streamlit/config.toml      # Theming, layout, and port settings\n"
        "├── requirements.txt            # Pinned package dependencies\n"
        "└── README.md                   # Setup guide and deployment runbook"
    )
    cr = cp.add_run(struct_text)
    cr.font.name = 'Consolas'
    cr.font.size = Pt(8.5)
    cr.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 3. Functional Requirements Matrix
    add_h1("3. Functional Requirements Specification (FR-01 to FR-12)")
    p = doc.add_paragraph("Technical mapping between Functional Requirements (FRs), upstream Business Requirements (BRs), and core Python modules:")
    p.paragraph_format.space_after = Pt(4)

    fr_table = doc.add_table(rows=13, cols=5)
    fr_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(fr_table, "CBD5E1")
    fr_headers = ["FR ID", "Functional Requirement Description", "Linked BR", "Target Python Module & Function", "Expected Output"]
    fr_widths = [Inches(0.7), Inches(2.2), Inches(0.7), Inches(1.8), Inches(1.1)]

    for j, h in enumerate(fr_headers):
        c = fr_table.rows[0].cells[j]
        c.width = fr_widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 80, 80, 70, 70)
        r = c.paragraphs[0].add_run(h)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    fr_data = [
        ("FR-01", "Ingest multi-channel marketing spend and enforce channel foreign key integrity.", "BR-01", "src/data_loader.py:load_campaign_spend", "Validated DataFrame of channel ad spends"),
        ("FR-02", "Aggregate monthly channel-specific and blended CAC metrics with 0% delta.", "BR-02", "src/cac_calculator.py:compute_cac_metrics", "Summary KPI dict + monthly CAC trends"),
        ("FR-03", "Track multi-stage funnel attrition and compute CPC, CPL, and conversion %.", "BR-03", "src/cac_calculator.py:compute_funnel_metrics", "Funnel milestone conversion metrics"),
        ("FR-04", "Model capital payback duration in months and trigger 12-month threshold alerts.", "BR-04", "src/cac_calculator.py:compute_payback_period", "Payback months per channel + alert flags"),
        ("FR-05", "Compute LTV:CAC ratios across customer subscription tiers with <3.0 risk flags.", "BR-05", "src/cac_calculator.py:compute_ltv_cac_ratio", "Tiered LTV:CAC matrix & health status"),
        ("FR-06", "Calculate customer segment net contribution margin and recovery velocity.", "BR-06", "src/cac_calculator.py:compute_segment_profitability", "Segment contribution table & payback bar"),
        ("FR-07", "Execute Shapiro-Wilk, One-Way ANOVA, and Kruskal-Wallis hypothesis tests.", "BR-07", "src/stats_engine.py:run_channel_anova_kruskal", "F/H statistics, p-values, decision strings"),
        ("FR-08", "Simulate marginal ROAS and customer acquisition under budget adjustments (+/-50%).", "BR-08", "src/simulator.py:simulate_budget_reallocation", "Simulated revenue curves & ROAS deltas"),
        ("FR-09", "Generate 24-month longitudinal seasonal indices and detect quarterly ad price surges.", "BR-09", "src/analytics.py:calculate_seasonality_index", "Monthly seasonality index & surge flags"),
        ("FR-10", "Render 2x2 interactive scatter matrix evaluating volume vs. cost efficiency.", "BR-10", "src/ui_components.py:plot_volume_efficiency_matrix", "Interactive Plotly scatter quadrant chart"),
        ("FR-11", "Expose auditable metric catalog with LaTeX mathematical expressions and lineage.", "BR-11", "pages/6_Formula_Catalog.py", "Searchable formula documentation UI"),
        ("FR-12", "Provide dynamic churn sensitivity slider dynamically updating LTV:CAC in real time.", "BR-12", "src/cac_calculator.py:recompute_ltv_sensitivity", "Real-time reactive LTV recalculation")
    ]
    for i, row in enumerate(fr_data, start=1):
        for j, text in enumerate(row):
            c = fr_table.rows[i].cells[j]
            c.width = fr_widths[j]
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            set_cell_margins(c, 60, 60, 70, 70)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 4. Data Layer & Caching
    add_h1("4. Data Access Layer & Caching Strategy")
    doc.add_paragraph(
        "To ensure sub-second response times on Streamlit Community Cloud without memory bloat, data access follows these rules:\n"
        "• Connection Management: SQLite connections are instantiated via context managers (`with sqlite3.connect(...) as conn:`) in read-only URI mode (`file:data/app.db?mode=ro`).\n"
        "• In-Memory Caching: All query execution functions are decorated with `@st.cache_data(ttl=3600, show_spinner=False)`. DataFrames are cached in memory as immutable Arrow tables, eliminating repeated disk queries during user UI interactions.\n"
        "• Cache Invalidation: Session state parameter filters pass discrete argument hashes to cached functions; when user modifies date ranges or channel selections, Streamlit selectively recalculates only the affected subsets.\n"
        "• Data Validation Guards: All numeric columns are validated for non-negativity; string categoricals are asserted against white-listed channel and tier sets."
    )
    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 5. Statistical Methods for OQs
    add_h1("5. Statistical Methodology Mapped to Operational Questions")
    p = doc.add_paragraph("Rigorous mathematical formulations and statistical tests addressing each Operational Question (OQ-01 through OQ-09):")
    p.paragraph_format.space_after = Pt(4)

    stat_table = doc.add_table(rows=10, cols=4)
    stat_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(stat_table, "CBD5E1")
    stat_headers = ["OQ ID", "Operational Question", "Mathematical Formulation / Statistical Test", "Analytical Package / Function"]
    stat_widths = [Inches(0.8), Inches(2.1), Inches(2.2), Inches(1.4)]

    for j, h in enumerate(stat_headers):
        c = stat_table.rows[0].cells[j]
        c.width = stat_widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 80, 80, 70, 70)
        r = c.paragraphs[0].add_run(h)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    stat_data = [
        ("OQ-01", "Blended Monthly CAC", "CAC_blended = Total_Spend_m / Total_New_Customers_m", "pandas: df.groupby('month').agg()"),
        ("OQ-02", "Channel CAC vs Conversion", "CAC_c = Spend_c / Conv_c; CR_c = Conv_c / Leads_c", "numpy / pandas vector division"),
        ("OQ-03", "LTV:CAC by Customer Tier", "LTV = (ARPU * Gross_Margin) / Churn_Rate; Ratio = LTV / CAC", "src/cac_calculator.py"),
        ("OQ-04", "Payback Period in Months", "Payback_Months = CAC_c / (ARPU_c * Gross_Margin_%)", "numpy.where() with 12m alert"),
        ("OQ-05", "Channel Cost Significance", "Shapiro-Wilk normality; One-Way ANOVA F-test; Kruskal-Wallis H-test", "scipy.stats.shapiro, f_oneway, kruskal"),
        ("OQ-06", "Seasonal Price Surges", "Seasonality_Index_m = CAC_m / (Rolling_Mean_12m); MoM growth %", "pandas: rolling(12).mean(), pct_change()"),
        ("OQ-07", "Funnel Drop-Off Leakage", "DropOff_k = 1 - (Stage_{k+1} / Stage_k); Cost_Per_Stage_k", "src/cac_calculator.py funnel logic"),
        ("OQ-08", "Marginal ROAS Curves", "ROAS_marginal = d(Revenue) / d(Spend) via logarithmic saturation model", "scipy.optimize / numpy.polyfit"),
        ("OQ-09", "Segment Profitability", "Net_Margin = Revenue_tier - (CAC_tier + Direct_Servicing_Cost)", "pandas cohort aggregation")
    ]
    for i, row in enumerate(stat_data, start=1):
        for j, text in enumerate(row):
            c = stat_table.rows[i].cells[j]
            c.width = stat_widths[j]
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            set_cell_margins(c, 60, 60, 70, 70)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 6. UI Notes: Tailwind CDN
    add_h1("6. User Interface Architecture & Tailwind CSS Integration")
    doc.add_paragraph(
        "• Tailwind CDN Injection: Standard Streamlit widgets have limited visual customization. To achieve an executive, enterprise-grade aesthetic, custom KPI summary cards and hero banners are authored in pure HTML5 styled with Tailwind CSS (loaded via CDN) and rendered directly using `st.components.v1.html(..., height=...)`.\n"
        "• Theme Palette: Follows an executive Slate/Navy color tokens scheme:\n"
        "  - Primary Brand Navy: `#1E3A8A` (Tailwind `blue-900`)\n"
        "  - Accent Vibrant Blue: `#2563EB` (Tailwind `blue-600`)\n"
        "  - Card Background: `#FFFFFF` on `#F8FAFC` (Tailwind `slate-50`)\n"
        "  - Alert Red: `#DC2626` (Tailwind `red-600`); Success Green: `#16A34A` (Tailwind `green-600`)\n"
        "• Responsive Layout: CSS flexbox grids dynamically adjust between 1-column mobile layouts and 3-to-4 column desktop dashboard views."
    )
    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 7. Performance, Error Handling & Security
    add_h1("7. Performance Optimization, Error Handling & Security")
    doc.add_paragraph(
        "• Query Optimization: SQLite database columns frequently filtered or joined (`channel_id`, `created_date`, `plan_tier`) are explicitly indexed using B-tree indices (`CREATE INDEX idx_...`).\n"
        "• Graceful Degradation: All analytical and statistical calculations are wrapped in defensive `try...except` blocks. If sample size in a filtered subset is insufficient for ANOVA ($n < 3$), the system gracefully displays an informative `st.warning()` callout rather than crashing with an unhandled exception.\n"
        "• Security & SQL Injection Prevention: All database queries utilize parameterized SQL placeholders (`?`) preventing SQL injection. The SQLite connection is opened in strict read-only mode during web runtime."
    )
    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 8. Pinned requirements.txt
    add_h1("8. Pinned Software Dependencies (requirements.txt)")
    doc.add_paragraph("Strictly pinned production dependencies ensuring reproducible build environments:")
    
    req_box = doc.add_table(rows=1, cols=1)
    req_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    rc = req_box.rows[0].cells[0]
    rc.width = Inches(6.5)
    set_cell_background(rc, "F8FAFC")
    set_cell_margins(rc, 80, 80, 100, 100)
    rp = rc.paragraphs[0]
    reqs_text = (
        "streamlit==1.32.0\n"
        "pandas==2.2.1\n"
        "numpy==1.26.4\n"
        "scipy==1.12.0\n"
        "statsmodels==0.14.1\n"
        "plotly==5.20.0\n"
        "openpyxl==3.1.2\n"
        "pytest==8.1.1\n"
        "pytest-cov==5.0.0"
    )
    rr = rp.add_run(reqs_text)
    rr.font.name = 'Consolas'
    rr.font.size = Pt(9)
    rr.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 9. Deployment Requirements
    add_h1("9. Deployment Requirements (Streamlit Community Cloud)")
    doc.add_paragraph(
        "• Continuous Deployment: Repository connected to Streamlit Community Cloud with auto-deploy tracking the `main` branch on GitHub.\n"
        "• Hardware Constraints: Execution budget operates strictly within 1 GB RAM and shared vCPU limits. Baseline in-memory footprint is constrained to ~220 MB.\n"
        "• Configuration: `.streamlit/config.toml` configures `headless = true`, `enableCORS = false`, and server port bindings."
    )
    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 10. Technical Risks & Mitigations
    add_h1("10. Technical Risks & Architectural Mitigations")
    trsk_table = doc.add_table(rows=5, cols=4)
    trsk_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(trsk_table, "CBD5E1")
    trsk_headers = ["Risk ID", "Technical Risk Description", "Impact Severity", "Architectural Mitigation"]
    trsk_widths = [Inches(0.9), Inches(2.4), Inches(1.1), Inches(2.1)]

    for j, h in enumerate(trsk_headers):
        c = trsk_table.rows[0].cells[j]
        c.width = trsk_widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 80, 80, 70, 70)
        r = c.paragraphs[0].add_run(h)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    trsk_data = [
        ("TRSK-01", "Memory exhaustion during continuous user filtering on Streamlit Cloud.", "High", "Enforce @st.cache_data with maximum 1-hour TTL; downsample high-cardinality charts."),
        ("TRSK-02", "Concurrent database write contention or lock errors in SQLite.", "Medium", "Keep SQLite in read-only mode during web app runtime; pre-seed database statically."),
        ("TRSK-03", "Statistical tests fail due to extreme outliers or zero-variance slices.", "Medium", "Implement variance pre-checks; conditionally route to Kruskal-Wallis if normality fails."),
        ("TRSK-04", "Streamlit UI rerun lag caused by heavy Plotly JSON serialization.", "Medium", "Pre-aggregate DataFrames in Python before passing to Plotly express figure builders.")
    ]
    for i, row in enumerate(trsk_data, start=1):
        for j, text in enumerate(row):
            c = trsk_table.rows[i].cells[j]
            c.width = trsk_widths[j]
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            set_cell_margins(c, 60, 60, 70, 70)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True

    target_path = os.path.abspath("docs/07_TRD.docx")
    doc.save(target_path)
    print(f"Generated: {target_path}")

if __name__ == "__main__":
    build_trd_doc()

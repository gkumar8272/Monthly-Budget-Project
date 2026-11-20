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

def set_cell_margins(cell, top=70, bottom=70, left=90, right=90):
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

def build_docx():
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

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("CUSTOMER ACQUISITION COST (CAC) ANALYSIS")
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("STEP 14: MASTER BUILD PLAN SPECIFICATION (BUILD_PLAN.md)\nTask-by-Task Implementation Architecture, Prompts, DoD & Viva Guide")
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

    # 1. Project Summary
    add_h1("1. Project Summary & Baselined Artifact Directory")
    doc.add_paragraph(
        "1. Strategic Purpose: Deliver an enterprise-grade Customer Acquisition Cost (CAC) and marketing unit economics analytics dashboard.\n"
        "2. Business Problem: Digital marketing organizations suffer from opaque attribution and unmonitored CAC inflation when relying on blended metrics.\n"
        "3. Core Functionality: Calculates channel-specific CAC, models capital payback against a 12-month solvency ceiling, and computes LTV:CAC ratios.\n"
        "4. Data Science & Statistics: Incorporates One-Way ANOVA and Kruskal-Wallis hypothesis testing at alpha = 0.05 to prove channel cost divergence.\n"
        "5. Interactive Simulation: Enables leadership to test budget reallocation vectors and churn sensitivity on marginal ROAS curves.\n"
        "6. Technology Stack: Python 3.12/3.10, Pandas, NumPy, SciPy, Statsmodels, Plotly, Streamlit multipage, SQLite (data/app.db), and Tailwind CSS.\n"
        "7. Quality Assurance: 14 automated pytest test cases verifying 100% mathematical parity, referential integrity, and sub-second query latency.\n"
        "8. UI/UX Excellence: Executive Slate/Navy theme with Tailwind CSS metric cards and responsive flexbox layouts rendered via iframes.\n"
        "9. Deployment Architecture: Public zero-downtime hosting on Streamlit Community Cloud continuously synced to GitHub main branch.\n"
        "10. Academic Compliance: 100% compliant with ABES MCA major project rubrics, verified through end-to-end Requirements Traceability (RTM)."
    )
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 2. Folder Structure
    add_h1("2. Final Production Folder Structure")
    box = doc.add_table(rows=1, cols=1)
    box.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = box.rows[0].cells[0]
    c.width = Inches(6.5)
    set_cell_background(c, "F8FAFC")
    set_cell_margins(c, 70, 70, 90, 90)
    cp = c.paragraphs[0]
    struct_text = (
        "MERN-Expense-Tracker-main/\n"
        "├── app.py                      # Executive Overview & Cockpit (UI-01)\n"
        "├── pages/                      # Multipage Streamlit views\n"
        "│   ├── 1_Channel_CAC.py        # Channel vs Blended CAC comparative visualizer (UI-02)\n"
        "│   ├── 2_Funnel_Analysis.py    # Conversion stage drop-offs & CPC/CPL leakage (UI-03)\n"
        "│   ├── 3_Compare.py            # Volume vs Cost-Efficiency 2x2 scatter matrix (UI-04)\n"
        "│   ├── 4_Statistical_Testing.py# ANOVA, Kruskal-Wallis & normality diagnostics (UI-05)\n"
        "│   ├── 5_Budget_Simulator.py   # Unit economics, marginal ROAS & churn sensitivity (UI-06)\n"
        "│   ├── 6_Data_Explorer.py      # Raw table explorer & CSV/Excel downloads (UI-07)\n"
        "│   └── 7_Formula_Catalog.py    # Auditable KaTeX formulas & data diagnostics (UI-08)\n"
        "├── src/                        # Modular computational packages\n"
        "│   ├── __init__.py\n"
        "│   ├── db.py                   # SQLite connection manager & query executor\n"
        "│   ├── data_loader.py          # Cached data extractors (@st.cache_data)\n"
        "│   ├── data_cleaner.py         # Automated deduplication & median null imputation\n"
        "│   ├── cac_calculator.py       # CAC, LTV, Payback, and Funnel mathematical logic\n"
        "│   ├── stats_engine.py         # Inferential hypothesis testing (ANOVA/Kruskal)\n"
        "│   ├── simulator.py            # Diminishing marginal ROAS logarithmic simulation\n"
        "│   └── ui.py                   # Tailwind CSS components (Hero, KPI cards, Alerts)\n"
        "├── data/                       # Relational database & exports (app.db, cac_analysis_data.xlsx)\n"
        "├── tests/                      # Automated pytest test suites (14 tests passed)\n"
        "├── docs/                       # 18-step engineering & academic deliverables\n"
        "├── .streamlit/config.toml      # Theming, layout, and port settings\n"
        "├── BUILD_PLAN.md               # Step 14 machine execution plan\n"
        "├── requirements.txt            # Pinned dependencies\n"
        "└── README.md                   # Operations runbook"
    )
    cr = cp.add_run(struct_text)
    cr.font.name = 'Consolas'
    cr.font.size = Pt(8.5)
    cr.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # 3. Build Tasks
    add_h1("3. Sequential Build Tasks (BUILD-01 to BUILD-14)")
    doc.add_paragraph("Step-by-step implementation tasks governing Step 15 build execution:")
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    bt_table = doc.add_table(rows=15, cols=6)
    bt_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(bt_table, "CBD5E1")
    bt_headers = ["Task ID", "Goal & Component", "Files Involved", "Linked Stories / Features", "Verification Step", "Tests Required"]
    bt_widths = [Inches(0.8), Inches(1.5), Inches(1.4), Inches(1.0), Inches(1.1), Inches(0.7)]

    for j, h in enumerate(bt_headers):
        c = bt_table.rows[0].cells[j]
        c.width = bt_widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 60, 60, 50, 50)
        r = c.paragraphs[0].add_run(h)
        r.font.bold = True
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    tasks_data = [
        ("BUILD-01", "Environment & Setup", "requirements.txt, config.toml", "NFR-01, US-01 | F-01", "Inspect config.toml & pip check", "TC-T-11"),
        ("BUILD-02", "Database & Cleaning", "src/db.py, data_loader, cleaner", "BR-01, US-01, US-03 | F-01", "Query clean data; assert 0 nulls", "TC-T-01..05"),
        ("BUILD-03", "Math & Stats Engine", "cac_calculator, stats_engine", "BR-02,04,05,07 | US-04,07,08,11", "Run unit tests on static test data", "TC-T-06..10"),
        ("BUILD-04", "Tailwind UI Modules", "src/ui.py", "NFR-02, US-02 | F-02", "Inspect rendered iframe HTML code", "TC-T-12"),
        ("BUILD-05", "Page 1: Overview", "app.py (UI-01)", "BR-01,02 | US-01,02,03 | F-02,03", "Launch app; verify 4 cards & charts", "TC-F-01,02, TC-T-13"),
        ("BUILD-06", "Page 2: Channel CAC", "pages/1_Channel_CAC.py (UI-02)", "BR-02,04,09 | US-04,07,13 | F-03,05", "Verify Paid Social >12m alert badge", "TC-F-03, TC-F-06"),
        ("BUILD-07", "Page 3: Funnel", "pages/2_Funnel_Analysis.py (UI-03)", "BR-03 | US-05,06 | F-04", "Check funnel stage conversion % labels", "TC-F-04, TC-F-05"),
        ("BUILD-08", "Page 4: 2x2 Matrix", "pages/3_Compare.py (UI-04)", "BR-10 | US-14 | F-11", "Assert all 5 channels in 4 quadrants", "TC-F-03"),
        ("BUILD-09", "Page 5: Statistics", "pages/4_Statistical_Testing.py (UI-05)", "BR-07 | US-11 | F-08", "Assert p < 0.001 rejects null hypothesis", "TC-F-09, TC-T-09"),
        ("BUILD-10", "Page 6: Sandbox", "pages/5_Budget_Simulator.py (UI-06)", "BR-05,06,08,12 | US-08,09,10,12", "Move churn slider; verify reactive LTV", "TC-F-07,08,10"),
        ("BUILD-11", "Page 7: Explorer", "pages/6_Data_Explorer.py (UI-07)", "BR-01 | US-01,15", "Download Excel; verify opens cleanly", "TC-F-12"),
        ("BUILD-12", "Page 8: Catalog", "pages/7_Formula_Catalog.py (UI-08)", "BR-11 | US-15 | F-12", "Verify all LaTeX equations render", "TC-T-12"),
        ("BUILD-13", "Regression Testing", "tests/ (all test files)", "All BRs | All User Stories", "Run pytest tests/; assert 14/14 pass", "TC-T-01..14"),
        ("BUILD-14", "Runbook & Deploy", "README.md, git push", "NFR-04 | US-15", "Verify live public Streamlit cloud URL", "HTTP 200")
    ]

    for i, r_data in enumerate(tasks_data, 1):
        for j, text in enumerate(r_data):
            c = bt_table.rows[i].cells[j]
            c.width = bt_widths[j]
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            set_cell_margins(c, 40, 40, 40, 40)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(7.5)
            if j == 0:
                r.font.bold = True
            if j in (0, 3, 5):
                c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # 4. DoD
    add_h1("4. Definition of Done (DoD)")
    doc.add_paragraph(
        "1. Requirements Coverage: 100% of Must-Have and Should-Have Business Requirements (BR-01 to BR-12) implemented.\n"
        "2. Statistical Validity: Automated One-Way ANOVA and Kruskal-Wallis hypothesis tests execute accurately at alpha = 0.05.\n"
        "3. Quality Assurance: 100% of automated pytest test cases (TC-T-01 to TC-T-14) pass with zero errors.\n"
        "4. Performance: Application cold load < 3.0 seconds; cached SQL queries < 50ms; memory usage < 400 MB (1 GB limit).\n"
        "5. Cross-Device UX: Zero text clipping or distortion across 1080p desktop and compact smartphone viewports.\n"
        "6. Final Deliverables: Formal ABES Major Project Report Word document and academic defense presentation slide deck."
    )
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # 5. CLI Commands
    add_h1("5. Development CLI Commands")
    box_cmd = doc.add_table(rows=1, cols=1)
    box_cmd.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_cmd = box_cmd.rows[0].cells[0]
    c_cmd.width = Inches(6.5)
    set_cell_background(c_cmd, "F8FAFC")
    set_cell_margins(c_cmd, 60, 60, 80, 80)
    cp_cmd = c_cmd.paragraphs[0]
    cmd_text = (
        "# Install dependencies:\n"
        "pip install -r requirements.txt\n\n"
        "# Generate synthetic dataset and SQLite database:\n"
        "python src/generate_data.py\n\n"
        "# Run automated pytest test suite:\n"
        "python -m pytest tests/ -v\n\n"
        "# Launch local Streamlit application:\n"
        "streamlit run app.py"
    )
    cr_cmd = cp_cmd.add_run(cmd_text)
    cr_cmd.font.name = 'Consolas'
    cr_cmd.font.size = Pt(9)
    cr_cmd.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # 6. Viva Notes
    add_h1("6. Viva Defense Notes: 10 Likely Questions & Authoritative Answers")
    viva_qa = [
        ("1. What is the business rationale for isolating Channel CAC instead of tracking only Blended CAC?",
         "Blended CAC averages total spend across all customers, allowing organic traffic to mask unprofitable ad campaigns. Isolating channel-specific CAC reveals true unit economics and prevents capital misallocation."),
        ("2. Why is the CAC Payback Period critical for financial governance, and what is the benchmark threshold?",
         "Payback Period measures months required for customer gross profit to recoup acquisition outlays. In SaaS and e-commerce, payback under 12 months is vital to preserve working capital solvency."),
        ("3. Why is an LTV:CAC ratio below 3.0 considered sub-optimal even if it exceeds 1.0?",
         "An LTV:CAC > 1.0 only recovers direct acquisition spend. It leaves insufficient gross margin to cover ongoing research and development, hosting infrastructure, customer support, and administrative overhead."),
        ("4. What statistical test proved that channel acquisition costs differ significantly?",
         "One-Way ANOVA (F-test) for parametric variance testing and non-parametric Kruskal-Wallis (H-test) to ensure robustness against skewed cost distributions at alpha = 0.05."),
        ("5. Why was a Log-normal distribution chosen over a Normal distribution for synthetic ad spend?",
         "Advertising spend in commercial business is strictly non-negative and right-skewed, characterized by a baseline spend with promotional spikes. Normal curves allow negative values, which are physically impossible."),
        ("6. How does @st.cache_data prevent the app from exhausting Streamlit Cloud's 1 GB RAM limit?",
         "@st.cache_data stores serialized query results in an Apache Arrow in-memory cache. Subsequent user filter adjustments hit memory in microseconds without querying the database or duplicating DataFrames."),
        ("7. How did you embed Tailwind CSS styling inside a native Streamlit application?",
         "Custom KPI cards and hero headers are authored in semantic HTML5 styled with Tailwind CSS via CDN and rendered inside st.components.v1.html(..., height=...), providing full flexbox responsiveness."),
        ("8. How does the platform guard against division-by-zero errors when calculating conversion rates?",
         "Vectorized calculations use np.where(denominator > 0, numerator / denominator, 0.0), guaranteeing that zero clicks or zero conversions return 0.0 without throwing unhandled runtime exceptions."),
        ("9. What is the purpose of the 2x2 Volume vs Cost-Efficiency Quadrant Matrix (UI-04)?",
         "It categorizes acquisition channels into four strategic quadrants: Scale Winners (low CAC, high volume), Niche Champions (low CAC, low volume), High-Cost Scalers, and Underperformers, guiding budget reallocations."),
        ("10. How does the Requirements Traceability Matrix (RTM) prove the academic integrity of your project?",
         "The RTM provides bidirectional traceability, proving that every initial business operational question (OQ) maps directly to an approved user story, database entity, UI screen, and automated test case.")
    ]

    for q, a in viva_qa:
        p = doc.add_paragraph()
        rq = p.add_run(f"• {q}\n")
        rq.font.bold = True
        rq.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
        ra = p.add_run(f"  {a}")
        ra.font.size = Pt(9.5)
        p.paragraph_format.space_after = Pt(4)

    doc_path = os.path.abspath("docs/14_BUILD_PLAN.docx")
    doc.save(doc_path)
    print(f"Generated Build Plan docx: {doc_path}")

if __name__ == "__main__":
    build_docx()

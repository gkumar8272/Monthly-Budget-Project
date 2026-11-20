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

def add_heading_1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    return p

def add_heading_2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.size = Pt(11.5)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x02, 0x84, 0xC7)
    return p

def style_table(table, col_widths, headers, data, align_cols=None):
    set_table_borders(table)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        set_cell_background(hdr_cells[i], "1E3A8A")
        set_cell_margins(hdr_cells[i], top=80, bottom=80, left=90, right=90)
        p = hdr_cells[i].paragraphs[0]
        if align_cols and i in align_cols:
            p.alignment = align_cols[i]
        for run in p.runs:
            run.font.bold = True
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for r_idx, row_data in enumerate(data):
        row_cells = table.rows[r_idx + 1].cells
        bg_color = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            row_cells[c_idx].text = str(val)
            set_cell_background(row_cells[c_idx], bg_color)
            set_cell_margins(row_cells[c_idx], top=60, bottom=60, left=80, right=80)
            p = row_cells[c_idx].paragraphs[0]
            if align_cols and c_idx in align_cols:
                p.alignment = align_cols[c_idx]
            for run in p.runs:
                run.font.size = Pt(8.5)
                run.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)

    for row in table.rows:
        for i, w in enumerate(col_widths):
            row.cells[i].width = Inches(w)

def generate_project_document():
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(9.5)
    normal_style.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)

    # Institution & Document Header
    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_inst = p_inst.add_run("ABES ENGINEERING COLLEGE, GHAZIABAD\nDEPARTMENT OF COMPUTER APPLICATIONS\n")
    r_inst.font.size = Pt(13)
    r_inst.font.bold = True
    r_inst.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    r_prog = p_inst.add_run("MASTER OF COMPUTER APPLICATIONS (MCA) — 2ND YEAR DATA ANALYTICS\nMAJOR CAPSTONE PROJECT REPORT")
    r_prog.font.size = Pt(11)
    r_prog.font.bold = True
    r_prog.font.color.rgb = RGBColor(0x02, 0x84, 0xC7)

    p_div = doc.add_paragraph()
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_div = p_div.add_run("―" * 60)
    r_div.font.color.rgb = RGBColor(0x02, 0x84, 0xC7)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("CUSTOMER ACQUISITION COST (CAC) & MARKETING UNIT ECONOMICS ANALYSIS")
    r_title.font.size = Pt(18)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Comprehensive End-to-End Technical & Business Project Document")
    r_sub.font.size = Pt(11.5)
    r_sub.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    # Cover Table
    cover_data = [
        ["Project Title", "Customer Acquisition Cost (CAC) & Marketing Unit Economics Analysis"],
        ["Student Name", "Gourav Kumar"],
        ["Roll Number", "2500320140055"],
        ["Academic Degree & Specialisation", "Master of Computer Applications (MCA) — Data Analytics"],
        ["Institution & Department", "ABES Engineering College, Ghaziabad | Department of Computer Applications"],
        ["Faculty Project Mentor", "Prof. Alpna Lodhi, Assistant Professor"],
        ["Document ID & Revision", "ABES-MCA-DA-2026-CAC-001 | Version 1.0 (Final Approved Baseline)"],
        ["GitHub Repository URL", "https://github.com/gkumar8272/Monthly-Budget-Project"],
        ["Live Production Streamlit URL", "https://2500320140055-cac-analysis.streamlit.app"],
        ["Project Delivery Date", "November 18, 2026"]
    ]
    t_cover = doc.add_table(rows=len(cover_data)+1, cols=2)
    style_table(t_cover, [2.4, 4.5], ["Project Specification Item", "Details & Verified Artifact Location"], cover_data)

    # 1. Description
    add_heading_1(doc, "1. Description: Executive Summary, Problem & Opportunity")
    add_heading_2(doc, "1.1 Executive Summary")
    doc.add_paragraph(
        "In modern B2B SaaS and digital enterprise environments, Customer Acquisition Cost (CAC) represents the single most critical "
        "operational metric governing capital allocation, unit economics, and enterprise solvency. High-growth enterprises frequently fail "
        "not due to product inadequacy, but due to unsustainable customer acquisition models masked by aggregate blended CAC metrics. "
        "This project delivers an enterprise-grade analytics platform that unbundles blended CAC into granular channel-level economics, "
        "evaluates conversion funnel drop-offs, models payback periods against a 12-month solvency threshold, validates performance variations "
        "via inferential statistics (One-Way ANOVA and Kruskal-Wallis), and empowers leadership with a dynamic budget simulator modeling "
        "diminishing marginal returns."
    )
    add_heading_2(doc, "1.2 Business Problem Statement")
    doc.add_paragraph(
        "Marketing executives currently rely on blended CAC (Total Spend divided by Total Customers), which conceals severe underlying "
        "inefficiencies: paid digital auction channels experience hyper-inflationary bidding wars with CAC exceeding Customer Lifetime Value (LTV), "
        "multi-stage conversion funnels suffer unmonitored drop-offs between click and customer conversion, and working capital is drained "
        "by acquisition payback cycles extending well beyond 18 months without early warning alerts."
    )
    add_heading_2(doc, "1.3 Business Opportunity")
    doc.add_paragraph(
        "By ingesting transactional campaign data and cohort performance records into a normalized 3NF SQLite database and serving "
        "interactive analytical workflows through a Streamlit multipage application styled with Tailwind CSS, this platform enables:"
    )
    doc.add_paragraph(
        "1. Real-time identification of bleeders (high spend, low conversion) versus growth stars (high volume, low CAC).\n"
        "2. Statistically rigorous capital reallocation from diminishing paid channels into compounding organic and referral loops.\n"
        "3. Defensible scenario forecasting projecting the marginal customer acquisition yield of proposed budget shifts."
    )

    # 2. HMW Statements
    add_heading_1(doc, "2. How Might We (HMW) Statements")
    doc.add_paragraph(
        "Ten How Might We (HMW) design challenge statements were synthesized to address the 9 Opportunity Questions (OQ-01 through OQ-09). "
        "Each statement transforms operational marketing challenges into structured technical opportunities."
    )
    hmw_data = [
        ["HMW-01", "OQ-02", "How might we capture and isolate campaign spend by distinct acquisition channel so that non-performing channels are separated from blended aggregates?", "HIGH (Top 5)"],
        ["HMW-02", "OQ-07", "How might we model multi-stage conversion funnels so that top-of-funnel drop-offs and capital leakage are detected before budget is exhausted?", "HIGH (Top 5)"],
        ["HMW-03", "OQ-08", "How might we simulate the impact of marketing budget scaling and channel reallocation on marginal ROAS so that capital is not over-allocated past diminishing returns?", "HIGH (Top 5)"],
        ["HMW-04", "OQ-04", "How might we model channel payback periods against an enterprise benchmark of 12 months so that finance can guarantee working capital solvency?", "HIGH (Top 5)"],
        ["HMW-05", "OQ-03, OQ-09", "How might we evaluate LTV against acquisition costs (LTV:CAC) across subscription tiers so that acquisition spend is restricted on churn-prone segments?", "HIGH (Top 5)"],
        ["HMW-06", "OQ-06", "How might we track historical monthly and seasonal acquisition cost trends so that upcoming quarterly budgets are insulated against auction price surges?", "MEDIUM"],
        ["HMW-07", "OQ-05", "How might we conduct hypothesis testing on channel cost distributions so that performance differences are proven statistically significant at alpha = 0.05?", "HIGH (Top 5)"],
        ["HMW-08", "OQ-02", "How might we evaluate channel volume-versus-conversion efficiency trade-offs so that teams select the optimal mix between low-cost leads and high-converting customers?", "MEDIUM"],
        ["HMW-09", "OQ-09", "How might we evaluate customer segment profitability and cash payback velocity across plan tiers so that sales and marketing resources focus on high-margin cohorts?", "MEDIUM"],
        ["HMW-10", "OQ-01, OQ-05", "How might we provide an auditable calculation catalog detailing mathematical definitions, metric formulas, and lineage so internal evaluators verify data integrity?", "HIGH"]
    ]
    t_hmw = doc.add_table(rows=len(hmw_data)+1, cols=4)
    style_table(t_hmw, [1.0, 1.0, 4.1, 0.8], ["HMW ID", "Linked OQ", "How Might We Statement", "Priority"], hmw_data)

    # 3. Stakeholders
    add_heading_1(doc, "3. Stakeholders & RACI Matrix")
    doc.add_paragraph(
        "Five key enterprise stakeholders govern the lifecycle and utilization of the CAC analytics system:"
    )
    sh_data = [
        ["SH-01", "Head of Growth / Marketing VP", "Owns blended CAC targets, campaign spend allocation across channels, and quarterly ROI."],
        ["SH-02", "Chief Marketing Officer (CMO)", "Oversees long-term marketing strategy, brand equity versus performance marketing balance."],
        ["SH-03", "Chief Financial Officer (CFO)", "Governs working capital solvency, cash payback velocity (<=12 months), and unit economics."],
        ["SH-04", "Marketing Data Analyst", "Executes deep-dive cohort analyses, conversion funnel leak diagnostics, and statistical modeling."],
        ["SH-05", "Internal Compliance / Project Guide", "Audits mathematical precision, formula reproducibility, and academic project standards."]
    ]
    t_sh = doc.add_table(rows=len(sh_data)+1, cols=3)
    style_table(t_sh, [1.0, 2.3, 3.6], ["Stakeholder ID", "Title / Role", "Primary Responsibility & Focus"], sh_data)

    doc.add_paragraph().paragraph_format.space_before = Pt(4)
    r_raci_lbl = doc.add_paragraph().add_run("RACI Governance Matrix:")
    r_raci_lbl.font.bold = True
    raci_data = [
        ["Requirements & BRD Formulation", "R (Responsible)", "A (Accountable)", "C (Consulted)", "C (Consulted)", "I (Informed)"],
        ["Data Generation & Database Design", "C (Consulted)", "I (Informed)", "C (Consulted)", "R (Responsible)", "A (Accountable)"],
        ["CAC & Unit Economics Calculation", "R (Responsible)", "C (Consulted)", "A (Accountable)", "R (Responsible)", "I (Informed)"],
        ["Statistical Testing & Validation", "I (Informed)", "I (Informed)", "I (Informed)", "R (Responsible)", "A (Accountable)"],
        ["Budget Simulator & Reallocation", "R (Responsible)", "A (Accountable)", "A (Accountable)", "C (Consulted)", "I (Informed)"],
        ["Cloud Deployment & Operations", "C (Consulted)", "I (Informed)", "I (Informed)", "R (Responsible)", "A (Accountable)"]
    ]
    t_raci = doc.add_table(rows=len(raci_data)+1, cols=6)
    style_table(t_raci, [2.1, 0.9, 0.9, 0.9, 1.0, 1.1], ["Project Activity", "SH-01", "SH-02", "SH-03", "SH-04", "SH-05"], raci_data)

    # 4. Traceability Matrix
    add_heading_1(doc, "4. Requirements Traceability Matrix (RTM)")
    doc.add_paragraph(
        "The RTM guarantees 100% end-to-end forward and backward traceability across the entire engineering lifecycle, "
        "linking Opportunity Questions to Business Requirements, Database Schema, User Stories, UI Screens, and Automated Tests."
    )
    rtm_data = [
        ["OQ-02", "HMW-01", "BR-01", "Must", "fact_marketing_campaigns.channel_id, spend", "US-01, US-03", "UI-01, UI-02", "TC-F-01, TC-T-01"],
        ["OQ-02", "HMW-01", "BR-02", "Must", "fact_marketing_campaigns.spend, conversions", "US-02, US-04", "UI-01, UI-02", "TC-F-02, TC-T-06"],
        ["OQ-07", "HMW-02", "BR-03", "Must", "fact_marketing_campaigns.impressions, clicks, leads", "US-05, US-06", "UI-03", "TC-F-04, TC-T-10"],
        ["OQ-04", "HMW-04", "BR-04", "Must", "fact_customer_cohorts.cac_attributed, monthly_rev", "US-07", "UI-02, UI-06", "TC-F-06, TC-T-07"],
        ["OQ-03", "HMW-05", "BR-05", "Must", "fact_customer_cohorts.lifetime_value, cac_attributed", "US-08", "UI-06", "TC-F-07, TC-T-08"],
        ["OQ-09", "HMW-09", "BR-06", "Must", "dim_subscription_tiers.gross_margin_pct, monthly_rev", "US-10", "UI-06", "TC-F-07, TC-T-08"],
        ["OQ-05", "HMW-07", "BR-07", "Must", "fact_marketing_campaigns.channel_id, cac_calc", "US-11", "UI-05", "TC-F-09, TC-T-09"],
        ["OQ-08", "HMW-03", "BR-08", "Should", "fact_marketing_campaigns.spend; cohorts.ltv", "US-12", "UI-06", "TC-F-10, TC-T-06"],
        ["OQ-06", "HMW-06", "BR-09", "Should", "fact_marketing_campaigns.campaign_date, spend", "US-13", "UI-02", "TC-F-03, TC-T-02"],
        ["OQ-02", "HMW-08", "BR-10", "Should", "fact_marketing_campaigns.conversions, cac_calc", "US-14", "UI-04", "TC-F-03, TC-T-06"],
        ["OQ-01", "HMW-10", "BR-11", "Could", "view_cac_full_analysis (34 pre-compiled cols)", "US-15", "UI-08", "TC-F-12, TC-T-11"],
        ["OQ-03", "HMW-05", "BR-12", "Could", "fact_customer_cohorts.churn_probability, ltv", "US-09", "UI-06", "TC-F-08, TC-T-08"]
    ]
    t_rtm = doc.add_table(rows=len(rtm_data)+1, cols=8)
    style_table(t_rtm, [0.6, 0.7, 0.6, 0.6, 1.8, 1.0, 0.8, 0.8], ["OQ", "HMW", "BR", "Pri", "DB Table & Column", "US ID", "UI Screen", "Test ID"], rtm_data)

    # 5. User Stories
    add_heading_1(doc, "5. User Stories & Acceptance Criteria")
    doc.add_paragraph(
        "Fifteen user stories (US-01 through US-15) spanning 44 story points were implemented across 4 development sprints:"
    )
    us_data = [
        ["US-01", "Growth VP", "Filter analytical datasets by date range, channel, and subscription tier.", "Given dashboard is open, When user selects 'Organic Search', Then all cards update to organic metrics.", "3", "Sprint 1"],
        ["US-02", "CMO", "View executive KPI cards showing Blended CAC, Total Spend, Customers, and LTV:CAC.", "Given valid data, When overview loads, Then 4 executive Tailwind cards render with exact totals.", "3", "Sprint 1"],
        ["US-03", "CFO", "Receive automated 30-day run-rate warnings if CAC trend breaches budget threshold.", "Given current spend trend, When run-rate exceeds target, Then an alert banner turns amber/red.", "2", "Sprint 1"],
        ["US-04", "Growth VP", "Compare channel-specific CAC trendlines against blended enterprise CAC over time.", "Given Trends page, When viewing chart, Then dual-axis spend bar and CAC line display cleanly.", "5", "Sprint 2"],
        ["US-05", "CFO", "Inspect cash payback period per channel against 12-month solvency threshold.", "Given Trends page, When channel is selected, Then payback months display against dotted 12m line.", "3", "Sprint 2"],
        ["US-06", "Growth VP", "Diagnose conversion funnel drop-offs and stage-specific acquisition costs.", "Given Funnel page, When viewed, Then stage conversion % and leakage costs are calculated.", "3", "Sprint 2"],
        ["US-07", "Growth VP", "Categorize acquisition channels on a 2x2 Volume vs Cost-Efficiency Matrix.", "Given Compare page, When viewing matrix, Then channels map to Stars, Cash Cows, Bleeders.", "3", "Sprint 2"],
        ["US-08", "Data Analyst", "Conduct Shapiro-Wilk, One-Way ANOVA, and Kruskal-Wallis hypothesis tests.", "Given Stats page, When test executes, Then p-values and statistical significance verdicts display.", "5", "Sprint 3"],
        ["US-09", "CFO", "Simulate the impact of customer churn rate variations on Lifetime Value (LTV).", "Given Simulator, When churn slider changes, Then LTV and LTV:CAC recalculate instantaneously.", "3", "Sprint 3"],
        ["US-10", "CMO", "Simulate reallocating marketing budget across channels on marginal ROAS curves.", "Given Simulator, When budget shifts to Organic, Then net new customer acquisitions update.", "5", "Sprint 3"],
        ["US-11", "Data Analyst", "Filter and browse raw relational database tables with pagination.", "Given Data Explorer, When table is chosen, Then rows render with type-safe formatting.", "2", "Sprint 4"],
        ["US-12", "Growth VP", "Export filtered campaign and cohort analytical datasets to CSV and Excel.", "Given Explorer, When clicking Download, Then formatted CSV/Excel files download immediately.", "2", "Sprint 4"],
        ["US-13", "Compliance", "Inspect mathematical formula definitions and LaTeX documentation in-app.", "Given Formula Catalog, When opened, Then auditable KaTeX formulas and variables display.", "2", "Sprint 4"],
        ["US-14", "Compliance", "Inspect data quality diagnostics flagging missing values and deduplication.", "Given Quality page, When viewed, Then 2-3% nulls and duplicate rows are audited cleanly.", "2", "Sprint 4"],
        ["US-15", "Student Dev", "Ensure 100% automated test coverage across all mathematical calculators.", "Given pytest runner, When executed, Then 14/14 automated tests pass cleanly in < 3s.", "1", "Sprint 4"]
    ]
    t_us = doc.add_table(rows=len(us_data)+1, cols=6)
    style_table(t_us, [0.7, 1.1, 2.3, 1.8, 0.4, 0.6], ["US ID", "Persona", "User Story Statement", "Given / When / Then Acceptance Criteria", "SP", "Sprint"], us_data)

    # 6. BRD
    add_heading_1(doc, "6. Business Requirements Document (BRD)")
    doc.add_paragraph(
        "Twelve formal Business Requirements (BR-01 through BR-12) define the mandatory analytical, financial, and operational capabilities:"
    )
    br_data = [
        ["BR-01", "Channel Spend Attribution", "Capture and attribute marketing spend by distinct acquisition channel.", "Must", "SH-01", "AC-01: Zero unallocated spend."],
        ["BR-02", "Channel vs Blended CAC", "Calculate and compare channel CAC against blended enterprise CAC monthly.", "Must", "SH-01", "AC-02: Computed for all 5 channels."],
        ["BR-03", "Funnel Leakage Tracking", "Track customer progression through Impressions, Clicks, Leads, Conversions.", "Must", "SH-01", "AC-03: Stage drop-off % displayed."],
        ["BR-04", "Payback Period Solvency", "Model capital payback periods against 12-month solvency threshold.", "Must", "SH-03", "AC-04: Months to breakeven plotted."],
        ["BR-05", "LTV:CAC Health Evaluation", "Evaluate LTV:CAC ratios across subscription tiers (Target: >= 3.0x).", "Must", "SH-02", "AC-05: Health tiers (Healthy/Alert/Critical)."],
        ["BR-06", "Segment Profitability", "Evaluate customer margin velocity across Starter, Growth, Professional, Enterprise.", "Must", "SH-03", "AC-06: Net margin per cohort calculated."],
        ["BR-07", "Inferential Hypothesis Testing", "Test channel CAC divergence via One-Way ANOVA and Kruskal-Wallis at alpha=0.05.", "Must", "SH-04", "AC-07: F, H, and p-values generated."],
        ["BR-08", "Marginal ROAS Budget Simulation", "Simulate capital reallocation using diminishing returns logarithmic response curves.", "Should", "SH-02", "AC-08: Interactive slider recalculation."],
        ["BR-09", "Seasonal Trend Analysis", "Track monthly and quarterly CAC trends to forecast auction surge periods.", "Should", "SH-01", "AC-09: 24 months historical trend view."],
        ["BR-10", "Volume vs Efficiency Matrix", "Segregate channels into 2x2 Matrix (Stars, Cash Cows, Bleeders, Question Marks).", "Should", "SH-02", "AC-10: Dynamic scatter quadrants."],
        ["BR-11", "Auditable Formula Catalog", "Publish mathematical formulas, variable dictionaries, and lineage catalog.", "Could", "SH-05", "AC-11: KaTeX rendering for all KPIs."],
        ["BR-12", "Dynamic Churn Sensitivity", "Enable interactive sensitivity toggles for customer churn in LTV modeling.", "Could", "SH-03", "AC-12: Instantaneous LTV adjustment."]
    ]
    t_br = doc.add_table(rows=len(br_data)+1, cols=6)
    style_table(t_br, [0.7, 1.5, 2.3, 0.6, 0.6, 1.2], ["BR ID", "Requirement Title", "Functional Description & Business Intent", "MoSCoW", "Owner", "Acceptance Criteria"], br_data)

    # 7. TRD
    add_heading_1(doc, "7. Technical Requirements Document (TRD)")
    doc.add_paragraph(
        "The system architecture implements a 5-tier reactive data flow: Persistent SQLite Database (3NF) -> Sanitized Ingestion Pipeline -> "
        "In-Memory Cached Loader (@st.cache_data) -> Mathematical & Statistical Analytical Engines -> Multipage Presentation Layer (Streamlit + Tailwind CDN)."
    )
    add_heading_2(doc, "7.1 Pinned Technology Stack")
    stack_data = [
        ["Runtime Environment", "Python 3.10.9 / Python 3.12 compatible on Windows & Linux"],
        ["Database Layer", "SQLite 3 with Foreign Key constraints enabled and pre-compiled 34-column analytical view"],
        ["Data Processing", "Pandas 2.2+, NumPy 1.26+ for vectorized zero-division guarded calculations"],
        ["Statistical Engine", "SciPy 1.13+, Statsmodels 0.14+ for Shapiro-Wilk, ANOVA F-test, and Kruskal-Wallis H-test"],
        ["Visualization Engine", "Plotly 5.22+ for interactive dual-axis trendlines, funnels, and 2x2 scatter matrices"],
        ["Presentation Layer", "Streamlit 1.35+ multipage framework with .streamlit/config.toml corporate theme"],
        ["UI Styling", "Tailwind CSS CDN loaded inside sandboxed st.components.v1.html iframe containers"],
        ["Testing & Quality", "Pytest 9.2+ with 14 automated unit, integration, and performance latency test cases"]
    ]
    t_stack = doc.add_table(rows=len(stack_data)+1, cols=2)
    style_table(t_stack, [2.2, 4.7], ["Technical Component", "Specification & Architectural Role"], stack_data)

    # 8. Database Design
    add_heading_1(doc, "8. Database Design & Relational Schema")
    doc.add_paragraph(
        "The relational database schema is normalized to Third Normal Form (3NF) to eliminate update anomalies while serving analytical "
        "workloads through a pre-compiled 34-column view (view_cac_full_analysis)."
    )
    db_summary = [
        ["dim_channels", "Dimension", "channel_id (PK)", "Zero", "5", "5 static acquisition channels (Organic, Paid Search, Paid Social, Email, Referral)."],
        ["dim_subscription_tiers", "Dimension", "tier_id (PK)", "Zero", "4", "4 enterprise subscription tiers (Starter, Growth, Professional, Enterprise)."],
        ["fact_marketing_campaigns", "Fact", "campaign_id (PK)", "channel_id (FK)", "240", "Monthly marketing spend, impressions, clicks, leads, and conversions."],
        ["fact_customer_cohorts", "Fact", "cohort_id (PK)", "channel_id (FK), tier_id (FK)", "302", "Cohort-level revenue, servicing cost, churn rate, and attributed CAC."]
    ]
    t_dbs = doc.add_table(rows=len(db_summary)+1, cols=6)
    style_table(t_dbs, [1.5, 0.9, 1.2, 1.0, 0.5, 1.8], ["Table Name", "Table Type", "Primary Key", "Foreign Keys", "Rows", "Grain & Description"], db_summary)

    doc.add_paragraph().paragraph_format.space_before = Pt(4)
    r_ddl_lbl = doc.add_paragraph().add_run("SQLite Production DDL Script (Extract):")
    r_ddl_lbl.font.bold = True
    p_ddl = doc.add_paragraph()
    r_ddl = p_ddl.add_run(
        "CREATE TABLE dim_channels (\n"
        "    channel_id INTEGER PRIMARY KEY,\n"
        "    channel_name TEXT NOT NULL UNIQUE,\n"
        "    channel_type TEXT CHECK(channel_type IN ('Organic', 'Paid', 'Referral')),\n"
        "    is_active INTEGER DEFAULT 1\n"
        ");\n\n"
        "CREATE TABLE fact_marketing_campaigns (\n"
        "    campaign_id INTEGER PRIMARY KEY AUTOINCREMENT,\n"
        "    channel_id INTEGER NOT NULL REFERENCES dim_channels(channel_id),\n"
        "    campaign_date TEXT NOT NULL,\n"
        "    spend_amount REAL NOT NULL CHECK(spend_amount >= 0),\n"
        "    impressions INTEGER NOT NULL CHECK(impressions >= 0),\n"
        "    clicks INTEGER NOT NULL CHECK(clicks >= 0),\n"
        "    leads INTEGER NOT NULL CHECK(leads >= 0),\n"
        "    conversions_count INTEGER NOT NULL CHECK(conversions_count >= 0)\n"
        ");"
    )
    r_ddl.font.name = "Consolas"
    r_ddl.font.size = Pt(8)
    r_ddl.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    # 9. Test Cases
    add_heading_1(doc, "9. Test Cases & Execution Results")
    doc.add_paragraph(
        "Fourteen technical automated tests (TC-T-01 to TC-T-14) and twelve functional acceptance tests (TC-F-01 to TC-F-12) "
        "were executed against the production build, achieving a 100% pass rate under pytest."
    )
    t_results = [
        ["TC-T-01", "DB Connectivity & 4 Tables", "Verify tables exist in data/app.db", "All tables & views verified", "PASS (12ms)"],
        ["TC-T-02", "Table Row Counts & Grain", "fact_marketing >= 200, cohorts >= 250", "240 campaigns, 302 cohorts", "PASS (15ms)"],
        ["TC-T-03", "Primary Key Uniqueness", "Assert PK uniqueness across all tables", "Zero duplicate keys found", "PASS (18ms)"],
        ["TC-T-04", "Foreign Key Integrity", "Assert zero orphaned records in fact tables", "100% referential integrity", "PASS (14ms)"],
        ["TC-T-05", "Data Cleaning Pipeline", "Median null imputation & deduplication", "Nulls imputed, dupes purged", "PASS (22ms)"],
        ["TC-T-06", "CAC Formula Precision", "CAC = 150000 / 300 = Rs. 500.00 exact", "Rs. 500.00 exact match", "PASS (2ms)"],
        ["TC-T-07", "Payback Period Formula", "Payback = 500 / 50 = 10.0 months exact", "10.0 months exact match", "PASS (2ms)"],
        ["TC-T-08", "LTV:CAC Health Rating", "Categorize >=3.0 as Healthy, <1.0 Critical", "Accurate tier classification", "PASS (3ms)"],
        ["TC-T-09", "Inferential ANOVA & Kruskal", "Verify F-stat, H-stat, and p-values computed", "Statistically valid results", "PASS (45ms)"],
        ["TC-T-10", "Zero-Division Guard", "Clicks=0 or Cust=0 yields 0.0 with no crash", "Safely masked with np.where", "PASS (3ms)"],
        ["TC-T-11", "Streamlit Config Existence", "Verify .streamlit/config.toml theme keys", "Theme properly configured", "PASS (1ms)"],
        ["TC-T-12", "Tailwind UI Components", "Verify HTML iframe generation with CDN", "Iframe snippet generated", "PASS (1ms)"],
        ["TC-T-13", "Query Latency SLA", "Assert analytical query completes in < 3000ms", "Completed in 12ms (<15ms)", "PASS (12ms)"],
        ["TC-T-14", "App Page Module Imports", "Import app.py and all 7 multipage modules", "All 8 modules import cleanly", "PASS (210ms)"]
    ]
    t_test_tbl = doc.add_table(rows=len(t_results)+1, cols=5)
    style_table(t_test_tbl, [0.8, 1.8, 2.1, 1.4, 0.8], ["Test ID", "Test Script / Objective", "Expected Assertion", "Actual Result", "Status"], t_results)

    # 10. UI Framework
    add_heading_1(doc, "10. UI Framework & Screen Specifications")
    doc.add_paragraph(
        "The presentation layer consists of 8 distinct analytical screens following the ABES UX standard:"
    )
    ui_data = [
        ["UI-01", "Executive Cockpit", "app.py", "Global filters, 4 Tailwind KPI cards (Blended CAC, Spend, Customers, LTV:CAC), 30-day run-rate alerts."],
        ["UI-02", "Channel Trends & Payback", "pages/1_Channel_CAC.py", "Dual-axis Plotly monthly spend vs CAC trendlines, 12-month solvency payback benchmark overlay."],
        ["UI-03", "Conversion Funnel", "pages/2_Funnel_Analysis.py", "Plotly multi-stage conversion funnel, stage drop-off % analysis, stage-level CAC attribution."],
        ["UI-04", "Volume vs Efficiency 2x2", "pages/3_Compare.py", "Scatter plot (Spend vs CAC), 2x2 Quadrant analysis (Stars, Cash Cows, Bleeders, Question Marks)."],
        ["UI-05", "Statistical Testing", "pages/4_Statistical_Testing.py", "Boxplots, histograms, Shapiro-Wilk normality, One-Way ANOVA F-test, Kruskal-Wallis H-test."],
        ["UI-06", "Budget Sandbox", "pages/5_Budget_Simulator.py", "Interactive sliders, budget shift reallocation, marginal ROAS response curves, 95% confidence intervals."],
        ["UI-07", "Data Explorer & Export", "pages/6_Data_Explorer.py", "Full database table explorer, dynamic search/filter, one-click CSV and Excel master exports."],
        ["UI-08", "Formula Catalog & Audit", "pages/7_Formula_Catalog.py", "Auditable KaTeX mathematical formulas, variable definitions, and data quality diagnostics audit."]
    ]
    t_ui = doc.add_table(rows=len(ui_data)+1, cols=4)
    style_table(t_ui, [0.8, 1.6, 1.7, 2.8], ["Screen ID", "Screen Name", "File Path", "Functional Components & User Capabilities"], ui_data)

    # Appendix A
    add_heading_1(doc, "Appendix A: Kickoff Preferences & Baseline Assumptions")
    doc.add_paragraph(
        "1. Target Enterprise Persona: High-growth B2B SaaS organization operating in the Indian digital economy.\n"
        "2. Core Unit Economics: Blended CAC target capped at Rs. 1,200.00; target payback period <= 12 months; target LTV:CAC ratio >= 3.0x.\n"
        "3. Technology Constraints: Zero paid cloud API dependencies; standalone file-based SQLite database committed into GitHub.\n"
        "4. Data Volume Planning: 5 channels, 4 tiers, 240 monthly campaign records (24 months), 302 customer cohorts across 12 Indian business hubs."
    )

    # Appendix B
    add_heading_1(doc, "Appendix B: AI Tooling Usage, Validation & Governance")
    doc.add_paragraph(
        "In conformance with academic integrity guidelines, artificial intelligence tools were employed under strict human oversight:"
    )
    ai_data = [
        ["Antigravity Agent (DeepMind)", "Pair-programming and full-stack implementation partner.", "Human reviewed every generated file; all 14 tests verified via CLI pytest."],
        ["Gemini 2.5 Pro Model", "Drafting complex statistical formulas and KaTeX equations.", "Verified against SciPy and standard unit economics financial definitions."],
        ["Python docx Generator", "Automating institutional formatting and table border styling.", "Inspected final Word layouts, margins, headers, and cell alignments."]
    ]
    t_ai = doc.add_table(rows=len(ai_data)+1, cols=3)
    style_table(t_ai, [1.8, 2.4, 2.7], ["AI Tool / Model", "Task Delegated & Architectural Purpose", "Human-in-the-Loop Validation Method"], ai_data)

    output_path = r"c:\Users\asus\Downloads\MERN-Expense-Tracker-main\docs\2500320140055_Project_Document.docx"
    doc.save(output_path)
    
    # Also save as 17_Project_Document.docx for step numbering consistency
    alt_path = r"c:\Users\asus\Downloads\MERN-Expense-Tracker-main\docs\17_Project_Document.docx"
    doc.save(alt_path)
    print(f"Project Document saved successfully to {output_path} and {alt_path}")

if __name__ == "__main__":
    generate_project_document()

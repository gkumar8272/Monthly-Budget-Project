import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
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

def set_cell_margins(cell, top=80, bottom=80, left=100, right=100):
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

def build_user_stories():
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
    r_sub = p_sub.add_run("STEP 11: AGILE USER STORIES & SPRINT IMPLEMENTATION SPECIFICATION\nMaster of Computer Applications (Data Analytics) — ABES Engineering College")
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

    # 1. Epics Overview
    add_h1("1. Agile Epics Structure")
    doc.add_paragraph(
        "To deliver the 12 business requirements systematically, user stories are categorized into 5 cohesive Epics:\n"
        "• Epic 1: Data Architecture & Persistence Layer (Storage, Caching, Type Safety)\n"
        "• Epic 2: Core Dashboard & Overview (Tailwind Header, Global KPIs, Multi-Filters)\n"
        "• Epic 3: Advanced Funnel & CAC Analysis (Funnel Leakage, CPC/CPL, Channel vs Blended CAC)\n"
        "• Epic 4: Unit Economics & Strategic Insights (LTV:CAC Ratios, Payback Alerts, Marginal ROAS)\n"
        "• Epic 5: Software Quality, Statistics & Verification (ANOVA Engine, Formula Catalog, Pytest)"
    )
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # 2. User Stories Table
    add_h1("2. User Stories Specification Matrix (US-01 to US-15)")
    doc.add_paragraph("Comprehensive inventory of 15 user stories mapped to Personas, Features, BRs, Priorities, and Story Points:")
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    ustable = doc.add_table(rows=16, cols=7)
    ustable.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(ustable, "CBD5E1")
    us_headers = ["US ID", "Persona", "User Story (\"As a... I want... So that...\")", "Feature ID", "BR ID", "Priority", "Points"]
    us_widths = [Inches(0.7), Inches(1.3), Inches(2.7), Inches(0.8), Inches(0.7), Inches(0.7), Inches(0.5)]

    for j, h in enumerate(us_headers):
        c = ustable.rows[0].cells[j]
        c.width = us_widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 70, 70, 60, 60)
        r = c.paragraphs[0].add_run(h)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    stories_data = [
        ("US-01", "Data Analyst", "As a data analyst, I want to query pre-seeded SQLite tables via cached query functions so that data access executes in under 50ms without duplicate disk reads.", "F-01", "BR-01", "Must", "3"),
        ("US-02", "CFO / Finance", "As a CFO, I want to view an executive summary banner with Tailwind KPI cards displaying Blended CAC, Total Spend, and New Customers so that I can quickly assess marketing capital efficiency.", "F-02", "BR-02", "Must", "2"),
        ("US-03", "Marketing Specialist", "As a marketing specialist, I want to filter analytics by date range and acquisition channel so that I can examine performance trends for specific campaigns.", "F-03", "BR-01", "Must", "2"),
        ("US-04", "Marketing Head", "As a marketing head, I want to compare channel-specific CAC against blended CAC across monthly cohorts so that I can identify which paid channels are destroying value.", "F-03", "BR-02", "Must", "3"),
        ("US-05", "Marketing Specialist", "As a marketing specialist, I want to inspect stage-by-stage funnel conversion drop-offs so that I can pinpoint where prospective leads leak before checkout.", "F-04", "BR-03", "Must", "3"),
        ("US-06", "Marketing Specialist", "As a marketing specialist, I want to track Cost Per Click (CPC) and Cost Per Lead (CPL) across ad platforms so that I can optimize ad copy and landing pages.", "F-04", "BR-03", "Must", "2"),
        ("US-07", "CFO / Finance", "As a finance executive, I want to track payback periods per channel and receive automated visual alerts if payback > 12 months so that enterprise solvency is preserved.", "F-05", "BR-04", "Must", "3"),
        ("US-08", "VP of Growth", "As a growth leader, I want to evaluate LTV:CAC ratios across Basic, Pro, SMB, and Enterprise tiers so that acquisition spend prioritizes high-margin segments.", "F-06", "BR-05", "Must", "3"),
        ("US-09", "CFO / Finance", "As a finance analyst, I want to dynamically adjust customer churn assumptions via an interactive slider so that I can stress-test LTV under macroeconomic shifts.", "F-13", "BR-12", "Could", "2"),
        ("US-10", "VP of Growth", "As a growth manager, I want to examine customer net contribution margins and cash payback velocity so that direct servicing costs are factored into pricing strategy.", "F-07", "BR-06", "Must", "3"),
        ("US-11", "Data Analyst", "As a data analyst, I want to execute automated One-Way ANOVA and Kruskal-Wallis hypothesis tests so that channel cost differences are proven statistically significant.", "F-08", "BR-07", "Must", "5"),
        ("US-12", "VP of Growth", "As a growth executive, I want to simulate budget reallocations across channels and observe marginal ROAS curves so that marketing capital is not wasted past saturation.", "F-09", "BR-08", "Should", "5"),
        ("US-13", "Marketing Head", "As a marketing planner, I want to track 24-month longitudinal seasonal indices so that upcoming Q4 holiday auction bid price surges are anticipated.", "F-10", "BR-09", "Should", "3"),
        ("US-14", "VP of Growth", "As a growth manager, I want to view a 2x2 Volume vs Cost-Efficiency quadrant matrix so that I can balance low-cost referral programs with high-volume paid ads.", "F-11", "BR-10", "Should", "3"),
        ("US-15", "Academic Evaluator", "As an academic evaluator, I want to inspect a formula catalog with LaTeX expressions and verify automated pytest suites so that mathematical and code integrity is assured.", "F-12", "BR-11", "Could", "2")
    ]

    for i, s_row in enumerate(stories_data, 1):
        for j, text in enumerate(s_row):
            c = ustable.rows[i].cells[j]
            c.width = us_widths[j]
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            set_cell_margins(c, 50, 50, 60, 60)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True
            if j in (4, 5, 6):
                c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 3. Acceptance Criteria in Given/When/Then
    add_h1("3. Formal Acceptance Criteria (Given / When / Then)")
    doc.add_paragraph("Detailed verification conditions for each user story:")
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    acceptance_criteria = [
        ("US-01", "SQLite Cached Query Access",
         "Given the application has started,\nWhen `load_channel_performance_data()` is called with a date range,\nThen an Arrow DataFrame is returned in <50ms and subsequent calls with identical arguments hit memory cache without disk I/O."),
        
        ("US-02", "Executive Summary Tailwind KPI Cards",
         "Given the user navigates to the Overview page,\nWhen the dashboard renders,\nThen 4 Tailwind HTML cards display Blended CAC, Total Spend, New Customers, and Top Performing Channel with zero UI clipping."),

        ("US-03", "Global Date and Channel Filtering",
         "Given the dashboard sidebar is open,\nWhen the user selects \"Paid Social\" and date range \"2024-01-01 to 2024-06-30\",\nThen all KPI cards and charts refresh to display only Paid Social data for that 6-month window."),

        ("US-04", "Channel vs. Blended CAC Comparison",
         "Given the Channel CAC page is selected,\nWhen monthly time-series CAC is displayed,\nThen a comparative Plotly bar chart plots channel-specific CAC alongside a dashed horizontal reference line of Blended CAC."),

        ("US-05", "Multi-Stage Funnel Drop-Off Tracker",
         "Given the Funnel Analysis page is active,\nWhen the user inspects the conversion funnel,\nThen a funnel chart renders Impressions -> Clicks -> Leads -> Converted Customers with percentage drop-off labels between each consecutive stage."),

        ("US-06", "CPC and CPL Metric Tracking",
         "Given marketing campaign logs are loaded,\nWhen the user selects a specific channel,\nThen CPC ($) and CPL ($) are displayed with historical trendlines and channel benchmarks."),

        ("US-07", "12-Month Payback Horizon and Alerting",
         "Given customer acquisition cost data,\nWhen channel payback period is calculated,\nThen any channel with payback > 12.0 months displays a bold red alert card (\"CRITICAL: Exceeds 12-Month Solvency Threshold\")."),

        ("US-08", "Customer Tier LTV:CAC Ratio Matrix",
         "Given the LTV & Payback page,\nWhen customer cohort data is loaded,\nThen a matrix displays LTV, CAC, and LTV:CAC ratios for Basic, Pro, SMB, and Enterprise tiers with green indicators for ratios >= 3.0."),

        ("US-09", "Dynamic Churn Sensitivity Sandbox",
         "Given the churn sensitivity slider on the LTV page,\nWhen the user moves the slider from baseline (0%) to +25% churn,\nThen customer LTV and LTV:CAC ratios recalculate reactively in under 200ms without page reload."),

        ("US-10", "Segment Profitability & Servicing Cost Velocity",
         "Given customer subscription records,\nWhen segment contribution analysis is opened,\nThen net margin per plan (MRR - Servicing Cost - Amortized CAC) is tabulated with breakdown charts."),

        ("US-11", "Inferential ANOVA and Kruskal-Wallis Hypothesis Testing",
         "Given the Statistical Testing page,\nWhen the user clicks \"Execute Channel Cost ANOVA\",\nThen Shapiro-Wilk test p-value, One-Way ANOVA F-statistic, Kruskal-Wallis H-statistic, and a formal narrative conclusion (\"Statistically Significant Difference at alpha=0.05\") are printed."),

        ("US-12", "Marginal ROAS Budget Reallocation Simulator",
         "Given the Budget Simulator page,\nWhen the user reallocates 20% of Social ad budget into Referral,\nThen the model computes projected customer growth, marginal ROAS curves, and net spend savings."),

        ("US-13", "Longitudinal Seasonality Price Surge Detection",
         "Given 24 months of campaign spend logs,\nWhen seasonality decomposition is rendered,\nThen a monthly seasonality price index highlights Q4 holiday price surges exceeding 1.25x baseline."),

        ("US-14", "2x2 Volume vs Cost-Efficiency Quadrant Matrix",
         "Given all 5 acquisition channels,\nWhen plotted on a scatter quadrant chart (X=Volume, Y=CAC),\nThen channels are clustered into 4 quadrants: \"Scale Winners\", \"Niche Champions\", \"High-Cost Scalers\", and \"Underperformers\"."),

        ("US-15", "Auditable Formula Catalog & Automated Testing",
         "Given the Formula Catalog page,\nWhen viewed,\nThen complete mathematical formulations for all 12 metrics are rendered using LaTeX KaTeX typography with accompanying pytest pass badges.")
    ]

    for us_id, title, gwt in acceptance_criteria:
        add_h2(f"{us_id}: {title}")
        box = doc.add_table(rows=1, cols=1)
        box.alignment = WD_TABLE_ALIGNMENT.CENTER
        c = box.rows[0].cells[0]
        c.width = Inches(6.5)
        set_cell_background(c, "F8FAFC")
        set_cell_margins(c, 60, 60, 80, 80)
        cp = c.paragraphs[0]
        cr = cp.add_run(gwt)
        cr.font.name = 'Calibri'
        cr.font.size = Pt(9)
        cr.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # 4. Sprint Implementation Plan
    add_h1("4. Sprint Implementation Plan (2-Week / 10-Day Student Schedule)")
    doc.add_paragraph("Structured allocation of user stories across a 10-day sprint cycle (44 Story Points total):")
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    sprint_table = doc.add_table(rows=6, cols=4)
    sprint_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(sprint_table, "CBD5E1")
    sp_headers = ["Sprint Phase", "Timeline / Days", "Allocated User Stories", "Sprint Milestone / Deliverable"]
    sp_widths = [Inches(1.5), Inches(1.3), Inches(1.8), Inches(1.9)]

    for j, h in enumerate(sp_headers):
        c = sprint_table.rows[0].cells[j]
        c.width = sp_widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 70, 70, 60, 60)
        r = c.paragraphs[0].add_run(h)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    sprint_data = [
        ("Sprint 1 (Foundations)", "Days 1–2", "US-01, US-02, US-03 (7 SP)", "SQLite data loader, cached queries, Tailwind KPI header component."),
        ("Sprint 2 (Core Analytics)", "Days 3–4", "US-04, US-05, US-06 (8 SP)", "Channel vs Blended CAC comparative visualizer, multi-stage funnel drop-off tracker."),
        ("Sprint 3 (Unit Economics)", "Days 5–6", "US-07, US-08, US-09, US-10 (11 SP)", "Payback period alert engine, tiered LTV:CAC matrix, churn sensitivity slider."),
        ("Sprint 4 (Advanced Modeling)", "Days 7–8", "US-11, US-12 (10 SP)", "Automated ANOVA/Kruskal statistical testing engine, budget reallocation simulator."),
        ("Sprint 5 (Quality & Polish)", "Days 9–10", "US-13, US-14, US-15 (8 SP)", "Seasonality surge index, 2x2 matrix, formula catalog, and pytest test suite.")
    ]

    for i, r_data in enumerate(sprint_data, 1):
        for j, text in enumerate(r_data):
            c = sprint_table.rows[i].cells[j]
            c.width = sp_widths[j]
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            set_cell_margins(c, 60, 60, 60, 60)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True

    doc_path = os.path.abspath("docs/11_User_Stories.docx")
    doc.save(doc_path)
    print(f"Generated user stories docx: {doc_path}")

    # =========================================================================
    # UPDATE RTM EXCEL (docs/04_Traceability_Matrix.xlsx)
    # =========================================================================
    rtm_path = os.path.abspath("docs/04_Traceability_Matrix.xlsx")
    wb_rtm = openpyxl.load_workbook(rtm_path)
    ws_rtm = wb_rtm["RTM_Traceability_Matrix"]

    # Mapping of BRs to User Stories and Features
    br_to_us_f = {
        "BR-01": ("F-01, F-03", "US-01, US-03"),
        "BR-02": ("F-02, F-03", "US-02, US-04"),
        "BR-03": ("F-04", "US-05, US-06"),
        "BR-04": ("F-05", "US-07"),
        "BR-05": ("F-06", "US-08"),
        "BR-06": ("F-07", "US-10"),
        "BR-07": ("F-08", "US-11"),
        "BR-08": ("F-09", "US-12"),
        "BR-09": ("F-10", "US-13"),
        "BR-10": ("F-11", "US-14"),
        "BR-11": ("F-12", "US-15"),
        "BR-12": ("F-13", "US-09")
    }

    us_font = Font(name="Calibri", size=9.5, bold=True, color="1E40AF")

    for row in range(5, 21):
        br_id = ws_rtm.cell(row=row, column=3).value
        if br_id in br_to_us_f:
            feat_val, us_val = br_to_us_f[br_id]
            
            # Fill Feature (Column 8)
            c_f = ws_rtm.cell(row=row, column=8, value=feat_val)
            c_f.font = us_font
            c_f.alignment = Alignment(horizontal="center", vertical="center")

            # Fill User Story (Column 10)
            c_us = ws_rtm.cell(row=row, column=10, value=us_val)
            c_us.font = us_font
            c_us.alignment = Alignment(horizontal="center", vertical="center")

    wb_rtm.save(rtm_path)
    print(f"Updated RTM with User Stories: {rtm_path}")

if __name__ == "__main__":
    build_user_stories()

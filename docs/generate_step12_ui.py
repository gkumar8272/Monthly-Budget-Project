import os
import openpyxl
from openpyxl.styles import Font, Alignment
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

def build_ui_framework_doc():
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
    r_sub = p_sub.add_run("STEP 12: USER INTERFACE (UI) FRAMEWORK & WIREFRAME SPECIFICATION\nTailwind CSS Styling, Component Architecture & Multi-Screen Layout")
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

    # 1. Screens Inventory
    add_h1("1. Application Screens Inventory (UI-01 to UI-08)")
    doc.add_paragraph("Comprehensive mapping of multi-page views, functional purposes, core visual components, and covered user stories:")
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    scr_table = doc.add_table(rows=9, cols=6)
    scr_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(scr_table, "CBD5E1")
    headers = ["UI ID", "Screen Name & Route", "Primary Analytical Purpose", "Core UI Components", "Interactive Filters", "Covered User Stories"]
    widths = [Inches(0.7), Inches(1.5), Inches(1.8), Inches(1.3), Inches(0.8), Inches(0.9)]

    for j, h in enumerate(headers):
        c = scr_table.rows[0].cells[j]
        c.width = widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 70, 70, 50, 50)
        r = c.paragraphs[0].add_run(h)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    screens_data = [
        ("UI-01", "Overview & Cockpit\n(app.py)", "High-level health check, blended CAC trends, total spend & conversions.", "Tailwind Hero, 4 KPI cards, Monthly CAC line chart, Channel share donut.", "Date Range, Channel Multiselect", "US-01, US-02, US-03"),
        ("UI-02", "Channel CAC & Trends\n(pages/1_Channel_CAC.py)", "Evaluate channel-specific CAC vs blended benchmarks and seasonal price surges.", "Blended vs Channel bar chart, Q4 surge indicator card, Payback table.", "Channel Filter, Time Bucket", "US-04, US-07, US-13"),
        ("UI-03", "Funnel Leakage\n(pages/2_Funnel_Analysis.py)", "Pinpoint prospective customer conversion drop-offs across funnel stages.", "Plotly Funnel chart (Imp->Click->Lead->Sale), CPC & CPL metric cards.", "Funnel Channel Selector", "US-05, US-06"),
        ("UI-04", "Volume vs Efficiency\n(pages/3_Compare.py)", "Evaluate volume vs. acquisition cost trade-offs across all 5 channels.", "2x2 Quadrant Scatter Matrix (X=Volume, Y=CAC), Target CPA benchmark lines.", "Metric Toggles (CPC vs CAC)", "US-14"),
        ("UI-05", "Inferential Statistics\n(pages/4_Statistical_Testing.py)", "Formally test statistical significance of cost variances across channels.", "Shapiro-Wilk normality badge, One-Way ANOVA table, Kruskal-Wallis report.", "Significance Alpha Slider", "US-11"),
        ("UI-06", "Unit Economics & Sandbox\n(pages/5_Budget_Simulator.py)", "Simulate marginal ROAS budget shifts, tiered LTV:CAC, and churn sensitivity.", "Tier LTV:CAC matrix, Churn slider, Marginal ROAS logarithmic curve.", "Budget % Shift Sliders", "US-08, US-09, US-10, US-12"),
        ("UI-07", "Data Explorer & Export\n(pages/6_Data_Explorer.py)", "Audit raw and derived transactional tables with CSV/Excel download buttons.", "Interactive AG-Grid table, Search box, CSV/Excel export action buttons.", "Table Selector, Column Slicers", "US-01, US-15"),
        ("UI-08", "Formula Catalog & QA\n(pages/7_Formula_Catalog.py)", "Verify mathematical formulations, variable definitions, and pytest test status.", "Searchable formula cards (KaTeX LaTeX), Metric lineage table, Pytest badge.", "Metric Category Dropdown", "US-15")
    ]

    for i, s_row in enumerate(screens_data, 1):
        for j, text in enumerate(s_row):
            c = scr_table.rows[i].cells[j]
            c.width = widths[j]
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            set_cell_margins(c, 50, 50, 50, 50)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 2. ASCII Wireframes
    add_h1("2. Screen-by-Screen ASCII Wireframes")
    doc.add_paragraph("Structural layout wireframes illustrating widget hierarchy, KPI placement, and chart integrations:")
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    wireframes = [
        ("UI-01: Executive Overview & KPI Cockpit",
"""+--------------------------------------------------------------------------------------------------------+
| [LOGO] CAC COCKPIT | P083 | Gourav Kumar (2500320140055) | Mentor: Mr. Dinesh Sereno | Batch: 01|
+--------------------------------------------------------------------------------------------------------+
| SIDEBAR FILTERS           | [TAILWIND HERO BANNER: Executive Unit Economics & Portfolio Health Summary] |
| ------------------------- |                                                                            |
| Date Range:               | [ KPI 1: Blended CAC ]   [ KPI 2: Total Spend ]   [ KPI 3: New Customers ]  |
| [ 2024-01-01 to 2025-12 ] | [ $71.40 | -12% vs Q3]   [ $248.5K | 24 Mo    ]   [ 3,480 | +18% YoY     ]  |
|                           |                                                                            |
| Channel Select:           | +------------------------------------------------------------------------+ |
| [x] Search  [x] Social    | | CHART 1 (Plotly Line): Monthly Blended CAC vs Target CPA ($55.00)      | |
| [x] Referral[x] Organic   | | (X: Year-Month [24 Months], Y: Blended CAC in $, Red dashed target line) |
|                           | +------------------------------------------------------------------------+ |
| Plan Tier:                | [ CHART 2 (Plotly Bar): Spend by Channel ] | [ CHART 3 (Donut): Customer Share ] |
| [All Tiers Selected    v] | (Search $78K, Social $92K, Referral $22K)  | (Search 30%, Social 24%, Ref 18%)   |
+--------------------------------------------------------------------------------------------------------+"""),

        ("UI-03: Multi-Stage Conversion Funnel & Leakage Analyzer",
"""+--------------------------------------------------------------------------------------------------------+
| CONVERSION FUNNEL & LEAKAGE ANALYZER (UI-03)                                                           |
+--------------------------------------------------------------------------------------------------------+
| SIDEBAR FILTERS           | [ KPI 1: Total Leads ]   [ KPI 2: Avg CPC ]   [ KPI 3: Avg CPL ]           |
| Channel: [ Paid Social v] | [ 4,820 | Qualified  ]   [ $1.84 | Low CPC]   [ $19.37 | Mid-Funnel]       |
| Cohort:  [ All Periods v] |                                                                            |
|                           | +------------------------------------------------------------------------+ |
| Metric Focus:             | | CHART (Plotly Funnel): Stage-by-Stage Attrition Leakage Analysis       | |
| (*) Drop-Off Rate (%)     | | Stage 1: Impressions (1,240,000) ===================================== | |
| ( ) Cumulative Conversion | | Stage 2: Ad Clicks     (28,400)    [ -97.7% drop-off at Click ]        | |
|                           | | Stage 3: Leads          (3,820)    [ -86.5% drop-off at Lead Capture]  | |
| [x] Highlight Leakage     | | Stage 4: Paying Cust      (142)    [ -96.3% CRITICAL LEAKAGE AT CHECKOUT] |
|     Threshold (> 90%)     | +------------------------------------------------------------------------+ |
|                           | [ INSIGHT BOX (Alert): Paid Social suffers severe checkout friction (<4% conv) ] |
+--------------------------------------------------------------------------------------------------------+"""),

        ("UI-05: Inferential EDA & Statistical Hypothesis Testing",
"""+--------------------------------------------------------------------------------------------------------+
| INFERENTIAL STATISTICAL TESTING & COST VARIANCE ANALYSIS (UI-05)                                       |
+--------------------------------------------------------------------------------------------------------+
| TEST CONTROLS             | [ STATISTICAL TEST 1: Normality Check (Shapiro-Wilk) ]                     |
| Hypothesis:               | W-Statistic: 0.842 | p-value: 0.00012 | Verdict: Non-Normal Distribution (Log) |
| H0: Channel CAC means =   |                                                                            |
| H1: Channel CAC differs   | [ STATISTICAL TEST 2: One-Way ANOVA (Parametric) ]                         |
|                           | F-Statistic: 38.45 | p-value: < 0.00001 | F-Critical: 2.41 (alpha = 0.05)       |
| Significance Level:       | Verdict: REJECT NULL HYPOTHESIS. Channel cost differences are genuine.     |
| Slider: [ alpha = 0.05  ] |                                                                            |
|                           | [ STATISTICAL TEST 3: Kruskal-Wallis (Non-Parametric Rank Sum Test) ]      |
| [ Run Statistical Engine ]| H-Statistic: 114.28 | p-value: < 0.00001 | Robust to non-normal distributions   |
|                           | +------------------------------------------------------------------------+ |
| Output Mode:              | | CHART (Plotly Boxplot): CAC Distribution & Outlier Spread by Channel   | |
| (*) Detailed Table        | | (Shows Organic CAC median $24 vs Paid Social median $94 with whiskers)  | |
| ( ) Academic Summary Text | +------------------------------------------------------------------------+ |
+--------------------------------------------------------------------------------------------------------+"""),

        ("UI-06: Strategic Unit Economics & Budget Reallocation Sandbox",
"""+--------------------------------------------------------------------------------------------------------+
| STRATEGIC UNIT ECONOMICS & BUDGET ALLOCATION SANDBOX (UI-06)                                          |
+--------------------------------------------------------------------------------------------------------+
| SIMULATOR SLIDERS         | [ LTV:CAC HEALTH MATRIX BY PLAN TIER ]                                     |
| Reallocate Social Spend:  | Basic: CAC $62 | LTV $180 | Ratio 2.9x [!] SUB-OPTIMAL (< 3.0x Threshold)  |
| [ -20% (-$18,400)       ] | Pro:   CAC $98 | LTV $420 | Ratio 4.3x [v] HEALTHY                         |
|                           | SMB:   CAC $195| LTV $1,240| Ratio 6.4x [v] EXCELLENT                       |
| Reallocate to Referral:   | Ent:   CAC $340| LTV $2,850| Ratio 8.4x [v] OUTSTANDING                     |
| [ +20% (+$18,400)       ] |                                                                            |
|                           | [ CHURN SENSITIVITY SLIDER ]                                               |
| Churn Sensitivity Slider: | Adjust Churn: [======|===== +10% Churn ] -> Recalculated LTV drops 8.5%    |
| [ Base Rate (0%)       v] |                                                                            |
|                           | +------------------------------------------------------------------------+ |
| [ Run Reallocation Sim ]  | | CHART (Plotly Line): Diminishing Marginal ROAS vs Budget Scaling       | |
|                           | | (Displays revenue saturation inflection point at $12,000 monthly spend)| |
+--------------------------------------------------------------------------------------------------------+""")
    ]

    for title, wireframe_text in wireframes:
        add_h2(title)
        box = doc.add_table(rows=1, cols=1)
        box.alignment = WD_TABLE_ALIGNMENT.CENTER
        c = box.rows[0].cells[0]
        c.width = Inches(6.5)
        set_cell_background(c, "F8FAFC")
        set_cell_margins(c, 60, 60, 80, 80)
        cp = c.paragraphs[0]
        cr = cp.add_run(wireframe_text)
        cr.font.name = 'Consolas'
        cr.font.size = Pt(8)
        cr.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # 3. Global Layout & Theme
    add_h1("3. Global Layout Architecture & Design System")
    doc.add_paragraph(
        "• State Management via `st.session_state`: Global filters (`date_range`, `selected_channels`, `selected_tiers`) are stored in session state, ensuring user selections persist across multi-page navigation without resetting.\n"
        "• Color Token System:\n"
        "  - Brand Navy: `#1E3A8A` (Tailwind `blue-900`) for headers and primary action states\n"
        "  - Accent Cobalt: `#2563EB` (Tailwind `blue-600`) for chart elements and active tabs\n"
        "  - Background Slate: `#F8FAFC` (Tailwind `slate-50`) for clean eye fatigue reduction\n"
        "  - Solvency Safe Green: `#16A34A` (Tailwind `green-600`) for payback < 12 months & LTV:CAC >= 3.0\n"
        "  - Critical Alert Red: `#DC2626` (Tailwind `red-600`) for payback > 12 months & value-destroying tiers\n"
        "• Mobile Responsiveness: Uses Tailwind flexbox wrap classes (`flex flex-wrap gap-4 min-w-[200px]`), ensuring metric cards gracefully wrap into a single vertical column on smartphones and 2x2 or 4x1 grids on wide screens."
    )
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # 4. Best Chart per OQ
    add_h1("4. Best Chart Type Selected per Operational Question (OQ)")
    doc.add_paragraph("Analytical rationale justifying each visual encoding:")
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    oq_chart_table = doc.add_table(rows=10, cols=4)
    oq_chart_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(oq_chart_table, "CBD5E1")
    ch_headers = ["OQ ID", "Operational Question", "Best Chart Type & Data Fields", "Analytical & Visual Rationale"]
    ch_widths = [Inches(0.8), Inches(2.0), Inches(2.0), Inches(1.7)]

    for j, h in enumerate(ch_headers):
        c = oq_chart_table.rows[0].cells[j]
        c.width = ch_widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 70, 70, 60, 60)
        r = c.paragraphs[0].add_run(h)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    oq_chart_data = [
        ("OQ-01", "Blended Monthly CAC Trend", "Line Chart with Reference Threshold\n(X: Month, Y: Blended CAC, dashed line: Target)", "Line charts accurately display continuous longitudinal time-series progression against fixed benchmarks."),
        ("OQ-02", "Channel CAC vs Conversion", "Dual-Axis Grouped Bar / Line Chart\n(Bar: CAC $, Line: Conversion %)", "Allows direct comparison of two differing units ($ and %) to spot low-cost, high-conversion channels."),
        ("OQ-03", "LTV:CAC by Customer Tier", "Horizontal Segmented Bar Chart\n(X: LTV:CAC Ratio, Y: Tiers, Red line at 3.0x)", "Horizontal bars provide clean readability for tier labels; reference line highlights viability gates."),
        ("OQ-04", "Capital Payback in Months", "Color-Coded Bullet / Progress Chart\n(X: Payback Months, Red > 12m, Green <= 12m)", "Instantly highlights capital solvency risk using intuitive traffic-light thresholds."),
        ("OQ-05", "Channel Cost Significance", "Box-and-Whisker Plot with Outliers\n(X: Channels, Y: CAC distribution, median notch)", "Visualizes group medians, interquartile ranges (IQR), and skewness required for ANOVA diagnostics."),
        ("OQ-06", "Seasonal Price Surges", "Heatmap / Rolling Index Area Chart\n(X: Year-Month, Y: Surge Index %)", "Heatmaps expose calendar clustering of Q4 holiday ad inflation across multiple years."),
        ("OQ-07", "Funnel Drop-Off Leakage", "Plotly Funnel Chart\n(Stages: Impressions -> Clicks -> Leads -> Sales)", "Standard visual grammar for pipeline conversion; width directly reflects volume attrition."),
        ("OQ-08", "Marginal ROAS Scaling", "Logarithmic Scatter with Fitted Curve\n(X: Spend Scaling %, Y: Marginal ROAS)", "Exposes the non-linear diminishing returns inflection point where extra budget becomes unprofitable."),
        ("OQ-09", "Segment Profitability", "Stacked Waterfall Chart\n(MRR -> Gross Profit -> Servicing Cost -> Net)", "Reveals exact dollar cost deductions between gross revenue and net cash contribution.")
    ]

    for i, r_data in enumerate(oq_chart_data, 1):
        for j, text in enumerate(r_data):
            c = oq_chart_table.rows[i].cells[j]
            c.width = ch_widths[j]
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            set_cell_margins(c, 50, 50, 60, 60)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 5. Reusable Components & Starter UI.py
    add_h1("5. Reusable Tailwind Components & Starter `src/ui.py`")
    doc.add_paragraph(
        "The starter module `src/ui.py` is implemented and verified. It provides three core visual primitives:\n"
        "1. `render_header(title, subtitle)`: Embeds an executive gradient hero banner with author metadata.\n"
        "2. `render_kpi_cards(kpi_data)`: Renders responsive Tailwind flexbox metric cards with status indicator pills.\n"
        "3. `render_insight_box(type, title, text)`: Renders contextual alerts (`warning`, `success`, `critical`, `info`).\n"
        "All components are rendered through `st.components.v1.html(..., height=...)` using Tailwind CSS loaded from CDN."
    )

    doc_path = os.path.abspath("docs/12_UI_Framework.docx")
    doc.save(doc_path)
    print(f"Generated UI Framework docx: {doc_path}")

    # =========================================================================
    # UPDATE RTM EXCEL (docs/04_Traceability_Matrix.xlsx) - Fill UI Screen (Col 11)
    # =========================================================================
    rtm_path = os.path.abspath("docs/04_Traceability_Matrix.xlsx")
    wb_rtm = openpyxl.load_workbook(rtm_path)
    ws_rtm = wb_rtm["RTM_Traceability_Matrix"]

    br_to_ui = {
        "BR-01": "UI-01, UI-02",
        "BR-02": "UI-01, UI-02",
        "BR-03": "UI-03",
        "BR-04": "UI-02, UI-06",
        "BR-05": "UI-06",
        "BR-06": "UI-06",
        "BR-07": "UI-05",
        "BR-08": "UI-06",
        "BR-09": "UI-02",
        "BR-10": "UI-04",
        "BR-11": "UI-08",
        "BR-12": "UI-06"
    }

    ui_font = Font(name="Calibri", size=9.5, bold=True, color="166534") # Dark Green

    for row in range(5, 21):
        br_id = ws_rtm.cell(row=row, column=3).value
        if br_id in br_to_ui:
            ui_val = br_to_ui[br_id]
            c_ui = ws_rtm.cell(row=row, column=11, value=ui_val)
            c_ui.font = ui_font
            c_ui.alignment = Alignment(horizontal="center", vertical="center")

    wb_rtm.save(rtm_path)
    print(f"Updated RTM with UI Screens: {rtm_path}")

if __name__ == "__main__":
    build_ui_framework_doc()

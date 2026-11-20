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

def build_needs_analysis_doc():
    doc = Document()

    for s in doc.sections:
        s.top_margin = Inches(1)
        s.bottom_margin = Inches(1)
        s.left_margin = Inches(1)
        s.right_margin = Inches(1)

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)

    # Header
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = title.add_run("CUSTOMER ACQUISITION COST (CAC) ANALYSIS")
    tr.font.size = Pt(22)
    tr.font.bold = True
    tr.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sub.add_run("STEP 3: BUSINESS NEEDS ANALYSIS & REQUIREMENTS (BR) SPECIFICATION\nMaster of Computer Applications (Data Analytics) — ABES Engineering College")
    sr.font.size = Pt(12)
    sr.font.italic = True
    sr.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Meta Table
    meta_table = doc.add_table(rows=5, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(meta_table, "CBD5E1")
    meta = [
        ("Candidate Name", "Gourav Kumar (Roll: 2500320140055)"),
        ("Academic Program", "MCA 2nd Year (Data Analytics), ABES Engineering College"),
        ("Faculty Mentor", "Prof. Alpna Lodhi"),
        ("Document Identifier", "DOC-STEP-03 | Business Requirements (BR) Specification"),
        ("Methodological Rule", "Technology-Agnostic: Defines WHAT the business needs, NOT HOW.")
    ]
    for i, (k, v) in enumerate(meta):
        row = meta_table.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = Inches(2.3), Inches(4.2)
        set_cell_background(c0, "F1F5F9")
        set_cell_margins(c0, 80, 80, 120, 120)
        set_cell_margins(c1, 80, 80, 120, 120)
        r0 = c0.paragraphs[0].add_run(k)
        r0.font.bold = True
        r0.font.size = Pt(9.5)
        r1 = c1.paragraphs[0].add_run(v)
        r1.font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    def add_h1(text):
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(14)
        h.paragraph_format.space_after = Pt(4)
        h.paragraph_format.keep_with_next = True
        r = h.add_run(text)
        r.font.size = Pt(14)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
        return h

    # 1. Executive Context & Methodology
    add_h1("1. Needs Analysis Methodology")
    p = doc.add_paragraph(
        "This document derives formal Business Requirements (BR-01 through BR-12) from the design-thinking HMW opportunities "
        "defined in Step 2. In strict accordance with business analysis standards (IIBA BABOK), all requirements are framed "
        "using the standardized capability-benefit template: \"The business needs to <capability> so that <benefit>.\" "
        "Requirements strictly define business capabilities (WHAT) and intentionally omit implementation technologies (HOW)."
    )
    p.paragraph_format.space_after = Pt(8)

    # 2. Business Requirements Table
    add_h1("2. Business Requirements Specification Matrix")
    p = doc.add_paragraph("Comprehensive inventory of 12 business requirements prioritized using the MoSCoW framework (~58% Must):")
    p.paragraph_format.space_after = Pt(6)

    br_table = doc.add_table(rows=13, cols=6)
    br_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(br_table, "CBD5E1")
    headers = ["BR ID", "Business Requirement Statement", "Primary Stakeholder", "Source HMW", "MoSCoW Priority", "Measurement & Verification Method"]
    widths = [Inches(0.7), Inches(2.4), Inches(0.9), Inches(0.7), Inches(0.7), Inches(1.4)]

    for j, h_text in enumerate(headers):
        c = br_table.rows[0].cells[j]
        c.width = widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 90, 90, 70, 70)
        r = c.paragraphs[0].add_run(h_text)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    br_data = [
        ("BR-01", "The business needs to capture and attribute marketing campaign expenditures by distinct acquisition channel so that non-performing channels can be isolated from blended aggregates.", "SH-01", "HMW-01", "Must", "100% of campaign spend records tagged to a validated acquisition channel."),
        ("BR-02", "The business needs to calculate and compare channel-specific CAC against blended CAC on monthly and quarterly timeframes so that underperforming paid campaigns are identified immediately.", "SH-01", "HMW-01", "Must", "Reconciled monthly channel CAC reports generated with zero unallocated cost variance."),
        ("BR-03", "The business needs to monitor prospective customer progression through multi-stage conversion funnels so that top-of-funnel drop-offs and capital leakage are detected.", "SH-01", "HMW-02", "Must", "Conversion rate and cost-per-stage calculated for every defined funnel milestone."),
        ("BR-04", "The business needs to model channel-specific capital payback periods against an enterprise benchmark of 12 months so that finance can guarantee working capital solvency.", "SH-03", "HMW-04", "Must", "Months-to-breakeven calculated per channel; visual flags triggered when payback > 12 months."),
        ("BR-05", "The business needs to evaluate Customer Lifetime Value against acquisition costs (LTV:CAC ratio) across distinct customer subscription tiers so that acquisition spend is restricted on churn-prone segments.", "SH-02, SH-03", "HMW-05", "Must", "LTV:CAC ratio generated per customer tier with threshold alerts for ratios falling below 3.0."),
        ("BR-06", "The business needs to evaluate customer segment profitability and cash payback velocity across plan tiers so that sales and marketing resources focus on high-margin customer cohorts.", "SH-03", "HMW-09", "Must", "Net contribution margin and payback velocity computed for each subscription plan category."),
        ("BR-07", "The business needs to conduct hypothesis testing on channel cost distributions so that performance differences are proven statistically significant at alpha = 0.05.", "SH-04, SH-05", "HMW-07, HMW-10", "Must", "Output of p-values, test statistics (F/H), and formal hypothesis rejection/acceptance reporting."),
        ("BR-08", "The business needs to simulate the impact of marketing budget scaling and channel reallocation on marginal ROAS so that capital is not over-allocated past the point of diminishing returns.", "SH-02", "HMW-03", "Should", "Simulated ROAS curves generated under varied budget shift scenarios (+/- 10% to 50%)."),
        ("BR-09", "The business needs to track historical monthly and seasonal acquisition cost trends so that upcoming quarterly campaign budgets can be insulated against high-demand auction price surges.", "SH-01", "HMW-06", "Should", "Month-over-month CAC change rates and seasonal index calculated across 24 historical months."),
        ("BR-10", "The business needs to evaluate channel volume-versus-conversion efficiency trade-offs so that marketing teams can select the optimal channel mix between low-cost leads and high-converting customers.", "SH-02", "HMW-08", "Should", "Matrix ranking channels on dual axes: cost per conversion vs total conversion volume."),
        ("BR-11", "The business needs to provide an auditable calculation catalog detailing mathematical definitions, metric formulas, and lineage so that internal stakeholders and evaluators can verify data integrity.", "SH-05", "HMW-10", "Could", "Complete documented catalog of all business formulas and data definitions accessible within the platform."),
        ("BR-12", "The business needs to dynamically adjust customer churn rate assumptions in LTV calculations so that scenario planning can account for macroeconomic changes.", "SH-03", "HMW-05", "Could", "Real-time recalculation of LTV and LTV:CAC when baseline churn parameter is modified by user.")
    ]

    for i, row in enumerate(br_data, start=1):
        for j, text in enumerate(row):
            c = br_table.rows[i].cells[j]
            c.width = widths[j]
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            set_cell_margins(c, 70, 70, 70, 70)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True
            if j == 4:
                r.font.bold = True
                if text == "Must":
                    r.font.color.rgb = RGBColor(0x99, 0x1B, 0x1B) # Dark Red
                elif text == "Should":
                    r.font.color.rgb = RGBColor(0x1E, 0x40, 0xAF) # Blue
                else:
                    r.font.color.rgb = RGBColor(0x4B, 0x55, 0x63) # Gray

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 3. Data Requirements Mapping
    add_h1("3. Required Data Entities & Attributes")
    p = doc.add_paragraph("Data assets strictly required to operationalize each business requirement:")
    p.paragraph_format.space_after = Pt(6)

    data_table = doc.add_table(rows=13, cols=3)
    data_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(data_table, "CBD5E1")
    d_headers = ["BR ID", "Core Requirement Focus", "Required Data Attributes & Entities"]
    d_widths = [Inches(0.9), Inches(2.3), Inches(3.6)]

    for j, h_text in enumerate(d_headers):
        c = data_table.rows[0].cells[j]
        c.width = d_widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 90, 90, 80, 80)
        r = c.paragraphs[0].add_run(h_text)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    data_map = [
        ("BR-01", "Channel Spend Attribution", "Campaign ID, Channel Name (Search, Social, Referral, Email, Organic), Start Date, End Date, Daily/Monthly Ad Spend."),
        ("BR-02", "Channel vs. Blended CAC", "Monthly Marketing Spend per Channel, Total Aggregate Spend, New Customers Converted per Channel, Calendar Month/Quarter."),
        ("BR-03", "Multi-stage Funnel Leakage", "Impressions, Ad Clicks, Leads Captured, Completed Sales / Conversions, Direct Stage Costs."),
        ("BR-04", "Capital Payback Modeling", "Channel-level CAC, Average Revenue Per User (ARPU), Gross Margin % by Plan, Subscription Duration."),
        ("BR-05", "Tiered LTV:CAC Ratios", "Customer ID, Subscription Plan Tier (Enterprise, SMB, Pro, Basic), Monthly Recurring Revenue (MRR), Churn Rate %, Attribution Channel."),
        ("BR-06", "Segment Contribution & Payback", "Plan Tier, Cohort Inception Date, Accumulated Revenue, Allocated Servicing Cost, Gross Margin Contribution."),
        ("BR-07", "Inferential Statistical Testing", "Per-campaign CAC sample observations across 5 channels, Variance data, Categorical Channel labels."),
        ("BR-08", "Budget Scaling & ROAS Simulation", "Historical Spend per Channel, Attributed Channel Gross Revenue, Scaled Budget Coefficient vectors."),
        ("BR-09", "Seasonal Trend Analysis", "24-Month Temporal Spend logs, Seasonal Indicator Flags (Holiday, Q4, Summer), Month-over-Month Acquisition counts."),
        ("BR-10", "Volume vs. Efficiency Trade-off", "Total Leads per Channel, Total Converted Customers, Lead-to-Customer Ratio, Channel Cost per Customer."),
        ("BR-11", "Formula Catalog & Audit Lineage", "Formula Dictionary, Mathematical Definitions, Input Variable Names, Lineage Cross-references."),
        ("BR-12", "Dynamic Churn Sensitivity", "Baseline Churn Rate %, User-selected Churn Multiplier (-50% to +50%), Adjusted Expected Customer Lifespan.")
    ]

    for i, row in enumerate(data_map, start=1):
        for j, text in enumerate(row):
            c = data_table.rows[i].cells[j]
            c.width = d_widths[j]
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            set_cell_margins(c, 70, 70, 80, 80)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 4. Gap Analysis
    add_h1("4. Traceability Gap Analysis")
    p = doc.add_paragraph(
        "A bidirectional verification audit was conducted between Step 2 HMW statements and Step 3 Business Requirements "
        "to ensure zero orphaned items:"
    )
    p.paragraph_format.space_after = Pt(6)

    gap_table = doc.add_table(rows=3, cols=4)
    gap_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(gap_table, "CBD5E1")
    g_headers = ["Audit Dimension", "Evaluated Scope", "Observed Gaps", "Audit Status"]
    g_widths = [Inches(1.8), Inches(1.8), Inches(1.8), Inches(1.1)]

    for j, h_text in enumerate(g_headers):
        c = gap_table.rows[0].cells[j]
        c.width = g_widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 90, 90, 80, 80)
        r = c.paragraphs[0].add_run(h_text)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    gap_data = [
        ("HMW Statements with No BR", "HMW-01 through HMW-10 (10 statements)", "0 Gaps (All 10 HMWs mapped to >=1 BR)", "PASSED (100%)"),
        ("BRs with No Source HMW", "BR-01 through BR-12 (12 requirements)", "0 Gaps (All 12 BRs originate from valid HMWs)", "PASSED (100%)")
    ]

    for i, row in enumerate(gap_data, start=1):
        for j, text in enumerate(row):
            c = gap_table.rows[i].cells[j]
            c.width = g_widths[j]
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            set_cell_margins(c, 70, 70, 80, 80)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(9)
            if j == 3:
                r.font.bold = True
                r.font.color.rgb = RGBColor(0x16, 0x65, 0x34)

    target_path = os.path.abspath("docs/03_Needs_Analysis.docx")
    doc.save(target_path)
    print(f"Generated: {target_path}")

if __name__ == "__main__":
    build_needs_analysis_doc()

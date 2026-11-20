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

def build_hmw_doc():
    doc = Document()

    # Margins
    for s in doc.sections:
        s.top_margin = Inches(1)
        s.bottom_margin = Inches(1)
        s.left_margin = Inches(1)
        s.right_margin = Inches(1)

    # Styles
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
    sr = sub.add_run("STEP 2: HOW MIGHT WE (HMW) DESIGN THINKING SPECIFICATION\nMaster of Computer Applications (Data Analytics) — ABES Engineering College")
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
        ("Document Identifier", "DOC-STEP-02 | HMW Statements & Pain Point Mapping"),
        ("System Baseline", "Linked directly to Step 1 OQs (OQ-01 through OQ-09) and SHs (SH-01 through SH-05)")
    ]
    for i, (k, v) in enumerate(meta):
        row = meta_table.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = Inches(2.3), Inches(4.2)
        set_cell_background(c0, "F1F5F9")
        set_cell_margins(c0, 80, 80, 120, 120)
        set_cell_margins(c1, 80, 80, 120, 120)
        p0 = c0.paragraphs[0]
        r0 = p0.add_run(k)
        r0.font.bold = True
        r0.font.size = Pt(9.5)
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(v)
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

    # 1. Stakeholder Pain Points
    add_h1("1. Stakeholder Pain Point Synthesis")
    p = doc.add_paragraph(
        "Through design-thinking discovery, each primary stakeholder's operational friction points were identified "
        "and mapped to ensure the project addresses concrete business challenges:"
    )
    p.paragraph_format.space_after = Pt(6)

    pp_table = doc.add_table(rows=6, cols=3)
    pp_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(pp_table, "CBD5E1")
    headers = ["Stakeholder ID & Name", "Organizational Role", "Core Operational Pain Points"]
    widths = [Inches(1.8), Inches(1.5), Inches(3.2)]

    for j, h_text in enumerate(headers):
        c = pp_table.rows[0].cells[j]
        c.width = widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 90, 90, 100, 100)
        r = c.paragraphs[0].add_run(h_text)
        r.font.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    pp_data = [
        ("SH-01: Head of Digital Marketing", "Strategic Sponsor", 
         "1. Blended CAC figures mask severe ad budget waste in specific paid channels.\n2. Lack of visibility into funnel leakage points between ad click, lead capture, and payment.\n3. Seasonal bid escalations cause unexpected quarterly budget burn."),
        ("SH-02: VP of Growth & Strategy", "Business User", 
         "1. Difficulty identifying which customer subscription tiers yield sustainable lifetime unit economics.\n2. Unclear diminishing marginal returns when aggressively scaling channel ad budgets.\n3. Friction in balancing aggressive acquisition volume against lead conversion quality."),
        ("SH-03: Finance Controller / CFO", "Financial Governance", 
         "1. Extended capital payback periods (>12 months) locking up organizational working capital.\n2. Customer cohorts exhibiting high early churn despite high initial acquisition spend (LTV:CAC < 3:1).\n3. Lack of unified analytical reconciliation between marketing reports and ledger entries."),
        ("SH-04: Data Analyst (Gourav Kumar)", "System Developer & QA", 
         "1. Disparate marketing logs with inconsistent attribution touchpoints across campaigns.\n2. Difficulty demonstrating statistical significance of cost variances to non-technical leaders.\n3. Need for automated testing to ensure math logic parity between database and UI layers."),
        ("SH-05: Academic Mentor (Prof. Alpna Lodhi)", "Faculty Evaluator", 
         "1. Academic analytics projects often stop at descriptive charts without formal hypothesis testing.\n2. Incomplete traceability between business operational questions and visual analytics deliverables.\n3. Lack of reproducible, automated unit test verification.")
    ]

    for i, row in enumerate(pp_data, start=1):
        for j, text in enumerate(row):
            c = pp_table.rows[i].cells[j]
            c.width = widths[j]
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            set_cell_margins(c, 80, 80, 90, 90)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(9)
            if j == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 2. HMW Table
    add_h1("2. Design Thinking 'How Might We' (HMW) Statements")
    p = doc.add_paragraph(
        "Each pain point is reframed into an actionable, problem-oriented HMW statement. "
        "Solutions are intentionally unconstrained to foster optimal analytical and visual design."
    )
    p.paragraph_format.space_after = Pt(6)

    hmw_table = doc.add_table(rows=11, cols=5)
    hmw_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(hmw_table, "CBD5E1")
    hmw_headers = ["HMW ID", "How Might We ...? Statement", "Stakeholder", "Addressed Pain Point", "Covered OQ"]
    hmw_widths = [Inches(0.8), Inches(2.4), Inches(1.0), Inches(1.5), Inches(0.8)]

    for j, h_text in enumerate(hmw_headers):
        c = hmw_table.rows[0].cells[j]
        c.width = hmw_widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 90, 90, 80, 80)
        r = c.paragraphs[0].add_run(h_text)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    hmw_rows = [
        ("HMW-01", "How might we enable marketing leaders to isolate channel-specific acquisition costs from blended averages so that underperforming paid ad campaigns can be adjusted in real time?", "SH-01", "Blended CAC masking channel inefficiencies", "OQ-01, OQ-02"),
        ("HMW-02", "How might we help digital marketing teams pinpoint the exact funnel stage where prospective customers drop off before converting so that wasted top-of-funnel spend can be minimized?", "SH-01", "Funnel budget leakage before checkout", "OQ-07"),
        ("HMW-03", "How might we assist growth leaders in identifying the optimal budget scaling threshold for each channel before diminishing marginal returns erode overall ROAS?", "SH-02", "Unclear diminishing returns when scaling ad budgets", "OQ-08"),
        ("HMW-04", "How might we empower finance executives to continuously track and benchmark capital payback duration across channels so that working capital is recovered within 12 months?", "SH-03", "Extended payback periods locking up capital", "OQ-04"),
        ("HMW-05", "How might we provide executives with transparent visibility into LTV:CAC ratios across distinct customer subscription tiers to prevent high-cost acquisition of churn-prone segments?", "SH-03, SH-02", "Unhealthy LTV:CAC and churn risk", "OQ-03, OQ-09"),
        ("HMW-06", "How might we help campaign planners anticipate and insulate marketing budgets against seasonal cost-per-click and acquisition spikes during high-demand quarters?", "SH-01", "Seasonal bid surges distorting quarterly CAC", "OQ-06"),
        ("HMW-07", "How might we enable data analysts to statistically validate whether observed channel cost variances represent genuine performance differences rather than random sampling noise?", "SH-04, SH-05", "Lack of statistical rigor in channel comparison", "OQ-05"),
        ("HMW-08", "How might we guide growth managers to evaluate trade-offs between low-cost/low-conversion channels versus high-cost/high-conversion channels when pursuing volume growth?", "SH-02", "Balancing volume targets against conversion efficiency", "OQ-02"),
        ("HMW-09", "How might we help finance and marketing stakeholders align on customer segment profitability so that sales efforts prioritize tiers with the fastest payback velocity?", "SH-03", "Segment profitability alignment", "OQ-09"),
        ("HMW-10", "How might we ensure academic and industry evaluators can verify end-to-end data integrity from raw transaction logs to statistical validation without ambiguity?", "SH-05", "Ensuring demonstrable analytical rigor", "OQ-01, OQ-05")
    ]

    for i, row in enumerate(hmw_rows, start=1):
        for j, text in enumerate(row):
            c = hmw_table.rows[i].cells[j]
            c.width = hmw_widths[j]
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            set_cell_margins(c, 80, 80, 80, 80)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 3. Priority HMW Statements
    add_h1("3. Top 5 Priority HMW Statements")
    p = doc.add_paragraph(
        "To maximize business impact and deliver the core analytical mandate, the top 5 statements are selected "
        "as architectural priorities with explicit one-line business rationale:"
    )
    p.paragraph_format.space_after = Pt(6)

    prio_table = doc.add_table(rows=6, cols=3)
    prio_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(prio_table, "CBD5E1")
    p_headers = ["Priority Rank", "HMW Statement", "Architectural & Business Rationale"]
    p_widths = [Inches(1.2), Inches(2.5), Inches(2.8)]

    for j, h_text in enumerate(p_headers):
        c = prio_table.rows[0].cells[j]
        c.width = p_widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 90, 90, 80, 80)
        r = c.paragraphs[0].add_run(h_text)
        r.font.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    prio_data = [
        ("PRIORITY 1\n(HMW-01)", "How might we enable marketing leaders to isolate channel-specific acquisition costs from blended averages so that underperforming paid ad campaigns can be adjusted in real time?", "Without granular channel-level CAC attribution, blended metrics obscure unprofitable ad campaigns and cause severe capital misallocation."),
        ("PRIORITY 2\n(HMW-04)", "How might we empower finance executives to continuously track and benchmark capital payback duration across channels so that working capital is recovered within 12 months?", "Payback duration is the decisive liquidity metric determining enterprise cash runway and marketing budget re-approval."),
        ("PRIORITY 3\n(HMW-05)", "How might we provide executives with transparent visibility into LTV:CAC ratios across distinct customer subscription tiers to prevent high-cost acquisition of churn-prone segments?", "Sustainable unit economics require an LTV:CAC ratio exceeding 3:1 to prevent high CAC expenditure on rapidly churning customer tiers."),
        ("PRIORITY 4\n(HMW-03)", "How might we assist growth leaders in identifying the optimal budget scaling threshold for each channel before diminishing marginal returns erode overall ROAS?", "Prevents marketing teams from overspending into saturation curves where customer acquisition yields diminishing marginal revenue."),
        ("PRIORITY 5\n(HMW-07)", "How might we enable data analysts to statistically validate whether observed channel cost variances represent genuine performance differences rather than random sampling noise?", "Elevates the project from a descriptive visualization dashboard into an academically sound, inferential data analytics system.")
    ]

    for i, row in enumerate(prio_data, start=1):
        for j, text in enumerate(row):
            c = prio_table.rows[i].cells[j]
            c.width = p_widths[j]
            if i % 2 == 1:
                set_cell_background(c, "EFF6FF")
            set_cell_margins(c, 80, 80, 80, 80)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(9)
            if j == 0:
                r.font.bold = True
                r.font.color.rgb = RGBColor(0x1E, 0x40, 0xAF)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 4. Coverage Matrix Check
    add_h1("4. Operational Question (OQ) Coverage Verification")
    p = doc.add_paragraph("Verification ensuring complete 100% mapping from Step 1 Operational Questions to Step 2 HMW Statements:")
    p.paragraph_format.space_after = Pt(6)

    cov_table = doc.add_table(rows=10, cols=3)
    cov_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(cov_table, "CBD5E1")
    c_headers = ["OQ ID", "Operational Question Core Theme", "Covering HMW Statement(s)"]
    c_widths = [Inches(0.9), Inches(3.6), Inches(2.0)]

    for j, h_text in enumerate(c_headers):
        c = cov_table.rows[0].cells[j]
        c.width = c_widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 90, 90, 80, 80)
        r = c.paragraphs[0].add_run(h_text)
        r.font.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    cov_data = [
        ("OQ-01", "Overall blended CAC trends across channels", "HMW-01, HMW-10"),
        ("OQ-02", "Channel CAC vs. conversion rate trade-off", "HMW-01, HMW-08"),
        ("OQ-03", "LTV:CAC ratio by customer tier", "HMW-05"),
        ("OQ-04", "Capital payback period in months", "HMW-04"),
        ("OQ-05", "Statistical significance of cost variance across channels", "HMW-07, HMW-10"),
        ("OQ-06", "Seasonal spikes and holiday quarter inflation", "HMW-06"),
        ("OQ-07", "Top-of-funnel drop-off and conversion leakage", "HMW-02"),
        ("OQ-08", "Marginal return on ad spend (ROAS) budget curves", "HMW-03"),
        ("OQ-09", "Segment-level unit economics and profitability", "HMW-05, HMW-09")
    ]

    for i, row in enumerate(cov_data, start=1):
        for j, text in enumerate(row):
            c = cov_table.rows[i].cells[j]
            c.width = c_widths[j]
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            set_cell_margins(c, 70, 70, 80, 80)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(9)
            if j == 0:
                r.font.bold = True
            if j == 2:
                r.font.bold = True
                r.font.color.rgb = RGBColor(0x16, 0x65, 0x34)

    target_path = os.path.abspath("docs/02_HMW_Statements.docx")
    doc.save(target_path)
    print(f"Generated: {target_path}")

if __name__ == "__main__":
    build_hmw_doc()

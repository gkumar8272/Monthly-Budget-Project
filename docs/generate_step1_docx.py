import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=120, bottom=120, left=150, right=150):
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

def set_table_borders(table, color="D3D3D3"):
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

def create_document():
    doc = Document()

    # Page Margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Styles
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(0x2D, 0x37, 0x48) # Slate 800

    # Header / Title Block
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run("CUSTOMER ACQUISITION COST (CAC) ANALYSIS")
    title_run.font.name = 'Calibri'
    title_run.font.size = Pt(24)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A) # Deep Navy

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = subtitle.add_run("STEP 1: PROJECT UNDERSTANDING DOCUMENT\nMaster of Computer Applications (Data Analytics) — 2nd Year")
    sub_run.font.size = Pt(13)
    sub_run.font.italic = True
    sub_run.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    metadata = [
        ("Project ID", "P083"),
        ("Candidate Name", "Gourav Kumar"),
        ("University Roll Number", "2500320140055"),
        ("Batch ID", "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"),
        ("Academic Program & Year", "MCA 2nd Year (Data Analytics), ABES Engineering College"),
        ("Faculty Mentor / Supervisor", "Mr. Dinesh Sereno"),
        ("Domain & Technology Stack", "Data Analytics | Python, Streamlit, Plotly, SQLite, Pytest")
    ]

    # Metadata Table
    meta_table = doc.add_table(rows=len(metadata), cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False
    set_table_borders(meta_table, "CBD5E1")

    for i, (k, v) in enumerate(metadata):
        row = meta_table.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.3)
        c1.width = Inches(4.2)
        set_cell_background(c0, "F1F5F9")
        set_cell_margins(c0, 100, 100, 140, 140)
        set_cell_margins(c1, 100, 100, 140, 140)
        p0 = c0.paragraphs[0]
        r0 = p0.add_run(k)
        r0.font.bold = True
        r0.font.size = Pt(10)
        r0.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(v)
        r1.font.size = Pt(10)

    doc.add_paragraph().paragraph_format.space_after = Pt(14)

    def add_h1(text):
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(16)
        h.paragraph_format.space_after = Pt(6)
        h.paragraph_format.keep_with_next = True
        r = h.add_run(text)
        r.font.name = 'Calibri'
        r.font.size = Pt(15)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
        return h

    def add_h2(text):
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(12)
        h.paragraph_format.space_after = Pt(4)
        h.paragraph_format.keep_with_next = True
        r = h.add_run(text)
        r.font.name = 'Calibri'
        r.font.size = Pt(12)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
        return h

    # 1. Executive Summary
    add_h1("1. Executive Summary")
    p = doc.add_paragraph(
        "In modern multi-channel digital commerce, businesses allocate substantial financial resources across search engine marketing (SEM), "
        "paid social, display advertising, email marketing, and referral channels. However, marketing teams frequently struggle with fragmented "
        "cost allocation and attribution opacity, leading to inflated Customer Acquisition Costs (CAC), sub-optimal channel spending, and extended payback periods. "
        "This project develops an enterprise-grade Customer Acquisition Cost Analysis application built with Python, Streamlit, Plotly, and SQLite. "
        "The system calculates blended and channel-specific CAC, correlates acquisition spend with Customer Lifetime Value (LTV), tracks payback cycles, "
        "and delivers actionable prescriptive insights through modern interactive visualizations."
    )
    p.paragraph_format.space_after = Pt(10)

    # 2. Problem Statement
    add_h1("2. Problem Statement")
    callout = doc.add_table(rows=1, cols=1)
    callout.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_cell = callout.rows[0].cells[0]
    c_cell.width = Inches(6.5)
    set_cell_background(c_cell, "EFF6FF")
    set_cell_margins(c_cell, 150, 150, 180, 180)
    cp = c_cell.paragraphs[0]
    cr = cp.add_run(
        "\"Modern digital enterprises fail to achieve sustainable unit economics due to fragmented marketing spend tracking and lack of granular attribution, "
        "resulting in unmonitored channel-level Customer Acquisition Cost (CAC) inflation, suboptimal capital allocation, and distorted Lifetime-Value-to-CAC (LTV:CAC) ratios.\""
    )
    cr.font.size = Pt(11.5)
    cr.font.bold = True
    cr.font.italic = True
    cr.font.color.rgb = RGBColor(0x1E, 0x40, 0xAF)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 3. Intended Audience and User Personas
    add_h1("3. Intended Audience & Users")
    p = doc.add_paragraph(
        "The application serves functional enterprise stakeholders across growth marketing, finance, and executive leadership. "
        "Usage patterns, frequencies, and workflows are structured as follows:"
    )
    p.paragraph_format.space_after = Pt(6)

    aud_table = doc.add_table(rows=4, cols=4)
    aud_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(aud_table, "CBD5E1")
    headers = ["User Role / Persona", "Access Frequency", "Primary Objective in App", "Key Interactivity Utilized"]
    col_widths = [Inches(1.6), Inches(1.2), Inches(2.2), Inches(1.5)]

    for j, h_text in enumerate(headers):
        cell = aud_table.rows[0].cells[j]
        cell.width = col_widths[j]
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 100, 100, 100, 100)
        p = cell.paragraphs[0]
        r = p.add_run(h_text)
        r.font.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    aud_data = [
        ("Performance Marketing Specialist", "Daily / Weekly", "Audit real-time campaign spend, click-through efficiency, and cost-per-lead across paid channels.", "Channel filter, date picker, campaign drilldown"),
        ("Growth Marketing Manager", "Weekly / Bi-weekly", "Evaluate channel conversion velocity, organic vs. paid CAC trends, and reallocate budget to top ROI channels.", "Comparative CAC bar charts, budget simulator"),
        ("Chief Financial Officer / VP Finance", "Monthly / Quarterly", "Assess payback period benchmarks, LTV:CAC health ratios, and ensure sustainable unit economic margins.", "Summary KPI cards, LTV:CAC ratio matrix, export reports")
    ]

    for i, row_data in enumerate(aud_data, start=1):
        for j, text in enumerate(row_data):
            cell = aud_table.rows[i].cells[j]
            cell.width = col_widths[j]
            if i % 2 == 1:
                set_cell_background(cell, "F8FAFC")
            set_cell_margins(cell, 90, 90, 100, 100)
            p = cell.paragraphs[0]
            r = p.add_run(text)
            r.font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 4. Stakeholder Analysis Matrix
    add_h1("4. Stakeholder Analysis Matrix")
    p = doc.add_paragraph("Enterprise stakeholders categorized with permanent project identifiers (SH-01 to SH-05):")
    p.paragraph_format.space_after = Pt(6)

    sh_table = doc.add_table(rows=6, cols=6)
    sh_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(sh_table, "CBD5E1")
    sh_headers = ["SH ID", "Stakeholder", "Organizational Role", "Core Interest", "Power / Influence", "Key Decisions Made with App"]
    sh_widths = [Inches(0.7), Inches(1.3), Inches(1.2), Inches(1.3), Inches(0.8), Inches(1.7)]

    for j, h_text in enumerate(sh_headers):
        cell = sh_table.rows[0].cells[j]
        cell.width = sh_widths[j]
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 100, 100, 80, 80)
        p = cell.paragraphs[0]
        r = p.add_run(h_text)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    sh_rows = [
        ("SH-01", "Head of Digital Marketing", "Strategic Sponsor", "Channel efficiency, blended CAC reduction, lead quality", "High", "Quarterly marketing spend allocation across Google Ads, Meta, and Organic SEO."),
        ("SH-02", "VP of Growth & Strategy", "Business User", "Customer acquisition volume, market expansion, LTV optimization", "High", "Setting customer acquisition targets and launching new marketing initiatives."),
        ("SH-03", "Finance Controller / CFO", "Governance & Finance", "Return on ad spend (ROAS), cash payback cycle, unit margin", "High", "Approving marketing budgets and enforcing maximum CAC thresholds (< 12-month payback)."),
        ("SH-04", "Data Analyst (Student Author)", "System Developer & QA", "Analytical integrity, schema normalization, data pipelines, visual UX", "Medium", "Pipeline design, metric formulas, anomaly thresholds, and feature rollouts."),
        ("SH-05", "Faculty Mentor / Project Evaluator", "Academic Supervisor", "Academic rigor, EDA depth, statistical testing, ABES rubric compliance", "High", "Evaluating project methodology, architectural soundness, and viva performance.")
    ]

    for i, row_data in enumerate(sh_rows, start=1):
        for j, text in enumerate(row_data):
            cell = sh_table.rows[i].cells[j]
            cell.width = sh_widths[j]
            if i % 2 == 1:
                set_cell_background(cell, "F8FAFC")
            set_cell_margins(cell, 80, 80, 80, 80)
            p = cell.paragraphs[0]
            r = p.add_run(text)
            r.font.size = Pt(9)
            if j == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 5. Operational Questions & Data Feasibility
    add_h1("5. Operational Questions (OQ) & Data Feasibility")
    p = doc.add_paragraph(
        "To ensure business utility, the application is engineered around 9 core operational questions (OQ-01 through OQ-09). "
        "Each question is verified against data availability to guarantee 100% computational feasibility."
    )
    p.paragraph_format.space_after = Pt(6)

    oq_table = doc.add_table(rows=10, cols=5)
    oq_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(oq_table, "CBD5E1")
    oq_headers = ["OQ ID", "Operational Business Question", "Primary Metric(s)", "Required Data Entities", "Feasibility Check"]
    oq_widths = [Inches(0.7), Inches(2.2), Inches(1.3), Inches(1.5), Inches(0.8)]

    for j, h_text in enumerate(oq_headers):
        cell = oq_table.rows[0].cells[j]
        cell.width = oq_widths[j]
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 100, 100, 80, 80)
        p = cell.paragraphs[0]
        r = p.add_run(h_text)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    oq_rows = [
        ("OQ-01", "What is the overall blended Customer Acquisition Cost (CAC) across all channels on a monthly basis?", "Blended CAC ($/customer)", "Total Marketing Spend, Total New Customers Acquired", "Yes (Spend / New Customers)"),
        ("OQ-02", "Which marketing acquisition channel delivers the lowest CAC versus the highest conversion rate?", "Channel CAC, Conversion Rate (%)", "Channel Spend, Leads Generated, Customers Converted", "Yes (Channel-grouped metrics)"),
        ("OQ-03", "How does Customer Lifetime Value (LTV) compare to CAC (LTV:CAC ratio) across distinct customer tiers?", "LTV:CAC Ratio, Margin Contribution", "Customer Revenue, Gross Margin, Cohort Spend", "Yes (LTV / Channel CAC)"),
        ("OQ-04", "What is the capital payback period (in months) required to recoup CAC per acquisition channel?", "Payback Period (Months)", "CAC, Average Monthly Recurring Revenue (MRR)", "Yes (CAC / [ARPU * Margin])"),
        ("OQ-05", "Are there statistically significant cost efficiency differences between Paid Search, Paid Social, Organic, and Referral?", "ANOVA F-Statistic, P-value, Kruskal-Wallis", "Per-campaign CAC distributions across channels", "Yes (Hypothesis testing on CAC distribution)"),
        ("OQ-06", "How do seasonal spikes and holiday quarters impact channel-level acquisition inflation?", "Month-over-Month CAC % Growth", "Time-series Date, Spend by Day/Week/Month", "Yes (Temporal aggregation & trends)"),
        ("OQ-07", "What proportion of marketing expenditure is consumed by top-of-funnel drop-offs before conversion?", "Cost Per Click (CPC), Cost Per Lead (CPL), Funnel Loss %", "Impressions, Clicks, Leads, Conversions, Channel Cost", "Yes (Funnel stage attrition analysis)"),
        ("OQ-08", "What is the marginal return on ad spend (ROAS) as monthly budget increases per channel?", "ROAS (Revenue / Spend), Marginal ROAS curve", "Attributed Customer Revenue, Channel Ad Spend", "Yes (Regression / ROAS modeling)"),
        ("OQ-09", "Which customer segments (e.g., Enterprise vs. SMB vs. Individual) yield the healthiest unit economics?", "Segment CAC, Segment LTV, Payback Speed", "Customer Demographic/Plan Type, Spend attribution", "Yes (Segment-sliced CAC & LTV calculations)")
    ]

    for i, row_data in enumerate(oq_rows, start=1):
        for j, text in enumerate(row_data):
            cell = oq_table.rows[i].cells[j]
            cell.width = oq_widths[j]
            if i % 2 == 1:
                set_cell_background(cell, "F8FAFC")
            set_cell_margins(cell, 80, 80, 80, 80)
            p = cell.paragraphs[0]
            r = p.add_run(text)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True
            if j == 4:
                r.font.color.rgb = RGBColor(0x16, 0x65, 0x34) # Dark Green

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 6. SMART Objectives
    add_h1("6. SMART Project Objectives")
    p = doc.add_paragraph("The technical and analytical deliverables of this project are structured around 5 SMART objectives:")
    p.paragraph_format.space_after = Pt(6)

    obj_table = doc.add_table(rows=6, cols=3)
    obj_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(obj_table, "CBD5E1")
    obj_headers = ["Objective ID", "SMART Objective Statement", "Target Metric & Timeline"]
    obj_widths = [Inches(1.0), Inches(4.0), Inches(1.5)]

    for j, h_text in enumerate(obj_headers):
        cell = obj_table.rows[0].cells[j]
        cell.width = obj_widths[j]
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 100, 100, 80, 80)
        p = cell.paragraphs[0]
        r = p.add_run(h_text)
        r.font.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    obj_rows = [
        ("OBJ-01", "Design and generate a realistic, synthetic multi-channel acquisition dataset spanning 24 months (minimum 15,000 lead records and campaign spend logs across 5 channels) with built-in seasonal patterns and funnel drop-offs.", "15,000+ records, 100% synthetic, completed by Step 10"),
        ("OBJ-02", "Implement an automated data pipeline using SQLite and Pandas that calculates Blended CAC, Channel CAC, Payback Period, and LTV:CAC ratios with zero reconciliation discrepancy.", "0% computational error vs. manual benchmarks, completed by Step 14"),
        ("OBJ-03", "Conduct rigorous exploratory data analysis (EDA) and inferential statistical testing (ANOVA and Kruskal-Wallis) to validate acquisition cost variances across channels at alpha = 0.05.", "P-values and effect sizes computed with automated reporting"),
        ("OBJ-04", "Develop a responsive, modern Streamlit web application styled with Tailwind CSS visual hierarchy, featuring dynamic Plotly dashboards, multi-parameter filters, and automated budget allocation simulations.", "Dashboard load time < 2.0s, 100% interactive filters"),
        ("OBJ-05", "Achieve comprehensive software quality assurance with >85% unit test coverage via pytest, followed by seamless deployment on Streamlit Community Cloud accompanied by ABES documentation.", ">85% test coverage, 0 failed tests, public cloud deployment URL")
    ]

    for i, row_data in enumerate(obj_rows, start=1):
        for j, text in enumerate(row_data):
            cell = obj_table.rows[i].cells[j]
            cell.width = obj_widths[j]
            if i % 2 == 1:
                set_cell_background(cell, "F8FAFC")
            set_cell_margins(cell, 80, 80, 80, 80)
            p = cell.paragraphs[0]
            r = p.add_run(text)
            r.font.size = Pt(9)
            if j == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 7. Project Scope: In-Scope vs. Out-of-Scope
    add_h1("7. Project Scope Boundary")
    p = doc.add_paragraph("Clear boundaries are established to ensure realistic delivery within the college academic timeline:")
    p.paragraph_format.space_after = Pt(6)

    scope_table = doc.add_table(rows=7, cols=3)
    scope_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(scope_table, "CBD5E1")
    scope_headers = ["Functional Area", "In-Scope (Deliverables Built)", "Out-of-Scope (Deferred / Excluded)"]
    scope_widths = [Inches(1.5), Inches(2.6), Inches(2.4)]

    for j, h_text in enumerate(scope_headers):
        cell = scope_table.rows[0].cells[j]
        cell.width = scope_widths[j]
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 100, 100, 80, 80)
        p = cell.paragraphs[0]
        r = p.add_run(h_text)
        r.font.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    scope_rows = [
        ("Data Ingestion & Storage", "Synthetic transactional SQLite relational database with seed scripts; CSV/Excel import module.", "Live API integrations with Facebook Graph API, Google Ads API, or live CRM webhooks."),
        ("CAC & Unit Metrics", "Blended CAC, Channel-specific CAC, LTV:CAC ratio, Payback period (months), Return on Ad Spend (ROAS).", "Algorithmic multi-touch Markov chain attribution or deep machine learning MTA modeling."),
        ("Statistical Analysis", "Descriptive statistics, distribution normality tests (Shapiro-Wilk), One-way ANOVA, Kruskal-Wallis tests.", "Predictive neural networks or deep auto-regressive spending forecasters."),
        ("User Interface", "Streamlit multi-page web application, interactive Plotly charts, Tailwind CSS cards, KPI summaries.", "Native mobile app (Android/iOS) or enterprise React microfrontend."),
        ("Security & Auth", "Local session state, basic role-selection simulation (Marketing vs. Finance view).", "Multi-tenant OAuth2/SAML SSO, RBAC permission microservices."),
        ("Testing & Deployment", "Automated pytest test suite (data transformations, calculation logic), Streamlit Community Cloud hosting.", "Kubernetes cluster orchestration or multi-region AWS high-availability failover.")
    ]

    for i, row_data in enumerate(scope_rows, start=1):
        for j, text in enumerate(row_data):
            cell = scope_table.rows[i].cells[j]
            cell.width = scope_widths[j]
            if i % 2 == 1:
                set_cell_background(cell, "F8FAFC")
            set_cell_margins(cell, 80, 80, 80, 80)
            p = cell.paragraphs[0]
            r = p.add_run(text)
            r.font.size = Pt(9)
            if j == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 8. Assumptions and Constraints
    add_h1("8. Assumptions and Constraints")
    p = doc.add_paragraph("Explicit operational parameters, assumptions, and engineering constraints:")
    p.paragraph_format.space_after = Pt(6)

    ac_table = doc.add_table(rows=6, cols=4)
    ac_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(ac_table, "CBD5E1")
    ac_headers = ["Category", "ID", "Description of Assumption / Constraint", "Architectural Impact & Mitigation"]
    ac_widths = [Inches(1.1), Inches(0.7), Inches(2.7), Inches(2.0)]

    for j, h_text in enumerate(ac_headers):
        cell = ac_table.rows[0].cells[j]
        cell.width = ac_widths[j]
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 100, 100, 80, 80)
        p = cell.paragraphs[0]
        r = p.add_run(h_text)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    ac_rows = [
        ("Assumption", "ASM-01", "Synthetic data statistically mirrors real-world SaaS/E-commerce acquisition distributions with log-normal spend and conversion behavior.", "Mitigates confidentiality risks; ensures student project complies with academic non-disclosure rules."),
        ("Assumption", "ASM-02", "Customer touchpoints are attributed using Last-Touch and First-Touch models for baseline CAC computations.", "Simplifies SQL aggregation pipelines while preserving robust analytical validity."),
        ("Constraint", "CST-01", "Hosting is constrained to Streamlit Community Cloud (free tier: 1 GB RAM, shared virtual CPU).", "Database queries and Plotly figures must be memory-efficient, utilizing cached queries (@st.cache_data)."),
        ("Constraint", "CST-02", "Single relational database instance (SQLite3) without distributed locking.", "Appropriate for analytical dashboards with read-heavy operations and pre-seeded demonstration data."),
        ("Constraint", "CST-03", "Academic semester timeline restriction (18 structured sequential phases).", "Scope is modularized; each step produces a standalone deliverable verified against ABES rubrics.")
    ]

    for i, row_data in enumerate(ac_rows, start=1):
        for j, text in enumerate(row_data):
            cell = ac_table.rows[i].cells[j]
            cell.width = ac_widths[j]
            if i % 2 == 1:
                set_cell_background(cell, "F8FAFC")
            set_cell_margins(cell, 80, 80, 80, 80)
            p = cell.paragraphs[0]
            r = p.add_run(text)
            r.font.size = Pt(8.5)
            if j in (0, 1):
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 9. Measurable Success Criteria
    add_h1("9. Measurable Success Criteria")
    p = doc.add_paragraph("Validation gates determining project acceptance:")
    p.paragraph_format.space_after = Pt(6)

    sc_table = doc.add_table(rows=6, cols=4)
    sc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(sc_table, "CBD5E1")
    sc_headers = ["Criteria ID", "Success Factor", "Target Benchmark / Metric", "Verification Method"]
    sc_widths = [Inches(0.9), Inches(1.8), Inches(2.1), Inches(1.7)]

    for j, h_text in enumerate(sc_headers):
        cell = sc_table.rows[0].cells[j]
        cell.width = sc_widths[j]
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 100, 100, 80, 80)
        p = cell.paragraphs[0]
        r = p.add_run(h_text)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    sc_rows = [
        ("SC-01", "Analytical Accuracy", "100% calculation parity between SQLite SQL queries and Pandas calculations for CAC, LTV, and Payback.", "Automated pytest mathematical assertion tests."),
        ("SC-02", "Application Performance", "Dashboard cold start < 3.0s; cached filter interactions and chart renders < 800ms.", "Chrome DevTools network audit & Streamlit runtime timing."),
        ("SC-03", "Test Suite Coverage", "Minimum 85% line and branch test coverage across core analytical and database modules.", "pytest --cov=src execution report."),
        ("SC-04", "UI Usability & Aesthetics", "Zero layout clipping across 1080p and 720p screens; modern Tailwind styling with clear visual hierarchy.", "Peer UI walkthrough and heuristic evaluation."),
        ("SC-05", "Deployment & Documentation", "Fully live, zero-downtime Streamlit Cloud deployment + 100% compliant ABES major project report.", "Mentor review, live URL verification, rubric grading.")
    ]

    for i, row_data in enumerate(sc_rows, start=1):
        for j, text in enumerate(row_data):
            cell = sc_table.rows[i].cells[j]
            cell.width = sc_widths[j]
            if i % 2 == 1:
                set_cell_background(cell, "F8FAFC")
            set_cell_margins(cell, 80, 80, 80, 80)
            p = cell.paragraphs[0]
            r = p.add_run(text)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True

    # Save to disk
    target_path_1 = os.path.abspath("docs/01_Project_Understanding.docx")
    target_path_2 = os.path.abspath("docs/01_Customer_Acquisition_Cost_Project_Understanding.docx")
    for tp in [target_path_1, target_path_2]:
        try:
            doc.save(tp)
            print(f"Generated: {tp}")
        except PermissionError:
            fallback = tp.replace(".docx", "_Updated.docx")
            doc.save(fallback)
            print(f"File was open in Word. Saved to: {fallback}")

if __name__ == "__main__":
    create_document()

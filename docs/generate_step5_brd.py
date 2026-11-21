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

def build_brd_doc():
    doc = Document()

    for s in doc.sections:
        s.top_margin = Inches(1)
        s.bottom_margin = Inches(1)
        s.left_margin = Inches(1)
        s.right_margin = Inches(1)

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(10.5)
    normal_style.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)

    # Document Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("CUSTOMER ACQUISITION COST (CAC) ANALYSIS")
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("STEP 5: BUSINESS REQUIREMENTS DOCUMENT (BRD)\nBaseline Version 1.0 — Enterprise Analytics Specification")
    r_sub.font.size = Pt(12)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    def add_h1(text):
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(14)
        h.paragraph_format.space_after = Pt(4)
        h.paragraph_format.keep_with_next = True
        r = h.add_run(text)
        r.font.size = Pt(13.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
        return h

    def add_h2(text):
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(10)
        h.paragraph_format.space_after = Pt(3)
        h.paragraph_format.keep_with_next = True
        r = h.add_run(text)
        r.font.size = Pt(11.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
        return h

    # 1. Document Control
    add_h1("1. Document Control & Revision History")
    p = doc.add_paragraph("Formal administrative metadata and version tracking:")
    p.paragraph_format.space_after = Pt(4)

    dc_table = doc.add_table(rows=6, cols=4)
    dc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(dc_table, "CBD5E1")
    dc_headers = ["Version", "Release Date", "Primary Author", "Summary of Changes / Baseline"]
    dc_widths = [Inches(0.9), Inches(1.2), Inches(1.8), Inches(2.6)]

    for j, h in enumerate(dc_headers):
        c = dc_table.rows[0].cells[j]
        c.width = dc_widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 80, 80, 80, 80)
        r = c.paragraphs[0].add_run(h)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    dc_data = [
        ("v0.1", "2026-11-18", "Gourav Kumar (Roll: 2500320140055)", "Draft Problem Statement and Stakeholder Discovery (Step 1)."),
        ("v0.2", "2026-11-18", "Gourav Kumar (Roll: 2500320140055)", "HMW Statements & Design Thinking Mapping (Step 2)."),
        ("v0.3", "2026-11-18", "Gourav Kumar (Roll: 2500320140055)", "Business Needs Analysis & MoSCoW Prioritization (Step 3)."),
        ("v0.4", "2026-11-18", "Gourav Kumar (Roll: 2500320140055)", "Requirements Traceability Matrix (RTM) Baseline (Step 4)."),
        ("v1.0", "2026-11-18", "Gourav Kumar (Roll: 2500320140055)", "Final Comprehensive BRD Baselined for PRD & Architecture.")
    ]
    for i, row in enumerate(dc_data, start=1):
        for j, text in enumerate(row):
            c = dc_table.rows[i].cells[j]
            c.width = dc_widths[j]
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            set_cell_margins(c, 70, 70, 80, 80)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 2. Executive Summary
    add_h1("2. Executive Summary")
    p = doc.add_paragraph(
        "Modern digital commerce and SaaS organizations invest heavily across diverse acquisition channels, including Search Engine Marketing (Google Ads), "
        "Paid Social (Meta, LinkedIn), Organic SEO, Email Marketing, and Partner Referrals. However, organizations frequently suffer from fragmented cost reporting, "
        "opaque channel attribution, and unmonitored acquisition inflation. Blended metrics obscure underperforming ad campaigns, leading to extended cash payback periods "
        "and degraded Lifetime-Value-to-Customer-Acquisition-Cost (LTV:CAC) ratios.\n\n"
        "This Business Requirements Document (BRD) defines the formal functional and business boundaries for an end-to-end Customer Acquisition Cost Analysis system. "
        "The platform ingests multi-channel campaign expenditures, isolates paid versus organic acquisition economics, models capital payback velocity against an enterprise "
        "12-month solvency ceiling, validates cost variances via inferential statistical testing (ANOVA/Kruskal-Wallis), and empowers stakeholders to simulate dynamic budget allocations."
    )
    p.paragraph_format.space_after = Pt(10)

    # 3. Business Problem & Opportunity
    add_h1("3. Business Problem & Market Opportunity")
    add_h2("3.1 Problem Statement")
    doc.add_paragraph(
        "\"Modern digital enterprises fail to achieve sustainable unit economics due to fragmented marketing spend tracking and lack of granular attribution, "
        "resulting in unmonitored channel-level Customer Acquisition Cost (CAC) inflation, suboptimal capital allocation, and distorted Lifetime-Value-to-CAC (LTV:CAC) ratios.\""
    )
    add_h2("3.2 Business Opportunity & Strategic Value")
    doc.add_paragraph(
        "By implementing a centralized, transparent CAC analytics platform, the organization captures significant strategic value:\n"
        "• Elimination of Unprofitable Ad Spend: Direct visibility into channel-specific CAC enables marketing leaders to trim budget from underperforming ad groups, recapturing 15-25% of misallocated capital.\n"
        "• Working Capital Protection: Enforcing automated alerts on channels exceeding a 12-month payback horizon protects enterprise cash reserves.\n"
        "• Alignment of Marketing and Finance: Bridges the gap between top-of-funnel lead generation metrics and bottom-line general ledger unit economics."
    )
    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 4. Objectives & Success Metrics
    add_h1("4. Business Objectives & Key Performance Indicators (KPIs)")
    p = doc.add_paragraph("The platform delivers against 5 core SMART objectives (OBJ-01 to OBJ-05) and quantifiable success metrics (SC-01 to SC-05):")
    p.paragraph_format.space_after = Pt(4)

    kpi_table = doc.add_table(rows=6, cols=4)
    kpi_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(kpi_table, "CBD5E1")
    kpi_headers = ["Metric ID", "Strategic Objective", "Quantitative Target", "Verification Mechanism"]
    kpi_widths = [Inches(1.0), Inches(2.2), Inches(1.8), Inches(1.5)]

    for j, h in enumerate(kpi_headers):
        c = kpi_table.rows[0].cells[j]
        c.width = kpi_widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 80, 80, 80, 80)
        r = c.paragraphs[0].add_run(h)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    kpi_data = [
        ("SC-01 / OBJ-02", "Calculation Accuracy", "100% mathematical parity across SQL views and Python pipelines (0% delta).", "Automated pytest reconciliation assertions."),
        ("SC-02 / OBJ-04", "Dashboard Query Latency", "Cold start < 3.0s; cached filter recalculations < 800ms.", "Chrome DevTools & Streamlit cache audit."),
        ("SC-03 / OBJ-05", "Automated QA Test Coverage", ">= 85% branch and line coverage across transformation logic.", "pytest --cov test execution report."),
        ("SC-04 / OBJ-03", "Inferential Statistical Rigor", "Automated ANOVA and Kruskal-Wallis testing at alpha = 0.05.", "Hypothesis test reporting module."),
        ("SC-05 / OBJ-01", "Synthetic Dataset Breadth", ">= 15,000 multi-channel lead logs spanning 24 historical months.", "Database record count and schema validation.")
    ]
    for i, row in enumerate(kpi_data, start=1):
        for j, text in enumerate(row):
            c = kpi_table.rows[i].cells[j]
            c.width = kpi_widths[j]
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            set_cell_margins(c, 70, 70, 80, 80)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 5. Stakeholders - RACI Table
    add_h1("5. Stakeholder RACI Matrix")
    p = doc.add_paragraph("RACI mapping (Responsible, Accountable, Consulted, Informed) across key project lifecycle phases:")
    p.paragraph_format.space_after = Pt(4)

    raci_table = doc.add_table(rows=7, cols=6)
    raci_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(raci_table, "CBD5E1")
    raci_headers = ["Project Lifecycle Phase", "SH-01\n(Marketing)", "SH-02\n(Growth)", "SH-03\n(CFO/Finance)", "SH-04\n(Analyst/Gourav)", "SH-05\n(Mentor/Dinesh)"]
    raci_widths = [Inches(2.5), Inches(0.8), Inches(0.8), Inches(0.8), Inches(0.8), Inches(0.8)]

    for j, h in enumerate(raci_headers):
        c = raci_table.rows[0].cells[j]
        c.width = raci_widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 80, 80, 60, 60)
        r = c.paragraphs[0].add_run(h)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    raci_data = [
        ("1. Requirements & Needs Analysis", "Accountable", "Consulted", "Consulted", "Responsible", "Informed"),
        ("2. Data Schema & Architecture", "Informed", "Informed", "Consulted", "Responsible", "Accountable"),
        ("3. Statistical Modeling & EDA", "Informed", "Consulted", "Informed", "Responsible", "Accountable"),
        ("4. Dashboard UI & Simulation Build", "Consulted", "Accountable", "Informed", "Responsible", "Consulted"),
        ("5. QA Testing & Test Cases", "Informed", "Informed", "Informed", "Responsible", "Accountable"),
        ("6. Cloud Deployment & Final Defense", "Informed", "Informed", "Informed", "Responsible", "Accountable")
    ]
    for i, row in enumerate(raci_data, start=1):
        for j, text in enumerate(row):
            c = raci_table.rows[i].cells[j]
            c.width = raci_widths[j]
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            set_cell_margins(c, 60, 60, 60, 60)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True
            else:
                c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                if text == "Accountable":
                    r.font.bold = True
                    r.font.color.rgb = RGBColor(0x99, 0x1B, 0x1B)
                elif text == "Responsible":
                    r.font.bold = True
                    r.font.color.rgb = RGBColor(0x1E, 0x40, 0xAF)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 6. Scope
    add_h1("6. Project Scope")
    scope_table = doc.add_table(rows=6, cols=3)
    scope_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(scope_table, "CBD5E1")
    s_headers = ["Functional Area", "In-Scope Capabilities", "Out-of-Scope Exclusions"]
    s_widths = [Inches(1.8), Inches(2.4), Inches(2.3)]

    for j, h in enumerate(s_headers):
        c = scope_table.rows[0].cells[j]
        c.width = s_widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 80, 80, 80, 80)
        r = c.paragraphs[0].add_run(h)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    s_data = [
        ("Data Capture & Storage", "Relational SQLite database with seed scripts; 24-month multi-channel marketing logs.", "Live external API sync with Google Ads / Meta Ads Manager."),
        ("Attribution & CAC", "Channel CAC, Blended CAC, LTV:CAC ratios, Payback period in months, ROAS curves.", "Algorithmic multi-touch Markov chain or Shapley machine learning attribution."),
        ("Statistical Testing", "Normality tests (Shapiro-Wilk), One-way ANOVA, and Kruskal-Wallis non-parametric tests.", "Deep recurrent time-series neural networks (LSTM) for predictive forecasting."),
        ("User Experience", "Multi-page interactive Streamlit dashboard with Tailwind-styled cards and Plotly figures.", "Native iOS/Android mobile apps or enterprise React/Angular microfrontends."),
        ("Security & Auth", "Local session state role toggle (Marketing view vs. Executive view).", "Enterprise OAuth2/SAML SSO, RBAC permission directory services.")
    ]
    for i, row in enumerate(s_data, start=1):
        for j, text in enumerate(row):
            c = scope_table.rows[i].cells[j]
            c.width = s_widths[j]
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            set_cell_margins(c, 70, 70, 80, 80)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 7. Business Requirements
    add_h1("7. Detailed Business Requirements Specification")
    p = doc.add_paragraph("Full inventory of 12 business requirements (BR-01 through BR-12) incorporating rationale and formal acceptance criteria:")
    p.paragraph_format.space_after = Pt(4)

    br_full = [
        ("BR-01", "Channel Spend Attribution Capture", 
         "The business needs to capture and attribute all marketing campaign expenditures by distinct acquisition channel so that non-performing channels can be isolated from blended aggregates.",
         "High aggregate ad spend obscures channels that consume budget without yielding paying users.",
         "Must", "HMW-01 (OQ-01, OQ-02)",
         "Given a campaign spend ledger, when records are ingested, then 100% of costs must be attributed to an authorized channel (Search, Social, Referral, Email, Organic) with $0 unallocated variance."),
        
        ("BR-02", "Channel-Specific vs. Blended CAC Computation",
         "The business needs to calculate and compare channel-specific CAC against blended CAC on monthly and quarterly timeframes so that underperforming paid campaigns are identified immediately.",
         "Relying solely on blended CAC allows low-cost organic traffic to subsidize unprofitable paid ad campaigns.",
         "Must", "HMW-01 (OQ-01, OQ-02)",
         "Given monthly channel spend and converted customer records, when CAC is computed, then channel CAC = (Channel Spend / Channel Customers) and Blended CAC = (Total Spend / Total Customers) with 0% reconciliation delta."),
        
        ("BR-03", "Multi-Stage Conversion Funnel Tracking",
         "The business needs to monitor prospective customer progression through multi-stage conversion funnels so that top-of-funnel drop-offs and capital leakage are detected.",
         "Understanding whether acquisition friction occurs at the click, lead capture, or checkout stage dictates marketing optimization.",
         "Must", "HMW-02 (OQ-07)",
         "Given funnel stage milestones (Impression -> Click -> Lead -> Customer), when viewed, then the system must output stage conversion rates and cost per milestone (CPC, CPL) across all channels."),

        ("BR-04", "Capital Payback Period Modeling",
         "The business needs to model channel-specific capital payback periods against an enterprise benchmark of 12 months so that finance can guarantee working capital solvency.",
         "Long payback periods deplete operational cash reserves, threatening organizational solvency.",
         "Must", "HMW-04 (OQ-04)",
         "Given CAC and monthly contribution margin (ARPU * Gross Margin %), when payback is calculated, then Payback Months = CAC / (ARPU * Margin %) and visual warning alerts trigger if Payback > 12.0 months."),

        ("BR-05", "Customer Lifetime Value to CAC (LTV:CAC) Ratio Evaluation",
         "The business needs to evaluate Customer Lifetime Value against acquisition costs (LTV:CAC ratio) across distinct customer subscription tiers so that acquisition spend is restricted on churn-prone segments.",
         "High acquisition expenditure is only justified if customer cohorts deliver long-term gross margin contribution exceeding 3x cost.",
         "Must", "HMW-05 (OQ-03, OQ-09)",
         "Given customer cohort revenue and churn rates, when LTV:CAC is evaluated, then the system displays LTV:CAC ratios per tier and flags any tier with a ratio < 3.0 as sub-optimal."),

        ("BR-06", "Segment Profitability & Payback Velocity",
         "The business needs to evaluate customer segment profitability and cash payback velocity across plan tiers so that sales and marketing resources focus on high-margin customer cohorts.",
         "Enterprise, SMB, Pro, and Basic tiers possess drastically different servicing costs and retention profiles.",
         "Must", "HMW-09 (OQ-09)",
         "Given customer tier subscription data, when segment economics are calculated, then net contribution margin and recovery velocity are generated for each plan type."),

        ("BR-07", "Inferential Statistical Hypothesis Testing",
         "The business needs to conduct hypothesis testing on channel cost distributions so that performance differences are proven statistically significant at alpha = 0.05.",
         "Prevents management from making high-stakes budget shifts based on random sampling variance.",
         "Must", "HMW-07, HMW-10 (OQ-05)",
         "Given per-campaign CAC distributions across 5 channels, when tested, then the system executes One-Way ANOVA and Kruskal-Wallis tests, reporting F/H statistics, p-values, and explicit hypothesis rejection/retention conclusions."),

        ("BR-08", "Budget Scaling & Marginal ROAS Simulation",
         "The business needs to simulate the impact of marketing budget scaling and channel reallocation on marginal ROAS so that capital is not over-allocated past the point of diminishing returns.",
         "Marketing channels exhibit non-linear saturation curves where marginal customer return decays.",
         "Should", "HMW-03 (OQ-08)",
         "Given user-adjusted channel budget sliders (+/- 10% to 50%), when simulated, then the system dynamically recalculates projected customers, revenue, and marginal ROAS curves."),

        ("BR-09", "Seasonal Trend & Price Spike Tracking",
         "The business needs to track historical monthly and seasonal acquisition cost trends so that upcoming quarterly campaign budgets can be insulated against high-demand auction price surges.",
         "Q4 holiday auctions cause severe ad cost inflation that must be anticipated in annual cash planning.",
         "Should", "HMW-06 (OQ-06)",
         "Given 24 months of longitudinal spend data, when rendered, then the system outputs month-over-month CAC growth % and flags seasonal price surges."),

        ("BR-10", "Channel Volume vs. Efficiency Trade-off Matrix",
         "The business needs to evaluate channel volume-versus-conversion efficiency trade-offs so that marketing teams can select the optimal channel mix between low-cost leads and high-converting customers.",
         "Low CAC channels often cannot scale volume, while high-volume channels carry premium ad costs.",
         "Should", "HMW-08 (OQ-02)",
         "Given channel metrics, when analyzed, then the system plots channels on a 2x2 matrix of Cost-Per-Customer versus Absolute Customer Volume."),

        ("BR-11", "Auditable Calculation & Formula Catalog",
         "The business needs to provide an auditable calculation catalog detailing mathematical definitions, metric formulas, and lineage so that internal stakeholders and evaluators can verify data integrity.",
         "Ensures academic and executive auditability without black-box opacity.",
         "Could", "HMW-10 (OQ-01, OQ-05)",
         "Given the system interface, when the user opens the documentation module, then exact mathematical formulas and variable definitions for all 12 metrics are displayed."),

        ("BR-12", "Dynamic Churn Sensitivity Scenario Planning",
         "The business needs to dynamically adjust customer churn rate assumptions in LTV calculations so that scenario planning can account for macroeconomic changes.",
         "Customer churn rates fluctuate during economic downturns, dramatically altering expected LTV.",
         "Could", "HMW-05 (OQ-03)",
         "Given baseline churn rates, when the user adjusts a sensitivity slider (+/- 10% to 50%), then LTV and LTV:CAC ratios immediately update in real time.")
    ]

    for br in br_full:
        add_h2(f"{br[0]}: {br[1]}")
        p = doc.add_paragraph()
        r = p.add_run("Requirement Statement: ")
        r.font.bold = True
        p.add_run(br[2])

        p2 = doc.add_paragraph()
        r2 = p2.add_run("Business Rationale: ")
        r2.font.bold = True
        p2.add_run(br[3])

        p3 = doc.add_paragraph()
        r3a = p3.add_run("Priority: ")
        r3a.font.bold = True
        p3.add_run(f"{br[4]}  |  ")
        r3b = p3.add_run("Source HMW (OQ): ")
        r3b.font.bold = True
        p3.add_run(br[5])

        p4 = doc.add_paragraph()
        r4 = p4.add_run("Acceptance Criteria: ")
        r4.font.bold = True
        p4.add_run(br[6])
        p4.paragraph_format.space_after = Pt(6)

    # 8. Business Rules
    add_h1("8. Business Rules Specification")
    p = doc.add_paragraph("Formal business logic constraints governing the analytical calculations:")
    p.paragraph_format.space_after = Pt(4)

    rules_table = doc.add_table(rows=6, cols=3)
    rules_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(rules_table, "CBD5E1")
    ru_headers = ["Rule ID", "Business Rule Name", "Operational Logic & Mathematical Constraint"]
    ru_widths = [Inches(1.0), Inches(2.2), Inches(3.3)]

    for j, h in enumerate(ru_headers):
        c = rules_table.rows[0].cells[j]
        c.width = ru_widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 80, 80, 80, 80)
        r = c.paragraphs[0].add_run(h)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    rules_data = [
        ("BRUL-01", "Attribution Model Primacy", "Customer acquisition attribution operates on Last-Touch Attribution (LTA) for direct conversions and First-Touch Attribution (FTA) for top-of-funnel lead generation."),
        ("BRUL-02", "Zero-Spend Attribution Integrity", "Organic Search and Referral channels with $0 direct media spend must have CAC computed using strictly allocated content/operational overhead costs to prevent divide-by-zero errors."),
        ("BRUL-03", "Payback Ceiling Threshold", "Any marketing channel exhibiting a calculated CAC payback period greater than 12.0 months is flagged with a high-priority warning state (Status: Critical/Review)."),
        ("BRUL-04", "LTV:CAC Viability Standard", "An LTV:CAC ratio >= 3.0 is classified as Healthy; between 1.0 and 2.9 is classified as Sub-Optimal; below 1.0 is classified as Value-Destroying."),
        ("BRUL-05", "Statistical Significance Gate", "Channel performance differences are declared statistically significant if and only if the calculated p-value < 0.05 (alpha = 0.05). If p >= 0.05, the null hypothesis of equal performance is retained.")
    ]
    for i, row in enumerate(rules_data, start=1):
        for j, text in enumerate(row):
            c = rules_table.rows[i].cells[j]
            c.width = ru_widths[j]
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            set_cell_margins(c, 70, 70, 80, 80)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 9. Data Requirements
    add_h1("9. Data Requirements")
    doc.add_paragraph(
        "• Ingestion Scope: Ingests 24 months of synthetic historical marketing logs across 5 channels (Search, Social, Email, Referral, Organic).\n"
        "• Schema Tables: 4 core relational entities: `channels`, `campaign_spend`, `leads_funnel`, `customer_subscriptions`.\n"
        "• Data Latency & Refresh: Real-time analytical query execution against pre-seeded SQLite database; session cache invalidation on filter change."
    )
    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 10. Assumptions, Constraints, Dependencies
    add_h1("10. Assumptions, Constraints, and Dependencies")
    doc.add_paragraph(
        "• Assumptions (ASM-01, ASM-02): Synthetic dataset mirrors log-normal marketing spend distributions; cohort spend and customer revenue are co-temporal.\n"
        "• Constraints (CST-01, CST-02, CST-03): Deployment constrained to Streamlit Community Cloud (1 GB RAM limit); single-file relational SQLite database; strict adherence to 18-step sequential academic delivery.\n"
        "• Dependencies (DEP-01, DEP-02): Delivery depends on verified Step 4 RTM baseline and successful completion of Step 10 synthetic seed generation."
    )
    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 11. Risks and Mitigation
    add_h1("11. Risks and Mitigation Matrix")
    rsk_table = doc.add_table(rows=5, cols=5)
    rsk_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(rsk_table, "CBD5E1")
    rsk_headers = ["Risk ID", "Risk Description", "Likelihood", "Impact", "Proactive Mitigation Strategy"]
    rsk_widths = [Inches(0.8), Inches(2.2), Inches(0.8), Inches(0.8), Inches(1.9)]

    for j, h in enumerate(rsk_headers):
        c = rsk_table.rows[0].cells[j]
        c.width = rsk_widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 80, 80, 70, 70)
        r = c.paragraphs[0].add_run(h)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    rsk_data = [
        ("RSK-01", "Synthetic data lacks realistic variance, invalidating ANOVA tests.", "Low", "High", "Parameterize data generator with log-normal noise and channel-specific mean offsets."),
        ("RSK-02", "Cloud memory exhaustion on Streamlit Community Cloud (1 GB cap).", "Medium", "High", "Implement lightweight SQLite indexed queries and Streamlit @st.cache_data decorators."),
        ("RSK-03", "Calculation discrepancies between SQL queries and Pandas front-end.", "Medium", "High", "Enforce automated pytest assertion test suites validating math parity before deployment."),
        ("RSK-04", "Academic rubric non-compliance during viva defense.", "Low", "High", "Directly map all deliverables to ABES guidelines with faculty mentor sign-offs at each step.")
    ]
    for i, row in enumerate(rsk_data, start=1):
        for j, text in enumerate(row):
            c = rsk_table.rows[i].cells[j]
            c.width = rsk_widths[j]
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            set_cell_margins(c, 60, 60, 70, 70)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 12. Glossary
    add_h1("12. Glossary of Terms")
    glossary = [
        ("Customer Acquisition Cost (CAC)", "The aggregate financial investment required to convince a prospective customer to purchase, calculated as total sales and marketing spend divided by new customers acquired."),
        ("Customer Lifetime Value (LTV)", "The total gross profit generated by a customer relationship over their entire active tenure with the enterprise."),
        ("LTV:CAC Ratio", "A fundamental unit economic indicator measuring ROI on customer acquisition. A ratio >= 3.0 indicates sustainable commercial scalability."),
        ("CAC Payback Period", "The duration in months required for a newly acquired customer to generate enough gross profit to recoup the initial CAC outlay."),
        ("Return on Ad Spend (ROAS)", "Gross revenue generated directly per dollar of advertising spend (Revenue / Media Spend)."),
        ("ANOVA (Analysis of Variance)", "A parametric statistical test used to determine if statistically significant differences exist between the means of three or more independent channel groups."),
        ("Kruskal-Wallis Test", "A non-parametric statistical counterpart to One-Way ANOVA evaluated on ranked data when distributions violate normality.")
    ]
    for term, definition in glossary:
        p = doc.add_paragraph()
        r1 = p.add_run(f"• {term}: ")
        r1.font.bold = True
        p.add_run(definition)
    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 13. Sign-off
    add_h1("13. Formal Stakeholder Sign-Off")
    p = doc.add_paragraph("By signing below, the stakeholders baseline this Business Requirements Document (BRD v1.0):")
    p.paragraph_format.space_after = Pt(6)

    so_table = doc.add_table(rows=4, cols=5)
    so_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(so_table, "CBD5E1")
    so_headers = ["Stakeholder Name", "Organizational Role", "Sign-Off Decision", "Approval Date", "Signature"]
    so_widths = [Inches(1.8), Inches(1.8), Inches(1.2), Inches(1.0), Inches(0.7)]

    for j, h in enumerate(so_headers):
        c = so_table.rows[0].cells[j]
        c.width = so_widths[j]
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 80, 80, 60, 60)
        r = c.paragraphs[0].add_run(h)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    so_data = [
        ("Gourav Kumar", "Lead Student Author / Developer", "APPROVED & BASELINED", "2026-11-18", "GK"),
        ("Mr. Dinesh Sereno", "Academic Mentor / Project Supervisor", "APPROVED & BASELINED", "2026-11-18", "DS"),
        ("Department Project Committee", "ABES MCA Academic Review Board", "APPROVED FOR DESIGN", "2026-11-18", "ABES")
    ]
    for i, row in enumerate(so_data, start=1):
        for j, text in enumerate(row):
            c = so_table.rows[i].cells[j]
            c.width = so_widths[j]
            if i % 2 == 1:
                set_cell_background(c, "F8FAFC")
            set_cell_margins(c, 70, 70, 60, 60)
            r = c.paragraphs[0].add_run(text)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True
            if j == 2:
                r.font.bold = True
                r.font.color.rgb = RGBColor(0x16, 0x65, 0x34)

    target_path = os.path.abspath("docs/05_BRD.docx")
    doc.save(target_path)
    print(f"Generated: {target_path}")

if __name__ == "__main__":
    build_brd_doc()

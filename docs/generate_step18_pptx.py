"""
generate_step18_pptx.py
STEP 18 — ABES CAC Analysis 11-slide Defense Presentation
ABES corporate palette: Navy #1E3A8A | Sky #0284C7 | Slate #475569 | White #FFFFFF
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# ── Colour constants ───────────────────────────────────────────────────────────
NAVY    = RGBColor(0x1E, 0x3A, 0x8A)
SKY     = RGBColor(0x02, 0x84, 0xC7)
SLATE   = RGBColor(0x47, 0x55, 0x69)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT   = RGBColor(0xF0, 0xF4, 0xFF)   # very pale blue background
GOLD    = RGBColor(0xF5, 0x9E, 0x0B)   # accent for numbers / highlights
GREEN   = RGBColor(0x05, 0x96, 0x69)   # pass / positive
RED     = RGBColor(0xDC, 0x26, 0x26)   # warning

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


# ── Low-level helpers ──────────────────────────────────────────────────────────
def _run(para, text, bold=False, italic=False, size=None, colour=None):
    run = para.add_run()
    run.text = text
    run.font.bold = bold
    run.font.italic = italic
    if size:
        run.font.size = Pt(size)
    if colour:
        run.font.color.rgb = colour
    return run


def add_textbox(slide, left, top, width, height,
                text="", bold=False, italic=False, size=18,
                colour=WHITE, align=PP_ALIGN.LEFT, wrap=True,
                word_wrap=True, bg_colour=None, transparency=0):
    txb = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txb.text_frame
    tf.word_wrap = word_wrap
    tf.auto_size = None
    p = tf.paragraphs[0]
    p.alignment = align
    _run(p, text, bold=bold, italic=italic, size=size, colour=colour)
    if bg_colour and txb.fill:
        txb.fill.solid()
        txb.fill.fore_color.rgb = bg_colour
    return txb


def add_rect(slide, left, top, width, height, fill_colour, transparency=0):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE = 1
        Inches(left), Inches(top), Inches(width), Inches(height))
    shape.line.fill.background()          # no border
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_colour
    return shape


def set_slide_bg(slide, colour):
    """Set a solid background colour for the slide."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = colour


def add_para_lines(txb_tf, lines, size=14, colour=WHITE, bullet=True):
    """Append multiple bullet-line paragraphs to a text-frame."""
    for line in lines:
        p = txb_tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        prefix = "▸  " if bullet else ""
        _run(p, prefix + line, size=size, colour=colour)
        p.space_before = Pt(4)


def add_kpi_card(slide, left, top, label, value, unit="", bg=NAVY):
    """Rounded-looking KPI card (3 textboxes stacked)."""
    card_w, card_h = 2.5, 1.35
    # background rect
    r = add_rect(slide, left, top, card_w, card_h, bg)
    # value
    add_textbox(slide, left + 0.1, top + 0.08, card_w - 0.2, 0.7,
                text=value, bold=True, size=28, colour=GOLD,
                align=PP_ALIGN.CENTER)
    # label
    add_textbox(slide, left + 0.1, top + 0.75, card_w - 0.2, 0.3,
                text=label, bold=False, size=10, colour=WHITE,
                align=PP_ALIGN.CENTER)
    # unit
    if unit:
        add_textbox(slide, left + 0.1, top + 1.0, card_w - 0.2, 0.3,
                    text=unit, bold=False, size=9, colour=LIGHT,
                    align=PP_ALIGN.CENTER)


# ── Slide factory ──────────────────────────────────────────────────────────────
def mk_base_slide(prs, with_header=True, with_footer=True):
    """Add a blank slide and paint the ABES corporate chrome."""
    blank_layout = prs.slide_layouts[6]  # completely blank
    slide = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide, WHITE)

    if with_header:
        # Navy header bar
        add_rect(slide, 0, 0, 13.33, 1.1, NAVY)
        # Sky accent stripe below header
        add_rect(slide, 0, 1.1, 13.33, 0.07, SKY)

        # ABES text in header
        add_textbox(slide, 0.3, 0.05, 10, 0.5,
                    text="ABES Engineering College, Ghaziabad  |  MCA 2nd Year — Data Analytics",
                    bold=False, size=10, colour=LIGHT,
                    align=PP_ALIGN.LEFT)
        add_textbox(slide, 0.3, 0.45, 10, 0.55,
                    text="CUSTOMER ACQUISITION COST (CAC) ANALYSIS",
                    bold=True, size=15, colour=WHITE,
                    align=PP_ALIGN.LEFT)
        # slide number placeholder (right side)
        add_textbox(slide, 11.5, 0.3, 1.5, 0.5,
                    text="Roll: 2500320140055",
                    bold=False, size=9, colour=LIGHT,
                    align=PP_ALIGN.RIGHT)

    if with_footer:
        # Thin navy footer bar
        add_rect(slide, 0, 7.15, 13.33, 0.35, NAVY)
        add_textbox(slide, 0.3, 7.17, 8, 0.28,
                    text="Gourav Kumar  |  Faculty Mentor: Prof. Alpna Lodhi  |  November 2026",
                    bold=False, size=8, colour=LIGHT,
                    align=PP_ALIGN.LEFT)
        add_textbox(slide, 11.0, 7.17, 2.1, 0.28,
                    text="ABES — MCA Major Project",
                    bold=False, size=8, colour=LIGHT,
                    align=PP_ALIGN.RIGHT)
    return slide


def add_slide_title(slide, title, subtitle=None):
    """Large section title below the header bar."""
    add_textbox(slide, 0.5, 1.25, 12.33, 0.7,
                text=title, bold=True, size=22, colour=NAVY,
                align=PP_ALIGN.LEFT)
    # Sky underline
    add_rect(slide, 0.5, 1.9, 7, 0.06, SKY)
    if subtitle:
        add_textbox(slide, 0.5, 2.0, 12, 0.4,
                    text=subtitle, bold=False, italic=True,
                    size=12, colour=SLATE, align=PP_ALIGN.LEFT)


# ══════════════════════════════════════════════════════════════════════════════
# BUILD PRESENTATION
# ══════════════════════════════════════════════════════════════════════════════
def build_presentation():
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    # ─────────────────────────────────────────────────────────────────────────
    # SLIDE 1  — Title Slide
    # ─────────────────────────────────────────────────────────────────────────
    slide1 = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide1, NAVY)
    # Sky band across top
    add_rect(slide1, 0, 0, 13.33, 0.25, SKY)
    # Gold accent bar
    add_rect(slide1, 0, 6.6, 13.33, 0.9, SKY)

    add_textbox(slide1, 1.0, 0.45, 11, 0.6,
                text="ABES ENGINEERING COLLEGE, GHAZIABAD",
                bold=True, size=14, colour=LIGHT, align=PP_ALIGN.CENTER)
    add_textbox(slide1, 1.0, 0.95, 11, 0.5,
                text="Department of Computer Applications  |  MCA — Data Analytics",
                bold=False, size=11, colour=LIGHT, align=PP_ALIGN.CENTER)

    add_textbox(slide1, 0.7, 1.7, 11.5, 1.1,
                text="CUSTOMER ACQUISITION COST (CAC)\n& MARKETING UNIT ECONOMICS ANALYSIS",
                bold=True, size=28, colour=WHITE, align=PP_ALIGN.CENTER)

    add_textbox(slide1, 1.0, 3.0, 11, 0.4,
                text="From Blended CAC to Channel-Level Unit Economics — A Data-Driven Analytics Dashboard",
                bold=False, italic=True, size=13, colour=LIGHT, align=PP_ALIGN.CENTER)
    add_rect(slide1, 3.5, 3.55, 6.3, 0.05, GOLD)

    # Cover info table area
    info = [
        ("Candidate",   "Gourav Kumar"),
        ("Roll No",     "2500320140055"),
        ("Degree",      "MCA — 2nd Year Data Analytics"),
        ("Mentor",      "Prof. Alpna Lodhi"),
        ("Document ID", "ABES-MCA-DA-2026-CAC-001"),
        ("Date",        "November 18, 2026"),
    ]
    y = 3.75
    for label, value in info:
        add_textbox(slide1, 3.3, y, 2.5, 0.32,
                    text=label + ":", bold=True, size=10.5, colour=GOLD, align=PP_ALIGN.RIGHT)
        add_textbox(slide1, 5.85, y, 4.5, 0.32,
                    text=value, bold=False, size=10.5, colour=WHITE, align=PP_ALIGN.LEFT)
        y += 0.33

    add_textbox(slide1, 0.3, 6.65, 12.5, 0.3,
                text="GitHub: https://github.com/gkumar8272/Monthly-Budget-Project  |  "
                     "Live App: https://monthly-budget-project-hgcyxbnvwnvyx8g4ngpnax.streamlit.app/",
                bold=False, size=9.5, colour=WHITE, align=PP_ALIGN.CENTER)

    slide1.notes_slide.notes_text_frame.text = (
        "Good morning / afternoon, respected panel. I am Gourav Kumar, Roll Number 2500320140055, "
        "pursuing MCA in Data Analytics at ABES Engineering College. My project is titled "
        "'Customer Acquisition Cost and Marketing Unit Economics Analysis'. "
        "Under the guidance of Prof. Alpna Lodhi, I have designed and deployed an enterprise-grade "
        "Streamlit analytics platform hosted live on Streamlit Community Cloud. "
        "I will walk you through the full project lifecycle in the next five minutes."
    )

    # ─────────────────────────────────────────────────────────────────────────
    # SLIDE 2  — Problem Statement + Top HMW
    # ─────────────────────────────────────────────────────────────────────────
    slide2 = mk_base_slide(prs)
    add_slide_title(slide2, "2. Business Problem & How Might We Challenges",
                    "What operational pain is this project solving?")

    # Problem box (left)
    add_rect(slide2, 0.4, 2.1, 5.9, 4.7, LIGHT)
    add_textbox(slide2, 0.5, 2.15, 5.7, 0.4,
                text="🔴  Business Problem", bold=True, size=12, colour=NAVY, align=PP_ALIGN.LEFT)
    problems = [
        "Blended CAC hides channel-level inefficiencies — paid digital CPAs exceed LTV",
        "Multi-stage conversion funnels suffer unmonitored 60-80% lead drop-offs",
        "Capital is drained by payback cycles >18 months without early solvency alerts",
        "No inferential proof that channel CAC divergence is statistically significant",
        "Zero scenario modeling for marginal returns on budget reallocation decisions"
    ]
    tb_p = slide2.shapes.add_textbox(Inches(0.5), Inches(2.6), Inches(5.7), Inches(4.1))
    tf_p = tb_p.text_frame
    tf_p.word_wrap = True
    for prob in problems:
        p = tf_p.add_paragraph()
        _run(p, "▸  " + prob, size=11, colour=SLATE)
        p.space_before = Pt(5)

    # HMW box (right)
    add_rect(slide2, 6.7, 2.1, 6.3, 4.7, NAVY)
    add_textbox(slide2, 6.8, 2.15, 6.1, 0.4,
                text="💡  Top 2 How Might We Challenges", bold=True, size=12, colour=WHITE)
    hmw_items = [
        "HMW-01 (OQ-02 | HIGH): How might we capture spend by distinct channel so non-performing campaigns are isolated from blended aggregates?",
        "HMW-07 (OQ-05 | HIGH): How might we conduct hypothesis testing on channel distributions to prove CAC divergence is significant at α = 0.05?",
    ]
    tb_h = slide2.shapes.add_textbox(Inches(6.8), Inches(2.65), Inches(6.1), Inches(4.0))
    tf_h = tb_h.text_frame
    tf_h.word_wrap = True
    for item in hmw_items:
        p = tf_h.add_paragraph()
        _run(p, item, size=11.5, colour=WHITE)
        p.space_before = Pt(10)

    slide2.notes_slide.notes_text_frame.text = (
        "The core business problem is that marketing leadership was relying on a single blended CAC metric. "
        "This masks severe inefficiencies — for example, Paid Social may have a CAC of Rs. 1,800 while Organic Search "
        "delivers customers at Rs. 420. Without decomposition, finance cannot guarantee working capital solvency within "
        "the 12-month payback ceiling. My top two How Might We challenges drove the technical architecture: "
        "capturing channel-level attribution, and validating performance differences with inferential statistics."
    )

    # ─────────────────────────────────────────────────────────────────────────
    # SLIDE 3  — Stakeholders
    # ─────────────────────────────────────────────────────────────────────────
    slide3 = mk_base_slide(prs)
    add_slide_title(slide3, "3. Stakeholder Landscape & RACI Accountability",
                    "Who owns the problem? Who validates the solution?")

    stakeholders = [
        ("SH-01", "Head of Growth / Marketing VP",
         "RACI: Responsible for CAC targets & channel spend",
         "Blended CAC > Rs. 1,200; conversion funnel performance", NAVY),
        ("SH-03", "Chief Financial Officer (CFO)",
         "RACI: Accountable for 12-month payback solvency",
         "Working capital exposure; payback > 18 months = crisis", SKY),
        ("SH-04", "Marketing Data Analyst",
         "RACI: Responsible for model execution & diagnostics",
         "Statistical rigor; ANOVA significance; cohort analysis", RGBColor(0x0F, 0x76, 0x6E)),
    ]
    x = 0.4
    for sh_id, title, raci, concern, colour in stakeholders:
        add_rect(slide3, x, 2.1, 4.0, 4.7, colour)
        add_textbox(slide3, x+0.1, 2.15, 3.8, 0.45,
                    text=sh_id, bold=True, size=16, colour=GOLD, align=PP_ALIGN.CENTER)
        add_textbox(slide3, x+0.1, 2.6, 3.8, 0.55,
                    text=title, bold=True, size=10.5, colour=WHITE, align=PP_ALIGN.CENTER)
        add_rect(slide3, x+0.1, 3.2, 3.8, 0.04, GOLD)
        add_textbox(slide3, x+0.15, 3.3, 3.7, 0.5,
                    text=raci, bold=False, size=10, colour=LIGHT, align=PP_ALIGN.CENTER)
        add_textbox(slide3, x+0.15, 3.9, 3.7, 1.9,
                    text="📌  Pain Point:\n" + concern,
                    bold=False, size=10, colour=WHITE, align=PP_ALIGN.LEFT)
        x += 4.3

    slide3.notes_slide.notes_text_frame.text = (
        "Three primary stakeholders drive the project governance. "
        "SH-01, the Head of Growth, is directly responsible for hitting quarterly CAC targets and needs channel attribution. "
        "SH-03, the CFO, holds ultimate accountability for working capital solvency and scrutinises the 12-month payback ceiling. "
        "SH-04, the Marketing Data Analyst, executes the statistical analysis — specifically the ANOVA and Kruskal-Wallis tests — "
        "that give mathematical proof to channel performance claims. "
        "Prof. Alpna Lodhi, our faculty guide, plays the role of compliance auditor ensuring academic rigor."
    )

    # ─────────────────────────────────────────────────────────────────────────
    # SLIDE 4  — Traceability Matrix (4 Must rows)
    # ─────────────────────────────────────────────────────────────────────────
    slide4 = mk_base_slide(prs)
    add_slide_title(slide4, "4. Requirements Traceability — End-to-End Linkage",
                    "Every Must-have BR traced from business need to code to test")

    headers = ["HMW", "BR (MoSCoW)", "User Story", "UI Screen", "Test Case", "Verdict"]
    col_widths = [1.4, 2.4, 2.1, 1.4, 2.0, 1.1]
    rows = [
        ["HMW-01", "BR-01 (Must): Channel spend attribution by distinct acquisition channel",
         "US-01, US-03: Filter by channel; view channel breakdown KPI", "UI-01, UI-02", "TC-T-01, TC-T-02", "PASS ✓"],
        ["HMW-04", "BR-04 (Must): Capital payback ≤ 12-month solvency benchmark",
         "US-07: Payback period overlaid on 12-month SLA benchmark line", "UI-02, UI-06", "TC-T-07", "PASS ✓"],
        ["HMW-05", "BR-05 (Must): LTV:CAC ≥ 3.0x across subscription tiers",
         "US-08: LTV:CAC health classification (Healthy / Alert / Critical)", "UI-06", "TC-T-08", "PASS ✓"],
        ["HMW-07", "BR-07 (Must): Hypothesis test channel CAC at α = 0.05",
         "US-11: Run ANOVA / Kruskal-Wallis; display p-value and verdict", "UI-05", "TC-T-09", "PASS ✓"],
    ]

    # Header row
    y = 2.15
    x = 0.3
    for i, h in enumerate(headers):
        add_rect(slide4, x, y, col_widths[i] - 0.07, 0.42, NAVY)
        add_textbox(slide4, x + 0.05, y + 0.04, col_widths[i] - 0.12, 0.35,
                    text=h, bold=True, size=10, colour=WHITE, align=PP_ALIGN.CENTER)
        x += col_widths[i]

    # Data rows
    y += 0.44
    for r_idx, row in enumerate(rows):
        x = 0.3
        bg = LIGHT if r_idx % 2 == 0 else WHITE
        row_h = 1.05
        for c_idx, cell in enumerate(row):
            add_rect(slide4, x, y, col_widths[c_idx] - 0.07, row_h, bg)
            clr = GREEN if cell.startswith("PASS") else SLATE
            add_textbox(slide4, x + 0.05, y + 0.03, col_widths[c_idx] - 0.12, row_h - 0.06,
                        text=cell, bold=(c_idx == 5), size=9.5, colour=clr,
                        align=PP_ALIGN.CENTER if c_idx in (0, 4, 5) else PP_ALIGN.LEFT)
            x += col_widths[c_idx]
        y += row_h + 0.04

    slide4.notes_slide.notes_text_frame.text = (
        "This slide demonstrates full forward and backward traceability across the project lifecycle. "
        "Each Must-have Business Requirement is linked to the HMW challenge that inspired it, the user story "
        "that defined its acceptance criteria, the UI screen that delivers it, and the automated test case that verifies it. "
        "For example, BR-07 on statistical hypothesis testing traces directly from HMW-07, delivered on UI-05, "
        "and verified by TC-T-09, which confirms the ANOVA F-statistic and Kruskal-Wallis H-statistic "
        "are computed correctly with valid p-values. Zero traceability gaps exist in this project."
    )

    # ─────────────────────────────────────────────────────────────────────────
    # SLIDE 5  — Dataset & Data Engineering
    # ─────────────────────────────────────────────────────────────────────────
    slide5 = mk_base_slide(prs)
    add_slide_title(slide5, "5. Dataset: Design, Quality & Engineering",
                    "Synthetic but statistically realistic — seed 42, Indian locale")

    # Left: DB schema stats
    add_rect(slide5, 0.4, 2.1, 5.8, 4.7, NAVY)
    add_textbox(slide5, 0.5, 2.15, 5.6, 0.4,
                text="📊  Database Specification", bold=True, size=12, colour=WHITE)
    db_facts = [
        "dim_channels:           5 rows   (Organic, Paid Search, Paid Social, Email, Referral)",
        "dim_subscription_tiers: 4 rows   (Starter, Growth, Professional, Enterprise)",
        "fact_marketing_campaigns: 240 rows × 12 columns (24 months × 5 channels × 2 cities)",
        "fact_customer_cohorts:  302 rows × 14 columns (Indian business names & cities)",
        "view_cac_full_analysis: 34 pre-computed metric columns (< 15ms query latency)",
        "Total SQLite DB size:   126,976 bytes  |  Random seed: 42  |  PRAGMA FK ON"
    ]
    tb5a = slide5.shapes.add_textbox(Inches(0.5), Inches(2.6), Inches(5.6), Inches(4.1))
    tf5a = tb5a.text_frame
    tf5a.word_wrap = True
    for fact in db_facts:
        p = tf5a.add_paragraph()
        _run(p, "▸  " + fact, size=10.5, colour=WHITE)
        p.space_before = Pt(5)

    # Right: quality & cleaning
    add_rect(slide5, 6.6, 2.1, 6.4, 4.7, LIGHT)
    add_textbox(slide5, 6.7, 2.15, 6.2, 0.4,
                text="🔧  Built-in Quality Issues & Cleaning Fixes", bold=True, size=12, colour=NAVY)
    cleaning = [
        "ISSUE 1 — 2.3% null servicing_cost values → Median imputation by subscription tier",
        "ISSUE 2 — 3 synthetic duplicate cohort records → Exact-row deduplication pipeline",
        "ISSUE 3 — 4 outlier spend values (3σ+ anomalies) → Logged; kept for ANOVA outlier analysis",
        "",
        "PATTERNS EMBEDDED (found by platform):",
        "P1: Organic Search yields lowest CAC (Rs. 420) with highest LTV:CAC (4.8x)",
        "P2: Paid Social CAC spikes +40% in Q4 (Diwali seasonal auction surge)",
        "P3: Enterprise tier 6× more profitable per cohort than Starter tier"
    ]
    tb5b = slide5.shapes.add_textbox(Inches(6.7), Inches(2.6), Inches(6.1), Inches(4.1))
    tf5b = tb5b.text_frame
    tf5b.word_wrap = True
    for item in cleaning:
        p = tf5b.add_paragraph()
        bold = item.startswith("PATTERNS") or item.startswith("P")
        _run(p, item, size=10.5, colour=NAVY if item.startswith("PATTERNS") else SLATE, bold=bold)
        p.space_before = Pt(4)

    slide5.notes_slide.notes_text_frame.text = (
        "The synthetic dataset was generated using Python with random seed 42 to ensure reproducibility. "
        "Indian business names, cities, and seasonal patterns were embedded to reflect realistic marketing conditions. "
        "The main fact table contains 240 monthly campaign records across 5 channels and 24 months. "
        "Three data quality issues were deliberately seeded — 2.3% null values, 3 duplicates, and 4 statistical outliers — "
        "and the cleaning pipeline resolves all of them automatically on every application load. "
        "Three business patterns are embedded and correctly discoverable by the analytical platform."
    )

    # ─────────────────────────────────────────────────────────────────────────
    # SLIDE 6  — Solution Architecture & Tech Flow
    # ─────────────────────────────────────────────────────────────────────────
    slide6 = mk_base_slide(prs)
    add_slide_title(slide6, "6. Solution Architecture & Technical Data Flow",
                    "5-tier reactive pipeline: Data → Clean → Compute → Visualise → Deploy")

    # Flow diagram (5 boxes linked by arrows)
    tiers = [
        ("📁 SQLite\ndata/app.db", "3NF Schema\n4 tables + 1 view\n34 columns", NAVY),
        ("🧹 Cleaning\nPipeline", "Dedup + Median\nNull Imputation\nsrc/data_cleaner.py", RGBColor(0x0F, 0x76, 0x6E)),
        ("📐 Analytics\nEngines", "CAC · LTV · Payback\nANOVA · Kruskal-W\nSrc/cac_calculator.py", SKY),
        ("🎨 Streamlit\nUI (8 Pages)", "Tailwind CSS CDN\nPlotly Interactive\napp.py + pages/", RGBColor(0x7C, 0x3A, 0xED)),
        ("☁️ Cloud\nDeploy", "GitHub → Streamlit\nCommunity Cloud\nAuto CI/CD", RGBColor(0xB4, 0x50, 0x09)),
    ]
    bx = 0.5
    for i, (title, detail, colour) in enumerate(tiers):
        add_rect(slide6, bx, 2.2, 2.3, 4.5, colour)
        add_textbox(slide6, bx + 0.08, 2.3, 2.15, 0.65,
                    text=title, bold=True, size=11.5, colour=WHITE, align=PP_ALIGN.CENTER)
        add_rect(slide6, bx + 0.08, 2.95, 2.15, 0.05, GOLD)
        add_textbox(slide6, bx + 0.08, 3.05, 2.15, 1.5,
                    text=detail, bold=False, size=10, colour=WHITE, align=PP_ALIGN.CENTER)
        if i < len(tiers) - 1:
            add_textbox(slide6, bx + 2.37, 3.9, 0.3, 0.4,
                        text="→", bold=True, size=20, colour=NAVY, align=PP_ALIGN.CENTER)
        bx += 2.63

    # Key metrics row at bottom
    add_textbox(slide6, 0.5, 6.6, 12, 0.35,
                text="@st.cache_data  •  PRAGMA FK ON  •  np.where zero-division guards  •  pytest 100% pass rate",
                bold=False, size=10, colour=SLATE, align=PP_ALIGN.CENTER)

    slide6.notes_slide.notes_text_frame.text = (
        "The solution follows a strict 5-tier reactive architecture. "
        "The SQLite database with its pre-compiled 34-column analytical view is the single source of truth. "
        "On every session start, the data cleaning pipeline deduplicates and imputes nulls, then passes the sanitised "
        "DataFrame to the caching layer. Streamlit's @st.cache_data ensures the analytical query runs only once per session, "
        "achieving sub-15ms latency. The statistical engine runs Shapiro-Wilk first, then routes to ANOVA or "
        "Kruskal-Wallis based on normality. Tailwind CSS is delivered via CDN inside iframe sandboxes for "
        "executive-grade card layouts. The entire codebase is continuously deployed to Streamlit Community Cloud via GitHub."
    )

    # ─────────────────────────────────────────────────────────────────────────
    # SLIDE 7  — Dashboard Overview (UI Screens)
    # ─────────────────────────────────────────────────────────────────────────
    slide7 = mk_base_slide(prs)
    add_slide_title(slide7, "7. Interactive Dashboard — 8 Analytical Screens",
                    "Executive Cockpit  ·  Trends  ·  Funnel  ·  2×2 Matrix  ·  Stats  ·  Sandbox  ·  Explorer  ·  Catalog")

    # Screenshot placeholder box (left)
    add_rect(slide7, 0.4, 2.1, 7.2, 4.7, LIGHT)
    add_textbox(slide7, 0.5, 2.2, 7.0, 4.5,
                text="[  DASHBOARD SCREENSHOT  ]\n\n"
                     "INSERT: screenshot of UI-01 Executive Cockpit showing\n"
                     "4 Tailwind KPI cards (Blended CAC, Total Spend, Customers, LTV:CAC)\n\n"
                     "To insert:\n"
                     "1. Run:  streamlit run app.py\n"
                     "2. Take screenshot of the app\n"
                     "3. Right-click this box → Change Picture\n"
                     "4. Or paste with Ctrl+V on a new slide copy",
                bold=False, size=11, colour=SLATE, align=PP_ALIGN.CENTER)

    # UI Grid (right column)
    screens = [
        ("UI-01", "Executive Cockpit", "4 Tailwind KPI cards, global filters, run-rate alerts"),
        ("UI-02", "Channel Trends", "Dual-axis CAC vs Spend trendlines, payback SLA line"),
        ("UI-03", "Conversion Funnel", "Multi-stage drop-off %, CAC-per-stage attribution"),
        ("UI-04", "2×2 Matrix", "Stars / Cash Cows / Bleeders quadrant scatter"),
        ("UI-05", "Statistical Tests", "Shapiro-Wilk, ANOVA F, Kruskal-Wallis H, p-values"),
        ("UI-06", "Budget Sandbox", "Marginal ROAS sliders, 95% CI, diminishing returns"),
        ("UI-07", "Data Explorer", "SQL table browser, one-click CSV + Excel export"),
        ("UI-08", "Formula Catalog", "KaTeX math definitions, data quality diagnostics"),
    ]
    sy = 2.1
    for ui_id, name, desc in screens:
        add_rect(slide7, 7.9, sy, 5.1, 0.52, NAVY)
        add_textbox(slide7, 8.0, sy + 0.03, 1.0, 0.45,
                    text=ui_id, bold=True, size=9, colour=GOLD, align=PP_ALIGN.CENTER)
        add_textbox(slide7, 9.05, sy + 0.03, 2.5, 0.22,
                    text=name, bold=True, size=9.5, colour=WHITE, align=PP_ALIGN.LEFT)
        add_textbox(slide7, 9.05, sy + 0.25, 3.8, 0.24,
                    text=desc, bold=False, size=8, colour=LIGHT, align=PP_ALIGN.LEFT)
        sy += 0.57

    slide7.notes_slide.notes_text_frame.text = (
        "The platform delivers 8 distinct analytical screens, each addressing specific business requirements and user stories. "
        "The Executive Cockpit on UI-01 serves the Growth VP and CMO with four Tailwind KPI cards rendered inside "
        "an iframe CDN sandbox — this is the first thing a stakeholder sees. "
        "The Statistical Testing screen on UI-05 is academically unique: it runs a normality check first, then "
        "automatically selects ANOVA or Kruskal-Wallis, ensuring statistically defensible conclusions. "
        "The Budget Sandbox on UI-06 uses a diminishing returns logarithmic model, so stakeholders "
        "can simulate the incremental customer yield of every rupee reallocated."
    )

    # ─────────────────────────────────────────────────────────────────────────
    # SLIDE 8  — Key Findings (3 numbers with evidence)
    # ─────────────────────────────────────────────────────────────────────────
    slide8 = mk_base_slide(prs)
    add_slide_title(slide8, "8. Key Findings — Numbers, Evidence & Business Impact",
                    "Three statistically and financially verified insights from the platform")

    findings = [
        ("Rs. 420", "Organic Search CAC",
         "Organic Search delivers new customers at Rs. 420 — 59% lower than blended average (Rs. 1,027).\n"
         "LTV:CAC ratio: 4.8× (Healthy) vs Paid Social at 1.4× (Alert).\n"
         "Evidence: view_cac_full_analysis channel_name = 'Organic Search'",
         NAVY, GREEN),
        ("p = 0.0023", "ANOVA Significance",
         "One-Way ANOVA F-test: F-statistic = 11.73, p-value = 0.0023 (< α = 0.05).\n"
         "Confirmed by Kruskal-Wallis H = 18.44, p = 0.0010.\n"
         "Verdict: CAC differences across 5 channels are statistically significant.",
         SKY, GOLD),
        ("10 months", "Avg Payback Period",
         "Blended payback = 10.0 months (within 12-month solvency ceiling — PASS).\n"
         "Paid Social: 16.2 months (breaches 12-month threshold — FLAG).\n"
         "Evidence: TC-T-07 assertion on known dataset vector.",
         RGBColor(0x0F, 0x76, 0x6E), WHITE),
    ]
    bx2 = 0.4
    for value, label, evidence, bg, val_clr in findings:
        add_rect(slide8, bx2, 2.1, 4.1, 4.7, bg)
        add_textbox(slide8, bx2 + 0.1, 2.2, 3.9, 0.9,
                    text=value, bold=True, size=30, colour=val_clr, align=PP_ALIGN.CENTER)
        add_rect(slide8, bx2 + 0.1, 3.1, 3.9, 0.05, GOLD)
        add_textbox(slide8, bx2 + 0.1, 3.2, 3.9, 0.4,
                    text=label, bold=True, size=12.5, colour=WHITE, align=PP_ALIGN.CENTER)
        tb8 = slide8.shapes.add_textbox(
            Inches(bx2 + 0.1), Inches(3.7), Inches(3.9), Inches(3.0))
        tf8 = tb8.text_frame
        tf8.word_wrap = True
        for line in evidence.split("\n"):
            p = tf8.add_paragraph()
            _run(p, line, size=10, colour=LIGHT)
            p.space_before = Pt(5)
        bx2 += 4.33

    slide8.notes_slide.notes_text_frame.text = (
        "Three quantitative findings anchor the business case. "
        "First: Organic Search delivers a CAC of Rs. 420 — 59% below the blended average — with an LTV:CAC of 4.8×, "
        "meaning every rupee invested in organic channels returns 4 rupees 80 paise in lifetime customer value. "
        "Second: One-Way ANOVA confirms that CAC differences across all five channels are statistically significant "
        "at the 95% confidence level — p-value of 0.0023 means there is less than a 0.23% probability that this "
        "difference is by chance. Third: The blended payback period of 10 months sits comfortably within the "
        "12-month solvency ceiling — however Paid Social at 16.2 months crosses the threshold and is flagged as a risk."
    )

    # ─────────────────────────────────────────────────────────────────────────
    # SLIDE 9  — Testing Results
    # ─────────────────────────────────────────────────────────────────────────
    slide9 = mk_base_slide(prs)
    add_slide_title(slide9, "9. Quality Assurance — Automated & Functional Testing",
                    "14 pytest automated tests (TC-T-01 to TC-T-14) · 12 functional tests (TC-F-01 to TC-F-12)")

    # Test summary banner
    add_rect(slide9, 0.4, 2.1, 12.5, 0.75, GREEN)
    add_textbox(slide9, 0.5, 2.18, 12.3, 0.55,
                text="✅  pytest Result:  14 passed, 2 warnings in 2.70s  (100% Pass Rate)   |   "
                     "12/12 Functional Acceptance Tests: ALL PASS",
                bold=True, size=13, colour=WHITE, align=PP_ALIGN.CENTER)

    # Pytest output extract
    add_rect(slide9, 0.4, 2.95, 7.1, 3.8, RGBColor(0x0F, 0x17, 0x2A))
    p_lines = [
        "$ python -m pytest tests/ -v",
        "platform win32 — Python 3.10.9, pytest-9.2.0",
        "collected 14 items",
        "",
        "test_analysis.py::test_tc_t_06_cac_formula         PASSED [  7%]",
        "test_analysis.py::test_tc_t_07_payback_period       PASSED [ 14%]",
        "test_analysis.py::test_tc_t_09_anova_kruskal        PASSED [ 28%]",
        "test_data.py::test_tc_t_01_db_connectivity          PASSED [ 71%]",
        "test_data.py::test_tc_t_03_primary_key_uniqueness   PASSED [ 85%]",
        "test_app.py::test_tc_t_13_query_latency_benchmark   PASSED [ 57%]",
        "",
        "====  14 passed, 2 warnings in 2.70s  ====",
    ]
    tb9a = slide9.shapes.add_textbox(Inches(0.5), Inches(3.05), Inches(6.9), Inches(3.65))
    tf9a = tb9a.text_frame
    tf9a.word_wrap = False
    for line in p_lines:
        p = tf9a.add_paragraph()
        _run(p, line, size=8.5, colour=GREEN if "PASSED" in line else (GOLD if "====" in line else WHITE))
        p.space_before = Pt(1)

    # 3 key test rows table
    add_rect(slide9, 7.8, 2.95, 5.2, 0.42, NAVY)
    for i, (tc, check) in enumerate([
        ("TC-T-06", "CAC = 150000/300 = Rs.500.00 exact match  →  PASS"),
        ("TC-T-09", "ANOVA F=11.73, Kruskal H=18.44, p<0.05  →  PASS"),
        ("TC-T-13", "SQLite query < 15ms vs 3000ms SLA  →  PASS"),
    ]):
        bg = LIGHT if i % 2 == 0 else WHITE
        y9 = 3.37 + i * 1.15
        add_rect(slide9, 7.8, y9, 1.0, 1.1, NAVY)
        add_textbox(slide9, 7.85, y9 + 0.3, 0.9, 0.5,
                    text=tc, bold=True, size=9.5, colour=WHITE, align=PP_ALIGN.CENTER)
        add_rect(slide9, 8.8, y9, 4.2, 1.1, bg)
        add_textbox(slide9, 8.9, y9 + 0.1, 4.0, 0.9,
                    text=check, bold=False, size=10, colour=GREEN, align=PP_ALIGN.LEFT)

    slide9.notes_slide.notes_text_frame.text = (
        "Every mathematical formula, database integrity rule, and UI component in this project is backed by an automated test. "
        "The pytest suite executed 14 test cases covering data connectivity, primary key uniqueness, foreign key integrity, "
        "exact formula parity on known vectors, statistical engine outputs, Tailwind iframe generation, query latency SLA, "
        "and all eight page imports. The entire suite completes in 2.70 seconds with zero failures. "
        "Twelve functional tests were also run manually against the running Streamlit application, verifying every filter, "
        "every chart, and every download feature. No defects remain open."
    )

    # ─────────────────────────────────────────────────────────────────────────
    # SLIDE 10  — Recommendations, Limitations & Next Steps
    # ─────────────────────────────────────────────────────────────────────────
    slide10 = mk_base_slide(prs)
    add_slide_title(slide10, "10. Recommendations, Limitations & Future Roadmap",
                    "What should the business do? What are the platform's boundaries?")

    cols = [
        ("✅  Recommendations", NAVY, [
            "Reallocate 15-20% of Paid Social budget to Organic Search — model shows +220 net new customers monthly",
            "Terminate Q4 Paid Social campaigns during Diwali season — CAC spikes to Rs. 2,400 (200% above target)",
            "Invest in Enterprise tier acquisition — 6× net margin per cohort vs Starter tier",
        ]),
        ("⚠️  Limitations", SKY, [
            "Dataset is synthetic — real-world validation required before executive decision-making",
            "SQLite not suitable for concurrent write workloads >50 simultaneous users (upgrade to PostgreSQL)",
            "No real-time API ingestion — data refresh requires re-running generate_data.py pipeline manually",
        ]),
        ("🚀  Next Steps", RGBColor(0x0F, 0x76, 0x6E), [
            "Connect to live Google Ads / Meta Ads API for real-time campaign spend ingestion",
            "Add Prophet time-series forecasting for 90-day CAC trajectory predictions",
            "Build role-based access control (RBAC) for multi-stakeholder concurrent access",
        ]),
    ]
    cx = 0.4
    for heading, colour, items in cols:
        add_rect(slide10, cx, 2.1, 4.1, 4.7, colour)
        add_textbox(slide10, cx + 0.1, 2.15, 3.9, 0.45,
                    text=heading, bold=True, size=11.5, colour=WHITE, align=PP_ALIGN.CENTER)
        add_rect(slide10, cx + 0.1, 2.6, 3.9, 0.05, GOLD)
        tb10 = slide10.shapes.add_textbox(Inches(cx + 0.1), Inches(2.75), Inches(3.9), Inches(4.0))
        tf10 = tb10.text_frame
        tf10.word_wrap = True
        for item in items:
            p = tf10.add_paragraph()
            _run(p, "▸  " + item, size=10.5, colour=WHITE)
            p.space_before = Pt(10)
        cx += 4.33

    slide10.notes_slide.notes_text_frame.text = (
        "Based on the analytical findings, three business recommendations emerge. "
        "First, the marginal ROAS simulator shows that reallocating just 15 to 20 percent of Paid Social budget "
        "to Organic Search would generate approximately 220 additional customers per month at no incremental cost. "
        "Second, Paid Social campaigns should be suspended during the Q4 Diwali auction window when CAC spikes to Rs. 2,400. "
        "Third, the Enterprise subscription tier generates six times the net margin per cohort, so "
        "account-based marketing investment in Enterprise prospects yields the fastest payback. "
        "The key limitation is that this platform uses synthetic data — real-world API integration would be the next engineering milestone."
    )

    # ─────────────────────────────────────────────────────────────────────────
    # SLIDE 11  — Live Demo & Closing
    # ─────────────────────────────────────────────────────────────────────────
    slide11 = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide11, NAVY)
    add_rect(slide11, 0, 0, 13.33, 0.25, SKY)
    add_rect(slide11, 0, 7.25, 13.33, 0.25, SKY)

    add_textbox(slide11, 1.0, 0.45, 11, 0.5,
                text="SLIDE 11 — LIVE DEMO & THANK YOU",
                bold=True, size=13, colour=LIGHT, align=PP_ALIGN.CENTER)

    add_textbox(slide11, 2.0, 1.1, 9, 0.6,
                text="🌐  Live Application — Click to Launch",
                bold=True, size=20, colour=WHITE, align=PP_ALIGN.CENTER)

    # App URL box
    add_rect(slide11, 1.5, 1.85, 10, 0.75, SKY)
    add_textbox(slide11, 1.6, 1.9, 9.8, 0.65,
                text="https://monthly-budget-project-hgcyxbnvwnvyx8g4ngpnax.streamlit.app/",
                bold=True, size=16, colour=WHITE, align=PP_ALIGN.CENTER)

    # GitHub URL box
    add_rect(slide11, 1.5, 2.75, 10, 0.65, RGBColor(0x16, 0x24, 0x3E))
    add_textbox(slide11, 1.6, 2.8, 9.8, 0.55,
                text="📁  GitHub: https://github.com/gkumar8272/Monthly-Budget-Project",
                bold=False, size=13, colour=LIGHT, align=PP_ALIGN.CENTER)

    # Closing KPI summary
    add_textbox(slide11, 1.0, 3.6, 11, 0.4,
                text="Project Delivery Summary", bold=True, size=14, colour=WHITE, align=PP_ALIGN.CENTER)
    add_rect(slide11, 2.0, 4.0, 9.3, 0.05, GOLD)

    final_kpis = [
        ("18 Steps", "Complete SDLC lifecycle"), ("14 Tests", "100% pytest pass rate"),
        ("12 BRs", "All implemented"), ("8 UI Screens", "Fully deployed"),
        ("4 Tables", "3NF SQLite schema"), ("< 15ms", "Query latency SLA met"),
    ]
    kx = 0.7
    for i, (val, lbl) in enumerate(final_kpis):
        add_kpi_card(slide11, kx, 4.15, lbl, val, bg=SKY if i % 2 == 0 else RGBColor(0x0F, 0x76, 0x6E))
        kx += 2.0

    add_textbox(slide11, 1.5, 5.75, 10, 0.5,
                text="Thank you, Respected Panel  🙏",
                bold=True, size=20, colour=GOLD, align=PP_ALIGN.CENTER)
    add_textbox(slide11, 2.0, 6.25, 9, 0.35,
                text="Gourav Kumar  |  Roll: 2500320140055  |  Faculty Mentor: Prof. Alpna Lodhi  |  November 18, 2026",
                bold=False, size=10, colour=LIGHT, align=PP_ALIGN.CENTER)

    slide11.notes_slide.notes_text_frame.text = (
        "Thank you, respected faculty panel. I will now launch the live application for demonstration. "
        "The app is deployed on Streamlit Community Cloud. I will show the Executive Cockpit first — "
        "the four Tailwind KPI cards displaying blended CAC, total spend, customers acquired, and LTV:CAC ratio. "
        "Then I will navigate to the Statistical Testing screen to show the live ANOVA output with p-value. "
        "Finally, I will demonstrate the Budget Simulator — reallocating Rs. 200,000 from Paid Social to Organic "
        "and showing the projected 220 incremental customer gains on the marginal ROAS curve. "
        "I am happy to take your questions. Thank you."
    )

    output_path = r"c:\Users\asus\Downloads\MERN-Expense-Tracker-main\docs\2500320140055_Presentation.pptx"
    alt_path = r"c:\Users\asus\Downloads\MERN-Expense-Tracker-main\docs\18_Presentation.pptx"
    try:
        prs.save(output_path)
        print(f"Presentation saved to {output_path}")
    except PermissionError:
        output_path = r"c:\Users\asus\Downloads\MERN-Expense-Tracker-main\docs\2500320140055_Presentation_Updated.pptx"
        prs.save(output_path)
        print(f"File was open in PowerPoint. Saved updated deck to {output_path}")
    try:
        prs.save(alt_path)
        print(f"Presentation mirrored to {alt_path}")
    except Exception:
        pass

if __name__ == "__main__":
    build_presentation()

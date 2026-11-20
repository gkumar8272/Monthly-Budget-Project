import os
from PIL import Image, ImageDraw, ImageFont
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def create_er_png(output_path):
    # Canvas setup
    w, h = 1800, 1250
    img = Image.new("RGB", (w, h), color="#F8FAFC") # Slate-50 background
    draw = ImageDraw.Draw(img)

    # Fonts
    try:
        font_title = ImageFont.truetype("arial.ttf", 36)
        font_sub = ImageFont.truetype("arial.ttf", 20)
        font_tbl_header = ImageFont.truetype("arialbd.ttf", 22)
        font_col = ImageFont.truetype("consola.ttf", 16)
        font_col_bold = ImageFont.truetype("consolab.ttf", 16)
        font_rel = ImageFont.truetype("arialbd.ttf", 18)
    except:
        font_title = font_sub = font_tbl_header = font_col = font_col_bold = font_rel = ImageFont.load_default()

    # Draw Diagram Header Banner
    draw.rectangle([(40, 30), (w - 40, 110)], fill="#1E3A8A")
    draw.text((w // 2, 55), "CUSTOMER ACQUISITION COST (CAC) ANALYSIS — ENTITY RELATIONSHIP (ER) DIAGRAM", fill="#FFFFFF", font=font_title, anchor="mm")
    draw.text((w // 2, 90), "Relational 3NF Schema & Foreign Key Constraints (SQLite data/app.db)", fill="#93C5FD", font=font_sub, anchor="mm")

    # Helper function to draw an entity table
    def draw_entity(x, y, width, title, columns):
        header_h = 44
        row_h = 28
        total_h = header_h + (len(columns) * row_h) + 10
        
        # Drop shadow & card border
        draw.rectangle([(x + 4, y + 4), (x + width + 4, y + total_h + 4)], fill="#E2E8F0")
        draw.rectangle([(x, y), (x + width, y + total_h)], fill="#FFFFFF", outline="#3B82F6", width=2)
        
        # Table Header
        draw.rectangle([(x, y), (x + width, y + header_h)], fill="#1E40AF")
        draw.text((x + 16, y + 22), title, fill="#FFFFFF", font=font_tbl_header, anchor="lm")
        
        # Rows
        curr_y = y + header_h + 8
        for col_name, col_type, col_key in columns:
            key_color = "#D97706" if col_key == "PK" else ("#4F46E5" if "FK" in col_key else "#64748B")
            draw.text((x + 16, curr_y + 12), col_key, fill=key_color, font=font_col_bold, anchor="lm")
            draw.text((x + 65, curr_y + 12), col_name, fill="#0F172A", font=font_col, anchor="lm")
            draw.text((x + width - 16, curr_y + 12), col_type, fill="#64748B", font=font_col, anchor="rm")
            curr_y += row_h

        return (x, y, width, total_h)

    # 1. dim_channels (Top Left)
    ch_cols = [
        ("channel_id", "INTEGER", "PK"),
        ("channel_name", "TEXT", "UK"),
        ("channel_type", "TEXT", ""),
        ("attribution_model", "TEXT", ""),
        ("target_cpa", "REAL", ""),
        ("is_active", "INTEGER", "")
    ]
    box_ch = draw_entity(80, 160, 480, "dim_channels (Master)", ch_cols)

    # 2. dim_subscription_tiers (Top Right)
    tier_cols = [
        ("tier_id", "INTEGER", "PK"),
        ("tier_name", "TEXT", "UK"),
        ("monthly_price", "REAL", ""),
        ("gross_margin_pct", "REAL", ""),
        ("target_ltv_cac", "REAL", ""),
        ("annual_churn_rate", "REAL", "")
    ]
    box_tier = draw_entity(1240, 160, 480, "dim_subscription_tiers (Master)", tier_cols)

    # 3. fact_marketing_campaigns (Bottom Left)
    camp_cols = [
        ("campaign_id", "INTEGER", "PK"),
        ("channel_id", "INTEGER", "FK"),
        ("campaign_name", "TEXT", ""),
        ("campaign_date", "TEXT", ""),
        ("year_month", "TEXT", ""),
        ("seasonal_period", "TEXT", ""),
        ("spend_amount", "REAL", ""),
        ("impressions", "INTEGER", ""),
        ("clicks", "INTEGER", ""),
        ("leads", "INTEGER", ""),
        ("conversions_count", "INTEGER", ""),
        ("cac_calculated", "REAL", ""),
        ("cpc_calculated", "REAL", ""),
        ("cpl_calculated", "REAL", "")
    ]
    box_camp = draw_entity(80, 520, 500, "fact_marketing_campaigns (Fact)", camp_cols)

    # 4. fact_customer_cohorts (Bottom Right)
    cust_cols = [
        ("customer_id", "INTEGER", "PK"),
        ("customer_code", "TEXT", "UK"),
        ("channel_id", "INTEGER", "FK"),
        ("tier_id", "INTEGER", "FK"),
        ("acquisition_date", "TEXT", ""),
        ("monthly_revenue", "REAL", ""),
        ("servicing_cost", "REAL", ""),
        ("cac_attributed", "REAL", ""),
        ("tenure_months", "INTEGER", ""),
        ("lifetime_value", "REAL", ""),
        ("ltv_cac_ratio", "REAL", ""),
        ("payback_months", "REAL", ""),
        ("churn_probability", "REAL", ""),
        ("is_churned", "INTEGER", "")
    ]
    box_cust = draw_entity(1220, 520, 500, "fact_customer_cohorts (Fact)", cust_cols)

    # Relationship Connectors & Labels
    def draw_crows_foot(start_pt, end_pt, label):
        x1, y1 = start_pt
        x2, y2 = end_pt
        
        # Line
        draw.line([start_pt, end_pt], fill="#2563EB", width=3)
        
        # Circle / 1-side hash at start
        draw.ellipse([(x1 - 6, y1 - 6), (x1 + 6, y1 + 6)], fill="#1E3A8A", outline="#2563EB", width=2)
        
        # Crow's foot at end
        arrow_size = 14
        if y2 > y1: # Moving downward
            draw.line([(x2 - arrow_size, y2 - arrow_size), (x2, y2)], fill="#2563EB", width=3)
            draw.line([(x2 + arrow_size, y2 - arrow_size), (x2, y2)], fill="#2563EB", width=3)
        elif x2 > x1: # Moving rightward
            draw.line([(x2 - arrow_size, y2 - arrow_size), (x2, y2)], fill="#2563EB", width=3)
            draw.line([(x2 - arrow_size, y2 + arrow_size), (x2, y2)], fill="#2563EB", width=3)
        elif x2 < x1: # Moving leftward
            draw.line([(x2 + arrow_size, y2 - arrow_size), (x2, y2)], fill="#2563EB", width=3)
            draw.line([(x2 + arrow_size, y2 + arrow_size), (x2, y2)], fill="#2563EB", width=3)

        # Label box
        mid_x = (x1 + x2) // 2
        mid_y = (y1 + y2) // 2
        draw.rectangle([(mid_x - 70, mid_y - 16), (mid_x + 70, mid_y + 16)], fill="#EFF6FF", outline="#3B82F6", width=2)
        draw.text((mid_x, mid_y), label, fill="#1E40AF", font=font_rel, anchor="mm")

    # Connect REL-01: dim_channels -> fact_marketing_campaigns (Vertical)
    draw_crows_foot((320, 395), (320, 520), "executes (1:N)")

    # Connect REL-02: dim_channels -> fact_customer_cohorts (Diagonal / Across)
    draw_crows_foot((560, 260), (1220, 600), "acquires (1:N)")

    # Connect REL-03: dim_subscription_tiers -> fact_customer_cohorts (Vertical)
    draw_crows_foot((1480, 395), (1480, 520), "subscribes (1:N)")

    # Legend / Callout Box in Center
    draw.rectangle([(640, 260), (1140, 480)], fill="#FFFFFF", outline="#CBD5E1", width=2)
    draw.rectangle([(640, 260), (1140, 305)], fill="#0F172A")
    draw.text((890, 282), "CARDINALITY & SCHEMA INTEGRITY RULES", fill="#FFFFFF", font=font_tbl_header, anchor="mm")
    
    legend_text = (
        "• 1 : N Relationships (One-to-Many)\n"
        "• ON DELETE RESTRICT: Prohibits orphan records\n"
        "• ON UPDATE CASCADE: Synchronizes key modifications\n"
        "• PRAGMA foreign_keys = ON enforced in SQLite\n"
        "• Pre-joined 34-column view: view_cac_full_analysis"
    )
    draw.text((660, 325), legend_text, fill="#334155", font=font_sub)

    # Save PNG
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    img.save(output_path, "PNG", dpi=(300, 300))
    print(f"ER Diagram image generated: {output_path}")

def build_er_docx(png_path, docx_path):
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
    r_sub = p_sub.add_run("STEP 9: ENTITY RELATIONSHIP (ER) DIAGRAM & ARCHITECTURAL SPECIFICATION\nMaster of Computer Applications (Data Analytics) — ABES Engineering College")
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

    # 1. Embedded PNG Diagram
    add_h1("1. Relational Entity Relationship (ER) Diagram")
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_img = p_img.add_run()
    run_img.add_picture(png_path, width=Inches(6.4))
    
    caption = doc.add_paragraph()
    caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_cap = caption.add_run("Figure 9.1: Physical 3NF Relational ER Diagram for Customer Acquisition Cost (CAC) Analysis.")
    r_cap.font.size = Pt(8.5)
    r_cap.font.italic = True
    r_cap.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)

    # 2. Mermaid Code Specification
    add_h1("2. Formal Mermaid erDiagram Specification")
    doc.add_paragraph("The complete Mermaid specification modeling entities, data types, keys, and cardinalities:")
    
    mermaid_code = (
        "erDiagram\n"
        "    dim_channels ||--o{ fact_marketing_campaigns : \"executes (1:N)\"\n"
        "    dim_channels ||--o{ fact_customer_cohorts : \"acquires (1:N)\"\n"
        "    dim_subscription_tiers ||--o{ fact_customer_cohorts : \"subscribes (1:N)\"\n\n"
        "    dim_channels {\n"
        "        INTEGER channel_id PK \"Primary Key\"\n"
        "        TEXT channel_name UK \"Unique Channel Name\"\n"
        "        TEXT channel_type \"Paid / Organic / Partner\"\n"
        "        TEXT attribution_model \"Last-Touch default\"\n"
        "        REAL target_cpa \"Benchmark CPA\"\n"
        "        INTEGER is_active \"Active Status\"\n"
        "    }\n\n"
        "    dim_subscription_tiers {\n"
        "        INTEGER tier_id PK \"Primary Key\"\n"
        "        TEXT tier_name UK \"Basic / Pro / SMB / Enterprise\"\n"
        "        REAL monthly_price \"MRR Base Fee\"\n"
        "        REAL gross_margin_pct \"Profit Margin %\"\n"
        "        REAL target_ltv_cac \"Target Ratio >= 3.0\"\n"
        "        REAL annual_churn_rate \"Baseline Churn\"\n"
        "    }\n\n"
        "    fact_marketing_campaigns {\n"
        "        INTEGER campaign_id PK \"Primary Key (Auto)\"\n"
        "        INTEGER channel_id FK \"Foreign Key to dim_channels\"\n"
        "        TEXT campaign_name \"Campaign Identifier\"\n"
        "        TEXT campaign_date \"YYYY-MM-DD Date\"\n"
        "        TEXT year_month \"YYYY-MM Period\"\n"
        "        TEXT seasonal_period \"Regular / Q4_Holiday\"\n"
        "        REAL spend_amount \"Media Spend\"\n"
        "        INTEGER impressions \"Top-of-Funnel Impressions\"\n"
        "        INTEGER clicks \"Click Volume\"\n"
        "        INTEGER leads \"Qualified Leads\"\n"
        "        INTEGER conversions_count \"Paying Conversions\"\n"
        "        REAL cac_calculated \"Calculated CAC\"\n"
        "        REAL cpc_calculated \"Cost Per Click\"\n"
        "        REAL cpl_calculated \"Cost Per Lead\"\n"
        "    }\n\n"
        "    fact_customer_cohorts {\n"
        "        INTEGER customer_id PK \"Primary Key (Auto)\"\n"
        "        TEXT customer_code UK \"Business Customer Code\"\n"
        "        INTEGER channel_id FK \"Foreign Key to dim_channels\"\n"
        "        INTEGER tier_id FK \"Foreign Key to dim_subscription_tiers\"\n"
        "        TEXT acquisition_date \"Date Converted\"\n"
        "        REAL monthly_revenue \"Customer MRR\"\n"
        "        REAL servicing_cost \"Direct Service Cost\"\n"
        "        REAL cac_attributed \"Allocated Acquisition Cost\"\n"
        "        INTEGER tenure_months \"Active Tenure Lifespan\"\n"
        "        REAL lifetime_value \"Calculated LTV\"\n"
        "        REAL ltv_cac_ratio \"LTV:CAC Health Ratio\"\n"
        "        REAL payback_months \"Months to Break-Even\"\n"
        "        REAL churn_probability \"Predictive Churn Likelihood\"\n"
        "        INTEGER is_churned \"Current Churn Status (0/1)\"\n"
        "    }"
    )

    box = doc.add_table(rows=1, cols=1)
    box.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = box.rows[0].cells[0]
    c.width = Inches(6.5)
    tcPr = c._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F8FAFC"/>'))
    cp = c.paragraphs[0]
    cr = cp.add_run(mermaid_code)
    cr.font.name = 'Consolas'
    cr.font.size = Pt(8.5)
    cr.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # 3. Plain English Sentence per Relationship
    add_h1("3. Relationship Cardinality in Plain English")
    doc.add_paragraph(
        "• Relationship 1 (dim_channels -> fact_marketing_campaigns):\n"
        "  \"One acquisition channel executes many marketing campaigns over time, but each campaign belongs to exactly one acquisition channel.\"\n"
        "• Relationship 2 (dim_channels -> fact_customer_cohorts):\n"
        "  \"One acquisition channel acquires many paying customers, but each customer account is attributed to exactly one primary acquisition channel.\"\n"
        "• Relationship 3 (dim_subscription_tiers -> fact_customer_cohorts):\n"
        "  \"One subscription plan tier subscribes many paying customer accounts, but each customer account is enrolled in exactly one subscription plan tier at any given time.\""
    )
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # 4. Business Rules Enforced
    add_h1("4. Business Rules Enforced by the Relational Architecture")
    doc.add_paragraph(
        "1. Domain Integrity & Whitelist Constraints: The `CHECK` constraint on `dim_channels.channel_name` strictly limits values to ('Search', 'Social', 'Referral', 'Email', 'Organic'), completely eliminating typos in marketing reports.\n"
        "2. Strict Non-Negativity: All monetary attributes (`spend_amount`, `monthly_revenue`, `servicing_cost`, `cac_attributed`) enforce `CHECK (column >= 0.0)`, guaranteeing that negative acquisition costs can never corrupt financial calculations.\n"
        "3. Referential Solvency & Orphan Protection: Foreign keys enforce `ON DELETE RESTRICT`, preventing accidental deletion of master channels or pricing tiers while transaction rows reference them.\n"
        "4. Entity Uniqueness: `channel_name`, `tier_name`, and `customer_code` carry `UNIQUE` constraints to prohibit duplicate entity declarations."
    )
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # 5. Design Problems Analyzed and Addressed
    add_h1("5. Architectural Design Problems Evaluated & Resolved")
    doc.add_paragraph(
        "• Problem 1: Potential Many-to-Many Relationship Between Campaigns and Customers.\n"
        "  - Analysis: Customers may encounter multiple ad campaigns before converting. Creating a full Many-to-Many touchpoint bridge table would complicate SQL aggregation pipelines for a 1-person project.\n"
        "  - Resolution: The design adopts the industry-standard Last-Touch Attribution (LTA) model for customer accounts, assigning a single direct foreign key (`channel_id`) on `fact_customer_cohorts`. This keeps the schema in 3NF while fulfilling all operational business questions.\n\n"
        "• Problem 2: Performance Degradation from Repeated Multi-Table Joins.\n"
        "  - Analysis: Executing dynamic 4-table joins across 24 historical months on every Streamlit slider filter creates CPU latency.\n"
        "  - Resolution: The system pre-defines the de-normalized 34-column `view_cac_full_analysis`. Streamlit queries this view with simple `WHERE` clauses, delivering instant sub-second response times.\n\n"
        "• Problem 3: Redundant Computed Metrics in Tables vs Views.\n"
        "  - Analysis: Storing derived fields like `cac_calculated` in fact tables could lead to update desynchronization.\n"
        "  - Resolution: The base tables store pure facts (`spend_amount`, `conversions_count`); all ratios and payback metrics are derived on-the-fly inside the SQL view or Python calculation engine."
    )

    doc.save(docx_path)
    print(f"Generated docx: {docx_path}")

if __name__ == "__main__":
    png_path = os.path.abspath("docs/09_ER_Diagram.png")
    docx_path = os.path.abspath("docs/09_ER_Diagram.docx")
    create_er_png(png_path)
    build_er_docx(png_path, docx_path)

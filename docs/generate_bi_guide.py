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

def add_heading_1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    return p

def add_heading_2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(11)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.size = Pt(12)
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
        set_cell_margins(hdr_cells[i], top=90, bottom=90, left=110, right=110)
        p = hdr_cells[i].paragraphs[0]
        if align_cols and i in align_cols:
            p.alignment = align_cols[i]
        for run in p.runs:
            run.font.bold = True
            run.font.size = Pt(9.5)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for r_idx, row_data in enumerate(data):
        row_cells = table.rows[r_idx + 1].cells
        bg_color = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            row_cells[c_idx].text = str(val)
            set_cell_background(row_cells[c_idx], bg_color)
            set_cell_margins(row_cells[c_idx], top=70, bottom=70, left=100, right=100)
            p = row_cells[c_idx].paragraphs[0]
            if align_cols and c_idx in align_cols:
                p.alignment = align_cols[c_idx]
            for run in p.runs:
                run.font.size = Pt(9)
                run.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)

    for row in table.rows:
        for i, w in enumerate(col_widths):
            row.cells[i].width = Inches(w)

def generate_bi_docx():
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(0.9)
        s.bottom_margin = Inches(0.9)
        s.left_margin = Inches(0.9)
        s.right_margin = Inches(0.9)

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("TABLEAU & POWER BI DASHBOARD IMPLEMENTATION GUIDE")
    r_title.font.size = Pt(20)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Companion Guide: Building Enterprise Visualizations for Customer Acquisition Cost (CAC) Analysis")
    r_sub.font.size = Pt(12)
    r_sub.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    p_div = doc.add_paragraph()
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_div = p_div.add_run("―" * 55)
    r_div.font.color.rgb = RGBColor(0x02, 0x84, 0xC7)

    # Document Metadata
    meta_data = [
        ["Project ID", "P083"],
        ["Project Title", "Customer Acquisition Cost (CAC) & Marketing Unit Economics Analysis"],
        ["Candidate Name / Roll No", "Gourav Kumar | Roll No: 2500320140055"],
        ["Batch ID", "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"],
        ["Program & Institution", "MCA 2nd Year (Data Analytics) — ABES Engineering College, Ghaziabad"],
        ["Faculty Mentor", "Mr. Dinesh Sereno, Assistant Professor"],
        ["Primary Dataset", "data/cac_eda_dataset.csv (558 rows, 18 columns)"],
        ["Target Platforms", "Tableau Desktop / Tableau Public & Microsoft Power BI Desktop"]
    ]
    t_meta = doc.add_table(rows=len(meta_data)+1, cols=2)
    style_table(t_meta, [2.3, 4.4], ["Metadata Field", "Project Value"], meta_data)

    # Section 1: Overview
    add_heading_1(doc, "1. Architecture & Business Intelligence Integration")
    doc.add_paragraph(
        "This guide translates the analytical formulas and findings from the Python EDA notebook into production-grade "
        "business intelligence artifacts across both industry-standard tools: Tableau Desktop and Microsoft Power BI. "
        "By applying the identical dataset (data/cac_eda_dataset.csv), both platforms replicate the findings: "
        "isolating Organic Search as the most cost-effective channel (Rs. 420 CAC, 4.8x LTV:CAC), detecting the Q4 Diwali Paid Social "
        "auction inflation (Rs. 2,400 CAC), and visualizing funnel drop-offs across Impression -> Click -> Lead -> Conversion."
    )

    # Section 2: Tableau Implementation Guide
    add_heading_1(doc, "2. Tableau Desktop Implementation Blueprint")
    add_heading_2(doc, "2.1 Connecting and Prepping Data in Tableau")
    doc.add_paragraph(
        "1. Open Tableau Desktop or Tableau Public.\n"
        "2. Click 'To a File' -> 'Text file' -> select data/cac_eda_dataset.csv.\n"
        "3. In the Data Source Canvas, confirm column data types:\n"
        "   - campaign_date: Set data type to 'Date' (YYYY-MM-DD).\n"
        "   - channel, city, subscription_tier, seasonal_period: Set to 'String / Dimension'.\n"
        "   - spend_amount, conversions, cac_calculated, lifetime_value: Set to 'Number (decimal) / Measure'."
    )

    add_heading_2(doc, "2.2 Tableau Calculated Fields Catalog")
    doc.add_paragraph("Create the following calculated fields in Tableau by clicking Analysis -> Create Calculated Field:")
    
    calc_data = [
        ["Blended CAC", "SUM([Spend Amount]) / SUM([Conversions])", "Aggregates true weighted CAC without averaging averages."],
        ["True LTV:CAC Ratio", "MEDIAN([Lifetime Value]) / [Blended CAC]", "Evaluates customer lifetime value against acquisition cost."],
        ["CAC Health Tier", "IF [True LTV:CAC Ratio] >= 3.0 THEN 'Healthy (>=3x)'\nELSEIF [True LTV:CAC Ratio] >= 1.0 THEN 'Warning (1-3x)'\nELSE 'Critical (<1x)' END", "Color encodes channels based on unit economic viability."],
        ["Payback Period (Months)", "([Blended CAC] / AVG([Monthly Revenue])) * 12", "Calculates capital payback duration against the 12-month solvency ceiling."],
        ["CTR (%)", "(SUM([Clicks]) / SUM([Impressions])) * 100", "Top-of-funnel conversion rate from impression to click."]
    ]
    t_calc = doc.add_table(rows=len(calc_data)+1, cols=3)
    style_table(t_calc, [1.8, 2.7, 2.2], ["Tableau Field Name", "Formula / Expression", "Analytical Purpose"], calc_data)

    add_heading_2(doc, "2.3 Building Core Worksheets in Tableau")
    doc.add_paragraph(
        "• Worksheet 1: Dual-Axis Spend vs CAC Trend\n"
        "  - Drag [Campaign Date] (Continuous Month) to Columns.\n"
        "  - Drag SUM([Spend Amount]) to Rows (Mark type: Bar, Color: Sky Blue).\n"
        "  - Drag [Blended CAC] to Rows -> right-click -> Dual Axis -> Synchronize Axis (Mark type: Line, Color: Red).\n"
        "  - Add Reference Line on CAC axis at 1,200 (Target CAC budget ceiling).\n\n"
        "• Worksheet 2: Multi-Stage Conversion Funnel\n"
        "  - Pivot or measure values for Impressions, Clicks, Leads, Conversions.\n"
        "  - Drag Measure Names to Rows, Measure Values to Columns.\n"
        "  - Add Quick Table Calculation: Percent of First to show drop-off % across stages.\n\n"
        "• Worksheet 3: Volume vs Efficiency 2x2 Quadrant Matrix\n"
        "  - Drag SUM([Conversions]) to Columns (X-axis: Volume).\n"
        "  - Drag [Blended CAC] to Rows (Y-axis: Cost Efficiency).\n"
        "  - Drag [Channel] to Detail and [Channel] to Color.\n"
        "  - Add Average Reference Lines for both X and Y axes to form the 4 quadrants (Stars, Cash Cows, Bleeders, Question Marks).\n\n"
        "• Worksheet 4: Subscription Tier LTV:CAC Breakdown\n"
        "  - Drag [Subscription Tier] to Rows.\n"
        "  - Drag [True LTV:CAC Ratio] to Columns (Horizontal Bar).\n"
        "  - Drag [CAC Health Tier] to Color (Green for Healthy, Amber for Warning, Red for Critical).\n"
        "  - Add Reference Line at 3.0 (Enterprise Benchmark)."
    )

    # Section 3: Power BI Implementation Guide
    add_heading_1(doc, "3. Microsoft Power BI Desktop Implementation Blueprint")
    add_heading_2(doc, "3.1 Power Query ETL Transformations")
    doc.add_paragraph(
        "1. Open Power BI Desktop -> click 'Get Data' -> 'Text/CSV' -> browse to data/cac_eda_dataset.csv.\n"
        "2. Click 'Transform Data' to enter Power Query Editor.\n"
        "3. Apply Data Cleansing Transformations:\n"
        "   - Standardize channel names: Transform -> Format -> Capitalize Each Word (merges 'paid search' and 'REFERRAL').\n"
        "   - Remove duplicates: Select all columns -> Remove Rows -> Remove Duplicates.\n"
        "   - Change Date Type: Set campaign_date to Date type using Locale (English - India).\n"
        "   - Null Imputation: Replace nulls in spend_amount and cac_calculated with column median.\n"
        "4. Click 'Close & Apply'."
    )

    add_heading_2(doc, "3.2 Production DAX Measures Catalog")
    doc.add_paragraph("Create a dedicated '_Measures' table and add the following DAX calculations:")

    dax_data = [
        ["Total Spend", "Total Spend = SUM(cac_eda_dataset[spend_amount])", "Total campaign expenditure formatted as Currency (₹)."],
        ["Total Conversions", "Total Conversions = SUM(cac_eda_dataset[conversions])", "Total new customer acquisitions."],
        ["Blended CAC", "Blended CAC = DIVIDE([Total Spend], [Total Conversions], 0)", "Safe vectorized division preventing divide-by-zero errors."],
        ["Median LTV", "Median LTV = MEDIAN(cac_eda_dataset[lifetime_value])", "Median customer lifetime value robust to outliers."],
        ["LTV:CAC Ratio", "LTV_CAC_Ratio = DIVIDE([Median LTV], [Blended CAC], 0)", "Key SaaS unit economic viability ratio."],
        ["Payback Months", "Payback_Months = DIVIDE([Blended CAC], AVERAGE(cac_eda_dataset[monthly_revenue]), 0) * 12", "Expected months to recover customer acquisition cost."],
        ["Solvency Flag", "Solvency_Flag = IF([Payback_Months] <= 12, \"Solvent (<=12m)\", \"High Risk (>12m)\")", "Conditional alert flag for executive risk management."]
    ]
    t_dax = doc.add_table(rows=len(dax_data)+1, cols=3)
    style_table(t_dax, [1.8, 3.2, 1.7], ["DAX Measure Name", "DAX Code / Expression", "Role & Formatting"], dax_data)

    add_heading_2(doc, "3.3 Power BI 3-Page Canvas Layout")
    doc.add_paragraph(
        "• Page 1: Executive Overview (Cockpit)\n"
        "  - Top Header: Department title, slicers for Date Range and Channel.\n"
        "  - 4 Card Visuals: [Total Spend], [Total Conversions], [Blended CAC], [LTV_CAC_Ratio].\n"
        "  - Line & Clustered Column Chart: X-axis = Month, Column Values = [Total Spend], Line Values = [Blended CAC].\n"
        "  - Donut Chart: Conversions by Subscription Tier.\n\n"
        "• Page 2: Funnel & Channel Leakage\n"
        "  - Funnel Visual: Category = Stages (Impressions, Clicks, Leads, Conversions), Values = Stage Count.\n"
        "  - Matrix Visual: Rows = Channel, Values = Clicks, Conversions, [Blended CAC], [Payback_Months].\n"
        "  - Conditional Formatting: Background color on [Blended CAC] (Green for < ₹800, Red for > ₹1,400).\n\n"
        "• Page 3: What-If Budget Allocation Simulator\n"
        "  - What-If Parameter: 'Budget Shift %' (Slider from -50% to +50%, increment 5%).\n"
        "  - Simulated Conversions DAX Measure incorporating diminishing marginal ROAS.\n"
        "  - Tornado / Clustered Bar: Current vs Simulated Customers by Channel."
    )

    # Section 4: Viva Defense Questions
    add_heading_1(doc, "4. Viva Defense Questions: Tableau & Power BI Analytics")
    doc.add_paragraph(
        "1. Q: Why should you calculate Blended CAC in Tableau and Power BI as SUM(Spend)/SUM(Conversions) rather than AVG(CAC)?\n"
        "   A: Averaging individual row-level CAC values creates a mathematically distorted 'average of averages', "
        "giving equal weight to small ₹5,000 test campaigns and large ₹2,50,000 core campaigns. Using SUM(Spend)/SUM(Conversions) "
        "computes true volume-weighted unit economics.\n\n"
        "2. Q: How does DIVIDE() in DAX prevent application crashes compared to the standard slash (/) operator?\n"
        "   A: Standard slash (/) throws an unhandled divide-by-zero error or returns infinity when conversions are zero. "
        "DIVIDE(Numerator, Denominator, 0) includes an alternate result parameter, safely returning 0 when the denominator is zero.\n\n"
        "3. Q: How do Tableau and Power BI replicate the EDA findings discovered in Python?\n"
        "   A: Both tools confirm that Organic Search and Referral deliver the lowest CAC (< ₹450) and highest LTV:CAC (> 4.0x), "
        "while Paid Social during Q4 incurs a severe cost surge exceeding the 12-month payback benchmark."
    )

    output_path = r"c:\Users\asus\Downloads\MERN-Expense-Tracker-main\docs\Tableau_and_PowerBI_Dashboard_Guide.docx"
    doc.save(output_path)
    print(f"BI Guide saved successfully to {output_path}")

if __name__ == "__main__":
    generate_bi_docx()

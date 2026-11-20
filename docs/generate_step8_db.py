import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def generate_db_design_and_update_rtm():
    # =========================================================================
    # 1. CREATE docs/08_Database_Design.xlsx
    # =========================================================================
    wb = openpyxl.Workbook()
    
    # Common Styling Elements
    title_font = Font(name="Calibri", size=15, bold=True, color="1E3A8A")
    subtitle_font = Font(name="Calibri", size=10, italic=True, color="475569")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    section_font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
    data_font = Font(name="Calibri", size=9.5, color="1E293B")
    data_bold = Font(name="Calibri", size=9.5, bold=True, color="1E293B")
    code_font = Font(name="Consolas", size=9, color="0F172A")
    
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    sub_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    pk_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid") # Amber
    fk_fill = PatternFill(start_color="E0E7FF", end_color="E0E7FF", fill_type="solid") # Indigo
    
    thin_border_side = Side(border_style="thin", color="CBD5E1")
    border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    # -------------------------------------------------------------
    # Sheet 1: Overview
    # -------------------------------------------------------------
    ws_ov = wb.active
    ws_ov.title = "Overview"
    ws_ov.views.sheetView[0].showGridLines = True
    
    ws_ov["A1"] = "CUSTOMER ACQUISITION COST (CAC) ANALYSIS — DATABASE DESIGN SPECIFICATION"
    ws_ov["A1"].font = title_font
    ws_ov["A2"] = "Relational 3NF Schema & Analytical Star-Snowflake Model | Baseline Step 8"
    ws_ov["A2"].font = subtitle_font
    
    meta_info = [
        ("Project Name", "Customer Acquisition Cost (CAC) Analysis"),
        ("Student Author", "Gourav Kumar (Roll: 2500320140055)"),
        ("Program / College", "MCA 2nd Year (Data Analytics), ABES Engineering College"),
        ("Faculty Mentor", "Prof. Alpna Lodhi"),
        ("RDBMS Platform", "SQLite 3.x (Embedded in data/app.db)"),
        ("Schema Architecture", "Normalized 3NF Relational Foundation with De-normalized Analytical View (34 Columns)"),
        ("Total Relational Tables", "4 Entities (2 Master / Dimension Tables + 2 Transactional Fact Tables)"),
        ("Total Analytical Views", "1 Comprehensive Joined Analysis View (view_cac_full_analysis)")
    ]
    
    ws_ov.row_dimensions[4].height = 20
    for idx, (label, val) in enumerate(meta_info, 5):
        ws_ov.row_dimensions[idx].height = 19
        c1 = ws_ov.cell(row=idx, column=1, value=label)
        c2 = ws_ov.cell(row=idx, column=2, value=val)
        c1.font = data_bold
        c1.fill = zebra_fill
        c1.border = border
        c2.font = data_font
        c2.border = border
    
    ws_ov.column_dimensions["A"].width = 26
    ws_ov.column_dimensions["B"].width = 75

    # -------------------------------------------------------------
    # Sheet 2: Tables_Summary
    # -------------------------------------------------------------
    ws_sum = wb.create_sheet(title="Tables_Summary")
    ws_sum.views.sheetView[0].showGridLines = True
    
    ws_sum["A1"] = "TABLES INVENTORY & ARCHITECTURAL SUMMARY"
    ws_sum["A1"].font = title_font
    ws_sum["A2"] = "Entity Categorization, Table Grain, Key Constraints, and Planned Row Volumes"
    ws_sum["A2"].font = subtitle_font
    
    sum_headers = ["Table Name", "Entity Type", "Business Grain / Unit of Analysis", "Primary Key (PK)", "Foreign Keys (FKs)", "Columns", "Target Row Count"]
    ws_sum.row_dimensions[4].height = 24
    for j, h in enumerate(sum_headers, 1):
        c = ws_sum.cell(row=4, column=j, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
        
    sum_rows = [
        ("dim_channels", "Dimension / Master", "One row per unique acquisition channel", "channel_id", "None", 6, "5 rows"),
        ("dim_subscription_tiers", "Dimension / Master", "One row per customer plan category", "tier_id", "None", 6, "4 rows"),
        ("fact_marketing_campaigns", "Fact Table 1", "One row per campaign execution per channel per month", "campaign_id", "channel_id -> dim_channels.channel_id", 14, "240 rows (10 camp x 24 mo)"),
        ("fact_customer_cohorts", "Fact Table 2", "One row per customer acquired with cohort attributes", "customer_id", "channel_id, tier_id", 14, "300 rows (SaaS client sample)"),
        ("view_cac_full_analysis", "Joined Analysis View", "Denormalized analytics record combining spend, funnel, and LTV", "N/A (Derived View)", "Joins across all 4 tables", 34, "300 combined analytical rows")
    ]
    
    for r_idx, r_data in enumerate(sum_rows, 5):
        ws_sum.row_dimensions[r_idx].height = 20
        fill = zebra_fill if r_idx % 2 == 1 else white_fill
        for c_idx, val in enumerate(r_data, 1):
            c = ws_sum.cell(row=r_idx, column=c_idx, value=val)
            c.border = border
            c.fill = fill
            c.font = data_font
            if c_idx in (1, 4):
                c.font = data_bold
            if c_idx in (1, 2, 4, 6, 7):
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")

    for col in ["A", "B", "C", "D", "E", "F", "G"]:
        ws_sum.column_dimensions[col].width = 24
    ws_sum.column_dimensions["A"].width = 26
    ws_sum.column_dimensions["C"].width = 38
    ws_sum.column_dimensions["E"].width = 32

    # -------------------------------------------------------------
    # Sheet 3A: Dict_dim_channels
    # -------------------------------------------------------------
    def style_dictionary_sheet(ws, title, columns_data):
        ws.views.sheetView[0].showGridLines = True
        ws["A1"] = title
        ws["A1"].font = title_font
        ws["A2"] = "Data Dictionary, Field Specifications, and Business Requirement Mapping"
        ws["A2"].font = subtitle_font
        
        dict_headers = ["Column Name", "Data Type", "Key", "Nullable", "Default / Allowed Values", "Description & Business Logic", "Supported OQ / BR"]
        ws.row_dimensions[4].height = 24
        for j, h in enumerate(dict_headers, 1):
            c = ws.cell(row=4, column=j, value=h)
            c.fill = header_fill
            c.font = header_font
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border
            
        for r_idx, r_data in enumerate(columns_data, 5):
            ws.row_dimensions[r_idx].height = 20
            fill = zebra_fill if r_idx % 2 == 1 else white_fill
            if r_data[2] == "PK":
                fill = pk_fill
            elif "FK" in r_data[2]:
                fill = fk_fill
                
            for c_idx, val in enumerate(r_data, 1):
                c = ws.cell(row=r_idx, column=c_idx, value=val)
                c.border = border
                c.fill = fill
                c.font = data_font
                if c_idx == 1:
                    c.font = data_bold
                if c_idx in (2, 3, 4):
                    c.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    c.alignment = Alignment(horizontal="left", vertical="center")
                    
        ws.column_dimensions["A"].width = 24
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["C"].width = 10
        ws.column_dimensions["D"].width = 12
        ws.column_dimensions["E"].width = 26
        ws.column_dimensions["F"].width = 44
        ws.column_dimensions["G"].width = 22

    ws_ch = wb.create_sheet(title="Dict_dim_channels")
    ch_cols = [
        ("channel_id", "INTEGER", "PK", "No", "1, 2, 3, 4, 5", "Unique surrogate key for acquisition channel", "OQ-01, OQ-02 / BR-01"),
        ("channel_name", "TEXT", "None", "No", "Search, Social, Referral, Email, Organic", "Standard marketing acquisition channel name", "OQ-02 / BR-01, BR-02"),
        ("channel_type", "TEXT", "None", "No", "Paid, Organic, Partner", "Strategic macro classification of acquisition media", "OQ-01, OQ-02 / BR-02"),
        ("attribution_model", "TEXT", "None", "No", "Last-Touch, First-Touch", "Default attribution model applied to this channel", "OQ-07 / BRUL-01"),
        ("target_cpa", "REAL", "None", "No", "> 0.0 ($15.00 to $120.00)", "Target maximum Cost Per Acquisition benchmark", "OQ-02 / BR-02"),
        ("is_active", "INTEGER", "None", "No", "1 (Active), 0 (Paused)", "Flag indicating whether channel currently runs campaigns", "OQ-01 / BR-01")
    ]
    style_dictionary_sheet(ws_ch, "DATA DICTIONARY: dim_channels (Master)", ch_cols)

    # -------------------------------------------------------------
    # Sheet 3B: Dict_dim_tiers
    # -------------------------------------------------------------
    ws_tr = wb.create_sheet(title="Dict_dim_tiers")
    tier_cols = [
        ("tier_id", "INTEGER", "PK", "No", "1, 2, 3, 4", "Unique surrogate key for subscription plan tier", "OQ-03, OQ-09 / BR-05"),
        ("tier_name", "TEXT", "None", "No", "Basic, Pro, SMB, Enterprise", "Commercial customer plan category name", "OQ-03, OQ-09 / BR-05, BR-06"),
        ("monthly_price", "REAL", "None", "No", "$29, $79, $199, $599", "Standard monthly subscription fee charged per account", "OQ-04 / BR-04"),
        ("gross_margin_pct", "REAL", "None", "No", "0.65 to 0.85 (65% - 85%)", "Direct gross profit margin percentage for the tier", "OQ-04 / BR-04, BR-06"),
        ("target_ltv_cac", "REAL", "None", "No", ">= 3.0", "Internal financial viability benchmark ratio", "OQ-03 / BR-05, BRUL-04"),
        ("annual_churn_rate", "REAL", "None", "No", "0.08 to 0.28 (8% - 28%)", "Expected baseline annual logo churn rate", "OQ-03 / BR-05, BR-12")
    ]
    style_dictionary_sheet(ws_tr, "DATA DICTIONARY: dim_subscription_tiers (Master)", tier_cols)

    # -------------------------------------------------------------
    # Sheet 3C: Dict_fact_campaigns
    # -------------------------------------------------------------
    ws_fc = wb.create_sheet(title="Dict_fact_campaigns")
    camp_cols = [
        ("campaign_id", "INTEGER", "PK", "No", "Auto-increment (1..240)", "Unique surrogate key for campaign spend entry", "OQ-01, OQ-05 / BR-01"),
        ("channel_id", "INTEGER", "FK", "No", "References dim_channels(channel_id)", "Foreign key link to acquisition channel entity", "OQ-01, OQ-02 / BR-01"),
        ("campaign_name", "TEXT", "None", "No", "Text (e.g. Q1_Search_Generic)", "Marketing campaign tracking label", "OQ-05 / BR-01"),
        ("campaign_date", "TEXT", "None", "No", "YYYY-MM-DD (2024-01-01 to 2025-12-31)", "Calendar date of spend / monthly cohort start", "OQ-06 / BR-09"),
        ("year_month", "TEXT", "None", "No", "YYYY-MM (e.g. 2024-03)", "Calendar month bucket for rolling time-series aggregation", "OQ-01, OQ-06 / BR-02, BR-09"),
        ("seasonal_period", "TEXT", "None", "No", "Regular, Q4_Holiday, Spring_Promo", "Seasonal indicator flag for price surge analysis", "OQ-06 / BR-09"),
        ("spend_amount", "REAL", "None", "No", ">= 0.0 ($500.00 to $18,000.00)", "Total media advertising spend incurred in period", "OQ-01, OQ-08 / BR-01, BR-08"),
        ("impressions", "INTEGER", "None", "No", ">= 0 (5,000 to 250,000)", "Total ad impressions served at top of funnel", "OQ-07 / BR-03"),
        ("clicks", "INTEGER", "None", "No", ">= 0 (150 to 12,000)", "Total click-throughs driven from impressions", "OQ-07 / BR-03"),
        ("leads", "INTEGER", "None", "No", ">= 0 (20 to 1,500)", "Total qualified sales leads captured", "OQ-07 / BR-03, BR-10"),
        ("conversions_count", "INTEGER", "None", "No", ">= 0 (5 to 280)", "Total newly acquired paying customer conversions", "OQ-01, OQ-02 / BR-02, BR-10"),
        ("cac_calculated", "REAL", "None", "No", "spend_amount / conversions_count", "Computed channel Customer Acquisition Cost in dollars", "OQ-01, OQ-05 / BR-02, BR-07"),
        ("cpc_calculated", "REAL", "None", "No", "spend_amount / clicks", "Cost Per Click metric for funnel leakage analysis", "OQ-07 / BR-03"),
        ("cpl_calculated", "REAL", "None", "No", "spend_amount / leads", "Cost Per Lead metric for mid-funnel audit", "OQ-07 / BR-03")
    ]
    style_dictionary_sheet(ws_fc, "DATA DICTIONARY: fact_marketing_campaigns (Fact)", camp_cols)

    # -------------------------------------------------------------
    # Sheet 3D: Dict_fact_customers
    # -------------------------------------------------------------
    ws_cu = wb.create_sheet(title="Dict_fact_customers")
    cust_cols = [
        ("customer_id", "INTEGER", "PK", "No", "Auto-increment (1..300)", "Unique surrogate key for acquired customer entity", "OQ-03, OQ-09 / BR-05"),
        ("customer_code", "TEXT", "None", "No", "CUST-0001 to CUST-0300", "Unique business alphanumeric identifier", "OQ-09 / BR-05"),
        ("channel_id", "INTEGER", "FK", "No", "References dim_channels(channel_id)", "Attributed acquisition channel key", "OQ-02 / BR-01, BR-02"),
        ("tier_id", "INTEGER", "FK", "No", "References dim_subscription_tiers(tier_id)", "Customer subscribed plan tier key", "OQ-03, OQ-09 / BR-05, BR-06"),
        ("acquisition_date", "TEXT", "None", "No", "YYYY-MM-DD", "Date customer converted into paying account", "OQ-01, OQ-04 / BR-02, BR-04"),
        ("monthly_revenue", "REAL", "None", "No", "$29.00 to $599.00", "Monthly Recurring Revenue (MRR) generated", "OQ-04 / BR-04, BR-06"),
        ("servicing_cost", "REAL", "None", "No", "$5.00 to $95.00", "Monthly direct customer support & hosting cost", "OQ-09 / BR-06"),
        ("cac_attributed", "REAL", "None", "No", "$35.00 to $450.00", "Individual allocated acquisition cost for this client", "OQ-03, OQ-04 / BR-04, BR-05"),
        ("tenure_months", "INTEGER", "None", "No", "1 to 24 months", "Number of consecutive months customer remained active", "OQ-03 / BR-05"),
        ("lifetime_value", "REAL", "None", "No", "Formula: (MRR * tenure) * margin", "Total realized gross profit across customer tenure", "OQ-03 / BR-05, BR-06"),
        ("ltv_cac_ratio", "REAL", "None", "No", "lifetime_value / cac_attributed", "Ratio determining customer unit economic return", "OQ-03 / BR-05"),
        ("payback_months", "REAL", "None", "No", "cac_attributed / (MRR * margin)", "Calculated months required to recoup acquisition spend", "OQ-04 / BR-04"),
        ("churn_probability", "REAL", "None", "No", "0.01 to 0.45", "Modeled probability of customer churning in next 90 days", "OQ-03 / BR-12"),
        ("is_churned", "INTEGER", "None", "No", "0 (Active), 1 (Churned)", "Current lifecycle status of customer relationship", "OQ-03, OQ-09 / BR-05")
    ]
    style_dictionary_sheet(ws_cu, "DATA DICTIONARY: fact_customer_cohorts (Fact)", cust_cols)

    # -------------------------------------------------------------
    # Sheet 3E: Dict_view_analysis (34 Columns)
    # -------------------------------------------------------------
    ws_va = wb.create_sheet(title="Dict_view_analysis")
    view_cols = [
        ("customer_id", "INTEGER", "Key", "No", "1..300", "Customer identifier from fact table", "OQ-09 / BR-05"),
        ("customer_code", "TEXT", "None", "No", "CUST-XXXX", "Business customer alphanumeric code", "OQ-09 / BR-05"),
        ("acquisition_date", "TEXT", "None", "No", "YYYY-MM-DD", "Conversion date timestamp", "OQ-01 / BR-02"),
        ("acquisition_year_month", "TEXT", "None", "No", "YYYY-MM", "Monthly cohort period bucket", "OQ-01, OQ-06 / BR-02, BR-09"),
        ("channel_id", "INTEGER", "FK", "No", "1..5", "Channel dimension key", "OQ-01 / BR-01"),
        ("channel_name", "TEXT", "None", "No", "Search, Social, etc.", "Acquisition channel name", "OQ-02 / BR-01"),
        ("channel_type", "TEXT", "None", "No", "Paid vs Organic", "Macro media category", "OQ-01 / BR-02"),
        ("tier_id", "INTEGER", "FK", "No", "1..4", "Subscription plan key", "OQ-03 / BR-05"),
        ("tier_name", "TEXT", "None", "No", "Basic, Pro, SMB, Enterprise", "Plan tier name", "OQ-03, OQ-09 / BR-05, BR-06"),
        ("monthly_revenue", "REAL", "None", "No", "$29 to $599", "Customer MRR", "OQ-04 / BR-04"),
        ("gross_margin_pct", "REAL", "None", "No", "0.65 to 0.85", "Plan gross margin percentage", "OQ-04 / BR-04"),
        ("monthly_gross_profit", "REAL", "None", "No", "MRR * margin", "Monthly gross profit contribution", "OQ-04, OQ-09 / BR-04"),
        ("servicing_cost", "REAL", "None", "No", "$5 to $95", "Direct account maintenance cost", "OQ-09 / BR-06"),
        ("net_contribution_margin", "REAL", "None", "No", "Gross profit - servicing cost", "Net cash margin contribution", "OQ-09 / BR-06"),
        ("cac_attributed", "REAL", "None", "No", "$35 to $450", "Acquisition cost allocated to customer", "OQ-01, OQ-02 / BR-02"),
        ("tenure_months", "INTEGER", "None", "No", "1..24", "Total active lifespan months", "OQ-03 / BR-05"),
        ("is_churned", "INTEGER", "None", "No", "0 or 1", "Churned indicator flag", "OQ-03 / BR-05"),
        ("churn_probability", "REAL", "None", "No", "0.01 to 0.45", "Predictive churn likelihood", "OQ-03 / BR-12"),
        ("realized_gross_revenue", "REAL", "None", "No", "MRR * tenure", "Cumulative topline billings", "OQ-03 / BR-05"),
        ("lifetime_value", "REAL", "None", "No", "Cumulative profit", "Realized customer lifetime value", "OQ-03 / BR-05"),
        ("ltv_cac_ratio", "REAL", "None", "No", "LTV / CAC", "Customer LTV:CAC health ratio", "OQ-03 / BR-05"),
        ("ltv_health_category", "TEXT", "None", "No", "Healthy, Sub-Optimal, Destructive", "Categorical economic classification", "OQ-03 / BR-05, BRUL-04"),
        ("payback_months", "REAL", "None", "No", "CAC / Gross profit", "Months to recover acquisition outlay", "OQ-04 / BR-04"),
        ("payback_status", "TEXT", "None", "No", "Safe (<12m), Critical (>12m)", "Solvency threshold warning indicator", "OQ-04 / BR-04, BRUL-03"),
        ("channel_monthly_spend", "REAL", "None", "No", "Campaign spend", "Aggregate monthly spend of channel", "OQ-01, OQ-08 / BR-01, BR-08"),
        ("channel_impressions", "INTEGER", "None", "No", "Impressions count", "Total impressions in cohort", "OQ-07 / BR-03"),
        ("channel_clicks", "INTEGER", "None", "No", "Clicks count", "Total clicks in cohort", "OQ-07 / BR-03"),
        ("channel_leads", "INTEGER", "None", "No", "Leads count", "Total leads in cohort", "OQ-07 / BR-03"),
        ("channel_conversions", "INTEGER", "None", "No", "Conversions count", "Total conversions in cohort", "OQ-02, OQ-07 / BR-02, BR-10"),
        ("channel_cac_monthly", "REAL", "None", "No", "Spend / conversions", "Channel monthly blended CAC", "OQ-01, OQ-02 / BR-02"),
        ("click_through_rate", "REAL", "None", "No", "Clicks / impressions", "Top-of-funnel engagement %", "OQ-07 / BR-03"),
        ("lead_conversion_rate", "REAL", "None", "No", "Conversions / leads", "Bottom-of-funnel velocity %", "OQ-02, OQ-07 / BR-03, BR-10"),
        ("seasonal_period", "TEXT", "None", "No", "Regular, Q4_Holiday", "Calendar seasonality flag", "OQ-06 / BR-09"),
        ("quarter_index", "TEXT", "None", "No", "Q1, Q2, Q3, Q4", "Fiscal quarter identifier", "OQ-06 / BR-09")
    ]
    style_dictionary_sheet(ws_va, "DATA DICTIONARY: view_cac_full_analysis (34 Columns)", view_cols)

    # -------------------------------------------------------------
    # Sheet 4: Relationships
    # -------------------------------------------------------------
    ws_rel = wb.create_sheet(title="Relationships")
    ws_rel.views.sheetView[0].showGridLines = True
    ws_rel["A1"] = "RELATIONAL SCHEMA INTEGRITY & CARDINALITY RULES"
    ws_rel["A1"].font = title_font
    ws_rel["A2"] = "Foreign Key Constraints, Referential Integrity, and Entity Cardinalities"
    ws_rel["A2"].font = subtitle_font
    
    rel_headers = ["Relationship ID", "Parent Table (Primary Key)", "Child Table (Foreign Key)", "Cardinality", "Referential Integrity Rule", "Business Justification"]
    ws_rel.row_dimensions[4].height = 24
    for j, h in enumerate(rel_headers, 1):
        c = ws_rel.cell(row=4, column=j, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
        
    rel_data = [
        ("REL-01", "dim_channels (channel_id)", "fact_marketing_campaigns (channel_id)", "1 : Many (1:N)", "ON DELETE RESTRICT\nON UPDATE CASCADE", "Every ad spend record must originate from a verified acquisition channel."),
        ("REL-02", "dim_channels (channel_id)", "fact_customer_cohorts (channel_id)", "1 : Many (1:N)", "ON DELETE RESTRICT\nON UPDATE CASCADE", "Every converted customer must be attributed to a valid acquisition channel."),
        ("REL-03", "dim_subscription_tiers (tier_id)", "fact_customer_cohorts (tier_id)", "1 : Many (1:N)", "ON DELETE RESTRICT\nON UPDATE CASCADE", "Every customer must subscribe to an authorized commercial plan tier.")
    ]
    for r_idx, r_item in enumerate(rel_data, 5):
        ws_rel.row_dimensions[r_idx].height = 28
        fill = zebra_fill if r_idx % 2 == 1 else white_fill
        for c_idx, val in enumerate(r_item, 1):
            c = ws_rel.cell(row=r_idx, column=c_idx, value=val)
            c.border = border
            c.fill = fill
            c.font = data_font
            if c_idx == 1:
                c.font = data_bold
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx == 4:
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
                
    ws_rel.column_dimensions["A"].width = 18
    ws_rel.column_dimensions["B"].width = 30
    ws_rel.column_dimensions["C"].width = 32
    ws_rel.column_dimensions["D"].width = 18
    ws_rel.column_dimensions["E"].width = 24
    ws_rel.column_dimensions["F"].width = 46

    # -------------------------------------------------------------
    # Sheet 5: KPIs_Measures
    # -------------------------------------------------------------
    ws_kpi = wb.create_sheet(title="KPIs_Measures")
    ws_kpi.views.sheetView[0].showGridLines = True
    ws_kpi["A1"] = "ANALYTICAL KPIS, METRICS & FORMULA DEFINITIONS"
    ws_kpi["A1"].font = title_font
    ws_kpi["A2"] = "Operational Question Alignment, Mathematical Logic, and Benchmark Thresholds"
    ws_kpi["A2"].font = subtitle_font
    
    kpi_headers = ["Metric ID", "KPI / Metric Name", "Exact Mathematical Formula", "Target Benchmark / Rule", "Covered OQ(s)"]
    ws_kpi.row_dimensions[4].height = 24
    for j, h in enumerate(kpi_headers, 1):
        c = ws_kpi.cell(row=4, column=j, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
        
    kpi_data = [
        ("MTR-01", "Blended Monthly CAC", "SUM(spend_amount) / SUM(conversions_count)", "<= $85.00 overall", "OQ-01"),
        ("MTR-02", "Channel-Specific CAC", "SUM(spend_amount_c) / SUM(conversions_count_c)", "Channel target CPA", "OQ-02"),
        ("MTR-03", "Lead Conversion Rate (LCR)", "SUM(conversions_count) / SUM(leads) * 100", ">= 12.0%", "OQ-02, OQ-07"),
        ("MTR-04", "Customer Lifetime Value (LTV)", "(monthly_revenue * gross_margin_pct) / annual_churn_rate", ">= $1,200 avg", "OQ-03"),
        ("MTR-05", "LTV:CAC Health Ratio", "LTV / CAC_attributed", ">= 3.0x (Healthy)", "OQ-03, OQ-09"),
        ("MTR-06", "Payback Period (Months)", "cac_attributed / (monthly_revenue * gross_margin_pct)", "<= 12.0 Months (Solvency)", "OQ-04"),
        ("MTR-07", "Channel ANOVA F-Statistic", "Between_Group_Variance / Within_Group_Variance", "p < 0.05 (Sig. Diff)", "OQ-05"),
        ("MTR-08", "Seasonality Price Surge Index", "monthly_cac / rolling_12m_avg_cac", "Index > 1.25 indicates surge", "OQ-06"),
        ("MTR-09", "Cost Per Click (CPC)", "SUM(spend_amount) / SUM(clicks)", "<= $2.50", "OQ-07"),
        ("MTR-10", "Cost Per Lead (CPL)", "SUM(spend_amount) / SUM(leads)", "<= $28.00", "OQ-07"),
        ("MTR-11", "Return on Ad Spend (ROAS)", "SUM(attributed_revenue) / SUM(spend_amount)", ">= 3.5x", "OQ-08"),
        ("MTR-12", "Segment Net Margin Contribution", "lifetime_value - (cac_attributed + servicing_cost_total)", "> 50% margin", "OQ-09")
    ]
    for r_idx, r_item in enumerate(kpi_data, 5):
        ws_kpi.row_dimensions[r_idx].height = 20
        fill = zebra_fill if r_idx % 2 == 1 else white_fill
        for c_idx, val in enumerate(r_item, 1):
            c = ws_kpi.cell(row=r_idx, column=c_idx, value=val)
            c.border = border
            c.fill = fill
            c.font = data_font
            if c_idx == 1:
                c.font = data_bold
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx in (4, 5):
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
                
    ws_kpi.column_dimensions["A"].width = 14
    ws_kpi.column_dimensions["B"].width = 30
    ws_kpi.column_dimensions["C"].width = 44
    ws_kpi.column_dimensions["D"].width = 28
    ws_kpi.column_dimensions["E"].width = 16

    # -------------------------------------------------------------
    # Sheet 6: SQLite_DDL
    # -------------------------------------------------------------
    ws_ddl = wb.create_sheet(title="SQLite_DDL")
    ws_ddl.views.sheetView[0].showGridLines = True
    ws_ddl["A1"] = "PRODUCTION SQLITE DDL SCHEMA SCRIPTS"
    ws_ddl["A1"].font = title_font
    ws_ddl["A2"] = "Complete Table DDL with PK, FK, NOT NULL, CHECK, UNIQUE, Indices & View Definition"
    ws_ddl["A2"].font = subtitle_font
    
    ddl_script = """-- ============================================================================
-- SQLITE RELATIONAL SCHEMA SPECIFICATION: CUSTOMER ACQUISITION COST ANALYSIS
-- Step 8 Baselined Architecture | Target: data/app.db
-- ============================================================================

PRAGMA foreign_keys = ON;

-- 1. DIMENSION TABLE: dim_channels
CREATE TABLE IF NOT EXISTS dim_channels (
    channel_id INTEGER PRIMARY KEY,
    channel_name TEXT NOT NULL UNIQUE CHECK (channel_name IN ('Search', 'Social', 'Referral', 'Email', 'Organic')),
    channel_type TEXT NOT NULL CHECK (channel_type IN ('Paid', 'Organic', 'Partner')),
    attribution_model TEXT NOT NULL DEFAULT 'Last-Touch',
    target_cpa REAL NOT NULL CHECK (target_cpa > 0.0),
    is_active INTEGER NOT NULL DEFAULT 1 CHECK (is_active IN (0, 1))
);

-- 2. DIMENSION TABLE: dim_subscription_tiers
CREATE TABLE IF NOT EXISTS dim_subscription_tiers (
    tier_id INTEGER PRIMARY KEY,
    tier_name TEXT NOT NULL UNIQUE CHECK (tier_name IN ('Basic', 'Pro', 'SMB', 'Enterprise')),
    monthly_price REAL NOT NULL CHECK (monthly_price > 0.0),
    gross_margin_pct REAL NOT NULL CHECK (gross_margin_pct BETWEEN 0.10 AND 0.99),
    target_ltv_cac REAL NOT NULL DEFAULT 3.0 CHECK (target_ltv_cac >= 1.0),
    annual_churn_rate REAL NOT NULL CHECK (annual_churn_rate BETWEEN 0.01 AND 0.90)
);

-- 3. FACT TABLE 1: fact_marketing_campaigns
CREATE TABLE IF NOT EXISTS fact_marketing_campaigns (
    campaign_id INTEGER PRIMARY KEY AUTOINCREMENT,
    channel_id INTEGER NOT NULL,
    campaign_name TEXT NOT NULL,
    campaign_date TEXT NOT NULL,
    year_month TEXT NOT NULL,
    seasonal_period TEXT NOT NULL DEFAULT 'Regular' CHECK (seasonal_period IN ('Regular', 'Q4_Holiday', 'Spring_Promo')),
    spend_amount REAL NOT NULL CHECK (spend_amount >= 0.0),
    impressions INTEGER NOT NULL CHECK (impressions >= 0),
    clicks INTEGER NOT NULL CHECK (clicks >= 0),
    leads INTEGER NOT NULL CHECK (leads >= 0),
    conversions_count INTEGER NOT NULL CHECK (conversions_count >= 0),
    cac_calculated REAL NOT NULL DEFAULT 0.0,
    cpc_calculated REAL NOT NULL DEFAULT 0.0,
    cpl_calculated REAL NOT NULL DEFAULT 0.0,
    FOREIGN KEY (channel_id) REFERENCES dim_channels (channel_id) ON DELETE RESTRICT ON UPDATE CASCADE
);

-- Indices for high-frequency queries
CREATE INDEX IF NOT EXISTS idx_camp_channel_date ON fact_marketing_campaigns (channel_id, campaign_date);
CREATE INDEX IF NOT EXISTS idx_camp_ym ON fact_marketing_campaigns (year_month);

-- 4. FACT TABLE 2: fact_customer_cohorts
CREATE TABLE IF NOT EXISTS fact_customer_cohorts (
    customer_id INTEGER PRIMARY KEY AUTOINCREMENT,
    customer_code TEXT NOT NULL UNIQUE,
    channel_id INTEGER NOT NULL,
    tier_id INTEGER NOT NULL,
    acquisition_date TEXT NOT NULL,
    monthly_revenue REAL NOT NULL CHECK (monthly_revenue > 0.0),
    servicing_cost REAL NOT NULL CHECK (servicing_cost >= 0.0),
    cac_attributed REAL NOT NULL CHECK (cac_attributed >= 0.0),
    tenure_months INTEGER NOT NULL CHECK (tenure_months >= 1),
    lifetime_value REAL NOT NULL DEFAULT 0.0,
    ltv_cac_ratio REAL NOT NULL DEFAULT 0.0,
    payback_months REAL NOT NULL DEFAULT 0.0,
    churn_probability REAL NOT NULL CHECK (churn_probability BETWEEN 0.0 AND 1.0),
    is_churned INTEGER NOT NULL DEFAULT 0 CHECK (is_churned IN (0, 1)),
    FOREIGN KEY (channel_id) REFERENCES dim_channels (channel_id) ON DELETE RESTRICT ON UPDATE CASCADE,
    FOREIGN KEY (tier_id) REFERENCES dim_subscription_tiers (tier_id) ON DELETE RESTRICT ON UPDATE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_cust_channel_tier ON fact_customer_cohorts (channel_id, tier_id);
CREATE INDEX IF NOT EXISTS idx_cust_acq_date ON fact_customer_cohorts (acquisition_date);

-- ============================================================================
-- 5. DE-NORMALIZED ANALYTICAL VIEW (34 COLUMNS): view_cac_full_analysis
-- ============================================================================
CREATE VIEW IF NOT EXISTS view_cac_full_analysis AS
SELECT 
    -- Customer Identity & Lifecycle
    cu.customer_id,
    cu.customer_code,
    cu.acquisition_date,
    strftime('%Y-%m', cu.acquisition_date) AS acquisition_year_month,
    
    -- Channel Dimensions
    ch.channel_id,
    ch.channel_name,
    ch.channel_type,
    
    -- Tier Dimensions
    tr.tier_id,
    tr.tier_name,
    
    -- Revenue & Unit Margins
    cu.monthly_revenue,
    tr.gross_margin_pct,
    ROUND(cu.monthly_revenue * tr.gross_margin_pct, 2) AS monthly_gross_profit,
    cu.servicing_cost,
    ROUND((cu.monthly_revenue * tr.gross_margin_pct) - cu.servicing_cost, 2) AS net_contribution_margin,
    
    -- Customer Acquisition & LTV
    cu.cac_attributed,
    cu.tenure_months,
    cu.is_churned,
    cu.churn_probability,
    ROUND(cu.monthly_revenue * cu.tenure_months, 2) AS realized_gross_revenue,
    cu.lifetime_value,
    cu.ltv_cac_ratio,
    CASE 
        WHEN cu.ltv_cac_ratio >= 3.0 THEN 'Healthy'
        WHEN cu.ltv_cac_ratio >= 1.0 THEN 'Sub-Optimal'
        ELSE 'Value-Destroying'
    END AS ltv_health_category,
    
    -- Payback Dynamics
    cu.payback_months,
    CASE 
        WHEN cu.payback_months <= 12.0 THEN 'Safe (<12m)'
        ELSE 'Critical (>12m)'
    END AS payback_status,
    
    -- Aggregated Monthly Channel Context
    COALESCE(camp.spend_amount, 0.0) AS channel_monthly_spend,
    COALESCE(camp.impressions, 0) AS channel_impressions,
    COALESCE(camp.clicks, 0) AS channel_clicks,
    COALESCE(camp.leads, 0) AS channel_leads,
    COALESCE(camp.conversions_count, 0) AS channel_conversions,
    COALESCE(camp.cac_calculated, 0.0) AS channel_cac_monthly,
    CASE WHEN camp.impressions > 0 THEN ROUND((CAST(camp.clicks AS REAL) / camp.impressions) * 100.0, 2) ELSE 0.0 END AS click_through_rate,
    CASE WHEN camp.leads > 0 THEN ROUND((CAST(camp.conversions_count AS REAL) / camp.leads) * 100.0, 2) ELSE 0.0 END AS lead_conversion_rate,
    COALESCE(camp.seasonal_period, 'Regular') AS seasonal_period,
    CASE 
        WHEN CAST(strftime('%m', cu.acquisition_date) AS INTEGER) BETWEEN 1 AND 3 THEN 'Q1'
        WHEN CAST(strftime('%m', cu.acquisition_date) AS INTEGER) BETWEEN 4 AND 6 THEN 'Q2'
        WHEN CAST(strftime('%m', cu.acquisition_date) AS INTEGER) BETWEEN 7 AND 9 THEN 'Q3'
        ELSE 'Q4'
    END AS quarter_index

FROM fact_customer_cohorts cu
JOIN dim_channels ch ON cu.channel_id = ch.channel_id
JOIN dim_subscription_tiers tr ON cu.tier_id = tr.tier_id
LEFT JOIN fact_marketing_campaigns camp 
    ON cu.channel_id = camp.channel_id 
    AND strftime('%Y-%m', cu.acquisition_date) = camp.year_month;
"""
    
    ws_ddl["A4"] = ddl_script
    ws_ddl["A4"].font = code_font
    ws_ddl["A4"].alignment = Alignment(vertical="top", wrap_text=True)
    ws_ddl.row_dimensions[4].height = 420
    ws_ddl.column_dimensions["A"].width = 110

    # Save Database Design Excel
    db_target = os.path.abspath("docs/08_Database_Design.xlsx")
    wb.save(db_target)
    print(f"Generated: {db_target}")

    # =========================================================================
    # 2. UPDATE docs/04_Traceability_Matrix.xlsx (FILL DB table.column)
    # =========================================================================
    rtm_path = os.path.abspath("docs/04_Traceability_Matrix.xlsx")
    rtm_wb = openpyxl.load_workbook(rtm_path)
    ws_rtm = rtm_wb["RTM_Traceability_Matrix"]
    
    db_col_mappings = {
        "BR-01": "fact_marketing_campaigns.channel_id, spend_amount",
        "BR-02": "fact_marketing_campaigns.spend_amount, conversions_count",
        "BR-03": "fact_marketing_campaigns.impressions, clicks, leads, conversions_count",
        "BR-04": "fact_customer_cohorts.cac_attributed, monthly_revenue; dim_subscription_tiers.gross_margin_pct",
        "BR-05": "fact_customer_cohorts.lifetime_value, cac_attributed; dim_subscription_tiers.tier_name",
        "BR-06": "dim_subscription_tiers.tier_name; fact_customer_cohorts.monthly_revenue, servicing_cost",
        "BR-07": "fact_marketing_campaigns.channel_id, cac_calculated",
        "BR-08": "fact_marketing_campaigns.spend_amount; fact_customer_cohorts.lifetime_value",
        "BR-09": "fact_marketing_campaigns.campaign_date, spend_amount, seasonal_period",
        "BR-10": "fact_marketing_campaigns.conversions_count, cac_calculated",
        "BR-11": "view_cac_full_analysis (34 computed metric columns)",
        "BR-12": "fact_customer_cohorts.churn_probability, lifetime_value"
    }

    designed_font = Font(name="Calibri", size=10, bold=True, color="1E40AF") # Blue
    designed_fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")

    for row in range(5, 21):
        br_id = ws_rtm.cell(row=row, column=3).value
        if br_id in db_col_mappings:
            # Update Column 7 (Status -> Designed)
            c_status = ws_rtm.cell(row=row, column=7, value="Designed")
            c_status.font = designed_font
            c_status.fill = designed_fill
            
            # Update Column 9 (DB Table.Column)
            c_db = ws_rtm.cell(row=row, column=9, value=db_col_mappings[br_id])
            c_db.font = Font(name="Consolas", size=8.5, color="1E293B")
            c_db.alignment = Alignment(horizontal="left", vertical="center")

    rtm_wb.save(rtm_path)
    print(f"Updated RTM: {rtm_path}")

if __name__ == "__main__":
    generate_db_design_and_update_rtm()
